import os
import asyncio
from typing import List, Optional, Dict, Any
from pydantic.v1 import BaseModel, Field
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.messages import HumanMessage
from langchain_core.output_parsers import JsonOutputParser
from langchain_openai import ChatOpenAI

from utils import (
    get_text_limiter,
    calibrate_and_enrich_scores,
    extract_first_json_object,
)

def _strip_code_fences(text: str) -> str:
    if not text:
        return ""
    t = text.strip()
    if t.startswith("```"):
        lines = t.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        return "\n".join(lines).strip()
    return t

def _to_int_1_10(x) -> int:
    try:
        v = int(round(float(x)))
        return max(1, min(10, v))
    except Exception:
        return 1

class ImageAnalysis(BaseModel):
    description: str = Field(description="Step-by-step description of the diagram.")
    type: str = Field(description="Type of diagram (e.g., Architecture, Flowchart).")

class WorkflowOutput(BaseModel):
    overall_summary: str = Field(description="A high-level summary of the entire workflow.")
    analyses: List[ImageAnalysis] = Field(description="A list of analyses for each image.")

class ScoringOutput(BaseModel):
    team_name: str
    scores: Dict[str, Any]
    summary: str
    workflow_analysis: Optional[WorkflowOutput]

class FeedbackOnly(BaseModel):
    positive: str
    criticism: str
    technical: str
    suggestions: str

class CombinedOutput(BaseModel):
    team_name: str
    scores: Dict[str, Any]
    summary: str
    workflow_analysis: Optional[WorkflowOutput]
    feedback: FeedbackOnly

class _LLMInvoker:
    def __init__(self):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY not found")
        self.api_key = api_key
        self.model = os.getenv("OPENAI_MODEL_TEXT", os.getenv("OPENAI_MODEL", "gpt-4o-mini"))
        self.timeout_s = int(os.getenv("LLM_TIMEOUT_S", "90"))
        self.max_retries = int(os.getenv("LLM_MAX_RETRIES", "2"))
        self.seed = os.getenv("OPENAI_SEED")
        self.limiter = get_text_limiter()

    async def ainvoke_json(self, messages, parser: JsonOutputParser):
        last_err = None
        for attempt in range(self.max_retries + 1):
            try:
                await self.limiter.acquire()
                gen_cfg = {"temperature": 0.0, "top_p": 0.0}
                if self.seed is not None:
                    gen_cfg["seed"] = int(self.seed) if str(self.seed).isdigit() else self.seed
                llm = ChatOpenAI(model=self.model, api_key=self.api_key, **gen_cfg)
                resp = await asyncio.wait_for(llm.ainvoke(messages), self.timeout_s)
                raw = _strip_code_fences(getattr(resp, "content", ""))
                raw = extract_first_json_object(raw) or raw
                return parser.parse(raw)
            except Exception as e:
                last_err = e
                await asyncio.sleep(1.5 * (2 ** attempt))
        raise last_err

STRICT_RUBRIC = """
Scoring rubric. Use INTEGER 1-10. Avoid default 10s.

Anchors:
- 10: Exceptional and proven in-deck with clear metrics, full architecture or demo.
- 8: Strong with one notable gap.
- 6: Adequate with multiple gaps; little hard evidence.
- 4: Minimal coverage; mostly claims.
- 2: Not addressed.

Rules:
- Treat diagram evidence equal to text evidence. If text and diagram conflict, prefer the diagram.
- Missing or vague -> 2-4; partial 5-7; low evidence cap 8.
- At most one criterion may be 10; bias downward if uncertain.

Checklist to consider:
problem framing, assumptions, baselines, datasets, metrics/KPIs & eval plan,
architecture & scalability, latency/cost estimates, risks & mitigations,
privacy/compliance, security, deployment plan, adoption path.
""".strip()

class ScoringAgent(_LLMInvoker):
    def __init__(self):
        super().__init__()
        self.parser = JsonOutputParser(pydantic_object=ScoringOutput)
        self.prompt = ChatPromptTemplate.from_template(
            """
You are a strict hackathon judge. Use BOTH sources of evidence with equal weight:
(A) Deck text
(B) Diagram summary extracted from images (only images classified as diagrams and important)

{strict_rubric}

Diagram Summary (evidence):
{workflow_report_text}

Evaluation:

1) Scoring & Summary:
   - Score each EXACT key (INTEGER 1-10):
     Problem Understanding, Innovation & Uniqueness, Technical Feasibility,
     Implementation Approach, Team Readiness, Potential Impact.
   - Provide a concise project summary grounded in diagram + text evidence.

2) Workflow Analysis:
   - If diagrams exist, describe them step-by-step and combine into an overall workflow.
   - Else, set workflow_analysis to null.

Tie-break order if totals match:
Innovation & Uniqueness > Technical Feasibility > Potential Impact
> Problem Understanding > Implementation Approach > Team Readiness.

Output:
- Return a SINGLE JSON object.

Format Instructions:
{format_instructions}

Deck Text:
{document_text}
"""
        )

    async def run(self, context):
        print(f"--- ScoringAgent: {context.file_path} ---")
        try:
            prompt_text = self.prompt.format(
                strict_rubric=STRICT_RUBRIC,
                workflow_report_text=context.workflow_report_text or "(no diagrams found)",
                document_text=context.raw_text or "",
                format_instructions=self.parser.get_format_instructions(),
            )
            messages = [HumanMessage(content=[{"type": "text", "text": prompt_text}])]
            parsed = await self.ainvoke_json(messages, self.parser)

            # Count only important diagrams as evidence
            diag_count = 0
            if context.workflow_report and context.workflow_report.get("image_analyses"):
                for a in context.workflow_report["image_analyses"]:
                    if a.get("is_diagram") and (a.get("importance","").lower() in {"critical","supporting"}):
                        diag_count += 1

            raw_scores = {k: _to_int_1_10(v) for k, v in (parsed.get("scores") or {}).items()}
            scores = calibrate_and_enrich_scores(context.raw_text or "", diag_count, raw_scores)

            context.update_scoring_results(
                parsed.get("team_name", "Unknown"),
                scores,
                parsed.get("summary", ""),
                parsed.get("workflow_analysis"),
            )
            print("  -> Scoring complete.")
        except Exception as e:
            print(f"  -> ERROR: {e}")
            context.set_error(f"Scoring Agent failed: {e}")

class CombinedAgent(_LLMInvoker):
    def __init__(self):
        super().__init__()
        self.parser = JsonOutputParser(pydantic_object=CombinedOutput)
        self.prompt = ChatPromptTemplate.from_template(
            """
You are a strict hackathon judge and mentor. Use deck text + diagram summary with equal weight.
Consider only images that are diagrams and marked critical/supporting as core evidence.

{strict_rubric}

Diagram Summary (evidence):
{workflow_report_text}

1) Scoring & Summary:
   - Score the six criteria and provide an evidence-backed summary.

2) Workflow Analysis:
   - Describe diagrams and overall workflow if present. Else null.

3) Feedback:
   - Fields: positive, criticism, technical, suggestions.
   - Make each field detailed, numbered, and reference slides/diagrams when possible.

Tie-break order:
Innovation & Uniqueness > Technical Feasibility > Potential Impact
> Problem Understanding > Implementation Approach > Team Readiness.

Output a SINGLE JSON.

Format Instructions:
{format_instructions}

Deck Text:
{document_text}
"""
        )

    async def run(self, context):
        print(f"--- CombinedAgent: {context.file_path} ---")
        try:
            prompt_text = self.prompt.format(
                strict_rubric=STRICT_RUBRIC,
                workflow_report_text=context.workflow_report_text or "(no diagrams found)",
                document_text=context.raw_text or "",
                format_instructions=self.parser.get_format_instructions(),
            )
            messages = [HumanMessage(content=[{"type": "text", "text": prompt_text}])]
            parsed = await self.ainvoke_json(messages, self.parser)

            diag_count = 0
            if context.workflow_report and context.workflow_report.get("image_analyses"):
                for a in context.workflow_report["image_analyses"]:
                    if a.get("is_diagram") and (a.get("importance","").lower() in {"critical","supporting"}):
                        diag_count += 1

            raw_scores = {k: _to_int_1_10(v) for k, v in (parsed.get("scores") or {}).items()}
            scores = calibrate_and_enrich_scores(context.raw_text or "", diag_count, raw_scores)

            context.update_scoring_results(
                parsed.get("team_name", "Unknown"),
                scores,
                parsed.get("summary", ""),
                parsed.get("workflow_analysis"),
            )
            context.update_feedback_results(parsed.get("feedback") or {})
            print("  -> Combined scoring + feedback complete.")
        except Exception as e:
            print(f"  -> ERROR: {e}")
            context.set_error(f"Combined Agent failed: {e}")
