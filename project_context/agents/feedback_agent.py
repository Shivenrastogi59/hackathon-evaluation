import os
import asyncio
from pydantic.v1 import BaseModel, Field
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.messages import HumanMessage
from langchain_core.output_parsers import JsonOutputParser
from langchain_openai import ChatOpenAI

from utils import get_text_limiter, extract_first_json_object

class FeedbackAgent:
    class FeedbackOutput(BaseModel):
        positive: str
        criticism: str
        technical: str
        suggestions: str

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
        self.parser = JsonOutputParser(pydantic_object=self.FeedbackOutput)
        self.prompt = ChatPromptTemplate.from_template(
            """
You are a hackathon mentor. Use BOTH the scoring summary and the diagram summary as evidence.
Return detailed, research-oriented guidance. Use numbered lists and reference slides/diagrams when possible.

Diagram Summary (evidence):
{workflow_report_text}

Inputs:
Summary: {scoring_summary}
Scores: {scores}

Output JSON:
- positive: 4–7 strengths with slide/diagram refs and impact.
- criticism: 4–7 gaps, risks, assumption checks.
- technical: architecture, data, eval plan, scalability, privacy/security, cost/latency, trade-offs.
- suggestions: prioritized next steps with milestones, acceptance criteria, fallbacks, risk mitigations.

Format Instructions: {format_instructions}
Deck Text: {document_text}
"""
        )

    async def _ainvoke_json(self, messages):
        last_err = None
        for attempt in range(self.max_retries + 1):
            try:
                await self.limiter.acquire()
                gen_cfg = {"temperature": 0.1, "top_p": 0.0}
                if self.seed is not None:
                    gen_cfg["seed"] = int(self.seed) if str(self.seed).isdigit() else self.seed
                llm = ChatOpenAI(model=self.model, api_key=self.api_key, **gen_cfg)
                resp = await asyncio.wait_for(llm.ainvoke(messages), self.timeout_s)
                raw = extract_first_json_object(resp.content or "") or (resp.content or "")
                return self.parser.parse(raw)
            except Exception as e:
                last_err = e
                await asyncio.sleep(1.5 * (2 ** attempt))
        raise last_err

    async def run(self, context):
        if context.evaluation_error:
            return
        print(f"--- FeedbackAgent: {context.team_name} ---")
        try:
            prompt_text = self.prompt.format(
                workflow_report_text=context.workflow_report_text or "(no diagrams found)",
                scoring_summary=context.scoring_summary,
                scores=str(context.scores),
                format_instructions=self.parser.get_format_instructions(),
                document_text=context.raw_text or ""
            )
            parts = [{"type": "text", "text": prompt_text}]
            parsed = await self._ainvoke_json([HumanMessage(content=parts)])
            context.update_feedback_results(parsed)
            print("  -> Feedback complete.")
        except Exception as e:
            print(f"  -> ERROR: {e}")
            context.set_error(f"Feedback Agent failed: {e}")
