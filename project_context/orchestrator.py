import os
import glob
import asyncio
from typing import List
from dotenv import load_dotenv

from project_context import ProjectAnalysisContext
from agents.scoring_agent import ScoringAgent, CombinedAgent
from agents.feedback_agent import FeedbackAgent
from agents.image_eval import WorkflowAnalysisAgent
from utils import load_document_content, display_consolidated_report, display_leaderboard, ALLOWED_EXTS, save_consolidated_reports_to_excel, save_leaderboard_to_excel

async def aload_document_content(file_path: str):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, load_document_content, file_path)

def _expand_team_glob(pattern: str) -> List[str]:
    if not pattern:
        return []
    raw = pattern.strip()
    parts = [p.strip().strip('"').strip("'") for p in raw.split(",") if p.strip()]
    found: List[str] = []
    for p in parts:
        p_norm = os.path.abspath(p.replace("\\", "/"))
        found.extend(glob.glob(p_norm, recursive=True))
    found = [f for f in found if os.path.isfile(f) and os.path.splitext(f)[1].lower() in ALLOWED_EXTS]
    return sorted(set(found))

async def process_file(file_path: str, agent_mode: str, semaphore: asyncio.Semaphore):
    async with semaphore:
        print("\n" + "*" * 70)
        print(f"Processing: {file_path}")
        print("*" * 70)

        ctx = ProjectAnalysisContext(file_path)

        if not os.path.exists(file_path):
            ctx.set_error("File not found.")
            display_consolidated_report(ctx)
            return ctx

        try:
            # 1) Load text + quick images for evidence count
            ctx.raw_text, ctx.images_base64 = await aload_document_content(file_path)

            # 2) Diagram summary from images (robust agent)
            try:
                img_agent = WorkflowAnalysisAgent()
                report = img_agent.analyze_workflows(file_path)
                if report:
                    ctx.update_workflow_report(report.dict())
            except Exception as e:
                print(f"  -> Diagram summary skipped: {e}")

            # 3) Evaluate
            if agent_mode == "combined":
                agent = CombinedAgent()
                await agent.run(ctx)
            else:
                scoring = ScoringAgent()
                feedback = FeedbackAgent()
                await scoring.run(ctx)
                await feedback.run(ctx)

        except Exception as e:
            ctx.set_error(f"Unhandled error: {type(e).__name__}: {e}")

        display_consolidated_report(ctx)
        return ctx

async def main():
    load_dotenv()
    pattern = os.getenv("TEAM_GLOB", "").strip()
    TEAM_FILES = _expand_team_glob(pattern) if pattern else []
    if not TEAM_FILES:
        print("No input files found. Set TEAM_GLOB.")
        raise SystemExit(1)

    max_concurrency = int(os.getenv("MAX_CONCURRENCY", "2"))
    semaphore = asyncio.Semaphore(max_concurrency)
    agent_mode = "combined" if os.getenv("USE_COMBINED", "0").lower() in ("1", "true", "yes") else "separate"
    print(f"[info] Mode: {agent_mode} | Files: {len(TEAM_FILES)}")

    tasks = [asyncio.create_task(process_file(fp, agent_mode, semaphore)) for fp in TEAM_FILES]
    results = await asyncio.gather(*tasks, return_exceptions=False)
    contexts = [r for r in results if r is not None]
    if contexts:
        display_leaderboard(contexts)
         # Save all consolidated reports to Excel
        save_consolidated_reports_to_excel(contexts, "consolidated_reports.xlsx")
        # Save leaderboard to Excel
        save_leaderboard_to_excel(contexts, "leaderboard.xlsx")


if __name__ == "__main__":
    asyncio.run(main())
