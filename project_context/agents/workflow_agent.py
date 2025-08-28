# Optional legacy agent. Not used by orchestrator. Kept for completeness.
import os
from typing import List
from pydantic.v1 import BaseModel, Field
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.messages import HumanMessage
from langchain_core.output_parsers import JsonOutputParser
from langchain_openai import ChatOpenAI

class WorkflowAgent:
    def __init__(self):
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise ValueError("OPENAI_API_KEY not found")
        self.llm = ChatOpenAI(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            api_key=api_key,
            temperature=0.0,
            top_p=0.0,
        )
        self.parser = self._create_parser()
        self.prompt = self._create_prompt()

    class ImageAnalysis(BaseModel):
        description: str = Field(description="Step-by-step description of the diagram.")
        type: str = Field(description="Type of diagram (e.g., Architecture, Flowchart).")

    class WorkflowOutput(BaseModel):
        overall_summary: str = Field(description="A high-level summary of the entire workflow.")
        analyses: List["WorkflowAgent.ImageAnalysis"] = Field(description="A list of analyses for each image.")

    def _create_parser(self):
        return JsonOutputParser(pydantic_object=self.WorkflowOutput)

    def _create_prompt(self):
        return ChatPromptTemplate.from_template(
            """
You are a system analyst. Analyze the provided images, which are flowcharts or architecture diagrams.
Describe each diagram and provide an overall summary of the system's workflow.
Return ONLY a JSON object.

Format Instructions: {format_instructions}
"""
        )

    def run(self, context):
        if context.evaluation_error or not context.images_base64:
            return
        try:
            prompt_text = self.prompt.format(format_instructions=self.parser.get_format_instructions())
            message_parts = [{"type": "text", "text": prompt_text}]
            for img in context.images_base64:
                message_parts.append({"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img}"}})
            response = self.llm.invoke([HumanMessage(content=message_parts)])
            parsed_response = self.parser.parse(response.content)
            context.update_workflow_results(parsed_response)
        except Exception as e:
            context.set_error(f"Workflow Agent failed: {e}")
