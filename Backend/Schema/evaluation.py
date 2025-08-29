from pydantic import BaseModel, Field
from typing import List, Optional, Dict
from datetime import datetime
from db.mongo import db  # ✅ MongoDB instance

# ------------------ ✅ Pydantic Schemas ------------------

class EvaluationCriteria(BaseModel):
    criteria_id: str
    name: str
    weight: float
    description: Optional[str]

class JudgeEvaluationScore(BaseModel):
    problem_solution_fit: float = Field(..., ge=1, le=10, description="Problem-Solution Fit (1-10)")
    functionality_features: float = Field(..., ge=1, le=10, description="Functionality & Features (1-10)")
    technical_feasibility: float = Field(..., ge=1, le=10, description="Technical Feasibility & Robustness (1-10)")
    innovation_creativity: float = Field(..., ge=1, le=10, description="Innovation & Creativity (1-10)")
    user_experience: float = Field(..., ge=1, le=10, description="User Experience (UI/UX) (1-10)")
    impact_value: float = Field(..., ge=1, le=10, description="Impact & Value Proposition (1-10)")
    presentation_demo_quality: float = Field(..., ge=1, le=10, description="Presentation & Demo Quality (1-10)")
    team_collaboration: float = Field(..., ge=1, le=10, description="Team Collaboration (1-10)")

class TeamEvaluation(BaseModel):
    evaluation_id: str = Field(..., description="Unique evaluation identifier")
    judge_id: str = Field(..., description="ID of the judge who performed evaluation")
    team_id: str = Field(..., description="ID of the team being evaluated")
    team_name: str = Field(..., description="Name of the team")
    problem_statement: str = Field(..., description="Problem statement being addressed")
    category: str = Field(..., description="Project category")
    round_id: int = Field(..., description="Evaluation round number")
    scores: JudgeEvaluationScore = Field(..., description="Detailed scores for each criteria")
    total_score: float = Field(..., description="Calculated total score")
    average_score: float = Field(..., description="Average score across all criteria")
    personalized_feedback: str = Field(..., description="Detailed feedback from judge")
    evaluation_status: str = Field(default="draft", description="draft, submitted, finalized")
    evaluated_at: datetime = Field(default_factory=datetime.utcnow, description="When evaluation was performed")
    submitted_at: Optional[datetime] = Field(None, description="When evaluation was submitted")
    
    class Config:
        json_encoders = {
            datetime: lambda v: v.isoformat()
        }

class EvaluationSummary(BaseModel):
    team_id: str
    team_name: str
    round_id: int
    total_evaluations: int
    average_total_score: float
    average_problem_solution_fit: float
    average_functionality_features: float
    average_technical_feasibility: float
    average_innovation_creativity: float
    average_user_experience: float
    average_impact_value: float
    average_presentation_demo_quality: float
    average_team_collaboration: float
    last_updated: datetime

class JudgeEvaluationHistory(BaseModel):
    judge_id: str
    judge_name: str
    total_evaluations: int
    average_scores_given: float
    evaluations: List[TeamEvaluation]

# ------------------ 🔌 MongoDB Collections ------------------

# Create collections if they don't exist
evaluation_criteria_collection = db["evaluation_criteria"]
team_evaluations_collection = db["team_evaluations"]
evaluation_summary_collection = db["evaluation_summary"]
judge_evaluation_history_collection = db["judge_evaluation_history"]

# Initialize default evaluation criteria
default_criteria = [
    {
        "criteria_id": "problem_solution_fit",
        "name": "Problem-Solution Fit",
        "weight": 10.0,
        "description": "How well the prototype addresses the problem statement; alignment with ideation phase."
    },
    {
        "criteria_id": "functionality_features",
        "name": "Functionality & Features",
        "weight": 20.0,
        "description": "Prototype actually works; number of implemented features; handling of real-world cases."
    },
    {
        "criteria_id": "technical_feasibility",
        "name": "Technical Feasibility & Robustness",
        "weight": 20.0,
        "description": "System design, architecture, performance, scalability, basic security."
    },
    {
        "criteria_id": "innovation_creativity",
        "name": "Innovation & Creativity",
        "weight": 15.0,
        "description": "Unique features, creative use of technology, disruptive potential."
    },
    {
        "criteria_id": "user_experience",
        "name": "User Experience (UI/UX)",
        "weight": 15.0,
        "description": "Prototype is easy to use, visually clear, accessible, intuitive."
    },
    {
        "criteria_id": "impact_value",
        "name": "Impact & Value Proposition",
        "weight": 10.0,
        "description": "Social/economic/environmental benefits; practical usefulness."
    },
    {
        "criteria_id": "presentation_demo_quality",
        "name": "Presentation & Demo Quality",
        "weight": 5.0,
        "description": "Clarity of demo, ability to answer judges’ questions, professional presentation."
    },
    {
        "criteria_id": "team_collaboration",
        "name": "Team Collaboration",
        "weight": 5.0,
        "description": "How well the team executed roles, solved challenges, and collaborated."
    }
]

# Insert default criteria if collection is empty
if evaluation_criteria_collection.count_documents({}) == 0:
    evaluation_criteria_collection.insert_many(default_criteria)

