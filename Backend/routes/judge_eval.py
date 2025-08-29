# judge_routes.py
from fastapi import APIRouter, HTTPException, Depends, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import List, Optional
from datetime import datetime
import uuid
from bson import ObjectId

from Schema.evaluation import (
    TeamEvaluation, 
    JudgeEvaluationScore, 
    EvaluationSummary,
    JudgeEvaluationHistory,
    team_evaluations_collection,
    evaluation_summary_collection,
    judge_evaluation_history_collection
)
from Schema.judge import JudgeModel
from auth.auth_middleware import get_current_judge
from db.mongo import db

router = APIRouter(tags=["Judge Evaluation"])
security = HTTPBearer()

# ==================== EVALUATION CRUD OPERATIONS ====================

@router.post("/submit", response_model=dict)
async def submit_evaluation(
    evaluation_data: dict,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """
    Submit a complete evaluation for a team
    """
    try:
        # Get current judge from token
        current_judge = await get_current_judge(credentials.credentials)
        
        # Generate unique evaluation ID
        evaluation_id = str(uuid.uuid4())
        
        # Extract scores from the evaluation data (new 8-criteria)
        scores = JudgeEvaluationScore(
            problem_solution_fit=float(evaluation_data.get("problem_solution_fit", 5)),
            functionality_features=float(evaluation_data.get("functionality_features", 5)),
            technical_feasibility=float(evaluation_data.get("technical_feasibility", 5)),
            innovation_creativity=float(evaluation_data.get("innovation_creativity", 5)),
            user_experience=float(evaluation_data.get("user_experience", 5)),
            impact_value=float(evaluation_data.get("impact_value", 5)),
            presentation_demo_quality=float(evaluation_data.get("presentation_demo_quality", 5)),
            team_collaboration=float(evaluation_data.get("team_collaboration", 5))
        )
        
        # Calculate weighted total based on rubric weights
        weights = {
            "problem_solution_fit": 10,
            "functionality_features": 20,
            "technical_feasibility": 20,
            "innovation_creativity": 15,
            "user_experience": 15,
            "impact_value": 10,
            "presentation_demo_quality": 5,
            "team_collaboration": 5,
        }
        raw_scores = [
            scores.problem_solution_fit,
            scores.functionality_features,
            scores.technical_feasibility,
            scores.innovation_creativity,
            scores.user_experience,
            scores.impact_value,
            scores.presentation_demo_quality,
            scores.team_collaboration,
        ]
        average_score = sum(raw_scores) / 8
        total_score = (
            scores.problem_solution_fit * weights["problem_solution_fit"] +
            scores.functionality_features * weights["functionality_features"] +
            scores.technical_feasibility * weights["technical_feasibility"] +
            scores.innovation_creativity * weights["innovation_creativity"] +
            scores.user_experience * weights["user_experience"] +
            scores.impact_value * weights["impact_value"] +
            scores.presentation_demo_quality * weights["presentation_demo_quality"] +
            scores.team_collaboration * weights["team_collaboration"]
        ) / 10.0
        
        # Create evaluation object
        evaluation = TeamEvaluation(
            evaluation_id=evaluation_id,
            judge_id=current_judge["id"],
            team_id=evaluation_data["team_id"],
            team_name=evaluation_data["team_name"],
            problem_statement=evaluation_data["problem_statement"],
            category=evaluation_data.get("category", "General"),
            round_id=evaluation_data.get("round_id", 1),
            scores=scores,
            total_score=total_score,
            average_score=round(average_score, 2),
            personalized_feedback=evaluation_data.get("personalized_feedback", ""),
            evaluation_status="submitted",
            evaluated_at=datetime.utcnow(),
            submitted_at=datetime.utcnow()
        )
        
        # Save to database
        result = await team_evaluations_collection.insert_one(evaluation.dict())
        
        if result.inserted_id:
            # Update evaluation summary
            await update_evaluation_summary(evaluation.team_id, evaluation.round_id)
            
            return {
                "success": True,
                "message": "Evaluation submitted successfully",
                "evaluation_id": evaluation_id,
                "total_score": total_score,
                "average_score": average_score
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to save evaluation")
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error submitting evaluation: {str(e)}")

@router.post("/save-draft", response_model=dict)
async def save_evaluation_draft(
    evaluation_data: dict,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """
    Save evaluation as draft (can be updated later)
    """
    try:
        current_judge = await get_current_judge(credentials.credentials)
        
        # Check if draft already exists
        existing_draft = await team_evaluations_collection.find_one({
            "judge_id": current_judge["id"],
            "team_id": evaluation_data["team_id"],
            "round_id": evaluation_data.get("round_id", 1),
            "evaluation_status": "draft"
        })
        
        if existing_draft:
            # Update existing draft
            evaluation_id = existing_draft["evaluation_id"]
            update_data = {
                "scores": {
                    "problem_solution_fit": float(evaluation_data.get("problem_solution_fit", 5)),
                    "functionality_features": float(evaluation_data.get("functionality_features", 5)),
                    "technical_feasibility": float(evaluation_data.get("technical_feasibility", 5)),
                    "innovation_creativity": float(evaluation_data.get("innovation_creativity", 5)),
                    "user_experience": float(evaluation_data.get("user_experience", 5)),
                    "impact_value": float(evaluation_data.get("impact_value", 5)),
                    "presentation_demo_quality": float(evaluation_data.get("presentation_demo_quality", 5)),
                    "team_collaboration": float(evaluation_data.get("team_collaboration", 5))
                },
                "personalized_feedback": evaluation_data.get("personalized_feedback", ""),
                "evaluated_at": datetime.utcnow()
            }
            
            # Calculate scores
            scores = evaluation_data["scores"]
            weights = {
                "problem_solution_fit": 10,
                "functionality_features": 20,
                "technical_feasibility": 20,
                "innovation_creativity": 15,
                "user_experience": 15,
                "impact_value": 10,
                "presentation_demo_quality": 5,
                "team_collaboration": 5,
            }
            total_score = sum([scores[k] * weights[k] for k in weights]) / 10.0
            average_score = sum(scores.values()) / 8
            
            update_data["total_score"] = total_score
            update_data["average_score"] = round(average_score, 2)
            
            await team_evaluations_collection.update_one(
                {"evaluation_id": evaluation_id},
                {"$set": update_data}
            )
            
            return {
                "success": True,
                "message": "Draft updated successfully",
                "evaluation_id": evaluation_id
            }
        else:
            # Create new draft
            evaluation_id = str(uuid.uuid4())
            
            scores = {
                "problem_solution_fit": float(evaluation_data.get("problem_solution_fit", 5)),
                "functionality_features": float(evaluation_data.get("functionality_features", 5)),
                "technical_feasibility": float(evaluation_data.get("technical_feasibility", 5)),
                "innovation_creativity": float(evaluation_data.get("innovation_creativity", 5)),
                "user_experience": float(evaluation_data.get("user_experience", 5)),
                "impact_value": float(evaluation_data.get("impact_value", 5)),
                "presentation_demo_quality": float(evaluation_data.get("presentation_demo_quality", 5)),
                "team_collaboration": float(evaluation_data.get("team_collaboration", 5))
            }
            
            weights = {
                "problem_solution_fit": 10,
                "functionality_features": 20,
                "technical_feasibility": 20,
                "innovation_creativity": 15,
                "user_experience": 15,
                "impact_value": 10,
                "presentation_demo_quality": 5,
                "team_collaboration": 5,
            }
            total_score = sum([scores[k] * weights[k] for k in weights]) / 10.0
            average_score = sum(scores.values()) / 8
            
            draft_evaluation = TeamEvaluation(
                evaluation_id=evaluation_id,
                judge_id=current_judge["id"],
                team_id=evaluation_data["team_id"],
                team_name=evaluation_data["team_name"],
                problem_statement=evaluation_data["problem_statement"],
                category=evaluation_data.get("category", "General"),
                round_id=evaluation_data.get("round_id", 1),
                scores=JudgeEvaluationScore(**scores),
                total_score=total_score,
                average_score=round(average_score, 2),
                personalized_feedback=evaluation_data.get("personalized_feedback", ""),
                evaluation_status="draft",
                evaluated_at=datetime.utcnow()
            )
            
            result = await team_evaluations_collection.insert_one(draft_evaluation.dict())
            
            if result.inserted_id:
                return {
                    "success": True,
                    "message": "Draft saved successfully",
                    "evaluation_id": evaluation_id
                }
            else:
                raise HTTPException(status_code=500, detail="Failed to save draft")
                
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error saving draft: {str(e)}")

@router.get("/my-evaluations", response_model=List[dict])
async def get_my_evaluations(
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """
    Get all evaluations submitted by the current judge
    """
    try:
        current_judge = await get_current_judge(credentials.credentials)
        
        evaluations = list(await team_evaluations_collection.find({
            "judge_id": current_judge["id"]
        }).to_list(None))
        
        # Convert ObjectId to string for JSON serialization
        for eval in evaluations:
            eval["_id"] = str(eval["_id"])
            eval["evaluated_at"] = eval["evaluated_at"].isoformat()
            if eval.get("submitted_at"):
                eval["submitted_at"] = eval["submitted_at"].isoformat()
        
        return evaluations
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error fetching evaluations: {str(e)}")

@router.get("/team/{team_id}", response_model=dict)
async def get_team_evaluation(
    team_id: str,
    round_id: int = 1,
    credentials: HTTPAuthorizationCredentials = Depends(security)
):
    """
    Get evaluation for a specific team by the current judge
    """
    try:
        current_judge = await get_current_judge(credentials.credentials)
        
        evaluation = await team_evaluations_collection.find_one({
            "judge_id": current_judge["id"],
            "team_id": team_id,
            "round_id": round_id
        })
        
        if evaluation:
            evaluation["_id"] = str(evaluation["_id"])
            evaluation["evaluated_at"] = evaluation["evaluated_at"].isoformat()
            if evaluation.get("submitted_at"):
                evaluation["submitted_at"] = evaluation["submitted_at"].isoformat()
            return evaluation
        else:
            return {"message": "No evaluation found for this team"}
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error fetching team evaluation: {str(e)}")

@router.get("/summary/{team_id}", response_model=dict)
async def get_team_evaluation_summary(
    team_id: str,
    round_id: int = 1
):
    """
    Get evaluation summary for a team (aggregated scores from all judges)
    """
    try:
        summary = await evaluation_summary_collection.find_one({
            "team_id": team_id,
            "round_id": round_id
        })
        
        if summary:
            summary["_id"] = str(summary["_id"])
            summary["last_updated"] = summary["last_updated"].isoformat()
            return summary
        else:
            return {"message": "No evaluation summary found for this team"}
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error fetching evaluation summary: {str(e)}")

# ==================== HELPER FUNCTIONS ====================

async def update_evaluation_summary(team_id: str, round_id: int):
    """
    Update evaluation summary for a team after new evaluation is submitted
    """
    try:
        # Get all evaluations for this team and round
        evaluations = list(await team_evaluations_collection.find({
            "team_id": team_id,
            "round_id": round_id,
            "evaluation_status": "submitted"
        }).to_list(None))
        
        if not evaluations:
            return
        
        # Calculate averages
        total_evaluations = len(evaluations)
        avg_total_score = sum(eval["total_score"] for eval in evaluations) / total_evaluations
        avg_problem_solution_fit = sum(eval["scores"]["problem_solution_fit"] for eval in evaluations) / total_evaluations
        avg_functionality_features = sum(eval["scores"]["functionality_features"] for eval in evaluations) / total_evaluations
        avg_technical_feasibility = sum(eval["scores"]["technical_feasibility"] for eval in evaluations) / total_evaluations
        avg_innovation_creativity = sum(eval["scores"]["innovation_creativity"] for eval in evaluations) / total_evaluations
        avg_user_experience = sum(eval["scores"]["user_experience"] for eval in evaluations) / total_evaluations
        avg_impact_value = sum(eval["scores"]["impact_value"] for eval in evaluations) / total_evaluations
        avg_presentation_demo_quality = sum(eval["scores"]["presentation_demo_quality"] for eval in evaluations) / total_evaluations
        avg_team_collaboration = sum(eval["scores"]["team_collaboration"] for eval in evaluations) / total_evaluations
        
        # Get team name from first evaluation
        team_name = evaluations[0]["team_name"]
        
        summary_data = {
            "team_id": team_id,
            "team_name": team_name,
            "round_id": round_id,
            "total_evaluations": total_evaluations,
            "average_total_score": round(avg_total_score, 2),
            "average_problem_solution_fit": round(avg_problem_solution_fit, 2),
            "average_functionality_features": round(avg_functionality_features, 2),
            "average_technical_feasibility": round(avg_technical_feasibility, 2),
            "average_innovation_creativity": round(avg_innovation_creativity, 2),
            "average_user_experience": round(avg_user_experience, 2),
            "average_impact_value": round(avg_impact_value, 2),
            "average_presentation_demo_quality": round(avg_presentation_demo_quality, 2),
            "average_team_collaboration": round(avg_team_collaboration, 2),
            "last_updated": datetime.utcnow()
        }
        
        # Upsert summary
        await evaluation_summary_collection.update_one(
            {"team_id": team_id, "round_id": round_id},
            {"$set": summary_data},
            upsert=True
        )
        
    except Exception as e:
        print(f"Error updating evaluation summary: {str(e)}")

# ==================== ADMIN ENDPOINTS ====================

@router.get("/admin/all-evaluations", response_model=List[dict])
async def get_all_evaluations(
    team_id: Optional[str] = None,
    round_id: Optional[int] = None,
    judge_id: Optional[str] = None
):
    """
    Admin endpoint to get all evaluations with optional filters
    """
    try:
        filter_query = {}
        if team_id:
            filter_query["team_id"] = team_id
        if round_id:
            filter_query["round_id"] = round_id
        if judge_id:
            filter_query["judge_id"] = judge_id
            
        evaluations = list(await team_evaluations_collection.find(filter_query).to_list(None))
        
        # Convert ObjectId to string
        for eval in evaluations:
            eval["_id"] = str(eval["_id"])
            eval["evaluated_at"] = eval["evaluated_at"].isoformat()
            if eval.get("submitted_at"):
                eval["submitted_at"] = eval["submitted_at"].isoformat()
        
        return evaluations
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error fetching all evaluations: {str(e)}")

@router.get("/admin/leaderboard", response_model=List[dict])
async def get_evaluation_leaderboard(round_id: int = 1):
    """
    Get leaderboard based on evaluation scores
    """
    try:
        pipeline = [
            {"$match": {"round_id": round_id, "evaluation_status": "submitted"}},
            {"$group": {
                "_id": "$team_id",
                "team_name": {"$first": "$team_name"},
                "total_score": {"$avg": "$total_score"},
                "average_score": {"$avg": "$average_score"},
                "evaluation_count": {"$sum": 1}
            }},
            {"$sort": {"total_score": -1}},
            {"$project": {
                "team_id": "$_id",
                "team_name": 1,
                "total_score": {"$round": ["$total_score", 2]},
                "average_score": {"$round": ["$average_score", 2]},
                "evaluation_count": 1
            }}
        ]
        
        leaderboard = list(await evaluation_summary_collection.aggregate(pipeline).to_list(None))
        
        # Add rank
        for i, entry in enumerate(leaderboard):
            entry["rank"] = i + 1
            
        return leaderboard
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error fetching leaderboard: {str(e)}")

@router.get("/count-by-team-name/{team_name}", response_model=dict)
async def count_evaluations_by_team_name(team_name: str):
    """
    Count all evaluations for a specific team name
    """
    try:
        # Count evaluations for the specific team name
        count = await team_evaluations_collection.count_documents({
            "team_name": team_name
        })
        
        return {
            "team_name": team_name,
            "evaluation_count": count,
            "message": f"Found {count} evaluations for team '{team_name}'"
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error counting evaluations: {str(e)}")
