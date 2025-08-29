import React, { useState, useEffect } from 'react';
import { 
  ExternalLink, 
  Download
} from 'lucide-react';
import { useLocation, useNavigate } from 'react-router-dom';
import { getPptReport } from '../utils/api';
import './EvaluateSubmission.css';

const EvaluateSubmission = () => {
  const location = useLocation();
  const navigate = useNavigate();
  const selectedTeam = location.state?.selectedTeam;
  
  const [evaluation, setEvaluation] = useState({
    problemSolutionFit: 5,
    functionalityFeatures: 5,
    technicalFeasibility: 5,
    innovationCreativity: 5,
    userExperience: 5,
    impactValue: 5,
    presentationDemoQuality: 5,
    teamCollaboration: 5,
    finalRecommendation: '',
    feedback: ''
  });

  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitStatus, setSubmitStatus] = useState('');
  const [pptReport, setPptReport] = useState(null);
  const [pptReportError, setPptReportError] = useState('');

  // Load existing evaluation if available
  useEffect(() => {
    if (selectedTeam) {
      loadExistingEvaluation(selectedTeam.team_id);
    }
  }, [selectedTeam]);

  // Load PPT report by team name
  useEffect(() => {
    const fetchPpt = async () => {
      if (!selectedTeam?.team_name) return;
      try {
        setPptReportError('');
        const data = await getPptReport(selectedTeam.team_name);
        setPptReport(data);
      } catch (err) {
        setPptReport(null);
        setPptReportError(err.message || 'Failed to load PPT report');
      }
    };
    fetchPpt();
  }, [selectedTeam?.team_name]);

  const loadExistingEvaluation = async (teamId) => {
    try {
      const judgeToken = localStorage.getItem('judgeToken') || 'your_judge_token_here';
      
      const response = await fetch(`http://localhost:8000/judge/evaluation/${teamId}`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${judgeToken}`
        }
      });

      if (response.ok) {
        const existingEval = await response.json();
        setEvaluation({
          problemSolutionFit: existingEval.problem_solution_fit || 5,
          functionalityFeatures: existingEval.functionality_features || 5,
          technicalFeasibility: existingEval.technical_feasibility || 5,
          innovationCreativity: existingEval.innovation_creativity || 5,
          userExperience: existingEval.user_experience || 5,
          impactValue: existingEval.impact_value || 5,
          presentationDemoQuality: existingEval.presentation_demo_quality || 5,
          teamCollaboration: existingEval.team_collaboration || 5,
          finalRecommendation: existingEval.final_recommendation || '',
          feedback: existingEval.personalized_feedback || ''
        });
      }
    } catch (error) {
      console.log('No existing evaluation found or error loading it');
    }
  };

  // If no team is selected, show selection message
  if (!selectedTeam) {
    return (
      <div className="evaluate-submission">
        <div className="page-header">
          <h1 className="page-title">Evaluate Submission</h1>
          <p className="page-subtitle">Please select a team to evaluate</p>
        </div>
        <div className="no-team-selected">
          <p>No team selected. Please go back and select a team to evaluate.</p>
          <button className="btn btn-primary" onClick={() => navigate('/dashboard')}>
            Go Back
          </button>
        </div>
      </div>
    );
  }

  // Debug: Log the team data structure
  console.log('Selected Team Data:', selectedTeam);
  console.log('Problem Statement Object:', selectedTeam.problem_statement);

  const handleSliderChange = (parameter, value) => {
    setEvaluation(prev => ({
      ...prev,
      [parameter]: parseInt(value)
    }));
  };

  const handleInputChange = (field, value) => {
    setEvaluation(prev => ({
      ...prev,
      [field]: value
    }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    
    if (!evaluation.feedback.trim()) {
      alert('Please provide personalized feedback before submitting.');
      return;
    }

    setIsSubmitting(true);
    setSubmitStatus('Submitting evaluation...');

    try {
      // Prepare data for backend API
      const evaluationData = {
        team_id: selectedTeam.team_id,
        team_name: selectedTeam.team_name,
        problem_statement: selectedTeam.problem_statement?.title || 'No problem statement',
        category: selectedTeam.problem_statement?.category || 'General',
        round_id: 1,
        problem_solution_fit: evaluation.problemSolutionFit,
        functionality_features: evaluation.functionalityFeatures,
        technical_feasibility: evaluation.technicalFeasibility,
        innovation_creativity: evaluation.innovationCreativity,
        user_experience: evaluation.userExperience,
        impact_value: evaluation.impactValue,
        presentation_demo_quality: evaluation.presentationDemoQuality,
        team_collaboration: evaluation.teamCollaboration,
        personalized_feedback: evaluation.feedback
      };

      console.log('Sending evaluation data:', evaluationData);

      // Get judge token from localStorage or context
      const judgeToken = localStorage.getItem('judgeToken') || 'your_judge_token_here';
      
      const response = await fetch('http://localhost:8000/judge/evaluation/submit', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${judgeToken}`
        },
        body: JSON.stringify(evaluationData)
      });

      if (response.ok) {
        const result = await response.json();
        console.log('Evaluation submitted successfully:', result);
        setSubmitStatus('✅ Evaluation submitted successfully!');
        
        // Show success message
        alert(`Evaluation submitted successfully!\nWeighted Total: ${result.total_score}/100\nAverage Score: ${result.average_score}/10`);
        
        // Wait for 2 seconds to show success message, then navigate to dashboard
        setTimeout(() => {
          navigate('/dashboard');
        }, 2000);
        
      } else {
        const errorData = await response.json();
        console.error('Failed to submit evaluation:', errorData);
        setSubmitStatus(`❌ Failed to submit: ${errorData.detail || 'Unknown error'}`);
        alert(`Failed to submit evaluation: ${errorData.detail || 'Unknown error'}`);
      }

    } catch (error) {
      console.error('Error submitting evaluation:', error);
      setSubmitStatus(`❌ Error: ${error.message}`);
      alert(`Error submitting evaluation: ${error.message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleSaveDraft = async () => {
    if (!evaluation.feedback.trim()) {
      alert('Please provide personalized feedback before saving draft.');
      return;
    }

    setIsSubmitting(true);
    setSubmitStatus('Saving draft...');

    try {
      const evaluationData = {
        team_id: selectedTeam.team_id,
        team_name: selectedTeam.team_name,
        problem_statement: selectedTeam.problem_statement?.title || 'No problem statement',
        category: selectedTeam.problem_statement?.category || 'General',
        round_id: 1,
        problem_solution_fit: evaluation.problemSolutionFit,
        functionality_features: evaluation.functionalityFeatures,
        technical_feasibility: evaluation.technicalFeasibility,
        innovation_creativity: evaluation.innovationCreativity,
        user_experience: evaluation.userExperience,
        impact_value: evaluation.impactValue,
        presentation_demo_quality: evaluation.presentationDemoQuality,
        team_collaboration: evaluation.teamCollaboration,
        personalized_feedback: evaluation.feedback
      };

      const judgeToken = localStorage.getItem('judgeToken') || 'your_judge_token_here';
      
      const response = await fetch('http://localhost:8000/judge/evaluation/save-draft', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${judgeToken}`
        },
        body: JSON.stringify(evaluationData)
      });

      if (response.ok) {
        const result = await response.json();
        console.log('Draft saved successfully:', result);
        setSubmitStatus('✅ Draft saved successfully!');
        alert('Draft saved successfully!');
        
        // Wait for 2 seconds to show success message, then navigate to dashboard
        setTimeout(() => {
          navigate('/dashboard');
        }, 2000);
      } else {
        const errorData = await response.json();
        setSubmitStatus(`❌ Failed to save draft: ${errorData.detail || 'Unknown error'}`);
        alert(`Failed to save draft: ${errorData.detail || 'Unknown error'}`);
      }

    } catch (error) {
      console.error('Error saving draft:', error);
      setSubmitStatus(`❌ Error: ${error.message}`);
      alert(`Error saving draft: ${error.message}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const getScoreColor = (score) => {
    if (score >= 8) return 'success';
    if (score >= 6) return 'warning';
    return 'error';
  };

  const getUniquenessColor = (score) => {
    if (score >= 80) return 'success';
    if (score >= 60) return 'warning';
    return 'error';
  };

  const getPlagiarismColor = (score) => {
    if (score <= 15) return 'success';
    if (score <= 30) return 'warning';
    return 'error';
  };

  return (
    <div className="evaluate-submission">
      <div className="page-header">
        <h1 className="page-title">Evaluate Submission</h1>
        <p className="page-subtitle">Review and evaluate team submission</p>
        <button className="btn btn-secondary back-btn" onClick={() => navigate('/dashboard')}>
          ← Back to Teams
        </button>
      </div>

      <div className="evaluation-container">
        {/* Project Metadata Section */}
        <div className="metadata-section card">
          <div className="metadata-header">
            <div className="team-info">
              <h2>{selectedTeam.team_name}</h2>
              <p className="team-id">ID: {selectedTeam.team_id}</p>
              <span className="category-badge">{selectedTeam.problem_statement?.category || 'General'}</span>
            </div>
            <div className="submission-meta">
              <p className="submission-date">Submitted: {selectedTeam.submission_date || 'N/A'}</p>
            </div>
          </div>

          <div className="metadata-content">
            <div className="metadata-item">
              <h3>Problem Statement</h3>
              <p className="problem-title">{selectedTeam.problem_statement?.title || 'No problem statement available'}</p>
              {selectedTeam.problem_statement?.description && (
                <div className="problem-description">
                  <p>{selectedTeam.problem_statement.description}</p>
                </div>
              )}
            </div>

            <div className="metadata-item">
              <h3>Team Information</h3>
              <div className="team-details-grid">
                <div className="detail-item">
                  <strong>Department:</strong> {selectedTeam.department || 'Not specified'}
                </div>
                <div className="detail-item">
                  <strong>Year:</strong> {selectedTeam.year || 'Not specified'}
                </div>
                <div className="detail-item">
                  <strong>Status:</strong> <span className={`status-badge ${selectedTeam.status || 'inactive'}`}>{selectedTeam.status || 'Inactive'}</span>
                </div>
              </div>
            </div>

            <div className="metadata-item">
              <h3>Team Leader</h3>
              {selectedTeam.team_leader ? (
                <div className="team-leader-info">
                  <p><strong>Name:</strong> {selectedTeam.team_leader.name}</p>
                  <p><strong>Roll No:</strong> {selectedTeam.team_leader.roll_no}</p>
                  <p><strong>Email:</strong> {selectedTeam.team_leader.email}</p>
                  <p><strong>Contact:</strong> {selectedTeam.team_leader.contact}</p>
                  <p><strong>Role:</strong> {selectedTeam.team_leader.role}</p>
                </div>
              ) : (
                <p>No team leader information available</p>
              )}
            </div>

            <div className="metadata-item">
              <h3>Team Members</h3>
              {selectedTeam.team_members && selectedTeam.team_members.length > 0 ? (
                <div className="team-members-list">
                  {selectedTeam.team_members.map((member, index) => (
                    <div key={index} className="member-card">
                      <div className="member-header">
                        <h4>{member.name}</h4>
                        <span className="member-role">{member.role}</span>
                      </div>
                      <div className="member-details">
                        <p><strong>Roll No:</strong> {member.roll_no}</p>
                        <p><strong>Email:</strong> {member.email}</p>
                        <p><strong>Contact:</strong> {member.contact}</p>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <p>No team members information available</p>
              )}
            </div>

            <div className="metadata-item">
              <h3>Project Details</h3>
              <div className="project-details">
                <div className="detail-row">
                  <strong>Category:</strong> {selectedTeam.problem_statement?.category || 'Not specified'}
                </div>
                <div className="detail-row">
                  <strong>Domain:</strong> {selectedTeam.problem_statement?.domain || 'Not specified'}
                </div>
                
                <div className="detail-row">
                  <strong>Problem ID:</strong> {selectedTeam.problem_statement?.ps_id || 'Not specified'}
                </div>
              </div>
            </div>


           

            {/* Debug: Show raw data for development */}
            
          </div>
        </div>

        {/* PPT Report Summary (from MongoDB ppt_reports) */}
        <div className="metadata-section card">
          <div className="metadata-header">
            <div className="team-info">
              <h2>PPT Report By AI</h2>
              <p className="team-id">Auto-analysis from uploaded PPT report</p>
            </div>
          </div>
          <div className="metadata-content">
            {pptReportError && (
              <div className="submit-status error">{pptReportError}</div>
            )}
            {pptReport ? (
              <>
                <div className="metadata-item">
                  <h3>Summary</h3>
                  <p>{pptReport.summary || 'No summary available'}</p>
                </div>
                <div className="metadata-item">
                  <h3>Key Scores</h3>
                  <div className="team-details-grid">
                    <div className="detail-item"><strong>Total Weighted:</strong> {pptReport.scores?.total_weighted ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Total Raw:</strong> {pptReport.scores?.total_raw ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Problem Understanding:</strong> {pptReport.scores?.['Problem Understanding'] ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Innovation & Uniqueness:</strong> {pptReport.scores?.['Innovation & Uniqueness'] ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Technical Feasibility:</strong> {pptReport.scores?.['Technical Feasibility'] ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Implementation Approach:</strong> {pptReport.scores?.['Implementation Approach'] ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Team Readiness:</strong> {pptReport.scores?.['Team Readiness'] ?? 'N/A'}</div>
                    <div className="detail-item"><strong>Potential Impact:</strong> {pptReport.scores?.['Potential Impact'] ?? 'N/A'}</div>
                  </div>
                </div>
                <div className="metadata-item">
                  <h3>Feedback Highlights</h3>
                  <div className="project-details">
                    <div className="detail-row"><strong>Positive:</strong> {pptReport.feedback_positive || 'N/A'}</div>
                    <div className="detail-row"><strong>Criticism:</strong> {pptReport.feedback_criticism || 'N/A'}</div>
                    <div className="detail-row"><strong>Technical:</strong> {pptReport.feedback_technical || 'N/A'}</div>
                    <div className="detail-row"><strong>Suggestions:</strong> {pptReport.feedback_suggestions || 'N/A'}</div>
                  </div>
                </div>
              </>
            ) : !pptReportError ? (
              <p>Loading PPT report...</p>
            ) : (
              <p>No PPT report found for this team.</p>
            )}
          </div>
        </div>

        {/* Evaluation Form */}
        <form className="evaluation-form card" onSubmit={handleSubmit}>
          <div className="form-header">
            <h2>Judging Parameters</h2>
            <p>Rate each parameter on a scale of 1-10</p>
          </div>

          <div className="parameters-grid">
            {[
              { key: 'problemSolutionFit', label: 'Problem-Solution Fit (10%)', description: 'How well the prototype addresses the problem statement; alignment with ideation phase.' },
              { key: 'functionalityFeatures', label: 'Functionality & Features (20%)', description: 'Prototype actually works; number of implemented features; handling of real-world cases.' },
              { key: 'technicalFeasibility', label: 'Technical Feasibility & Robustness (20%)', description: 'System design, architecture, performance, scalability, basic security.' },
              { key: 'innovationCreativity', label: 'Innovation & Creativity (15%)', description: 'Unique features, creative use of technology, disruptive potential.' },
              { key: 'userExperience', label: 'User Experience (UI/UX) (15%)', description: 'Ease of use, visual clarity, accessibility, intuitiveness.' },
              { key: 'impactValue', label: 'Impact & Value Proposition (10%)', description: 'Social/economic/environmental benefits; practical usefulness.' },
              { key: 'presentationDemoQuality', label: 'Presentation & Demo Quality (5%)', description: 'Clarity of demo, ability to answer questions, professionalism.' },
              { key: 'teamCollaboration', label: 'Team Collaboration (5%)', description: 'Execution of roles, solving challenges, collaboration.' }
            ].map((param) => (
              <div key={param.key} className="parameter-item">
                <div className="parameter-header">
                  <label className="parameter-label">{param.label}</label>
                  <span className={`parameter-score score-${getScoreColor(evaluation[param.key])}`}>
                    {evaluation[param.key]}/10
                  </span>
                </div>
                <p className="parameter-description">{param.description}</p>
                <div className="slider-container">
                  <input
                    type="range"
                    min="1"
                    max="10"
                    value={evaluation[param.key]}
                    onChange={(e) => handleSliderChange(param.key, e.target.value)}
                    className="slider"
                  />
                  <div className="slider-labels">
                    <span>Poor</span>
                    <span>Excellent</span>
                  </div>
                </div>
              </div>
            ))}
          </div>

          <div className="form-section">
            <label className="form-label">Personalized Feedback</label>
            <textarea
              value={evaluation.feedback}
              onChange={(e) => handleInputChange('feedback', e.target.value)}
              className="form-textarea"
              placeholder="Provide detailed feedback for the team..."
              rows="4"
              required
            />
          </div>

          <div className="form-actions">
            <button 
              type="button" 
              className="btn btn-secondary"
              onClick={handleSaveDraft}
              disabled={isSubmitting}
            >
              {isSubmitting ? 'Saving...' : 'Save Draft'}
            </button>
            <button 
              type="submit" 
              className="btn btn-primary"
              disabled={isSubmitting}
            >
              {isSubmitting ? 'Submitting...' : 'Submit Evaluation'}
            </button>
          </div>

          {submitStatus && (
            <div className={`submit-status ${submitStatus.includes('✅') ? 'success' : submitStatus.includes('❌') ? 'error' : 'info'}`}>
              {submitStatus}
            </div>
          )}
        </form>
      </div>
    </div>
  );
};

export default EvaluateSubmission;