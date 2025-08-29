const API_BASE_URL = 'http://localhost:8000';

// Helper function to get auth headers
export const getAuthHeaders = () => {
    const token = localStorage.getItem('judgeToken');
    return {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
    };
};

// Judge authentication
export const judgeLogin = async (username, password) => {
    try {
        const formData = new FormData();
        formData.append('username', username);
        formData.append('password', password);

        console.log('Attempting login with:', { username, password: '***' });
        console.log('API URL:', `${API_BASE_URL}/auth/judge/login`);

        const response = await fetch(`${API_BASE_URL}/auth/judge/login`, {
            method: 'POST',
            body: formData,
        });

        console.log('Response status:', response.status);
        console.log('Response headers:', response.headers);

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ detail: 'Unknown error' }));
            console.error('Backend error:', errorData);
            throw new Error(errorData.detail || `HTTP ${response.status}: Login failed`);
        }

        const data = await response.json();
        console.log('Login successful, received token');
        return data;
    } catch (error) {
        console.error('Login error details:', error);
        if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
            throw new Error('Cannot connect to backend server. Please check if the server is running on port 8000.');
        }
        throw error;
    }
};

// Get judge profile
export const getJudgeProfile = async () => {
    const response = await fetch(`${API_BASE_URL}/judge/profile`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch profile');
    }

    return response.json();
};

// Get assigned teams
export const getAssignedTeams = async (roundId = null) => {
    const url = roundId 
        ? `${API_BASE_URL}/judge/assigned-teams?round_id=${roundId}`
        : `${API_BASE_URL}/judge/assigned-teams`;
    
    const response = await fetch(url, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch assigned teams');
    }

    return response.json();
};

// Submit evaluation
export const submitEvaluation = async (teamId, evaluation) => {
    const response = await fetch(`${API_BASE_URL}/judge/evaluate/${teamId}`, {
        method: 'POST',
        headers: getAuthHeaders(),
        body: JSON.stringify(evaluation),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to submit evaluation');
    }

    return response.json();
};

// Get evaluations
export const getEvaluations = async (roundId = null, teamId = null) => {
    let url = `${API_BASE_URL}/judge/evaluations`;
    const params = new URLSearchParams();
    
    if (roundId) params.append('round_id', roundId);
    if (teamId) params.append('team_id', teamId);
    
    if (params.toString()) {
        url += `?${params.toString()}`;
    }

    const response = await fetch(url, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch evaluations');
    }

    return response.json();
};

// Get judge's own evaluations
export const getMyEvaluations = async () => {
    const response = await fetch(`${API_BASE_URL}/judge/evaluation/my-evaluations`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch my evaluations');
    }

    return response.json();
};

// Get team details by team ID
export const getTeamDetails = async (teamId) => {
    const response = await fetch(`${API_BASE_URL}/team-ps/teams/${teamId}`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch team details');
    }

    return response.json();
};

// Get all teams with problem statement details
export const getAllTeams = async () => {
    const response = await fetch(`${API_BASE_URL}/judge/all-teams`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to fetch teams');
    }

    return response.json();
};

// Count evaluations by team name
export const countEvaluationsByTeamName = async (teamName) => {
    const response = await fetch(`${API_BASE_URL}/judge/evaluation/count-by-team-name/${encodeURIComponent(teamName)}`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to count evaluations');
    }

    return response.json();
};

// Check if user is authenticated
export const isAuthenticated = () => {
    const token = localStorage.getItem('judgeToken');
    return !!token;
};

// Logout
export const logout = () => {
    localStorage.removeItem('judgeToken');
    localStorage.removeItem('judgeUsername');
    window.location.href = '/';
};

// Fetch PPT report by team name (from hackathon_evaluation.ppt_reports)
export const getPptReport = async (teamName) => {
    const response = await fetch(`${API_BASE_URL}/api/ppt-report/${encodeURIComponent(teamName)}`, {
        headers: getAuthHeaders(),
    });

    if (!response.ok) {
        const errorData = await response.json().catch(() => ({ detail: 'Unknown error' }));
        throw new Error(errorData.detail || 'Failed to fetch PPT report');
    }

    return response.json();
};

// Search PPT reports by team name (partial, case-insensitive)
export const searchPptReports = async (teamNameQuery) => {
    const url = `${API_BASE_URL}/api/ppt-reports?team_name=${encodeURIComponent(teamNameQuery)}`;
    const response = await fetch(url, {
        headers: getAuthHeaders(),
    });
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({ detail: 'Unknown error' }));
        throw new Error(errorData.detail || 'Failed to search PPT reports');
    }
    return response.json();
};