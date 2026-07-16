import { useEffect, useState } from 'react';
import { createClient } from '@supabase/supabase-js';
import './App.css';

// Retrieve config from environment variables
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const apiUrl = import.meta.env.VITE_API_URL || 'http://localhost:8000';

// Check if Supabase keys have been configured (not placeholders)
const isConfigured =
  supabaseUrl &&
  supabaseAnonKey &&
  supabaseAnonKey !== 'your-supabase-anon-key-placeholder' &&
  !supabaseAnonKey.includes('placeholder');

// Initialize Supabase Client if configured
const supabase = isConfigured ? createClient(supabaseUrl, supabaseAnonKey) : null;

function App() {
  const [session, setSession] = useState(null);
  const [authError, setAuthError] = useState(null);
  const [testResult, setTestResult] = useState(null);
  const [testing, setTesting] = useState(false);
  const [testError, setTestError] = useState(null);

  // Monitor auth status on load
  useEffect(() => {
    if (!supabase) return;

    // Get current session
    supabase.auth.getSession().then(({ data: { session }, error }) => {
      if (error) {
        setAuthError(error.message);
      } else {
        setSession(session);
      }
    });

    // Listen for auth state changes
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      // Reset verification test state when user logs out/in
      if (!session) {
        setTestResult(null);
        setTestError(null);
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  // Handle Login with Microsoft
  const handleLogin = async () => {
    if (!supabase) return;
    setAuthError(null);

    const { error } = await supabase.auth.signInWithOAuth({
      provider: 'azure',
      options: {
        // Scopes requested from Microsoft Entra ID
        scopes: 'User.Read Mail.Read offline_access openid profile email',
        redirectTo: window.location.origin,
      },
    });

    if (error) {
      setAuthError(error.message);
    }
  };

  // Handle Logout
  const handleLogout = async () => {
    if (!supabase) return;
    await supabase.auth.signOut();
  };

  // Call the backend to verify Mailbox access
  const handleVerifyMailbox = async () => {
    if (!session?.provider_token) {
      setTestError('No Microsoft Access Token available in the current session. Try logging in again.');
      return;
    }

    setTesting(true);
    setTestError(null);
    setTestResult(null);

    try {
      const response = await fetch(`${apiUrl}/api/test-mail`, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${session.provider_token}`,
          'Content-Type': 'application/json',
        },
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Server returned error status (${response.status}): ${errText}`);
      }

      const data = await response.json();
      setTestResult(data);
      if (!data.success) {
        setTestError(data.error || 'Failed to verify mailbox access.');
      }
    } catch (err) {
      setTestError(err.message || 'Failed to connect to backend endpoint.');
    } finally {
      setTesting(false);
    }
  };

  // Helper to extract user initials for avatar
  const getInitials = (name) => {
    if (!name) return 'U';
    return name
      .split(' ')
      .map((n) => n[0])
      .slice(0, 2)
      .join('')
      .toUpperCase();
  };

  return (
    <div className="app-container">
      {/* Background Radial Glow */}
      <div className="background-glow" />

      <header className="header animate-fade-in">
        <div className="logo-container">
          <span className="logo-tag">POC</span>
          <h1 className="logo-title">Provident</h1>
        </div>
        <p className="header-subtitle">Microsoft 365 Mailbox Verification System</p>
      </header>

      <main className="main-content">
        {/* If Supabase configuration is missing */}
        {!isConfigured && (
          <div className="card glass warning-card animate-slide-up">
            <div className="warning-icon-wrapper">⚠️</div>
            <h3>Configuration Required</h3>
            <p>
              Please edit the <code>dashboard/.env</code> file and replace the placeholder keys with your actual
              Supabase Project credentials:
            </p>
            <div className="code-block">
              <code>VITE_SUPABASE_URL=https://vkensxazltfyweijkywi.supabase.co</code>
              <code>VITE_SUPABASE_ANON_KEY=your-supabase-anon-key-here</code>
            </div>
            <p className="hint">
              Once configured, restart the Vite dev server to apply variables.
            </p>
          </div>
        )}

        {/* If Configured but Not Signed In */}
        {isConfigured && !session && (
          <div className="card glass login-card animate-slide-up">
            <h2>Verify Mailbox Access</h2>
            <p className="card-description">
              Please authenticate with your Microsoft 365 account to verify if the application has permission to
              connect to your mailbox and read messages.
            </p>

            <div className="permissions-box">
              <h4>Requested Delegated Permissions:</h4>
              <ul className="permissions-list">
                <li>
                  <span className="bullet">✓</span>
                  <div>
                    <strong>User.Read</strong> - Retrieve profile name and photo
                  </div>
                </li>
                <li>
                  <span className="bullet">✓</span>
                  <div>
                    <strong>Mail.Read</strong> - Read unread and read emails in your mailbox
                  </div>
                </li>
                <li>
                  <span className="bullet">✓</span>
                  <div>
                    <strong>offline_access</strong> - Maintain mailbox sync permissions
                  </div>
                </li>
                <li>
                  <span className="bullet">✓</span>
                  <div>
                    <strong>openid, profile, email</strong> - Standard identity information
                  </div>
                </li>
              </ul>
            </div>

            {authError && <div className="error-banner">⚠️ Authentication Error: {authError}</div>}

            <button className="btn btn-microsoft" onClick={handleLogin}>
              <svg className="microsoft-logo" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
                <rect x="0" y="0" width="11" height="11" fill="#f25022" />
                <rect x="12" y="0" width="11" height="11" fill="#7fba00" />
                <rect x="0" y="12" width="11" height="11" fill="#00a4ef" />
                <rect x="12" y="12" width="11" height="11" fill="#ffb900" />
              </svg>
              <span>Sign in with Microsoft</span>
            </button>
          </div>
        )}

        {/* If Signed In */}
        {isConfigured && session && (
          <div className="dashboard-layout animate-slide-up">
            {/* User Profile Card */}
            <div className="card glass user-card">
              <div className="user-profile-header">
                <div className="avatar">
                  {getInitials(session.user?.user_metadata?.full_name || session.user?.email)}
                </div>
                <div className="user-info">
                  <h3>{session.user?.user_metadata?.full_name || 'Microsoft User'}</h3>
                  <span className="user-email">{session.user?.email}</span>
                </div>
              </div>

              <div className="token-details">
                <div className="detail-row">
                  <span className="label">Auth Provider:</span>
                  <span className="val badge">Microsoft (Azure)</span>
                </div>
                <div className="detail-row">
                  <span className="label">Access Token:</span>
                  <span className="val token-status">
                    {session.provider_token ? '✓ Available (Delegated)' : '❌ Missing'}
                  </span>
                </div>
              </div>

              <div className="actions">
                <button
                  className="btn btn-primary"
                  onClick={handleVerifyMailbox}
                  disabled={testing}
                >
                  {testing ? (
                    <>
                      <span className="spinner" /> Testing Access...
                    </>
                  ) : (
                    'Run Mailbox Access Test'
                  )}
                </button>
                <button className="btn btn-secondary" onClick={handleLogout}>
                  Sign Out
                </button>
              </div>
            </div>

            {/* Test Verification Results */}
            {(testResult || testError) && (
              <div className="card glass results-card animate-fade-in">
                <h2>Verification Results</h2>

                {testError && (
                  <div className="result-status failed animate-pulse">
                    <div className="status-header">
                      <span className="status-badge failed-badge">FAILED</span>
                      <h4>Connection Denied</h4>
                    </div>
                    <p className="error-message">{testError}</p>
                  </div>
                )}

                {testResult?.success && (
                  <div className="result-status success animate-fade-in">
                    <div className="status-header">
                      <span className="status-badge success-badge">SUCCESSFUL</span>
                      <h4>Delegated Mailbox Access Verified</h4>
                    </div>

                    <div className="graph-user-info">
                      <h5>Connected Microsoft Graph User:</h5>
                      <div className="graph-user-details">
                        <div className="detail-item">
                          <span className="detail-lbl">Name:</span>
                          <span className="detail-val">{testResult.user?.displayName}</span>
                        </div>
                        <div className="detail-item">
                          <span className="detail-lbl">Mailbox UPN:</span>
                          <span className="detail-val">{testResult.user?.mail}</span>
                        </div>
                      </div>
                    </div>

                    <div className="emails-list-container">
                      <h5>First 5 Message Subjects:</h5>
                      {testResult.emails && testResult.emails.length > 0 ? (
                        <ul className="emails-list">
                          {testResult.emails.map((subject, index) => (
                            <li key={index} className="email-item">
                              <span className="email-index">{index + 1}</span>
                              <span className="email-subject">{subject}</span>
                            </li>
                          ))}
                        </ul>
                      ) : (
                        <p className="no-emails">No messages found in mailbox.</p>
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
