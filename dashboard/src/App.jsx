import { useEffect, useState, useRef } from 'react';
import { createClient } from '@supabase/supabase-js';
import './App.css';

// Retrieve config from environment variables with dynamic fallback
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

const getApiUrl = () => {
  if (import.meta.env.VITE_API_URL) {
    return import.meta.env.VITE_API_URL;
  }
  // Auto-detect production vs development environment
  if (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1') {
    return 'http://localhost:8002';
  }
  return 'https://providentemail.onrender.com';
};
const apiUrl = getApiUrl();

const isConfigured =
  supabaseUrl &&
  supabaseAnonKey &&
  supabaseAnonKey !== 'your-supabase-anon-key-placeholder' &&
  !supabaseAnonKey.includes('placeholder');

const supabase = isConfigured ? createClient(supabaseUrl, supabaseAnonKey) : null;

// Categories definition matching backend + add-in icons/labels
const CATEGORIES = {
  purchase_order: { name: 'Purchase Orders', hint: 'New POs ready for ERP entry', icon: '📦', colorClass: 'po' },
  enquiry: { name: 'Quote Enquiries', hint: 'RFQ details, pricing & samples', icon: '❓', colorClass: 'enq' },
  invoice: { name: 'Vendor Invoices', hint: 'Verify & route to accounts payable', icon: '💰', colorClass: 'inv' },
  shipping: { name: 'Shipping & Logistics', hint: 'Track dispatch, ETA & delays', icon: '🚚', colorClass: 'shp' },
  general: { name: 'General Operations', hint: 'Internal correspondence & misc', icon: '📧', colorClass: 'gen' },
  junk: { name: 'Junk Email', hint: 'Spam, newsletters & marketing', icon: '🗑️', colorClass: 'jnk' },
};

function App() {
  const [session, setSession] = useState(null);
  const [authError, setAuthError] = useState(null);

  // App States
  const [triageSummary, setTriageSummary] = useState({ total: 0, counts: {} });
  const [staleAlerts, setStaleAlerts] = useState({ stale_count: 0, stale: [] });
  const [selectedCategory, setSelectedCategory] = useState('purchase_order');
  const [emails, setEmails] = useState([]);
  const [selectedEmail, setSelectedEmail] = useState(null);
  const [selectedEmailDetail, setSelectedEmailDetail] = useState(null);
  const [statusFlows, setStatusFlows] = useState({});
  const [nextStatusMap, setNextStatusMap] = useState({});

  // Loading States
  const [loadingTriage, setLoadingTriage] = useState(false);
  const [loadingEmails, setLoadingEmails] = useState(false);
  const [loadingDetail, setLoadingDetail] = useState(false);

  // Form / Action States
  const [autoReplyEnabled, setAutoReplyEnabled] = useState(false);
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [priorityFilter, setPriorityFilter] = useState('all');
  const [activeTab, setActiveTab] = useState('triage');

  // Calculator States
  const [calcType, setCalcType] = useState('3ply');
  const [calcQty, setCalcQty] = useState(1000);
  const [calcL, setCalcL] = useState(300);
  const [calcW, setCalcW] = useState(200);
  const [calcH, setCalcH] = useState(200);
  const [calcPrint, setCalcPrint] = useState('none');
  const [calcResult, setCalcResult] = useState('$0.00');

  // Draft Reply / Status Actions
  const [draftReplyText, setDraftReplyText] = useState('');
  const [sendingReply, setSendingReply] = useState(false);
  const [updatingStatus, setUpdatingStatus] = useState(false);

  // Feedback State
  const [feedbackSuccess, setFeedbackSuccess] = useState(null); // 'thumb_up' or 'thumb_down'

  // Mobile Navigation & Sync States
  const [activeMobilePane, setActiveMobilePane] = useState('sidebar');
  const [syncing, setSyncing] = useState(false);

  // Toast State
  const [toast, setToast] = useState(null);
  const toastTimeoutRef = useRef(null);

  // Sync Mailbox on demand using delegated M365 token
  const handleSyncMailbox = async () => {
    setSyncing(true);
    try {
      const res = await apiFetch('/api/addin/sync', { method: 'POST' });
      if (res.success) {
        showToast(
          `Sync complete! Processed ${res.processed_count} new emails (skipped ${res.skipped_count}, failed ${res.failed_count}).`,
          'success'
        );

        // Refresh triage summary counts
        const summary = await apiFetch('/api/addin/triage_summary');
        setTriageSummary(summary);

        // Refresh active emails list
        let path = `/api/addin/category_emails?category=${selectedCategory}&limit=50`;
        if (startDate) path += `&start_date=${new Date(startDate).toISOString()}`;
        if (endDate) path += `&end_date=${new Date(endDate).toISOString()}`;
        const data = await apiFetch(path);
        setEmails(data.emails || []);
      }
    } catch (err) {
      showToast(`Sync failed: ${err.message}`, 'error');
    } finally {
      setSyncing(false);
    }
  };

  // Monitor auth status on load
  useEffect(() => {
    if (!supabase) return;

    supabase.auth.getSession().then(({ data: { session }, error }) => {
      if (error) {
        setAuthError(error.message);
      } else {
        setSession(session);
        if (session?.provider_token) {
          localStorage.setItem('m365_provider_token', session.provider_token);
        }
      }
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session?.provider_token) {
        localStorage.setItem('m365_provider_token', session.provider_token);
      }
      if (!session) {
        // Reset states on logout
        localStorage.removeItem('m365_provider_token');
        setTriageSummary({ total: 0, counts: {} });
        setStaleAlerts({ stale_count: 0, stale: [] });
        setEmails([]);
        setSelectedEmail(null);
        setSelectedEmailDetail(null);
      }
    });

    return () => {
      subscription.unsubscribe();
    };
  }, []);

  // Show Toast helper
  const showToast = (message, type = 'success') => {
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    setToast({ message, type });
    toastTimeoutRef.current = setTimeout(() => {
      setToast(null);
    }, 3000);
  };

  // Authenticate helper for API calls
  const apiFetch = async (path, options = {}) => {
    const token = session?.provider_token || localStorage.getItem('m365_provider_token');
    if (!token) {
      throw new Error('Access token is missing. Please log in again.');
    }
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    };
    const response = await fetch(`${apiUrl}${path}`, { ...options, headers });
    if (!response.ok) {
      if (response.status === 401) {
        // Clear stored token on 401 unauthorized to force re-auth
        localStorage.removeItem('m365_provider_token');
      }
      const msg = await response.text();
      throw new Error(msg || `Request failed with status ${response.status}`);
    }
    return response.json();
  };

  // Fetch Triage Summary & Settings & Flows
  useEffect(() => {
    if (!session) return;

    const loadInitialAppData = async () => {
      setLoadingTriage(true);
      try {
        // Load Triage Counts
        const summary = await apiFetch('/api/addin/triage_summary');
        setTriageSummary(summary);

        // Load Stale Alerts
        const stale = await apiFetch('/api/addin/stale_alerts');
        setStaleAlerts(stale);

        // Load Settings (Auto-Reply)
        const settings = await apiFetch('/api/settings');
        setAutoReplyEnabled(settings.auto_reply);

        // Load Status Flows
        const flowsData = await apiFetch('/api/addin/status_flows');
        if (flowsData.success) {
          setStatusFlows(flowsData.flows || {});
          setNextStatusMap(flowsData.next_status || {});
        }
      } catch (err) {
        showToast(`Failed to load app data: ${err.message}`, 'error');
      } finally {
        setLoadingTriage(false);
      }
    };

    loadInitialAppData();
  }, [session]);

  // Fetch Emails in Category when category or date filter changes
  useEffect(() => {
    if (!session) return;

    const loadEmails = async () => {
      setLoadingEmails(true);
      try {
        let path = `/api/addin/category_emails?category=${selectedCategory}&limit=50`;
        if (startDate) {
          path += `&start_date=${new Date(startDate).toISOString()}`;
        }
        if (endDate) {
          path += `&end_date=${new Date(endDate).toISOString()}`;
        }
        const data = await apiFetch(path);
        setEmails(data.emails || []);
      } catch (err) {
        showToast(`Failed to load emails: ${err.message}`, 'error');
      } finally {
        setLoadingEmails(false);
      }
    };

    loadEmails();
  }, [session, selectedCategory, startDate, endDate]);

  // Fetch full details of selected email
  useEffect(() => {
    if (!session || !selectedEmail) {
      setSelectedEmailDetail(null);
      return;
    }

    const loadEmailDetail = async () => {
      setLoadingDetail(true);
      setFeedbackSuccess(null);
      try {
        const data = await apiFetch(`/api/addin/email_detail?message_id=${encodeURIComponent(selectedEmail.message_id)}`);
        setSelectedEmailDetail(data);
        setDraftReplyText(data.response_draft || '');
        // Default to pricing tab if quote enquiry
        if (data.category === 'enquiry') {
          setActiveTab('pricing');
        } else {
          setActiveTab('triage');
        }
      } catch (err) {
        showToast(`Failed to load email details: ${err.message}`, 'error');
      } finally {
        setLoadingDetail(false);
      }
    };

    loadEmailDetail();
  }, [session, selectedEmail]);

  // Handle Login
  const handleLogin = async () => {
    if (!supabase) return;
    setAuthError(null);
    const { error } = await supabase.auth.signInWithOAuth({
      provider: 'azure',
      options: {
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

  // Toggle Auto-Reply settings
  const handleToggleAutoReply = async (e) => {
    const checked = e.target.checked;
    setAutoReplyEnabled(checked);
    try {
      await apiFetch('/api/settings', {
        method: 'POST',
        body: JSON.stringify({ auto_reply: checked }),
      });
      showToast(checked ? 'Auto-Reply activated' : 'Auto-Reply deactivated', 'success');
    } catch (err) {
      setAutoReplyEnabled(!checked); // Revert
      showToast('Failed to save settings', 'error');
    }
  };

  // Handle manual category reclassification
  const handleOverrideCategory = async (e) => {
    const newCategory = e.target.value;
    if (!selectedEmail) return;

    try {
      const res = await apiFetch('/api/addin/reclassify', {
        method: 'POST',
        body: JSON.stringify({ message_id: selectedEmail.message_id, category: newCategory }),
      });
      if (res.success) {
        showToast('Category updated successfully', 'success');

        // Update active email state locally
        const updatedEmail = { ...selectedEmail, category: newCategory };
        setSelectedEmail(updatedEmail);

        // Update items list
        setEmails((prev) =>
          prev.map((em) => (em.message_id === selectedEmail.message_id ? { ...em, category: newCategory } : em))
        );

        // Refresh triage summary counts
        const summary = await apiFetch('/api/addin/triage_summary');
        setTriageSummary(summary);
      }
    } catch (err) {
      showToast(`Reclassification failed: ${err.message}`, 'error');
    }
  };

  // Handle advances on status stepper
  const handleAdvanceStatus = async () => {
    if (!selectedEmailDetail) return;
    const currentStatus = selectedEmailDetail.email_status || 'po_received';
    const nextStatus = nextStatusMap[currentStatus];
    if (!nextStatus) return;

    setUpdatingStatus(true);
    try {
      const res = await apiFetch('/api/addin/update_status', {
        method: 'POST',
        body: JSON.stringify({ message_id: selectedEmailDetail.message_id, status: nextStatus }),
      });
      if (res.success) {
        // Reload details
        const data = await apiFetch(`/api/addin/email_detail?message_id=${encodeURIComponent(selectedEmailDetail.message_id)}`);
        setSelectedEmailDetail(data);

        // Update list status locally
        setEmails((prev) =>
          prev.map((em) => (em.message_id === selectedEmailDetail.message_id ? { ...em, email_status: nextStatus } : em))
        );

        showToast(`Moved to next stage: ${nextStatus}`, 'success');
      }
    } catch (err) {
      showToast(`Failed to update status: ${err.message}`, 'error');
    } finally {
      setUpdatingStatus(false);
    }
  };

  // Submit classification feedback
  const handleFeedback = async (isCorrect) => {
    if (!selectedEmail) return;
    setFeedbackSuccess(isCorrect ? 'thumb_up' : 'thumb_down');
    try {
      await apiFetch('/api/addin/feedback', {
        method: 'POST',
        body: JSON.stringify({
          message_id: selectedEmail.message_id,
          is_correct: isCorrect,
          note: isCorrect ? 'Looks good' : 'Incorrect category classification',
        }),
      });
      showToast('Feedback submitted', 'success');
    } catch (err) {
      showToast(`Failed to submit feedback: ${err.message}`, 'error');
    }
  };

  // Pricing Calculator live computation
  useEffect(() => {
    const base = { '3ply': 1.5e-7, '5ply': 2.5e-7, '7ply': 3.5e-7, 'duplex': 1.2e-7 }[calcType] || 1.5e-7;
    const pMul = { none: 1, '1color': 1.1, '2color': 1.2, 'multi': 1.4 }[calcPrint] || 1;
    let unit = Math.max(0.07, calcL * calcW * calcH * base * pMul);
    if (calcQty >= 5000) unit *= 0.8;
    else if (calcQty >= 2000) unit *= 0.9;
    const total = unit * calcQty;
    setCalcResult(`$${total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  }, [calcType, calcQty, calcL, calcW, calcH, calcPrint]);

  // Insert estimated price into smart reply draft text
  const handleInsertPrice = () => {
    const typeLabel = {
      '3ply': '3-Ply Corrugated Box',
      '5ply': '5-Ply Corrugated Box',
      '7ply': '7-Ply Corrugated Box',
      duplex: 'Duplex Printed Carton',
    }[calcType];
    const dims = `${calcL}×${calcW}×${calcH}mm`;
    const insertText = `\n\nPricing estimate based on your specifications:\n• Product: ${typeLabel}\n• Dimensions: ${dims}\n• Quantity: ${calcQty} units\n• Estimated Total: ${calcResult} (excl. tax & freight)\n\nPlease confirm to proceed with a formal quotation.`;
    setDraftReplyText((prev) => prev.trim() + insertText);
    setActiveTab('reply');
    showToast('Pricing inserted into draft reply', 'success');
  };

  // Submit AI smart reply (Records draft and advances status in DB)
  const handleSendReply = async () => {
    if (!selectedEmailDetail) return;
    setSendingReply(true);
    try {
      const res = await apiFetch('/api/send_reply', {
        method: 'POST',
        body: JSON.stringify({ message_id: selectedEmailDetail.message_id, reply_text: draftReplyText }),
      });
      if (res.success) {
        showToast('Reply recorded & status advanced', 'success');

        // Reload details to show updated status and reply indicators
        const data = await apiFetch(`/api/addin/email_detail?message_id=${encodeURIComponent(selectedEmailDetail.message_id)}`);
        setSelectedEmailDetail(data);

        // Update locally in lists
        setEmails((prev) =>
          prev.map((em) =>
            em.message_id === selectedEmailDetail.message_id
              ? { ...em, reply_sent: true, email_status: res.new_status }
              : em
          )
        );

        // Refresh triage summary counts
        const summary = await apiFetch('/api/addin/triage_summary');
        setTriageSummary(summary);
      }
    } catch (err) {
      showToast(`Failed to send reply: ${err.message}`, 'error');
    } finally {
      setSendingReply(false);
    }
  };

  // Reset filter inputs
  const handleResetFilters = () => {
    setStartDate('');
    setEndDate('');
    setPriorityFilter('all');
  };

  // Filtered emails in frontend
  const filteredEmails = emails.filter((em) => {
    if (priorityFilter === 'all') return true;
    return em.priority_tier?.toLowerCase() === priorityFilter.toLowerCase();
  });

  // Initials generator
  const getInitials = (name) => {
    if (!name) return 'U';
    return name
      .split(' ')
      .map((n) => n[0])
      .slice(0, 2)
      .join('')
      .toUpperCase();
  };

  // Date formatter helper
  const formatDate = (dateStr) => {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return (
      d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) +
      ' ' +
      d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })
    );
  };

  return (
    <div className="app-shell">
      {/* Toast Notification Portal */}
      {toast && (
        <div className="toast-container">
          <div className={`toast ${toast.type} animate-fade-in`}>{toast.message}</div>
        </div>
      )}

      {/* Main Login Screen */}
      {!session && (
        <div className="login-screen">
          <div className="background-glow" />
          <div className="card glass login-card animate-slide-up">
            <img src="/logo.jpg" alt="Provident Logo" className="login-logo" />
            <h1 className="login-title">Provident Copilot</h1>
            <p className="card-description">
              Sign in with your Microsoft 365 account to load your mailbox and manage customer orders, enquiries,
              and logistics.
            </p>

            {authError && <div className="error-banner">⚠️ Authentication Error: {authError}</div>}
            {!isConfigured && (
              <div className="warning-banner">
                ⚠️ Supabase credentials not fully configured. Please check your <code>.env</code> file.
              </div>
            )}

            <button className="btn btn-microsoft" onClick={handleLogin} disabled={!isConfigured}>
              <svg className="microsoft-logo" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
                <rect x="0" y="0" width="11" height="11" fill="#f25022" />
                <rect x="12" y="0" width="11" height="11" fill="#7fba00" />
                <rect x="0" y="12" width="11" height="11" fill="#00a4ef" />
                <rect x="12" y="12" width="11" height="11" fill="#ffb900" />
              </svg>
              <span>Sign in with Microsoft</span>
            </button>
          </div>
        </div>
      )}

      {/* Authenticated Workspace */}
      {session && (
        <div className={`workspace-container mobile-active-${activeMobilePane}`}>
          {/* ══════════════════════════════════════════════════
               LEFT PANE: SIDEBAR NAVIGATION
          ══════════════════════════════════════════════════ */}
          <aside className={`sidebar ${activeMobilePane === 'sidebar' ? 'mobile-show' : 'mobile-hide'}`}>
            <div className="sidebar-header">
              <div className="logo-area">
                <img src="/logo.jpg" alt="Provident Logo" className="app-logo" />
                <span className="app-title">Provident Copilot</span>
              </div>
              <div className="auto-reply-wrapper" title="Global Auto-Reply Mode">
                <label className="toggle-label">
                  <input
                    type="checkbox"
                    checked={autoReplyEnabled}
                    onChange={handleToggleAutoReply}
                    className="toggle-checkbox"
                  />
                  <span className="toggle-text">Auto-Reply</span>
                </label>
              </div>
            </div>

            <div className="sidebar-scroll">
              {/* Category selector grid */}
              <div className="cat-section-header">MAILBOX CATEGORIES</div>
              <nav className="cat-list">
                {Object.keys(CATEGORIES).map((key) => {
                  const cat = CATEGORIES[key];
                  const count = triageSummary.counts[key] || 0;
                  const isActive = selectedCategory === key;
                  return (
                    <button
                      key={key}
                      onClick={() => {
                        setSelectedCategory(key);
                        setSelectedEmail(null);
                        setActiveMobilePane('list');
                      }}
                      className={`cat-item ${isActive ? 'active' : ''}`}
                    >
                      <span className={`cat-icon-badge ${cat.colorClass}`}>{cat.icon}</span>
                      <div className="cat-item-body">
                        <span className="cat-item-name">{cat.name}</span>
                        <span className="cat-item-hint">{cat.hint}</span>
                      </div>
                      {count > 0 && <span className="cat-item-count">{count}</span>}
                    </button>
                  );
                })}
              </nav>

              {/* Stale email alerts section */}
              {staleAlerts.stale_count > 0 && (
                <div
                  className="stale-alert-box animate-pulse"
                  onClick={() => {
                    if (staleAlerts.stale[0]?.category) {
                      setSelectedCategory(staleAlerts.stale[0].category);
                    }
                    setActiveMobilePane('list');
                  }}
                  style={{ cursor: 'pointer' }}
                >
                  <div className="stale-alert-icon">⚠️</div>
                  <div className="stale-alert-body">
                    <strong>{staleAlerts.stale_count} email(s) need attention</strong>
                    <span>No action taken in 24 hours+</span>
                  </div>
                </div>
              )}
            </div>

            <div className="sidebar-footer">
              <div className="avatar">
                {getInitials(session.user?.user_metadata?.full_name || session.user?.email)}
              </div>
              <div className="user-info-area">
                <span className="user-name">{session.user?.user_metadata?.full_name || 'M365 User'}</span>
                <span className="user-email-text">{session.user?.email}</span>
              </div>
              <button className="btn-logout" onClick={handleLogout} title="Sign Out">
                ✕
              </button>
            </div>
          </aside>

          {/* ══════════════════════════════════════════════════
               MIDDLE PANE: EMAIL LIST
          ══════════════════════════════════════════════════ */}
          <section className={`inbox-pane ${activeMobilePane === 'list' ? 'mobile-show' : 'mobile-hide'}`}>
            <div className="inbox-header">
              <div className="mobile-nav-row">
                <button className="mobile-back-btn" onClick={() => setActiveMobilePane('sidebar')}>
                  ← Categories
                </button>
              </div>

              <div className="inbox-title-row">
                <div className="inbox-title-left">
                  <h2>{CATEGORIES[selectedCategory]?.name || 'Inbox'}</h2>
                  <span className="inbox-total-count">{filteredEmails.length} messages</span>
                </div>
                <button
                  className="btn-sync-inbox"
                  onClick={handleSyncMailbox}
                  disabled={syncing}
                  title="Fetch and classify new emails from M365 Mailbox"
                >
                  {syncing ? '🔄 Syncing...' : '🔄 Sync Inbox'}
                </button>
              </div>

              {/* Advanced date and range filtering */}
              <div className="filters-card">
                <div className="date-filter-group">
                  <div className="input-field">
                    <label>Start Date</label>
                    <input type="datetime-local" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
                  </div>
                  <div className="input-field">
                    <label>End Date</label>
                    <input type="datetime-local" value={endDate} onChange={(e) => setEndDate(e.target.value)} />
                  </div>
                  <button className="btn-reset-filters" onClick={handleResetFilters}>
                    Reset
                  </button>
                </div>

                <div className="priority-chips-row">
                  {['all', 'critical', 'high', 'medium', 'low'].map((tier) => {
                    const isActive = priorityFilter === tier;
                    return (
                      <button
                        key={tier}
                        onClick={() => setPriorityFilter(tier)}
                        className={`priority-chip ${tier} ${isActive ? 'active' : ''}`}
                      >
                        {tier === 'all' ? 'All' : tier.toUpperCase()}
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Email Rows List */}
            <div className="email-rows-list">
              {loadingEmails && (
                <div className="shimmer-list">
                  <div className="shimmer-row" />
                  <div className="shimmer-row" />
                  <div className="shimmer-row" />
                </div>
              )}

              {!loadingEmails && filteredEmails.length === 0 && (
                <div className="empty-list-placeholder">
                  <span className="empty-icon">📭</span>
                  <p>No emails found in this category.</p>
                </div>
              )}

              {!loadingEmails &&
                filteredEmails.map((em) => {
                  const isSelected = selectedEmail?.message_id === em.message_id;
                  const hasPriority = em.priority_tier && em.priority_tier !== 'low';
                  return (
                    <button
                      key={em.message_id}
                      onClick={() => {
                        setSelectedEmail(em);
                        setActiveMobilePane('detail');
                      }}
                      className={`email-row-card ${isSelected ? 'selected' : ''} ${em.reply_sent ? 'replied' : ''}`}
                    >
                      <div className="email-row-top">
                        <span className="email-row-sender" title={em.sender}>
                          {em.sender?.split('<')[0]?.trim() || em.sender}
                        </span>
                        <span className="email-row-date">{formatDate(em.received_at)}</span>
                      </div>
                      <h4 className="email-row-subject">{em.subject || '(No Subject)'}</h4>
                      <p className="email-row-snippet">{em.body_preview || 'No preview available.'}</p>
                      <div className="email-row-bottom-meta">
                        {em.reply_sent && <span className="replied-badge">✓ Replied</span>}
                        {em.priority_tier && (
                          <span className={`priority-badge-dot ${em.priority_tier?.toLowerCase()}`}>
                            {em.priority_tier}
                          </span>
                        )}
                        {em.email_status && (
                          <span className="stage-badge-pill">
                            {em.email_status.replace(/_/g, ' ').toUpperCase()}
                          </span>
                        )}
                      </div>
                    </button>
                  );
                })}
            </div>
          </section>

          {/* ══════════════════════════════════════════════════
               RIGHT PANE: DETAIL COPILOT WORKSPACE
          ══════════════════════════════════════════════════ */}
          <section className={`detail-pane ${activeMobilePane === 'detail' ? 'mobile-show' : 'mobile-hide'}`}>
            <div className="detail-mobile-header">
              <button className="mobile-back-btn" onClick={() => setActiveMobilePane('list')}>
                ← Inbox
              </button>
            </div>

            {/* Empty view state */}
            {!selectedEmail && (
              <div className="empty-detail-state">
                <div className="empty-mailbox-art">📧</div>
                <h3>Select an email to view details</h3>
                <p>Choose an item from the inbox pane to audit its classification, summary, and action stages.</p>
              </div>
            )}

            {/* Email details loading spinner */}
            {selectedEmail && loadingDetail && (
              <div className="empty-detail-state">
                <div className="spinner" />
                <p>Loading copilot workspace...</p>
              </div>
            )}

            {/* Loaded details content workspace */}
            {selectedEmail && !loadingDetail && selectedEmailDetail && (
              <div className="detail-workspace-content">
                {/* Classification Banner Alert Box */}
                <div className={`classification-banner-alert ${selectedEmailDetail.category}`}>
                  <div className="banner-icon-area">
                    {CATEGORIES[selectedEmailDetail.category]?.icon || '📧'}
                  </div>
                  <div className="banner-text-info">
                    <h4>{CATEGORIES[selectedEmailDetail.category]?.name || 'Unclassified'}</h4>
                    <p>{selectedEmailDetail.reason || 'Automatically parsed email message.'}</p>
                  </div>
                  <div className="banner-feedback-actions">
                    <button
                      className={`btn-feedback-thumb ${feedbackSuccess === 'thumb_up' ? 'active' : ''}`}
                      onClick={() => handleFeedback(true)}
                      title="Correct Category"
                    >
                      👍
                    </button>
                    <button
                      className={`btn-feedback-thumb ${feedbackSuccess === 'thumb_down' ? 'active' : ''}`}
                      onClick={() => handleFeedback(false)}
                      title="Wrong Category"
                    >
                      👎
                    </button>
                  </div>
                </div>

                {/* Main panel info section */}
                <div className="detail-panels-body">
                  {/* Top quick-meta bar */}
                  <div className="quick-meta-bar">
                    <div className="quick-meta-item">
                      <span className="lbl">From:</span>
                      <strong className="val">{selectedEmailDetail.sender}</strong>
                    </div>
                    <div className="quick-meta-item">
                      <span className="lbl">Received:</span>
                      <strong className="val">{formatDate(selectedEmailDetail.received_at)}</strong>
                    </div>
                    {selectedEmailDetail.priority_tier && (
                      <div className="quick-meta-item">
                        <span className="lbl">Priority:</span>
                        <strong className={`val priority-text-color ${selectedEmailDetail.priority_tier.toLowerCase()}`}>
                          {selectedEmailDetail.priority_tier} (Score: {selectedEmailDetail.priority_score})
                        </strong>
                      </div>
                    )}
                  </div>

                  {/* Manual Category Override dropdown */}
                  <div className="manual-override-card">
                    <label>Manual Category Override:</label>
                    <select
                      value={selectedEmailDetail.category}
                      onChange={handleOverrideCategory}
                      className="override-select-dropdown"
                    >
                      <option value="purchase_order">📦 Purchase Order</option>
                      <option value="enquiry">❓ Quote Enquiry</option>
                      <option value="invoice">💰 Vendor Invoice</option>
                      <option value="shipping">🚚 Shipping & Logistics</option>
                      <option value="general">📧 General Operations</option>
                      <option value="junk">🗑️ Junk Email</option>
                    </select>
                  </div>

                  {/* Thread Lifecycle Tracker */}
                  {selectedEmailDetail.thread_lifecycle && (
                    <div className="thread-lifecycle-section">
                      <h4 className="section-small-title">🧵 THREAD LIFECYCLE TRACKER</h4>
                      <div className="thread-stepper">
                        {[
                          { key: 'enquiry', label: 'Enquiry', icon: '❓', check: selectedEmailDetail.thread_lifecycle.has_enquiry },
                          { key: 'purchase_order', label: 'Order', icon: '📦', check: selectedEmailDetail.thread_lifecycle.has_order },
                          { key: 'invoice', label: 'Invoice', icon: '💰', check: selectedEmailDetail.thread_lifecycle.has_invoice },
                          { key: 'shipping', label: 'Shipment', icon: '🚚', check: selectedEmailDetail.thread_lifecycle.has_shipping },
                        ].map((stage) => {
                          const isDone = stage.check;
                          const isActive = selectedEmailDetail.thread_lifecycle.current_stage === stage.key;
                          return (
                            <div key={stage.key} className={`thread-node ${isDone ? 'done' : ''} ${isActive ? 'active' : ''}`}>
                              <div className="node-bubble">{isDone ? '✓' : stage.icon}</div>
                              <span className="node-label">{stage.label}</span>
                            </div>
                          );
                        })}
                      </div>

                      {/* Milestone Chronicles */}
                      <div className="milestone-chronicles-list">
                        {selectedEmailDetail.thread_lifecycle.milestones?.map((milestone, idx) => (
                          <div key={idx} className="chronicle-row">
                            <span className="chronicle-indicator">🟢 {milestone.label}</span>
                            <span className="chronicle-time">{formatDate(milestone.timestamp)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Status Progression Stepper */}
                  {statusFlows[selectedEmailDetail.category] && (
                    <div className="status-progression-section">
                      <div className="stepper-label-row">
                        <span className="lbl">Status Stage:</span>
                        <strong className="val">
                          {statusFlows[selectedEmailDetail.category]
                            .find((s) => s.id === selectedEmailDetail.email_status)
                            ?.label || selectedEmailDetail.email_status}
                        </strong>
                      </div>

                      <div className="status-progress-bar">
                        {statusFlows[selectedEmailDetail.category].map((step, idx) => {
                          const flow = statusFlows[selectedEmailDetail.category];
                          const currIdx = flow.findIndex((s) => s.id === selectedEmailDetail.email_status);
                          let stateClass = '';
                          if (idx < currIdx) stateClass = 'done';
                          else if (idx === currIdx) stateClass = 'active';
                          return (
                            <div key={step.id} className={`status-step-node ${stateClass}`}>
                              <div className="step-circle">{idx < currIdx ? '✓' : '●'}</div>
                              <span className="step-label" title={step.label}>
                                {step.label}
                              </span>
                            </div>
                          );
                        })}
                      </div>

                      {/* Advance Stage button */}
                      {nextStatusMap[selectedEmailDetail.email_status] && (
                        <button
                          className="btn-advance-stage"
                          onClick={handleAdvanceStatus}
                          disabled={updatingStatus}
                        >
                          {updatingStatus ? (
                            <span className="spinner" />
                          ) : (
                            `Move to "${
                              statusFlows[selectedEmailDetail.category]?.find(
                                (s) => s.id === nextStatusMap[selectedEmailDetail.email_status]
                              )?.label || 'Next'
                            }" →`
                          )}
                        </button>
                      )}
                    </div>
                  )}

                  {/* Tabs workspace layout */}
                  <div className="tabs-header-bar">
                    <button
                      className={`tab-workspace-pill ${activeTab === 'triage' ? 'active' : ''}`}
                      onClick={() => setActiveTab('triage')}
                    >
                      Email Summary
                    </button>
                    {selectedEmailDetail.category === 'enquiry' && (
                      <button
                        className={`tab-workspace-pill ${activeTab === 'pricing' ? 'active' : ''}`}
                        onClick={() => setActiveTab('pricing')}
                      >
                        Pricing Calculator
                      </button>
                    )}
                    <button
                      className={`tab-workspace-pill ${activeTab === 'reply' ? 'active' : ''}`}
                      onClick={() => setActiveTab('reply')}
                    >
                      Smart Reply
                    </button>
                  </div>

                  {/* Tab Workspace Panels */}
                  <div className="tab-workspace-panel-body">
                    {/* TAB 1: Summary */}
                    {activeTab === 'triage' && (
                      <div className="tab-panel-container">
                        <div className="summary-block">
                          <h5>AI Email Summary</h5>
                          <p className="summary-body-text">{selectedEmailDetail.reason || 'No summary available.'}</p>
                        </div>

                        <div className="extracted-fields-block">
                          <h5>Extracted Details</h5>
                          <div className="fields-keyval-grid">
                            {(() => {
                              let meta = {};
                              try {
                                meta =
                                  typeof selectedEmailDetail.extracted_data === 'string'
                                    ? JSON.parse(selectedEmailDetail.extracted_data)
                                    : selectedEmailDetail.extracted_data || {};
                              } catch {}
                              const keys = Object.keys(meta);
                              if (keys.length === 0) {
                                return <p className="no-fields-text">No fields extracted for this email.</p>;
                              }
                              return keys.map((key) => {
                                const val = String(meta[key] || 'N/A');
                                return (
                                  <div key={key} className="field-grid-row">
                                    <div className="field-kv-box">
                                      <span className="field-k">{key}</span>
                                      <span className="field-v">{val}</span>
                                    </div>
                                    <button
                                      className="btn-copy-field"
                                      onClick={() => {
                                        navigator.clipboard.writeText(val);
                                        showToast('Copied to clipboard');
                                      }}
                                    >
                                      Copy
                                    </button>
                                  </div>
                                );
                              });
                            })()}
                          </div>
                        </div>
                      </div>
                    )}

                    {/* TAB 2: Calculator */}
                    {activeTab === 'pricing' && selectedEmailDetail.category === 'enquiry' && (
                      <div className="tab-panel-container">
                        <div className="calculator-card">
                          <h4 className="calc-card-title">Estimated Pricing Tool</h4>
                          <div className="calc-inputs-grid">
                            <div className="input-field">
                              <label>Item Type</label>
                              <select value={calcType} onChange={(e) => setCalcType(e.target.value)}>
                                <option value="3ply">3-Ply Corrugated Box</option>
                                <option value="5ply">5-Ply Corrugated Box</option>
                                <option value="7ply">7-Ply Corrugated Box</option>
                                <option value="duplex">Duplex Printed Carton</option>
                              </select>
                            </div>
                            <div className="input-field">
                              <label>Quantity</label>
                              <input
                                type="number"
                                value={calcQty}
                                min="100"
                                onChange={(e) => setCalcQty(parseInt(e.target.value) || 1000)}
                              />
                            </div>
                            <div className="input-field">
                              <label>Length (mm)</label>
                              <input
                                type="number"
                                value={calcL}
                                min="50"
                                onChange={(e) => setCalcL(parseFloat(e.target.value) || 0)}
                              />
                            </div>
                            <div className="input-field">
                              <label>Width (mm)</label>
                              <input
                                type="number"
                                value={calcW}
                                min="50"
                                onChange={(e) => setCalcW(parseFloat(e.target.value) || 0)}
                              />
                            </div>
                            <div className="input-field">
                              <label>Height (mm)</label>
                              <input
                                type="number"
                                value={calcH}
                                min="50"
                                onChange={(e) => setCalcH(parseFloat(e.target.value) || 0)}
                              />
                            </div>
                            <div className="input-field">
                              <label>Printing</label>
                              <select value={calcPrint} onChange={(e) => setCalcPrint(e.target.value)}>
                                <option value="none">No Print</option>
                                <option value="1color">1-Color Flexo</option>
                                <option value="2color">2-Color Flexo</option>
                                <option value="multi">Offset Multicolor</option>
                              </select>
                            </div>
                          </div>

                          <div className="calc-result-box">
                            <span className="calc-lbl">Estimated Price:</span>
                            <span className="calc-val">{calcResult}</span>
                          </div>

                          <button className="btn-insert-calc-price" onClick={handleInsertPrice}>
                            📋 Insert Price into Draft Reply
                          </button>
                        </div>
                      </div>
                    )}

                    {/* TAB 3: Smart Reply */}
                    {activeTab === 'reply' && (
                      <div className="tab-panel-container">
                        <div className="reply-block">
                          <h5>AI Draft Reply</h5>
                          <textarea
                            className="reply-textarea-draft"
                            value={draftReplyText}
                            onChange={(e) => setDraftReplyText(e.target.value)}
                            spellCheck="false"
                            placeholder="AI generating draft reply..."
                          />

                          <div className="reply-actions-row">
                            <button
                              className="btn-reply-action copy"
                              onClick={() => {
                                navigator.clipboard.writeText(draftReplyText);
                                showToast('Copied draft reply');
                              }}
                            >
                              📋 Copy Draft
                            </button>
                            <button
                              className="btn-reply-action send"
                              onClick={handleSendReply}
                              disabled={sendingReply}
                            >
                              {sendingReply ? <span className="spinner" /> : '✉ Send Reply'}
                            </button>
                          </div>

                          {selectedEmailDetail.reply_sent && (
                            <div className="recorded-reply-box">
                              <h5>Last Sent/Recorded Reply:</h5>
                              <pre className="sent-reply-pre">{selectedEmailDetail.sent_reply}</pre>
                              {selectedEmailDetail.reply_sent_at && (
                                <span className="sent-time">
                                  Recorded at {formatDate(selectedEmailDetail.reply_sent_at)}
                                </span>
                              )}
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>
            )}
          </section>
        </div>
      )}
    </div>
  );
}

export default App;
