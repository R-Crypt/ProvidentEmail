// ─────────────────────────────────────────────────────────────────────────────
// PROVIDENT OPERATIONS COPILOT — Task Pane JS v2.2
// UX-first rewrite: toasts, breadcrumbs, smooth nav, clean state machine
// ─────────────────────────────────────────────────────────────────────────────

const API_HOST = window.PROVIDENT_API_HOST || window.location.origin;

// ─── Auth ─────────────────────────────────────────────────────────────────────
let _authToken = null;

async function getAuthToken() {
    if (_authToken) return _authToken;
    return new Promise((resolve) => {
        if (!Office?.context || !Office?.auth || typeof Office.auth.getAccessTokenAsync !== 'function') {
            resolve(null); return;
        }
        try {
            Office.auth.getAccessTokenAsync({ allowSignInPrompt: true }, (result) => {
                if (result.status === "succeeded" || (Office?.MailboxEnums?.AsyncResultStatus && result.status === Office.MailboxEnums.AsyncResultStatus.Succeeded)) {
                    _authToken = result.value;
                    resolve(_authToken);
                } else {
                    resolve(null);
                }
            });
        } catch (e) { resolve(null); }
    });
}

async function apiFetch(url, options = {}, retries = 3) {
    const token = await getAuthToken();
    const headers = { 'Content-Type': 'application/json', ...(options.headers || {}) };
    if (token) headers['Authorization'] = `Bearer ${token}`;
    for (let i = 1; i <= retries; i++) {
        try {
            const r = await fetch(url, { ...options, headers });
            if (r.status === 401 && i < retries) { _authToken = null; await sleep(500 * i); continue; }
            return r;
        } catch (e) {
            if (i === retries) throw e;
            await sleep(1000 * i);
        }
    }
}
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

// ─── Category / Status Maps ────────────────────────────────────────────────────
const CAT = {
    purchase_order: { name: 'Purchase Order',       icon: '📦', cls: 'purchase_order' },
    enquiry:        { name: 'Enquiry / RFQ',        icon: '❓', cls: 'enquiry' },
    invoice:        { name: 'Vendor Invoice',       icon: '💰', cls: 'invoice' },
    shipping:       { name: 'Shipping & Logistics', icon: '🚚', cls: 'shipping' },
    general:        { name: 'General Operations',   icon: '📧', cls: 'general' },
    junk:           { name: 'Junk Email',           icon: '🗑️', cls: 'junk' },
};
const PRIORITY_LABEL = { critical: '🔴 Critical', high: '🟠 High', medium: '🟡 Medium', low: '🟢 Low' };

let STATUS_FLOWS = {};
let NEXT_STATUS  = {};

// ─── App State ─────────────────────────────────────────────────────────────────
let _currentView     = 'home';    // 'home' | 'list' | 'detail'
let _currentCat      = null;
let _allEmails       = [];        // cached full email list for active category
let _activeFilter    = 'all';
let _currentEmail    = null;      // email data object in detail view
let _currentMsgId    = null;      // Outlook item ID

// ─── Office Init ───────────────────────────────────────────────────────────────
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Outlook) {
        await boot();
        if (Office.context.mailbox.item) loadCurrentOutlookEmail();
        else navTo('home');
        try {
            Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
                if (Office.context.mailbox.item) loadCurrentOutlookEmail();
                else navTo('home');
            });
        } catch {}
    } else {
        await boot();
        runMockDemo();
    }
    setupListeners();
});

async function boot() {
    updateGreeting();
    await loadStatusFlows();
    await refreshHome();
    await loadAddInAutoReplySetting();
}

function updateGreeting() {
    const greetingEl = el('greeting-title');
    if (!greetingEl) return;
    const hour = new Date().getHours();
    let greeting = "Good morning ☀️";
    if (hour >= 12 && hour < 17) {
        greeting = "Good afternoon 🌤️";
    } else if (hour >= 17 || hour < 4) {
        greeting = "Good evening 🌙";
    }
    greetingEl.textContent = greeting;
}

// ─── Navigation ────────────────────────────────────────────────────────────────
function navTo(view, title) {
    ['view-home','view-list','view-detail'].forEach(id => {
        const el = document.getElementById(id);
        el.classList.remove('active');
    });
    document.getElementById(`view-${view}`).classList.add('active');
    _currentView = view;
    updateBreadcrumb(view, title);
}

function updateBreadcrumb(view, label) {
    const bc = document.getElementById('breadcrumb');
    if (view === 'home') {
        bc.innerHTML = `<span class="crumb active">Home</span>`;
    } else if (view === 'list') {
        bc.innerHTML = `
            <span class="crumb crumb-link" onclick="navTo('home')">Home</span>
            <span class="sep">›</span>
            <span class="crumb active">${esc(label || 'Emails')}</span>`;
    } else if (view === 'detail') {
        if (_currentCat) {
            bc.innerHTML = `
                <span class="crumb crumb-link" onclick="navTo('home')">Home</span>
                <span class="sep">›</span>
                <span class="crumb crumb-link" onclick="navTo('list','${esc(CAT[_currentCat]?.name || _currentCat)}'); loadCategoryEmails(_currentCat);">${esc(CAT[_currentCat]?.name || _currentCat)}</span>
                <span class="sep">›</span>
                <span class="crumb active">Detail</span>`;
        } else {
            bc.innerHTML = `
                <span class="crumb crumb-link" onclick="navTo('home')">Home</span>
                <span class="sep">›</span>
                <span class="crumb active">Email Detail</span>`;
        }
    }
}

// ─── Home View ─────────────────────────────────────────────────────────────────
async function refreshHome() {
    try {
        const [summaryRes, staleRes] = await Promise.all([
            apiFetch(`${API_HOST}/api/addin/triage_summary`),
            apiFetch(`${API_HOST}/api/addin/stale_alerts`)
        ]);

        if (summaryRes?.ok) {
            const d = await summaryRes.json();
            const c = d.counts || {};
            const cc = d.critical_counts || {};
            let totalCrit = 0;
            Object.keys(CAT).forEach(cat => {
                el('cnt-' + cat).textContent = c[cat] || 0;
                const n = cc[cat] || 0;
                totalCrit += n;
                const pill = el('crit-' + cat);
                if (n > 0) { pill.textContent = `${n} critical`; pill.classList.add('visible'); }
                else { pill.classList.remove('visible'); }
            });
            el('total-today').textContent    = d.total || 0;
            el('total-critical').textContent = totalCrit;
        }

        if (staleRes?.ok) {
            const d = await staleRes.json();
            const box = el('alert-stale');
            if (d.stale_count > 0) {
                el('stale-num').textContent = d.stale_count;
                box.classList.remove('hidden');
            } else {
                box.classList.add('hidden');
            }
        }
    } catch (e) { console.warn('refreshHome failed:', e); }
}

// ─── Category List View ────────────────────────────────────────────────────────
function navToCategory(cat) {
    _currentCat  = cat;
    _activeFilter = 'all';
    _allEmails   = [];
    setChip('all');

    const catInfo = CAT[cat];
    el('list-title').textContent = catInfo?.name || cat;
    el('list-count').textContent = '';
    el('email-list').innerHTML   = '';
    el('list-loading').classList.remove('hidden');

    navTo('list', catInfo?.name || cat);
    loadCategoryEmails(cat);
}

async function loadCategoryEmails(cat) {
    try {
        const startVal = el('filter-start').value;
        const endVal = el('filter-end').value;
        
        let url = `${API_HOST}/api/addin/category_emails?category=${cat}`;
        if (startVal) {
            const startIso = new Date(startVal).toISOString();
            url += `&start_date=${encodeURIComponent(startIso)}`;
        }
        if (endVal) {
            const endIso = new Date(endVal).toISOString();
            url += `&end_date=${encodeURIComponent(endIso)}`;
        }

        const res = await apiFetch(url);
        el('list-loading').classList.add('hidden');
        if (!res?.ok) { renderListEmpty('Failed to load emails.'); return; }
        const d = await res.json();
        _allEmails = d.emails || [];
        renderEmailList(_allEmails);
    } catch (e) {
        el('list-loading').classList.add('hidden');
        renderListEmpty('Cannot reach server. Is it running?');
    }
}

function onDateFilterChange() {
    if (_currentCat) {
        el('list-loading').classList.remove('hidden');
        el('email-list').innerHTML = '';
        loadCategoryEmails(_currentCat);
    }
}

function clearDateFilters() {
    el('filter-start').value = '';
    el('filter-end').value = '';
    onDateFilterChange();
}

function renderEmailListEmpty(msg) {
    el('email-list').innerHTML = `<div class="empty-box"><p class="empty-msg">${esc(msg)}</p></div>`;
}

function applyFilter(tier) {
    _activeFilter = tier;
    setChip(tier);
    const filtered = tier === 'all' ? _allEmails : _allEmails.filter(e => e.priority_tier === tier);
    renderEmailList(filtered);
}

function setChip(tier) {
    document.querySelectorAll('.f-chip').forEach(c => c.classList.remove('on'));
    const chip = el('fc-' + tier);
    if (chip) chip.classList.add('on');
}

function renderEmailList(emails) {
    const container = el('email-list');
    container.innerHTML = '';

    if (!emails.length) {
        renderListEmpty('No emails match this filter.');
        return;
    }

    el('list-count').textContent = `${emails.length} email${emails.length !== 1 ? 's' : ''}`;

    emails.forEach(email => {
        const tier = email.priority_tier || 'none';
        const row  = document.createElement('div');
        row.className = 'email-row';
        row.onclick = () => openEmailDetail(email);

        const timeAgo   = fmtTimeAgo(email.received_at || email.processed_at);
        const flow      = STATUS_FLOWS[email.category] || [];
        const step      = flow.find(s => s.id === email.email_status);
        const statusLbl = step?.label || email.email_status || '—';
        const valueStr  = email.estimated_value > 0 ? `$${Math.round(email.estimated_value).toLocaleString()}` : null;
        const isJunk    = (email.source_folder || '').toLowerCase().includes('junk');

        row.innerHTML = `
            <div class="email-stripe ${tier}"></div>
            <div class="email-body">
                <div class="email-top">
                    <span class="email-sender">${esc(email.sender || 'Unknown Sender')}</span>
                    <span class="email-time">${timeAgo}</span>
                </div>
                <div class="email-subject">${esc(email.subject || 'No Subject')}</div>
                <div class="email-tags">
                    ${tier !== 'none' ? `<span class="tag priority-${tier}">${PRIORITY_LABEL[tier] || tier}</span>` : ''}
                    <span class="tag status">${esc(statusLbl)}</span>
                    ${valueStr ? `<span class="tag value">${valueStr}</span>` : ''}
                    ${isJunk   ? `<span class="tag junk">Junk</span>` : ''}
                </div>
            </div>`;

        container.appendChild(row);
    });
}

function renderListEmpty(msg) {
    el('email-list').innerHTML = `
        <div class="empty-box">
            <div class="empty-icon">📭</div>
            <p class="empty-msg">${esc(msg)}</p>
        </div>`;
    el('list-count').textContent = '';
}

// ─── Email Detail View ─────────────────────────────────────────────────────────
function openEmailDetail(emailData) {
    _currentEmail = emailData;
    _currentMsgId = emailData.message_id;
    navTo('detail');
    setDetailLoading(false);
    renderDetail(emailData);
}

function loadCurrentOutlookEmail() {
    const item = Office.context.mailbox.item;
    if (!item) { navTo('home'); return; }

    _currentMsgId = item.itemId;
    _currentEmail = null;
    _currentCat   = null;
    navTo('detail');
    setDetailLoading(true);

    const subject = item.subject || '';
    const sender  = item.from?.emailAddress || '';
    const convId  = item.conversationId || null;

    if (item.body) {
        item.body.getAsync("text", (result) => {
            const bodyText = (result.status === "succeeded" || (Office?.MailboxEnums?.AsyncResultStatus && result.status === Office.MailboxEnums.AsyncResultStatus.Succeeded)) ? result.value : '';
            classifyAndShow(_currentMsgId, subject, bodyText, sender, convId, 'Inbox');
        });
    } else {
        classifyAndShow(_currentMsgId, subject, '', sender, convId, 'Inbox');
    }
}

async function classifyAndShow(msgId, subject, body, sender, convId, folder) {
    try {
        if (!STATUS_FLOWS || !Object.keys(STATUS_FLOWS).length) {
            await loadStatusFlows();
        }
        const res = await apiFetch(`${API_HOST}/api/addin/classify`, {
            method: 'POST',
            body: JSON.stringify({ message_id: msgId, subject, body, sender, conversation_id: convId, source_folder: folder })
        });
        if (!res?.ok) throw new Error(`HTTP ${res?.status}`);
        const d = await res.json();
        _currentEmail = d.data || d;
        setDetailLoading(false);
        renderDetail(_currentEmail);
    } catch (e) {
        setDetailError('Cannot reach server: ' + e.message);
    }
}

function setDetailLoading(loading) {
    el('detail-loading').classList.toggle('hidden', !loading);
    el('detail-error').classList.add('hidden');
    el('detail-content').classList.toggle('hidden', loading);
}

function setDetailError(msg) {
    el('detail-loading').classList.add('hidden');
    el('detail-content').classList.add('hidden');
    el('detail-error').classList.remove('hidden');
    el('detail-error-msg').textContent = msg;
}

function renderDetail(data) {
    if (!data) return;

    const cat  = data.category || 'general';
    const catI = CAT[cat] || CAT.general;

    // ── Classification banner
    const banner = el('class-banner');
    banner.className = `class-banner ${cat}`;
    el('class-icon').textContent    = catI.icon;
    el('class-name').textContent    = catI.name;
    el('class-reason').textContent  = data.reason || 'AI classified';
    el('class-conf').style.display  = 'none';
    const cl = document.querySelector('.class-label');
    if (cl) cl.style.display = 'none';

    // ── Priority row
    const tier = data.priority_tier || 'low';
    const pRow = el('priority-row');
    if (data.priority_score != null) {
        pRow.className = `priority-row ${tier}`;
        el('pr-badge').textContent   = PRIORITY_LABEL[tier] || tier;
        el('pr-score').style.display  = 'none';
        el('pr-reasons').style.display = 'none';
        pRow.classList.remove('hidden');
    } else {
        pRow.classList.add('hidden');
    }

    // ── Status stepper
    renderStepper(cat, data.email_status);

    // ── Thread Lifecycle Tracker
    renderThreadLifecycle(data.thread_lifecycle);

    // ── Override selector
    el('override-select').value = cat;

    // ── Extracted fields
    const fieldsList = el('fields-list');
    fieldsList.innerHTML = '';
    let meta = {};
    try { meta = typeof data.extracted_data === 'string' ? JSON.parse(data.extracted_data) : (data.extracted_data || {}); } catch {}
    const keys = Object.keys(meta);
    if (!keys.length) {
        fieldsList.innerHTML = `<p style="font-size:12px;color:var(--ink-3);">No fields extracted for this email.</p>`;
    } else {
        keys.forEach(k => {
            const v = String(meta[k] || 'N/A');
            const row = document.createElement('div');
            row.className = 'field-row';
            row.innerHTML = `
                <div class="field-kv">
                    <div class="field-k">${esc(k)}</div>
                    <div class="field-v">${esc(v)}</div>
                </div>
                <button class="copy-btn" onclick="copyVal('${escAttr(v)}',this)">Copy</button>`;
            fieldsList.appendChild(row);
        });
    }

    // ── Email Summary
    el('email-summary-text').textContent = data.reason || 'No summary available.';

    // ── Pricing calculator visibility
    const calcPanel = el('calc-panel');
    const calcEmpty = el('calc-empty');
    if (cat === 'enquiry') {
        calcPanel.classList.remove('hidden');
        calcEmpty.classList.add('hidden');
        runCalc();
    } else {
        calcPanel.classList.add('hidden');
        calcEmpty.classList.remove('hidden');
    }

    // ── Draft reply
    el('draft-box').value = data.response_draft || '';

    // ── Switch to Triage tab by default
    switchTab('triage');
}

// ─── Status Stepper ────────────────────────────────────────────────────────────
function renderStepper(cat, currentStatus) {
    const flow = STATUS_FLOWS[cat] || [];
    const card = el('stepper-card');
    const btn  = el('advance-btn');

    if (!flow.length) { card.classList.add('hidden'); btn.classList.add('hidden'); return; }
    card.classList.remove('hidden');

    const currIdx = flow.findIndex(s => s.id === currentStatus);
    const termGood = new Set(['po_closed','enq_converted','inv_paid','ship_delivered','gen_read']);
    const termBad  = new Set(['po_cancelled','enq_lost','inv_disputed','ship_delayed']);

    const stepper = el('stepper');
    stepper.innerHTML = '';

    flow.forEach((step, i) => {
        const item = document.createElement('div');
        item.className = 'step';
        let stateCls = '';
        if (termGood.has(step.id) && i === currIdx)  stateCls = 'good';
        else if (termBad.has(step.id) && i === currIdx) stateCls = 'bad';
        else if (i < currIdx)  stateCls = 'done';
        else if (i === currIdx) stateCls = 'active';
        if (stateCls) item.classList.add(stateCls);

        item.innerHTML = `
            <div class="step-node">${i < currIdx ? '✓' : (i === currIdx ? '●' : '')}</div>
            <div class="step-lbl">${esc(step.label)}</div>`;
        stepper.appendChild(item);
    });

    // Current stage label
    const curStep = flow[currIdx];
    el('current-stage-name').textContent = curStep ? curStep.label : '';

    // Advance button
    const nextId   = NEXT_STATUS[currentStatus];
    const nextStep = nextId ? flow.find(s => s.id === nextId) : null;
    if (nextStep) {
        btn.disabled   = false;
        btn.className  = 'advance-btn';
        btn.textContent = `Move to "${nextStep.label}" →`;
        btn.classList.remove('hidden');
    } else if (currIdx >= 0) {
        btn.disabled   = true;
        btn.className  = 'advance-btn complete';
        btn.textContent = '✅ Stage Complete';
        btn.classList.remove('hidden');
    } else {
        btn.classList.add('hidden');
    }
}

function renderThreadLifecycle(info) {
    const card = el('thread-lifecycle-card');
    if (!info || !info.conversation_id) {
        card.classList.add('hidden');
        return;
    }
    card.classList.remove('hidden');

    const stepper = el('thread-lifecycle-stepper');
    const milestonesEl = el('thread-lifecycle-milestones');
    stepper.innerHTML = '';
    milestonesEl.innerHTML = '';

    const stages = [
        { key: 'enquiry', label: 'Enquiry', icon: '❓', check: info.has_enquiry },
        { key: 'purchase_order', label: 'Order', icon: '📦', check: info.has_order },
        { id: 'invoice', key: 'invoice', label: 'Invoice', icon: '💰', check: info.has_invoice },
        { key: 'shipping', label: 'Shipment', icon: '🚚', check: info.has_shipping }
    ];

    stages.forEach((stage, i) => {
        const node = document.createElement('div');
        node.className = 'step';
        
        let stateCls = '';
        if (stage.check) {
            stateCls = 'done';
            if (info.current_stage === stage.key) {
                stateCls = 'active';
            }
        }
        if (stateCls) node.classList.add(stateCls);

        node.innerHTML = `
            <div class="step-node">${stage.check ? '✓' : stage.icon}</div>
            <div class="step-lbl" style="font-size: 9px; margin-top: 4px; font-weight: 600;">${esc(stage.label)}</div>
        `;
        stepper.appendChild(node);
    });

    // Populate milestones list
    const milestones = info.milestones || [];
    if (!milestones.length) {
        milestonesEl.innerHTML = `<div style="color: #9ca3af; font-style: italic; text-align: center;">No lifecycle milestones detected yet.</div>`;
    } else {
        const list = document.createElement('div');
        list.style.display = 'flex';
        list.style.flexDirection = 'column';
        list.style.gap = '6px';

        milestones.forEach(m => {
            const dt = new Date(m.timestamp);
            const formattedDate = dt.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) + ' ' + 
                                  dt.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });

            const row = document.createElement('div');
            row.style.display = 'flex';
            row.style.justifyContent = 'space-between';
            row.style.alignItems = 'center';
            row.style.fontSize = '10px';
            row.innerHTML = `
                <span style="font-weight: 600; color: #374151; display: flex; align-items: center; gap: 4px;">
                    🟢 ${esc(m.label)}
                </span>
                <span style="color: #6b7280; font-size: 9px;">${formattedDate}</span>
            `;
            list.appendChild(row);
        });
        milestonesEl.appendChild(list);
    }
}

async function advanceStatus() {
    if (!_currentMsgId || !_currentEmail) return;
    const nextStatus = NEXT_STATUS[_currentEmail.email_status];
    if (!nextStatus) return;

    const btn = el('advance-btn');
    btn.disabled    = true;
    btn.textContent = 'Updating…';

    try {
        const res = await apiFetch(`${API_HOST}/api/addin/update_status`, {
            method: 'POST',
            body: JSON.stringify({ message_id: _currentMsgId, status: nextStatus })
        });
        if (res?.ok) {
            _currentEmail.email_status = nextStatus;
            renderStepper(_currentEmail.category, nextStatus);
            showToast('✓ Status updated', 'success');
        } else {
            throw new Error('Server error');
        }
    } catch (e) {
        btn.disabled    = false;
        btn.textContent = 'Retry →';
        showToast('Failed to update status', 'error');
    }
}

// ─── Reply Sending ────────────────────────────────────────────────────────────
async function sendAddInReply() {
    if (!_currentMsgId || !_currentEmail) return;
    const replyText = el('draft-box').value;
    const btn = el('send-reply');
    
    btn.textContent = 'Sending…';
    btn.disabled = true;

    try {
        const res = await apiFetch(`${API_HOST}/api/send_reply`, {
            method: 'POST',
            body: JSON.stringify({ message_id: _currentMsgId, reply_text: replyText })
        });
        if (res?.ok) {
            const data = await res.json();
            _currentEmail.email_status = data.new_status;
            _currentEmail.reply_sent = true;
            _currentEmail.sent_reply = replyText;
            
            renderStepper(_currentEmail.category, data.new_status);
            showToast('✓ Reply sent & status updated', 'success');
            
            // Re-render details to show updated state
            renderDetail(_currentEmail);
        } else {
            throw new Error('Server error');
        }
    } catch (e) {
        showToast('Failed to send reply', 'error');
    } finally {
        btn.textContent = '✉ Send Reply';
        btn.disabled = false;
    }
}

async function toggleAddInAutoReply() {
    const checked = el('add-in-auto-reply-toggle').checked;
    try {
        await apiFetch(`${API_HOST}/api/settings`, {
            method: 'POST',
            body: JSON.stringify({ auto_reply: checked })
        });
        showToast(checked ? 'Auto-Reply active' : 'Auto-Reply inactive', 'success');
    } catch {
        showToast('Failed to save settings', 'error');
    }
}

async function loadAddInAutoReplySetting() {
    try {
        const res = await apiFetch(`${API_HOST}/api/settings`);
        if (res?.ok) {
            const data = await res.json();
            el('add-in-auto-reply-toggle').checked = data.auto_reply;
        }
    } catch {}
}

// ─── Override Category ─────────────────────────────────────────────────────────
async function overrideCategory(newCat) {
    if (!_currentMsgId) return;
    try {
        await apiFetch(`${API_HOST}/api/addin/reclassify`, {
            method: 'POST',
            body: JSON.stringify({ message_id: _currentMsgId, category: newCat })
        });
        applyOutlookCategory(newCat);
        if (_currentEmail) { _currentEmail.category = newCat; renderDetail(_currentEmail); }
        showToast(`Reclassified as ${CAT[newCat]?.name || newCat}`, 'success');
    } catch (e) {
        showToast('Reclassify failed', 'error');
    }
}

function applyOutlookCategory(cat) {
    const item = Office.context.mailbox?.item;
    if (!item) return;
    const catInfo = CAT[cat];
    if (!catInfo) return;
    item.categories.getAsync(r => {
        if (r.status === "succeeded" || (Office?.MailboxEnums?.AsyncResultStatus && r.status === Office.MailboxEnums.AsyncResultStatus.Succeeded)) {
            const cur = r.value || [];
            Object.values(CAT).forEach(c => { if (c.name !== catInfo.name && cur.includes(c.name)) item.categories.removeAsync([c.name]); });
            if (!cur.includes(catInfo.name)) item.categories.addAsync([catInfo.name]);
        }
    });
}

// ─── Tab Switcher ──────────────────────────────────────────────────────────────
function switchTab(id) {
    ['triage','pricing','reply'].forEach(t => {
        el(`tab-${t}`).classList.toggle('on', t === id);
        el(`tab-btn-${t}`).classList.toggle('on', t === id);
    });
}

// ─── Pricing Calculator ────────────────────────────────────────────────────────
function runCalc() {
    const type  = el('calc-type').value;
    const qty   = parseInt(el('calc-qty').value) || 1000;
    const l     = parseFloat(el('calc-l').value) || 0;
    const w     = parseFloat(el('calc-w').value) || 0;
    const h     = parseFloat(el('calc-h').value) || 0;
    const print = el('calc-print').value;

    const base  = { '3ply': 1.5e-7, '5ply': 2.5e-7, '7ply': 3.5e-7, 'duplex': 1.2e-7 }[type] || 1.5e-7;
    const pMul  = { none: 1, '1color': 1.1, '2color': 1.2, 'multi': 1.4 }[print] || 1;
    let unit    = Math.max(0.07, l * w * h * base * pMul);
    if (qty >= 5000) unit *= 0.80;
    else if (qty >= 2000) unit *= 0.90;
    const total = unit * qty;
    el('calc-result').textContent = `$${total.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function insertPriceInDraft() {
    const typeEl = el('calc-type');
    const typeName = typeEl.options[typeEl.selectedIndex].text;
    const qty  = el('calc-qty').value;
    const dims = `${el('calc-l').value}×${el('calc-w').value}×${el('calc-h').value}mm`;
    const price = el('calc-result').textContent;

    const insert = `\n\nPricing estimate based on your specifications:\n• Product: ${typeName}\n• Dimensions: ${dims}\n• Quantity: ${qty} units\n• Estimated Total: ${price} (excl. tax & freight)\n\nPlease confirm to proceed with a formal quotation.`;
    el('draft-box').value = el('draft-box').value.trim() + insert;
    switchTab('reply');
    showToast('Price inserted into draft', 'success');
}

// ─── Toast Notification ────────────────────────────────────────────────────────
function showToast(msg, type = '') {
    const container = el('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 220);
    }, 2500);
}

// ─── Status Flows ──────────────────────────────────────────────────────────────
async function loadStatusFlows() {
    try {
        const res = await apiFetch(`${API_HOST}/api/addin/status_flows`);
        if (res?.ok) {
            const d  = await res.json();
            STATUS_FLOWS = d.flows || {};
            NEXT_STATUS  = d.next_status || {};
        }
    } catch {}
}

// ─── Listeners ─────────────────────────────────────────────────────────────────
function setupListeners() {
    el('override-select')?.addEventListener('change', e => overrideCategory(e.target.value));

    ['calc-type','calc-qty','calc-l','calc-w','calc-h','calc-print'].forEach(id => {
        el(id)?.addEventListener('input',  runCalc);
        el(id)?.addEventListener('change', runCalc);
    });
    el('calc-insert')?.addEventListener('click', insertPriceInDraft);

    el('copy-draft')?.addEventListener('click', function () {
        const text = el('draft-box').value;
        navigator.clipboard.writeText(text)
            .then(() => showToast('Draft copied to clipboard', 'success'))
            .catch(() => showToast('Copy failed — please copy manually', 'error'));
    });
}

// ─── Utility ───────────────────────────────────────────────────────────────────
const el = id => {
    const element = document.getElementById(id);
    if (element) return element;
    return {
        style: {},
        classList: {
            add: () => {},
            remove: () => {},
            toggle: () => {},
            contains: () => false
        },
        setAttribute: () => {},
        getAttribute: () => null,
        addEventListener: () => {},
        appendChild: () => {},
        innerHTML: "",
        textContent: ""
    };
};

function esc(str) {
    return String(str)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function escAttr(str) { return String(str).replace(/'/g, "\\'").replace(/"/g, '\\"'); }

function fmtTimeAgo(isoStr) {
    if (!isoStr) return '—';
    const secs = (Date.now() - new Date(isoStr).getTime()) / 1000;
    if (secs < 60)      return 'Just now';
    if (secs < 3600)    return `${Math.floor(secs/60)}m ago`;
    if (secs < 86400)   return `${Math.floor(secs/3600)}h ago`;
    if (secs < 604800)  return `${Math.floor(secs/86400)}d ago`;
    return new Date(isoStr).toLocaleDateString('en-GB', { day:'numeric', month:'short' });
}

function copyVal(text, btn) {
    navigator.clipboard.writeText(text)
        .then(() => { btn.textContent = 'Copied!'; setTimeout(() => btn.textContent = 'Copy', 1800); })
        .catch(() => showToast('Copy failed', 'error'));
}

// ─── Mock Demo (browser preview) ──────────────────────────────────────────────
async function runMockDemo() {
    STATUS_FLOWS = {
        enquiry: [
            { id: 'enq_new',      label: 'New Enquiry' },
            { id: 'enq_quoting',  label: 'Quoting' },
            { id: 'enq_quoted',   label: 'Quote Sent' },
            { id: 'enq_followup', label: 'Follow-up' },
            { id: 'enq_converted',label: 'Converted' },
        ],
        purchase_order: [
            { id: 'po_received',      label: 'Received' },
            { id: 'po_acknowledged',  label: 'Acknowledged' },
            { id: 'po_in_production', label: 'In Production' },
            { id: 'po_ready',         label: 'Ready' },
            { id: 'po_shipped',       label: 'Shipped' },
            { id: 'po_closed',        label: 'Delivered' },
        ],
        junk: [
            { id: 'junk_new',      label: 'New Junk' },
            { id: 'junk_review',   label: 'Under Review' },
            { id: 'junk_flagged',  label: 'Flagged' },
            { id: 'junk_archived', label: 'Archived' },
            { id: 'junk_deleted',  label: 'Deleted' },
        ],
    };
    NEXT_STATUS = {
        enq_new: 'enq_quoting', enq_quoting: 'enq_quoted',
        enq_quoted: 'enq_followup', enq_followup: 'enq_converted',
        po_received: 'po_acknowledged', po_acknowledged: 'po_in_production',
        po_in_production: 'po_ready', po_ready: 'po_shipped', po_shipped: 'po_closed',
        junk_new: 'junk_review', junk_review: 'junk_flagged', junk_flagged: 'junk_archived', junk_archived: 'junk_deleted',
    };

    // Simulate home data
    el('total-today').textContent    = '12';
    el('total-critical').textContent = '2';
    el('cnt-purchase_order').textContent = '4';
    el('cnt-enquiry').textContent        = '3';
    el('cnt-invoice').textContent        = '2';
    el('cnt-shipping').textContent       = '2';
    el('cnt-general').textContent        = '1';
    const critPill = el('crit-enquiry');
    critPill.textContent = '1 critical';
    critPill.classList.add('visible');
    navTo('home');

    // After 1.5s auto-navigate to enquiry list as a demo
    await sleep(1500);
    navToCategory('enquiry');
    await sleep(400);

    // Inject mock email rows
    _allEmails = [
        {
            message_id: 'demo-1', sender: 'ibrahim@quickbox.com', subject: 'URGENT: Request for 5000 units heavy duty cartons',
            priority_tier: 'critical', priority_score: 88, priority_reasons: ['48h without response', '$12,500 estimated value', 'Urgency keyword detected'],
            email_status: 'enq_quoting', category: 'enquiry', estimated_value: 12500,
            received_at: new Date(Date.now() - 49*3600*1000).toISOString(), source_folder: 'Inbox',
            confidence: 95, reason: 'Contains RFQ intent with urgency indicators and high quantity request.',
            extracted_data: JSON.stringify({ Customer: 'QuickBox Inc.', 'Requested Item': '5-Ply Heavy Duty Cartons', Quantity: '5,000 units', 'Custom Printing': 'Yes (Logo)', Urgency: 'High' }),
            response_draft: 'Hi QuickBox Team,\n\nThank you for reaching out regarding your requirement for 5,000 units of 5-Ply Heavy Duty Cartons with custom logo printing.\n\nWe are currently calculating pricing for your specifications and will send a formal quotation within 24 hours. We can also arrange samples upon request.\n\nBest regards,\nProvident Packaging Sales'
        },
        {
            message_id: 'demo-2', sender: 'sales@packplus.co', subject: 'Quote needed for duplex cartons — 2000 pcs',
            priority_tier: 'high', priority_score: 65, priority_reasons: ['No response in 24h', '$4,200 estimated value'],
            email_status: 'enq_new', category: 'enquiry', estimated_value: 4200,
            received_at: new Date(Date.now() - 26*3600*1000).toISOString(), source_folder: 'Inbox',
            confidence: 91, reason: 'RFQ intent with specific quantity and product type.',
            extracted_data: JSON.stringify({ Customer: 'PackPlus Co.', 'Requested Item': 'Duplex Printed Cartons', Quantity: '2,000 units', 'Custom Printing': 'Yes', Urgency: 'Medium' }),
            response_draft: 'Dear PackPlus Team,\n\nThank you for your enquiry for 2,000 Duplex Printed Cartons. We will prepare your pricing estimate and revert within 24 hours.\n\nBest regards,\nProvident Packaging Sales'
        },
        {
            message_id: 'demo-3', sender: 'buyer@ecoboxes.in', subject: 'Pricing enquiry — 3 ply corrugated boxes for Q3',
            priority_tier: 'medium', priority_score: 47, priority_reasons: ['$1,800 estimated value'],
            email_status: 'enq_quoted', category: 'enquiry', estimated_value: 1800,
            received_at: new Date(Date.now() - 4*3600*1000).toISOString(), source_folder: 'Inbox',
            confidence: 88, reason: 'General pricing enquiry for corrugated boxes.',
            extracted_data: JSON.stringify({ Customer: 'EcoBoxes India', 'Requested Item': '3-Ply Corrugated Boxes', Quantity: '1,000 units', Urgency: 'Low' }),
            response_draft: 'Dear EcoBoxes Team,\n\nThank you for your enquiry. Please find our pricing estimate attached.\n\nBest regards,\nProvident Packaging Sales'
        },
    ];
    renderEmailList(_allEmails);
}
