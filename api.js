/**
 * PKP Reimbursement — Frontend API Client
 * js/api.js
 *
 * Replace APPS_SCRIPT_URL with your deployed Apps Script Web App URL.
 */

const PKP_CONFIG = {
  // ⚠ Replace this with your Apps Script deployment URL after publishing
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbwbUXVKUKAFb7DWnkpsLDUEJsK4_cepH867tnCaZVdWrva6W94fBjpmAWToQbnVR7PIgg/exec',
};

// ── Internal fetch helpers ──────────────────────────────────────

async function _get(action, params = {}) {
  const token = Auth.getToken();
  const qs = new URLSearchParams({ action, token, ...params }).toString();
  const res = await fetch(`${PKP_CONFIG.APPS_SCRIPT_URL}?${qs}`);
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

async function _post(action, body = {}) {
  const token = Auth.getToken();
  const res = await fetch(PKP_CONFIG.APPS_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action, token, ...body }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error);
  return data;
}

// ── Auth ────────────────────────────────────────────────────────

const Auth = {
  getToken: () => sessionStorage.getItem('pkp_token') || '',
  getUser:  () => JSON.parse(sessionStorage.getItem('pkp_user') || 'null'),
  isLoggedIn: () => !!Auth.getToken(),
  requireLogin: () => { if (!Auth.isLoggedIn()) window.location.href = 'index.html'; },

  async login(email, password) {
    // Hash password client-side before sending
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(password));
    const passwordHash = Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
    const data = await _post('login', { email, passwordHash });
    sessionStorage.setItem('pkp_token', data.token);
    sessionStorage.setItem('pkp_user',  JSON.stringify(data.user));
    return data.user;
  },

  async logout() {
    try { await _post('logout'); } catch {}
    sessionStorage.clear();
    window.location.href = 'index.html';
  },
};

// ── Claims ──────────────────────────────────────────────────────

const Claims = {
  // filter: 'mine' | 'team' | 'all'
  getList: (filter = 'mine', status = '') =>
    _get('getClaims', { filter, status }),

  getDetail: (claimId) =>
    _get('getClaimDetail', { claimId }),

  submit: (claimData) =>
    _post('submitClaim', claimData),

  managerApprove: (claimId, status, remarks = '') =>
    _post('managerApprove', { claimId, status, remarks }),

  financeApprove: (claimId, status, remarks = '') =>
    _post('financeApprove', { claimId, status, remarks }),

  updateAccountsDetails: (claimId, details) =>
    _post('updateAccountsDetails', { claimId, ...details }),
};

// ── Attachments ─────────────────────────────────────────────────

const Attachments = {
  getNextRef: () =>
    _post('getAttachmentRef'),

  saveLink: (claimId, attachmentRef, attachmentURL) =>
    _post('saveAttachmentLink', { claimId, attachmentRef, attachmentURL }),
};

// ── Reference Data ──────────────────────────────────────────────

const RefData = {
  getTypes: () =>
    _get('getReimbursementTypes'),

  addType: (typeData) =>
    _post('addReimbursementType', typeData),
};

// ── Employees ───────────────────────────────────────────────────

const Employees = {
  getAll: () =>
    _get('getEmployees'),

  add: (empData) =>
    _post('addEmployee', empData),

  update: (employeeId, updates) =>
    _post('updateEmployee', { employeeId, ...updates }),
};

// ── Config ──────────────────────────────────────────────────────

const Config = {
  get: () => _get('getConfig'),
  set: (key, value) => _post('updateConfig', { key, value }),
};

// ── Reports ─────────────────────────────────────────────────────

const Reports = {
  // type: 'clientwise' | 'employeewise' | 'pending_payment' |
  //       'pending_client_receipt' | 'pending_invoice' | 'monthly' | 'all'
  get: (type, extraParams = {}) =>
    _get('getReport', { type, ...extraParams }),
};

// ── Excel Export (SheetJS) ──────────────────────────────────────

const Export = {
  /**
   * Export an array of objects to Excel
   * @param {Array}  data     - Array of row objects
   * @param {string} filename - e.g. 'claims_report.xlsx'
   */
  toExcel(data, filename = 'report.xlsx') {
    if (!window.XLSX) {
      console.error('SheetJS not loaded. Add <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>');
      return;
    }
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Report');
    XLSX.writeFile(wb, filename);
  },
};

// ── Formatting Helpers ──────────────────────────────────────────

const Fmt = {
  amount: (n) => '₹' + (parseFloat(n)||0).toLocaleString('en-IN', { minimumFractionDigits: 2 }),
  date:   (d) => d ? new Date(d).toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' }) : '—',
  status: (s) => ({
    PENDING_MANAGER:  'Pending Manager Approval',
    PENDING_FINANCE:  'Pending Finance Approval',
    APPROVED_FINANCE: 'Finance Approved',
    REJECTED_MANAGER: 'Rejected by Manager',
    REJECTED_FINANCE: 'Rejected by Finance',
    PAID:    'Paid',
    CLOSED:  'Closed',
  }[s] || s),
};
