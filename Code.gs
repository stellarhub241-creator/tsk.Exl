/** School project budget. Amounts are integer satang; Sheets is authoritative. */
const BUDGET = Object.freeze({
  spreadsheetId: '1J54XRY5QRG7Y-Ak7WTDtL-FnJiqWc4vUIo99AQGWZrQ',
  schemaVersion: 1,
  cacheSeconds: 120,
  maxAmount: 99999999900,
  revisionKey: 'budgetRevision',
  ownerKey: 'budgetSetupOwner',
  triggerKey: 'budgetMaintenanceTriggerId'
});
const BG_SCHEMA = Object.freeze({
  BG_Users: ['email', 'name', 'role', 'active'],
  BG_Funds: ['id', 'name'],
  BG_Projects: ['id', 'requestId', 'payloadHash', 'code', 'name', 'category', 'owner', 'fiscalYear', 'fundId', 'allocatedSatang', 'openingSatang', 'createdAt', 'createdBy'],
  BG_Transactions: ['id', 'requestId', 'payloadHash', 'projectId', 'fundId', 'fiscalYear', 'date', 'type', 'amountSatang', 'title', 'documentNo', 'note', 'createdAt', 'createdBy'],
  BG_Audit: ['id', 'recordId', 'requestId', 'action', 'createdAt', 'createdBy'],
  BG_Meta: ['key', 'value']
});
const BG_FUNDS = Object.freeze([
  { id: 'teaching', name: 'ค่าจัดการเรียนการสอน' },
  { id: 'student', name: 'กิจกรรมพัฒนาผู้เรียน' }
]);

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบงบประมาณโครงการโรงเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * กด Run ฟังก์ชันนี้ใน Apps Script editor เพียงครั้งแรกเพื่อขอสิทธิ์และสร้างระบบ
 * ชื่อลงท้าย _ ทำให้หน้าเว็บเรียกฟังก์ชันตั้งค่านี้ผ่าน google.script.run ไม่ได้
 */
function setupSystem_() {
  return publicResult_(function () {
    const email = activeEmail_();
    if (!email) fail_('IDENTITY_UNAVAILABLE', 'ไม่พบอีเมล Google ของบัญชีที่กำลังรัน กรุณาเปิด Apps Script ด้วยบัญชีเจ้าของระบบ');
    return withBudgetLock_(function () {
      const props = PropertiesService.getScriptProperties();
      const owner = props.getProperty(BUDGET.ownerKey);
      if (owner && owner !== email) fail_('FORBIDDEN', 'บัญชีนี้ไม่ใช่ผู้ตั้งค่าระบบ');
      const repaired = ensureSchema_();
      const users = readRows_('BG_Users');
      const own = users.filter(function (u) { return normalizeEmail_(u.email) === email; });
      if (own.length > 1) fail_('DATA_ERROR', 'พบอีเมลซ้ำใน BG_Users กรุณาตรวจสอบ');
      if (!own.length) appendRecord_('BG_Users', { email: email, name: email, role: 'admin', active: true });
      else if (own[0].role !== 'admin' || !isActive_(own[0].active)) fail_('FORBIDDEN', 'บัญชีผู้ตั้งค่าถูกระงับหรือไม่มีสิทธิ์ admin กรุณาตรวจ BG_Users');
      props.setProperty(BUDGET.ownerKey, email);
      bumpRevision_();
      const audit = repairAudit_();
      return { message: 'สร้างและตรวจโครงสร้างระบบเรียบร้อย ไม่มีข้อมูลโครงการจำลอง', repaired: repaired, audit: audit, bootstrap: buildBootstrap_(requireUser_()) };
    });
  });
}

/** Optional: install exactly one hourly maintenance trigger, using admin identity. */
function setupMaintenance() {
  return publicResult_(function () {
    requireRole_(requireUser_(), ['admin']);
    return withBudgetLock_(function () {
      const props = PropertiesService.getScriptProperties();
      const matches = ScriptApp.getProjectTriggers().filter(function (t) { return ['maintainSystem', 'maintainSystem_'].indexOf(t.getHandlerFunction()) >= 0; });
      let keep = matches.find(function (t) { return t.getHandlerFunction() === 'maintainSystem_'; });
      matches.forEach(function (t) { if (t !== keep) ScriptApp.deleteTrigger(t); });
      if (!keep) keep = ScriptApp.newTrigger('maintainSystem_').timeBased().everyHours(1).create();
      props.setProperty(BUDGET.triggerKey, keep.getUniqueId());
      return { message: 'ตั้งตรวจโครงสร้างและซ่อมประวัติรายชั่วโมงแล้ว' };
    });
  });
}

function maintainSystem() {
  return publicResult_(function () {
    requireRole_(requireUser_(), ['admin']);
    return runMaintenance_();
  });
}

/** Private to google.script.run: only the registered trigger uses this entry. */
function maintainSystem_(event) {
  return publicResult_(function () {
    const expected = PropertiesService.getScriptProperties().getProperty(BUDGET.triggerKey);
    const trustedTrigger = expected && event && String(event.triggerUid) === String(expected);
    if (!trustedTrigger) requireRole_(requireUser_(), ['admin']);
    return runMaintenance_();
  });
}

function runMaintenance_() {
  return withBudgetLock_(function () {
    const repaired = ensureSchema_();
    const audit = repairAudit_();
    // Also invalidate after direct edits by trusted spreadsheet editors.
    bumpRevision_();
    return { repaired: repaired, audit: audit, revision: getRevision_() };
  });
}

function apiGateway(action, payload) {
  return publicResult_(function () {
    if (typeof action !== 'string' || !payload || typeof payload !== 'object' || Array.isArray(payload)) fail_('INVALID_INPUT', 'คำสั่งหรือข้อมูลไม่ถูกต้อง');
    const user = requireUser_(); // Fresh whitelist check before cache or any schema mutation.
    if (action === 'GET_BOOTSTRAP') return cachedBootstrap_(user);
    if (action === 'CHECK_OPERATION') {
      const requestId = requestId_(payload.requestId);
      return withBudgetLock_(function () {
        const found = findOperation_(requestId);
        if (!found || normalizeEmail_(found.record.createdBy) !== user.email) return { found: false };
        return { found: true, record: publicRecord_(found.table, found.record), bootstrap: buildBootstrap_(requireUser_()) };
      });
    }
    if (action === 'REPAIR_SCHEMA') {
      requireRole_(user, ['admin']);
      return withBudgetLock_(function () {
        requireRole_(requireUser_(), ['admin']);
        const repaired = ensureSchema_();
        const audit = repairAudit_();
        bumpRevision_();
        return { record: { repaired: repaired, audit: audit }, bootstrap: buildBootstrap_(requireUser_()) };
      });
    }
    if (action !== 'CREATE_PROJECT' && action !== 'CREATE_TRANSACTION') fail_('UNKNOWN_ACTION', 'ไม่พบคำสั่งนี้');
    requireRole_(user, ['admin', 'writer']);
    const normalized = action === 'CREATE_PROJECT' ? projectInput_(payload) : transactionInput_(payload);
    const hash = payloadHash_({ action: action, payload: normalized });
    return withBudgetLock_(function () {
      const currentUser = requireUser_();
      requireRole_(currentUser, ['admin', 'writer']);
      const repairs = ensureSchema_();
      if (repairs.length) bumpRevision_();
      const existing = findOperation_(normalized.requestId);
      if (existing) {
        if (normalizeEmail_(existing.record.createdBy) !== currentUser.email || existing.record.payloadHash !== hash || existing.table !== (action === 'CREATE_PROJECT' ? 'BG_Projects' : 'BG_Transactions')) fail_('REQUEST_CONFLICT', 'รหัสคำขอนี้ถูกใช้กับข้อมูลอื่นแล้ว กรุณาตรวจสอบรายการเดิม');
        try { auditRecord_(existing.table, existing.record); } catch (ignored) { console.warn('Audit replay repair deferred'); }
        return { record: publicRecord_(existing.table, existing.record), bootstrap: buildBootstrap_(currentUser) };
      }
      const table = action === 'CREATE_PROJECT' ? 'BG_Projects' : 'BG_Transactions';
      const record = action === 'CREATE_PROJECT' ? createProjectRecord_(normalized, currentUser) : createTransactionRecord_(normalized, currentUser);
      record.payloadHash = hash;
      // Commit point: this ledger row. Audit/cache failures must never undo it.
      appendRecord_(table, record);
      try { bumpRevision_(); } catch (ignored) { console.warn('Revision recovery needed; ledger is committed'); }
      try { auditRecord_(table, record); } catch (ignored) { console.warn('Audit repair deferred; ledger is committed'); }
      return { record: publicRecord_(table, record), bootstrap: buildBootstrap_(currentUser) };
    });
  });
}

function spreadsheet_() { return SpreadsheetApp.openById(BUDGET.spreadsheetId); }
function withBudgetLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) fail_('SERVER_BUSY', 'ระบบกำลังบันทึกรายการอื่น กรุณารอสักครู่');
  try { return fn(); } finally { lock.releaseLock(); }
}
function fail_(code, message) { const error = new Error(message); error.publicCode = code; throw error; }
function publicResult_(fn) {
  try { return { status: 'success', data: fn() }; }
  catch (error) {
    if (error && error.publicCode) return { status: 'error', code: error.publicCode, message: error.message };
    console.error('Budget server failure: ' + String(error && error.message || error));
    return { status: 'error', code: 'SERVER_ERROR', message: 'ระบบขัดข้องชั่วคราว หากกำลังบันทึก กรุณาตรวจสอบผลรายการก่อนส่งซ้ำ' };
  }
}
function normalizeEmail_(email) { return String(email || '').trim().toLowerCase(); }
function activeEmail_() { return normalizeEmail_(Session.getActiveUser().getEmail()); }
function isActive_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function requireUser_() {
  const email = activeEmail_();
  if (!email) fail_('IDENTITY_UNAVAILABLE', 'ไม่พบอีเมล Google ที่ยืนยันตัวตน กรุณาเปิดด้วยบัญชีโรงเรียนและตรวจการตั้งค่า Web App');
  const users = readRows_('BG_Users');
  const matches = users.filter(function (u) { return normalizeEmail_(u.email) === email; });
  if (matches.length > 1) fail_('DATA_ERROR', 'พบอีเมลซ้ำใน BG_Users กรุณาแจ้งผู้ดูแล');
  if (!matches.length || !isActive_(matches[0].active) || ['admin', 'writer', 'viewer'].indexOf(matches[0].role) < 0) fail_('FORBIDDEN', 'บัญชีนี้ยังไม่ได้รับสิทธิ์ใช้งาน กรุณาติดต่อผู้ดูแล');
  return { email: email, name: String(matches[0].name || email), role: matches[0].role };
}
function requireRole_(user, allowed) { if (allowed.indexOf(user.role) < 0) fail_('FORBIDDEN', 'บัญชีนี้ไม่มีสิทธิ์บันทึกหรือแก้ไขข้อมูล'); }

function headerInfo_(sheet, table, allowMissing) {
  const expected = BG_SCHEMA[table];
  if (!expected) fail_('SCHEMA_ERROR', 'ไม่รู้จักโครงสร้างตาราง');
  const count = sheet.getLastColumn();
  const headers = count ? sheet.getRange(1, 1, 1, count).getValues()[0].map(function (v) { return String(v).trim(); }) : [];
  const seen = {};
  headers.forEach(function (h) {
    if (!h || seen[h] || expected.indexOf(h) < 0) fail_('SCHEMA_ERROR', 'หัวตาราง ' + table + ' ว่าง ซ้ำ หรือไม่ตรง schema กรุณาตรวจสอบโดยไม่ลบข้อมูลเดิม');
    seen[h] = true;
  });
  if (!allowMissing && expected.some(function (h) { return !seen[h]; })) fail_('SCHEMA_REQUIRED', 'ตาราง ' + table + ' ยังไม่ครบ กรุณาให้ผู้ดูแลใช้ตรวจและซ่อมโครงสร้าง');
  return headers;
}
function ensureSchema_() {
  const ss = spreadsheet_();
  const changed = [];
  // Validate every existing managed table before creating or repairing anything.
  Object.keys(BG_SCHEMA).forEach(function (table) {
    const existing = ss.getSheetByName(table);
    if (existing) headerInfo_(existing, table, true);
  });
  Object.keys(BG_SCHEMA).forEach(function (table) {
    let sheet = ss.getSheetByName(table);
    if (!sheet) { sheet = ss.insertSheet(table); changed.push('สร้าง ' + table); }
    const headers = headerInfo_(sheet, table, true);
    const missing = BG_SCHEMA[table].filter(function (h) { return headers.indexOf(h) < 0; });
    if (missing.length) {
      const need = headers.length + missing.length;
      if (sheet.getMaxColumns() < need) sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      changed.push('เติมหัวตาราง ' + table + ': ' + missing.join(', '));
    }
    sheet.setFrozenRows(1);
  });
  const funds = readRows_('BG_Funds');
  if (funds.some(function (f) { return !BG_FUNDS.some(function (known) { return known.id === f.id; }); }) || new Set(funds.map(function (f) { return f.id; })).size !== funds.length) fail_('DATA_ERROR', 'BG_Funds มีรหัสกองเงินซ้ำหรือไม่ตรงระบบ');
  BG_FUNDS.forEach(function (f) {
    if (!funds.some(function (r) { return r.id === f.id; })) { appendRecord_('BG_Funds', f); changed.push('เพิ่มกองเงิน ' + f.id); }
  });
  const meta = readRows_('BG_Meta');
  const versions = meta.filter(function (r) { return r.key === 'schemaVersion'; });
  if (!versions.length) { appendRecord_('BG_Meta', { key: 'schemaVersion', value: BUDGET.schemaVersion }); changed.push('ตั้ง schemaVersion'); }
  else if (versions.length !== 1 || Number(versions[0].value) !== BUDGET.schemaVersion) fail_('SCHEMA_ERROR', 'schemaVersion ไม่ตรงรุ่น กรุณาตรวจสอบก่อนซ่อม');
  SpreadsheetApp.flush();
  return changed;
}
function readRows_(table) {
  const sheet = spreadsheet_().getSheetByName(table);
  if (!sheet) fail_('SETUP_REQUIRED', 'ยังไม่มีตาราง ' + table + ' กรุณาให้เจ้าของกด Run ฟังก์ชัน setupSystem_ ใน Apps Script');
  const headers = headerInfo_(sheet, table, false);
  if (sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
    .filter(function (row) { return row.some(function (v) { return v !== ''; }); })
    .map(function (row) {
      const record = {};
      headers.forEach(function (h, i) { record[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]; });
      return record;
    });
}
function appendRecord_(table, record) {
  const sheet = spreadsheet_().getSheetByName(table);
  if (!sheet) fail_('SETUP_REQUIRED', 'ยังไม่มีตาราง ' + table);
  const headers = headerInfo_(sheet, table, false);
  const row = headers.map(function (h) { return record[h] === undefined ? '' : record[h]; });
  const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length);
  headers.forEach(function (h, index) {
    if (typeof row[index] === 'string') {
      sheet.getRange(range.getRow(), index + 1).setNumberFormat('@');
      // Leading apostrophe makes spreadsheet formula-like text literal.
      if (/^[=+\-@]/.test(row[index])) row[index] = "'" + row[index];
    }
  });
  range.setValues([row]);
  SpreadsheetApp.flush();
}

function requestId_(value) {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{8,100}$/.test(value)) fail_('INVALID_INPUT', 'รหัสคำขอไม่ถูกต้อง');
  return value;
}
function text_(value, label, max, optional) {
  if (typeof value !== 'string') { if (optional && (value === undefined || value === null)) return ''; fail_('INVALID_INPUT', label + ' ต้องเป็นข้อความ'); }
  const text = value.trim();
  if ((!optional && !text) || text.length > max || /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(text)) fail_('INVALID_INPUT', label + ' ไม่ถูกต้องหรือยาวเกินกำหนด');
  return text;
}
function amount_(value, label, positive) {
  if (!Number.isSafeInteger(value) || value < (positive ? 1 : 0) || value > BUDGET.maxAmount) fail_('INVALID_INPUT', label + ' ต้องเป็นจำนวนสตางค์เต็มที่ถูกต้อง');
  return value;
}
function addMoney_(left, right) {
  const result = left + right;
  if (!Number.isSafeInteger(result)) fail_('DATA_ERROR', 'ยอดรวมเกินขอบเขตจำนวนเงินที่ระบบรองรับ');
  return result;
}
function fiscalYear_(year) { if (!Number.isInteger(year) || year < 2500 || year > 2800) fail_('INVALID_INPUT', 'ปีงบประมาณต้องเป็น พ.ศ. ที่ถูกต้อง'); return year; }
function validDate_(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) fail_('INVALID_INPUT', 'วันที่ไม่ถูกต้อง');
  const parts = value.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (date.getUTCFullYear() !== parts[0] || date.getUTCMonth() !== parts[1] - 1 || date.getUTCDate() !== parts[2]) fail_('INVALID_INPUT', 'วันที่ไม่มีอยู่จริง');
  return value;
}
function projectInput_(p) {
  const fund = text_(p.fundId, 'กองเงิน', 40, false);
  if (!BG_FUNDS.some(function (f) { return f.id === fund; })) fail_('INVALID_INPUT', 'กองเงินไม่ถูกต้อง');
  return { requestId: requestId_(p.requestId), code: text_(p.code, 'รหัสโครงการ', 40, false), name: text_(p.name, 'ชื่อโครงการ', 200, false), category: text_(p.category, 'กลุ่มงาน', 100, false), owner: text_(p.owner, 'ผู้รับผิดชอบ', 160, false), fiscalYear: fiscalYear_(p.fiscalYear), fundId: fund, allocatedSatang: amount_(p.allocatedSatang, 'งบจัดสรร', false), openingSatang: amount_(p.openingSatang, 'เงินยกมา', false) };
}
function transactionInput_(p) {
  if (p.type !== 'income' && p.type !== 'expense') fail_('INVALID_INPUT', 'ประเภทรายการไม่ถูกต้อง');
  return { requestId: requestId_(p.requestId), projectId: text_(p.projectId, 'โครงการ', 100, false), date: validDate_(p.date), type: p.type, amountSatang: amount_(p.amountSatang, 'จำนวนเงิน', true), title: text_(p.title, 'ชื่อรายการ', 200, false), documentNo: text_(p.documentNo, 'เลขเอกสาร', 100, true), note: text_(p.note, 'หมายเหตุ', 500, true) };
}
function payloadHash_(payload) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(payload))
    .map(function (byte) { return ('0' + (byte < 0 ? byte + 256 : byte).toString(16)).slice(-2); }).join('');
}
function findOperation_(requestId) {
  const found = [];
  ['BG_Projects', 'BG_Transactions'].forEach(function (table) {
    readRows_(table).forEach(function (record) { if (record.requestId === requestId) found.push({ table: table, record: record }); });
  });
  if (found.length > 1) fail_('DATA_ERROR', 'พบรหัสคำขอซ้ำในข้อมูลจริง กรุณาแจ้งผู้ดูแล');
  return found[0] || null;
}
function createProjectRecord_(p, user) {
  const projects = readRows_('BG_Projects');
  if (projects.some(function (r) { return r.fiscalYear === p.fiscalYear && r.fundId === p.fundId && String(r.code).toLowerCase() === p.code.toLowerCase(); })) fail_('DUPLICATE_PROJECT', 'รหัสโครงการนี้มีแล้วในปีและกองเงินที่เลือก');
  return Object.assign({}, p, { id: Utilities.getUuid(), createdAt: new Date().toISOString(), createdBy: user.email });
}
function createTransactionRecord_(p, user) {
  const state = readState_();
  const project = state.projects.find(function (r) { return r.id === p.projectId; });
  if (!project) fail_('PROJECT_NOT_FOUND', 'ไม่พบโครงการนี้ กรุณาโหลดข้อมูลใหม่');
  const totals = projectTotals_(project, state.transactions);
  if (p.type === 'expense') {
    if (p.amountSatang > totals.budgetRemaining) fail_('BUDGET_EXCEEDED', 'ยอดจ่ายเกินงบจัดสรรคงเหลือของโครงการ');
    if (p.amountSatang > totals.cashBalance) fail_('CASH_EXCEEDED', 'เงินรับจริงและเงินยกมาคงเหลือไม่พอจ่าย');
  } else addMoney_(totals.cashBalance, p.amountSatang);
  return Object.assign({}, p, { id: Utilities.getUuid(), fundId: project.fundId, fiscalYear: project.fiscalYear, createdAt: new Date().toISOString(), createdBy: user.email });
}
function projectTotals_(project, transactions) {
  let income = 0, expense = 0;
  transactions.filter(function (t) { return t.projectId === project.id; }).forEach(function (t) {
    if (t.type === 'income') income = addMoney_(income, t.amountSatang);
    else expense = addMoney_(expense, t.amountSatang);
  });
  return { budgetRemaining: addMoney_(project.allocatedSatang, -expense), cashBalance: addMoney_(addMoney_(project.openingSatang, income), -expense) };
}

function publicRecord_(table, record) {
  const keys = table === 'BG_Projects' ? ['id', 'code', 'name', 'category', 'owner', 'fiscalYear', 'fundId', 'allocatedSatang', 'openingSatang'] : ['id', 'requestId', 'projectId', 'fundId', 'fiscalYear', 'date', 'type', 'amountSatang', 'title', 'documentNo', 'note', 'createdAt', 'createdBy'];
  const result = {};
  keys.forEach(function (key) { result[key] = record[key]; });
  return result;
}
function readState_() {
  const projects = readRows_('BG_Projects');
  const transactions = readRows_('BG_Transactions');
  const ids = {};
  const operations = {};
  projects.forEach(function (p) {
    if (!p.id || ids[p.id]) fail_('DATA_ERROR', 'รหัสโครงการว่างหรือซ้ำ');
    ids[p.id] = p;
    fiscalYear_(p.fiscalYear);
    amount_(p.allocatedSatang, 'งบจัดสรรในชีท', false);
    amount_(p.openingSatang, 'เงินยกมาในชีท', false);
    if (!BG_FUNDS.some(function (f) { return f.id === p.fundId; })) fail_('DATA_ERROR', 'โครงการอ้างอิงกองเงินไม่ถูกต้อง');
  });
  const txIds = {};
  transactions.forEach(function (t) {
    const project = ids[t.projectId];
    if (!t.id || txIds[t.id] || !project || project.fundId !== t.fundId || project.fiscalYear !== t.fiscalYear || ['income', 'expense'].indexOf(t.type) < 0) fail_('DATA_ERROR', 'รายการเงินจริงซ้ำหรืออ้างอิงโครงการ ปี หรือกองเงินไม่ถูกต้อง');
    txIds[t.id] = true;
    amount_(t.amountSatang, 'จำนวนเงินในชีท', true);
    validDate_(t.date);
  });
  projects.concat(transactions).forEach(function (r) {
    if (!r.requestId || operations[r.requestId] || !r.payloadHash || !r.createdBy) fail_('DATA_ERROR', 'ข้อมูล requestId ผู้บันทึก หรือ payloadHash ในบัญชีรายการไม่ครบหรือซ้ำ');
    operations[r.requestId] = true;
  });
  projects.forEach(function (p) { projectTotals_(p, transactions); });
  return { projects: projects, transactions: transactions };
}
function buildBootstrap_(user) {
  const state = readState_();
  const years = Array.from(new Set(state.projects.map(function (p) { return p.fiscalYear; }))).sort(function (a, b) { return b - a; });
  if (!years.length) {
    const date = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM').split('-').map(Number);
    years.push(date[0] + 543 + (date[1] >= 10 ? 1 : 0));
  }
  return { schemaVersion: BUDGET.schemaVersion, revision: getRevision_(), serverTime: new Date().toISOString(), user: user, funds: BG_FUNDS.map(function (f) { return { id: f.id, name: f.name }; }), projects: state.projects.map(function (p) { return publicRecord_('BG_Projects', p); }), transactions: state.transactions.map(function (t) { return publicRecord_('BG_Transactions', t); }), fiscalYears: years };
}
function getRevision_() {
  const value = PropertiesService.getScriptProperties().getProperty(BUDGET.revisionKey) || '0';
  if (!/^\d+$/.test(value) || !Number.isSafeInteger(Number(value))) fail_('DATA_ERROR', 'revision ระบบไม่ถูกต้อง กรุณาติดต่อผู้ดูแล');
  return value;
}
function bumpRevision_() {
  const props = PropertiesService.getScriptProperties();
  const next = Number(getRevision_()) + 1;
  if (!Number.isSafeInteger(next)) fail_('DATA_ERROR', 'revision เกินขอบเขตที่รองรับ');
  const revision = String(next);
  props.setProperty(BUDGET.revisionKey, revision);
  return revision;
}
function cachedBootstrap_(user) {
  const cache = CacheService.getScriptCache();
  const revision = getRevision_();
  const key = ['bg', BUDGET.spreadsheetId, BUDGET.schemaVersion, revision, user.email, user.role].join(':');
  try {
    const hit = cache.get(key);
    if (hit) {
      const data = JSON.parse(hit);
      if (data.revision === revision && data.user.email === user.email && data.user.role === user.role) { data.user = user; data.serverTime = new Date().toISOString(); return data; }
    }
  } catch (ignored) { /* Cache unavailable: Sheets remains authoritative. */ }
  const data = withBudgetLock_(function () { return buildBootstrap_(requireUser_()); });
  // A concurrent commit changes revision: cache this snapshot under its own revision.
  const snapshotKey = ['bg', BUDGET.spreadsheetId, BUDGET.schemaVersion, data.revision, user.email, user.role].join(':');
  try { const encoded = JSON.stringify(data); if (encoded.length < 30000) cache.put(snapshotKey, encoded, BUDGET.cacheSeconds); } catch (ignored) {}
  return data;
}
function auditRecord_(table, record) {
  const exists = readRows_('BG_Audit').some(function (a) { return a.recordId === record.id; });
  if (!exists) appendRecord_('BG_Audit', { id: Utilities.getUuid(), recordId: record.id, requestId: record.requestId, action: table === 'BG_Projects' ? 'CREATE_PROJECT' : 'CREATE_TRANSACTION', createdAt: record.createdAt, createdBy: record.createdBy });
}
function repairAudit_() {
  const existing = {};
  readRows_('BG_Audit').forEach(function (a) { existing[a.recordId] = true; });
  let repaired = 0;
  ['BG_Projects', 'BG_Transactions'].forEach(function (table) {
    readRows_(table).forEach(function (record) {
      if (!existing[record.id]) { auditRecord_(table, record); existing[record.id] = true; repaired++; }
    });
  });
  return { repaired: repaired };
}
