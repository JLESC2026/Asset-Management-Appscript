// ═══════════════════════════════════════════════════════════
//  ASSET MANAGEMENT SYSTEM — Code.gs  (v6.1 — Bulk Transfer + Accountability Forms)
// ═══════════════════════════════════════════════════════════

const SHEET_ID    = '18tuYQKH2OLLu1NqPJiA28n8n7GNN6XR_SSZXUO4XEe8';
const SH_ENTRY    = 'Asset Entry';
const SH_USERS    = 'Users';
const SH_ORG      = 'Org Structure';
const SH_MASTER   = 'Masterlist';
const SH_XFER     = 'Transfers';
const SH_BORROW   = 'Borrows';
const SH_DISPOSE  = 'Disposals';
const SH_LOG      = 'ActivityLog';
const SH_DROPDOWN = 'Drop down';
const SH_ALLOC    = 'Allocated';
const SH_SPARE    = 'Spare';

const AE_DATA_START  = 4;
const EVT_DATA_START = 4;

const C = {
  ENTRY_EMP_ID:  1,
  ENTRY_NAME:    2,
  EMP_ID:        3,
  STAFF:         4,
  DESIGNATION:   5,
  DEPARTMENT:    6,
  BASE_OFFICE:   7,
  DIVISION:      8,
  DISTRICT:      9,
  AREA:          10,
  BRANCH:        11,
  ASSIGNMENT:    12,
  EFF_DATE:      13,
  BARCODE:       14,
  TYPE:          15,
  BRAND:         16,
  SERIAL:        17,
  SPECS:         18,
  SUPPLIER:      19,
  CONDITION:     20,
  ASSET_LOCATION:21,
  LIFECYCLE:     22,
  STATUS_LABEL:  23,
  ASSET_STATUS:  24,
  PURCH_DATE:    25,
  WARRANTY_TERM: 26,
  WARRANTY_VAL:  27,
  REMARKS:       28,
  NOTES:         29,
  CREATED_AT:    30,
  LAST_UPDATED:  31,
};

const TOTAL_COLS = 31;

const AE_HEADERS = [
  'Inputter ID','Inputter Name',
  'Accountable Employee ID','Accountable Staff','Designation',
  'Department','Base Office',
  'Division','District','Area','Branch',
  'Assignment','Effectivity Date',
  'Barcode','Category','Brand','Serial','Specification',
  'Supplier','Condition','Asset Location',
  'Lifecycle Status','Status Label','Assignment Status',
  'Date of Purchase','Warranty Term','Warranty Validity','Remarks',
  'Notes','Timestamp','Last Updated'
];

// ─── APPROVAL COLUMNS (32-37) ─────────────────────────────────────────────────
const CAE = {
  APPROVAL_STATUS:    32,
  FORM_ID:            33,
  DRAFTED_BY:         34,
  DRAFTED_AT:         35,
  REJECTION_COMMENT:  36,
  GRANDFATHERED:      37
};

// ─── UTILITY ─────────────────────────────────────────────────────────────────
function _sanitize(val, maxLen) {
  maxLen = maxLen || 500;
  return String(val || '').trim().substring(0, maxLen);
}

function _normDiv(raw) {
  if (!raw) return '';
  // Normalise "div 1", "DIV 1", "Division 1", "DIVISION 1" → "Div 1"
  return String(raw).trim().replace(/^div(?:ision)?\s*/i, 'Div ').trim();
}

function _normDist(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  const m = s.match(/^district\s+0*(\d+)$/i);
  if (m) return 'District ' + parseInt(m[1]);
  return s;
}

function _numSort(a, b) {
  const na = parseInt((a || '').replace(/\D+/g, ''), 10) || 0;
  const nb = parseInt((b || '').replace(/\D+/g, ''), 10) || 0;
  return na !== nb ? na - nb : a.localeCompare(b);
}

function _blankOrDash(val) {
  const s = String(val || '').trim();
  return s || ' - ';
}

// ─── STATUS COMPUTATION ───────────────────────────────────────────────────────
function _computeStatus(lifecycle, assetStatus, empId) {
  const lc  = String(lifecycle    || '').trim().toLowerCase();
  const as  = String(assetStatus  || '').trim().toLowerCase();
  const hasEmp = empId && String(empId).trim()
    && !['', 'n/a', 'none', '#n/a'].includes(String(empId).trim().toLowerCase());

  if (as === 'borrowitem' || lc === 'borrowitem') return 'borrow-item';
  if (lc === 'borrow')                            return 'borrowed';
  if (lc === 'returned')                          return 'returned';
  if (lc === 'dispose' || lc === 'disposal' || lc === 'disposed' || as === 'disposal')
                                                  return 'disposal';
  if (lc === 'transfer')                          return 'transfer';
  if (lc === 'allocated' || as === 'assigned')    return 'allocated';
  if (lc === 'spare')                             return 'spare';
  return hasEmp ? 'allocated' : 'spare';
}

// ─── SHEET HELPERS ───────────────────────────────────────────────────────────
function _ss() { return SpreadsheetApp.openById(SHEET_ID); }

function _getOrCreate(name, headers) {
  const ss = _ss();
  let sh   = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers]).setFontWeight('bold')
        .setBackground('#0f0e1c').setFontColor('#a07ee0').setWrap(false);
      sh.setFrozenRows(1);
      sh.setColumnWidths(1, headers.length, 140);
    sh.insertRowsAfter(1, 2); // ensure data starts at row 4, matching EVT_DATA_START = 4
    }
  }
  return sh;
}

function _entrySheet() { return _getOrCreate(SH_ENTRY, AE_HEADERS); }
function _xferSheet() {
  return _getOrCreate(SH_XFER, [
    'Barcode','TransferType',
    'FromStaff','FromEmpID','FromDesig','FromDept','FromBaseOffice',
    'FromDiv','FromDist','FromArea','FromBranch','FromCondition','FromAssetLoc','FromRemarks',
    'ToStaff','ToEmpID','ToDesig','ToDept','ToBaseOffice',
    'ToDiv','ToDist','ToArea','ToBranch','ToCondition','ToAssetLoc','ToRemarks',
    'EffDate','Status','Timestamp'
  ]);
}
function _borrowSheet() {
  return _getOrCreate(SH_BORROW, [
    'Barcode','BorrowerName','EmpID','Designation','Base Office',
    'Division','District','Branch',
    'BorrowDate','ExpectedReturn','ActualReturn',
    'Status','Remarks','Timestamp'
  ]);
}
function _disposeSheet() {
  return _getOrCreate(SH_DISPOSE, [
    'Barcode','Serial Number','Reason','DisposedBy','DisposeDate',
    'Remarks','Timestamp','Location','Equipment Type','Description'
  ]);
}
function _logSheet() {
  return _getOrCreate(SH_LOG, [
    'Timestamp','Action','Barcode','Details','Performed By'
  ]);
}
function _allocLogSheet() {
  return _getOrCreate(SH_ALLOC, [
    'Barcode','Category','Brand','Serial No.','Employee ID',
    'Accountable Staff','Designation','Department','Base Office',
    'Division','District','Area','Branch','Effectivity Date',
    'Condition','Asset Location','Remarks','Timestamp','Allocated By'
  ]);
}
function _spareSheet() {
  return _getOrCreate(SH_SPARE, [
    'Accountable Person ID','Accountable Person','Assignment','Designation','Department',
    'Barcode','Category','Brand','Serial No.','Condition',
    'Purchase Date','Warranty Validity','Supplier','Base Office',
    'Division','District','Area','Branch','Asset Location',
    'Enrolled By','Created By'
  ]);
}

// ─── ROW HELPERS ─────────────────────────────────────────────────────────────
function _findRow(sheet, barcode) {
  if (!barcode) return -1;
  const bcStr = String(barcode).trim();

  // Synthetic key for no-barcode assets: NOBC-{sheetRowNumber}
  if (bcStr.startsWith('NOBC-')) {
    const rowNum = parseInt(bcStr.slice(5), 10);
    return (!isNaN(rowNum) && rowNum >= AE_DATA_START) ? rowNum : -1;
  }

  try {
    const finder = sheet.createTextFinder(bcStr)
      .matchEntireCell(true).matchCase(false);
    const range = finder.findNext();
    if (!range) return -1;
    const row = range.getRow();
    if (row < AE_DATA_START || range.getColumn() !== C.BARCODE) return -1;
    return row;
  } catch (e) {
    const last = sheet.getLastRow();
    if (last < AE_DATA_START) return -1;
    const vals = sheet.getRange(AE_DATA_START, C.BARCODE,
      last - AE_DATA_START + 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === bcStr) return i + AE_DATA_START;
    }
    return -1;
  }
}

function _setRow(sheet, rowIdx, updates) {
  // Dynamically expand to cover approval columns (32-37) when needed
  const maxCol = updates.reduce(function(m, u) { return Math.max(m, u[0]); }, TOTAL_COLS);
  const range  = sheet.getRange(rowIdx, 1, 1, maxCol);
  const rowVals = range.getValues()[0];
  // Pad the array if the row doesn't yet have those columns
  while (rowVals.length < maxCol) rowVals.push('');
  updates.forEach(function(u) { rowVals[u[0] - 1] = (u[1] != null ? u[1] : ''); });
  rowVals[C.LAST_UPDATED - 1] = new Date().toLocaleString('en-PH');
  range.setValues([rowVals]);
}

// ─── WEB APP ─────────────────────────────────────────────────────────────────
function doGet() {
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.SCRIPT_URL = ScriptApp.getService().getUrl();
  return tmpl.evaluate()
    .setTitle('Asset Management System')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }
function getScriptUrl() { return ScriptApp.getService().getUrl(); }

// ─── PASSWORD ─────────────────────────────────────────────────────────────────
function _hashPwd(pwd) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(pwd));
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}
function _isHashed(str) { return /^[0-9a-f]{64}$/.test(String(str)); }

// ─── ROLE CLASSIFICATION ──────────────────────────────────────────────────────
function _mapRoleTier(role) {
  const r = String(role || '').trim().toLowerCase();

  // Keyword-based detection for flexibility
  if (r.includes('admin') || r.includes('head') || r.includes('manager') ||
      r.includes('director') || r.includes('superuser') || r.includes('super user') ||
      r.includes('it head') || r.includes('department head') || r.includes('dept head') ||
      r.includes('iictd head') || r.includes('system admin') || r.includes('sysadmin')) {
    return 'ho';
  }
  if (r.includes('senior') || r.includes('supervisor') || r.includes('team lead') ||
      r.includes('lead engineer') || r.includes('sfe') || r.includes('sfr')) {
    return 'senior';
  }
  return 'fe';
}
function _parseRemarks(remarks, roleTier) {
  const r = String(remarks || '').toLowerCase().trim();

  var scopeType = null;
  if ([
    'dual scope','all scope','both scope','full visibility','all data',
    'all inventory access','all view access','all access','full access to all'
  ].some(k => r.includes(k)))
    scopeType = 'both';
  else if ([
    'can see field','field scope','field data','field only',
    'field office inventories','field office inventory','field office'
  ].some(k => r.includes(k)))
    scopeType = 'field';
  else if ([
    'can see ho','can see head office','ho scope','ho data','ho only','head office only',
    'central office inventories','central office inventory','central office'
  ].some(k => r.includes(k)))
    scopeType = 'ho';

  const DEFAULTS = {
    ho:     { canAdd:true,  canEdit:true, canDelete:true,  canAllocate:true, canDealloc:true, canDispose:true, canBorrow:true, canTransfer:true, viewOnly:false },
    senior: { canAdd:true,  canEdit:true, canDelete:false, canAllocate:true, canDealloc:true, canDispose:true, canBorrow:true, canTransfer:true, viewOnly:false },
    fe:     { canAdd:true,  canEdit:true, canDelete:false, canAllocate:true, canDealloc:true, canDispose:true, canBorrow:true, canTransfer:true, viewOnly:false }
  };
  const perms = Object.assign({}, DEFAULTS[roleTier] || DEFAULTS.fe);
  if (!r) return Object.assign({}, perms, { scopeType: scopeType });

  if (['view only','view-only','read only','read-only','view access only','can only view','view'].some(k => r.includes(k)))
    return Object.assign({}, { canAdd:false, canEdit:false, canDelete:false, canAllocate:false, canDealloc:false, canDispose:false, canBorrow:false, canTransfer:false, viewOnly:true }, { scopeType: scopeType });

  if (['full access','all access','full permission','all permissions','unrestricted','complete access'].some(k => r.includes(k)))
    return Object.assign({}, { canAdd:true, canEdit:true, canDelete:true, canAllocate:true, canDealloc:true, canDispose:true, canBorrow:true, canTransfer:true, viewOnly:false }, { scopeType: scopeType });

  if (r.includes('can delete')   || r.includes('can remove'))                             perms.canDelete   = true;
  if (r.includes('can add')      || r.includes('can enroll')  || r.includes('can create')) perms.canAdd     = true;
  if (r.includes('can edit')     || r.includes('can update')  || r.includes('can modify')) perms.canEdit    = true;
  if (r.includes('can allocate') || r.includes('can assign'))                             perms.canAllocate = true;
  if (r.includes('can return to spare') || r.includes('can deallocate'))                  perms.canDealloc  = true;
  if (r.includes('can dispose')  || r.includes('can decommission'))                       perms.canDispose  = true;
  if (r.includes('can borrow')   || r.includes('can lend'))                               perms.canBorrow   = true;
  if (r.includes('can transfer') || r.includes('can reassign'))                           perms.canTransfer = true;
  if (r.includes('no add')       || r.includes('cannot add')      || r.includes('no enroll'))  perms.canAdd      = false;
  if (r.includes('no edit')      || r.includes('cannot edit')     || r.includes('no modify'))  perms.canEdit     = false;
  if (r.includes('no delete')    || r.includes('cannot delete')   || r.includes('no remove'))  perms.canDelete   = false;
  if (r.includes('no allocate')  || r.includes('cannot allocate'))                             perms.canAllocate = false;
  if (r.includes('no dispose')   || r.includes('cannot dispose'))                              perms.canDispose  = false;
  if (r.includes('no borrow')    || r.includes('cannot borrow'))                               perms.canBorrow   = false;
  if (r.includes('no transfer')  || r.includes('cannot transfer'))                             perms.canTransfer = false;
  return Object.assign({}, perms, { scopeType: scopeType });
}

// ─── MASTERLIST LOOKUP ────────────────────────────────────────────────────────
// Masterlist columns:
//   Col A (0) = EmpID
//   Col C (2) = Name
//   Col E (4) = Division
//   Col F (5) = District
//   Col G (6) = Area
//   Col H (7) = Base Office / Branch
//   Col L (11) = Position / Designation
function _getMasterlistEntry(empId) {
  try {
    const sh   = _ss().getSheetByName(SH_MASTER);
    if (!sh) return {};
    const last = sh.getLastRow();
    if (last < 2) return {};
    const id   = String(empId).trim().toLowerCase();
    const ids  = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() !== id) continue;
      const row = sh.getRange(i + 2, 1, 1, 15).getValues()[0];
      return {
        name:       String(row[2]  || '').trim(),
        division:   _normDiv(_blankOrDash(String(row[4]  || ''))),
        district:   _normDist(_blankOrDash(String(row[5] || ''))),
        area:       _blankOrDash(String(row[6]  || '')),
        baseOffice: _blankOrDash(String(row[7]  || '')),
        position:   _blankOrDash(String(row[11] || ''))
      };
    }
  } catch(e) {}
  return {};
}

// ─── EMPLOYEE LOOKUP (public, called from frontend) ───────────────────────────
// Returns all fields needed for auto-fill in modals.
// area is now included and all blank fields return ' - '
function getEmployeeById(empId) {
  try {
    const sh = _ss().getSheetByName(SH_MASTER);
    if (!sh) return { ok: false, error: 'Masterlist not found.' };
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, error: 'No employees found.' };
    const id = String(empId).trim().toLowerCase();
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() !== id) continue;
      const row = sh.getRange(i + 2, 1, 1, 15).getValues()[0];
      return {
        ok:         true,
        empId:      String(row[0]  || '').trim(),
        name:       String(row[2]  || '').trim(),
        // ⚠ Verify these column indices match your actual Masterlist sheet:
        department: _blankOrDash(String(row[3]  || '')),  // Col D — adjust if needed
        division:   _normDiv(_blankOrDash(String(row[4]  || ''))),
        district:   _normDist(_blankOrDash(String(row[5] || ''))),
        area:       _blankOrDash(String(row[6]  || '')),
        branch:     _blankOrDash(String(row[7]  || '')),
        baseOffice: _blankOrDash(String(row[7]  || '')),
        position:   _blankOrDash(String(row[11] || ''))
      };
    }
    return { ok: false, error: 'Employee ID not found: ' + empId };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ─── ORG STRUCTURE SCOPE ──────────────────────────────────────────────────────
function _parseOrgStructure(empId, roleTier) {
  const DEFAULT = {
    userDivisions: [], userDistricts: [],
    seniorDistrictScope: [], divDistrictMap: {}
  };
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return DEFAULT;
    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, 12).getValues();
    const id   = String(empId || '').trim().toLowerCase();

    if (roleTier === 'ho') {
      const divSet = new Set(), distSet = new Set(), divMap = {};
      data.forEach(r => {
        const div  = _normDiv(String(r[10] || '').trim());
        const dist = _normDist(String(r[11] || '').trim());
        if (div)  divSet.add(div);
        if (dist) distSet.add(dist);
        if (div && dist) {
          if (!divMap[div]) divMap[div] = [];
          if (!divMap[div].includes(dist)) divMap[div].push(dist);
        }
      });
      const allDists = [...distSet].sort(_numSort);
      return {
        userDivisions:       [...divSet].sort(),
        userDistricts:       allDists,
        seniorDistrictScope: allDists,
        divDistrictMap:      divMap
      };
    }

    if (roleTier === 'senior') {
      const divSet = new Set(), distSet = new Set(), divMap = {};
      data.forEach(r => {
        const supId = String(r[6] || '').trim().toLowerCase();
        if (!id || supId !== id) return;
        const div  = _normDiv(String(r[10] || '').trim());
        const dist = _normDist(String(r[11] || '').trim());
        if (div)  divSet.add(div);
        if (dist) distSet.add(dist);
        if (div && dist) {
          if (!divMap[div]) divMap[div] = [];
          if (!divMap[div].includes(dist)) divMap[div].push(dist);
        }
      });
      const dists = [...distSet].sort(_numSort);
      return {
        userDivisions:       [...divSet].sort(),
        userDistricts:       dists,
        seniorDistrictScope: dists,
        divDistrictMap:      divMap
      };
    }

    const divSet = new Set(), distSet = new Set(), divMap = {};
    data.forEach(r => {
      const feId = String(r[3] || '').trim().toLowerCase();
      if (!id || feId !== id) return;
      const div  = _normDiv(String(r[10] || '').trim());
      const dist = _normDist(String(r[11] || '').trim());
      if (div)  divSet.add(div);
      if (dist) distSet.add(dist);
      if (div && dist) {
        if (!divMap[div]) divMap[div] = [];
        if (!divMap[div].includes(dist)) divMap[div].push(dist);
      }
    });
    return {
      userDivisions:       [...divSet].sort(),
      userDistricts:       [...distSet].sort(_numSort),
      seniorDistrictScope: [],
      divDistrictMap:      divMap
    };
  } catch(e) {
    Logger.log('_parseOrgStructure error: ' + e.message);
    return DEFAULT;
  }
}

function _buildFullDivDistrictMap() {
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return { divDistrictMap: {}, userDivisions: [], userDistricts: [] };
    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, 12).getValues();
    const divMap = {};
    data.forEach(r => {
      const div  = _normDiv(String(r[10] || '').trim());
      const dist = _normDist(String(r[11] || '').trim());
      if (!div || !dist) return;
      if (!divMap[div]) divMap[div] = [];
      if (!divMap[div].includes(dist)) divMap[div].push(dist);
    });
    const allDivs  = Object.keys(divMap).sort();
    const allDists = [...new Set(Object.values(divMap).flat())].sort(_numSort);
    return { divDistrictMap: divMap, userDivisions: allDivs, userDistricts: allDists };
  } catch(e) {
    return { divDistrictMap: {}, userDivisions: [], userDistricts: [] };
  }
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
function loginUser(empId, password) {
  try {
    const sh = _ss().getSheetByName(SH_USERS);
    if (!sh) return { ok: false, error: 'Users sheet not found.' };
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, error: 'No users registered.' };

    const data = sh.getRange(2, 1, last - 1, 9).getValues();
    for (let ri = 0; ri < data.length; ri++) {
      const row   = data[ri];
      const rowId = String(row[2] || '').trim();
      if (rowId.toLowerCase() !== String(empId).trim().toLowerCase()) continue;

      const role     = String(row[0] || 'User').trim();
      const pwd      = String(row[1] || '').trim();
      const empName  = String(row[3] || '').trim();
      const desig    = String(row[4] || '').trim();
      const supId    = String(row[5] || '').trim();
      const supName  = String(row[6] || '').trim();
      const supDesig = String(row[7] || '').trim();
      const remarks  = String(row[8] || '').trim();

      const inputHash = _hashPwd(password);
      if (_isHashed(pwd)) {
        if (inputHash !== pwd) return { ok: false, error: 'Incorrect password.' };
      } else {
        if (String(password) !== pwd) return { ok: false, error: 'Incorrect password.' };
        sh.getRange(ri + 2, 2).setValue(inputHash);
      }

      const firstLogin  = (pwd === '1234' || pwd === _hashPwd('1234'));
      const roleTier    = _mapRoleTier(role);
      const perms       = _parseRemarks(remarks, roleTier);
      const BOTH_ROLES = ['super admin', 'superadmin', 'super user', 'superuser'];
      const fallbackScope = BOTH_ROLES.includes(role.trim().toLowerCase()) ? 'both'
        : (roleTier === 'ho' ? 'ho' : 'field');
      const scopeType = perms.scopeType !== null ? perms.scopeType : fallbackScope;
      const scopeData   = _parseOrgStructure(rowId, roleTier);
      const mlData      = _getMasterlistEntry(rowId);

      return {
        ok: true, username: rowId, role, roleTier,
        name:                empName || mlData.name || rowId,
        designation:         desig   || mlData.position || '',
        supervisorId:        supId,
        supervisorName:      supName,
        supervisorDesig:     supDesig,
        remarks, perms, firstLogin,
        scopeType,
        division:            scopeData.userDivisions[0]  || mlData.division || '',
        district:            scopeData.userDistricts[0]  || mlData.district || '',
        userDivisions:       scopeData.userDivisions,
        userDistricts:       scopeData.userDistricts.length > 0 ? scopeData.userDistricts : (mlData.district ? [mlData.district] : []),
        seniorDistrictScope: scopeData.seniorDistrictScope,
        divDistrictMap:      scopeData.divDistrictMap,
        headOfficeDepts: [], area: mlData.area || '', branch: mlData.baseOffice || ''
      };
    }
    return { ok: false, error: 'Employee ID not found.' };
  } catch(e) { return { ok: false, error: e.message }; }
}

function changePassword(empId, newPwd) {
  try {
    if (!newPwd || newPwd.length < 4) return { ok: false, error: 'Minimum 4 characters.' };
    const sh   = _ss().getSheetByName(SH_USERS);
    if (!sh) return { ok: false, error: 'Users sheet not found.' };
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, error: 'No users found.' };
    const ids = sh.getRange(2, 1, last - 1, 3).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][2] || '').trim().toLowerCase() === String(empId).trim().toLowerCase()) {
        sh.getRange(i + 2, 2).setValue(_hashPwd(newPwd));
        return { ok: true };
      }
    }
    return { ok: false, error: 'Employee ID not found.' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ─── BARCODE GENERATION ───────────────────────────────────────────────────────
function generateNextBarcode(type) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const PFX = {
      'Laptop':'LTP','Laptop Adaptor':'LAD','CPU':'CPU',
      'Monitor':'MTR','Printer':'PTR','Scanner':'SCN',
      'Scansnap':'SCN','Keyboard':'KBD','Mouse':'MSE',
      'UPS':'UPS','Camera':'CAM','Speaker':'SPR','External Drive':'EXD'
    };
    const pre     = PFX[type] || 'AST';
    const yr      = new Date().getFullYear();
    const sh      = _entrySheet();
    const last    = sh.getLastRow();
    let   max     = 0;

    if (last >= AE_DATA_START) {
      const barcodes = sh.getRange(
        AE_DATA_START, C.BARCODE, last - AE_DATA_START + 1, 1).getValues();
      const pattern  = new RegExp('^' + pre + '-' + yr + '-(\\d+)$');
      barcodes.forEach(r => {
        const bc = String(r[0] || '').trim();
        const m  = bc.match(pattern);
        if (m) { const n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
      });
    }

    let seq = max + 1;
    let cand = pre + '-' + yr + '-' + String(seq).padStart(3, '0');
    while (_findRow(sh, cand) > 0) { seq++; cand = pre + '-' + yr + '-' + String(seq).padStart(3, '0'); }
    return cand;
  } finally { lock.releaseLock(); }
}

// ─── ASSETS: READ ─────────────────────────────────────────────────────────────
function getAllAssets() {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return { success: true, data: [] };

    const data   = sh.getRange(AE_DATA_START, 1, last - AE_DATA_START + 1, 37).getValues();
    const result = data
      .map((row, i) => [row, i + AE_DATA_START])          // pin sheet row BEFORE filter
      .filter(([row]) => {
        const bc = String(row[C.BARCODE - 1] || '').trim();
        const badBc = !bc || bc === '-' || bc === 'N/A' || bc === 'None' || bc === '#N/A';
        if (!badBc) return true;
        return [C.DIVISION,C.DISTRICT,C.BRANCH,C.STAFF,C.EMP_ID,C.DEPARTMENT,C.BASE_OFFICE,C.ASSET_LOCATION]
          .some(col => String(row[col - 1] || '').trim());
      })
      .map(([row, sheetRow]) => {                          // sheetRow is always correct now
        const get = col => String(row[col - 1] || '');
        const approvalStatus = get(CAE.APPROVAL_STATUS) || 'Confirmed';
        const grandfathered  = String(get(CAE.GRANDFATHERED)).toLowerCase() === 'true';
        const effectiveStatus = (approvalStatus === 'Confirmed' || grandfathered)
          ? _computeStatus(get(C.LIFECYCLE), get(C.ASSET_STATUS), get(C.EMP_ID))
          : 'pending-approval';

        const rawBc = String(row[C.BARCODE - 1] || '').trim();
        const badBc = !rawBc || rawBc === '-' || rawBc === 'N/A' || rawBc === 'None' || rawBc === '#N/A';
        const syntheticKey = badBc ? ('NOBC-' + sheetRow) : rawBc;  // correct row!

        const rawLC = get(C.LIFECYCLE);
        const displayLC = rawLC || {
          'allocated':'Allocated','spare':'Active','borrowed':'Borrow',
          'returned':'Returned','disposal':'Dispose','transfer':'Transfer',
          'borrow-item':'BorrowItem'
        }[effectiveStatus] || 'Active';

        const rawAssignment = get(C.ASSIGNMENT).trim();
        let effectiveAssignment = rawAssignment;
        if (!rawAssignment) {
          if (get(C.DEPARTMENT) || get(C.BASE_OFFICE)) effectiveAssignment = 'Central Office';
          else if (get(C.DIVISION) || get(C.DISTRICT) || get(C.BRANCH)) effectiveAssignment = 'Field Office';
          else effectiveAssignment = 'Unknown';
        }

        return {
          Barcode: badBc ? syntheticKey : rawBc,
          hasRealBarcode: !badBc,
          Type: get(C.TYPE), Brand: get(C.BRAND), Serial: get(C.SERIAL),
          Specs: get(C.SPECS), Supplier: get(C.SUPPLIER),
          Condition: get(C.CONDITION) || 'Good',
          AssetLocation: get(C.ASSET_LOCATION),
          Lifecycle: displayLC, AssetStatus: get(C.ASSET_STATUS) || 'Active',
          StatusLabel: get(C.STATUS_LABEL) || 'Unassigned',
          jsApprovalStatus: approvalStatus,
          FormID: get(CAE.FORM_ID) || '', Grandfathered: grandfathered,
          PurchDate: get(C.PURCH_DATE), WarrantyTerm: get(C.WARRANTY_TERM),
          WarrantyVal: get(C.WARRANTY_VAL), Remarks: get(C.REMARKS),
          EmpID: get(C.EMP_ID) || 'N/A', Staff: get(C.STAFF) || 'Unassigned',
          Designation: get(C.DESIGNATION), Department: get(C.DEPARTMENT),
          BaseOffice: get(C.BASE_OFFICE),
          Assignment: rawAssignment, EffectiveAssignment: effectiveAssignment,
          Division: _normDiv(get(C.DIVISION)), District: _normDist(get(C.DISTRICT)),
          Area: get(C.AREA), Branch: get(C.BRANCH), EffDate: get(C.EFF_DATE),
          CreatedAt: get(C.CREATED_AT), LastUpdated: get(C.LAST_UPDATED),
          EntryEmpId: get(C.ENTRY_EMP_ID), EntryName: get(C.ENTRY_NAME),
          BorName:'', BorEmpID:'', BorDesig:'', BorDiv:'', BorDist:'',
          BorBranch:'', BorDate:'', ExpReturn:'', ActReturn:'', BorRemarks:'',
          XferType:'', ToStaff:'', ToEmpID:'', ToDiv:'', ToBranch:'', XferDate:'',
          status: effectiveStatus
        };
      });

    return { success: true, data: result };
  } catch(e) { return { success: false, error: e.message }; }
}

// ─── GET ASSET BY BARCODE ─────────────────────────────────────────────────────
function getAssetByBarcode(barcode) {
  try {
    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return null;
    const row    = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
    const get    = col => String(row[col - 1] || '');
    return {
      barcode:          get(C.BARCODE), type:        get(C.TYPE),
      brand:            get(C.BRAND),   serial:      get(C.SERIAL),
      specs:            get(C.SPECS),   supplier:    get(C.SUPPLIER),
      condition:        get(C.CONDITION) || 'Good',
      lifecycle:        get(C.LIFECYCLE) || 'Active',
      statusLabel:      get(C.STATUS_LABEL) || 'Unassigned',
      purchaseDate:     get(C.PURCH_DATE),
      warrantyTerm:     get(C.WARRANTY_TERM),
      warrantyValidity: get(C.WARRANTY_VAL),
      remarks:          get(C.REMARKS),
      employeeId:       get(C.EMP_ID) || 'N/A',
      staff:            get(C.STAFF)  || 'Unassigned',
      designation:      get(C.DESIGNATION),
      department:       get(C.DEPARTMENT),
      baseOffice:       get(C.BASE_OFFICE),
      division:         _normDiv(get(C.DIVISION)),
      district:         _normDist(get(C.DISTRICT)),
      area:             get(C.AREA),
      branch:           get(C.BRANCH),
      effDate:          get(C.EFF_DATE),
      assetLocation:    get(C.ASSET_LOCATION),
      status:           _computeStatus(get(C.LIFECYCLE), get(C.ASSET_STATUS), get(C.EMP_ID))
    };
  } catch(e) { return null; }
}

// ─── DELETE ASSETS ────────────────────────────────────────────────────────────
function deleteAssets(barcodes, callerEmpId) {
  try {
    const userSh = _ss().getSheetByName(SH_USERS);
    if (!userSh || !callerEmpId) return 'Error: Unauthorized — caller ID required.';
    const last = userSh.getLastRow();
    if (last < 2) return 'Error: Unauthorized.';
    const rows = userSh.getRange(2, 1, last - 1, 3).getValues();
    const caller = rows.find(r =>
      String(r[2] || '').trim().toLowerCase() === String(callerEmpId).trim().toLowerCase()
    );
    if (!caller) return 'Error: Unauthorized — identity not verified.';
    const role = String(caller[0] || '').toLowerCase();
    if (!role.includes('supervisor') && !role.includes('admin') && !role.includes('super'))
      return 'Error: Unauthorized — insufficient permissions.';
  } catch(e) { return 'Error: ' + e.message; }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return 'Error: System busy — try again.'; }

  try {
    const sh = _entrySheet();
    const blocked = [], toDelete = [];

    barcodes.forEach(bc => {
      const r = _findRow(sh, bc);
      if (r > 0) {
        const lc = String(sh.getRange(r, C.LIFECYCLE).getValue() || '').toLowerCase();
        if (lc === 'borrow' || lc === 'transfer' || lc === 'borrowitem')
          blocked.push(bc + ' (' + lc + ')');
        else
          toDelete.push({ bc, r });
      }
    });

    if (blocked.length && !toDelete.length)
      return 'Error: Cannot delete — active lifecycle: ' + blocked.join(', ');

    toDelete.sort((a, b) => b.r - a.r).forEach(({ bc, r }) => {
      sh.deleteRow(r);
      _log('DELETE', bc, 'Deleted by ' + callerEmpId, callerEmpId);
    });

    let msg = 'Deleted ' + toDelete.length + ' asset(s)';
    if (blocked.length) msg += '. Skipped ' + blocked.length + ': ' + blocked.join(', ');
    return msg;
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

// ─── FIX SPARE SHEET HEADERS ──────────────────────────────────────────────────
// Run this once if the Spare sheet headers were misaligned (data pushed down)
function fixSpareSheetHeaders() {
  const sh = _ss().getSheetByName(SH_SPARE);
  if (!sh) return 'Spare sheet not found';
  const headers = [
    'Barcode','Category','Brand','Serial No.','Condition',
    'Purchase Date','Warranty Validity','Supplier',
    'Division','District','Area','Branch','Asset Location',
    'Enrolled By','Timestamp','Status'
  ];
  // Check if row 1 already has headers
  const row1 = sh.getRange(1,1,1,headers.length).getValues()[0];
  if (row1[0] === 'Barcode') return 'Headers already correct';
  // Insert a row at top and add headers
  sh.insertRowBefore(1);
  sh.getRange(1,1,1,headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#0f0e1c')
    .setFontColor('#a07ee0');
  sh.setFrozenRows(1);
  return 'Headers fixed — ' + headers.length + ' columns';
}

// ─── ALLOCATE ─────────────────────────────────────────────────────────────────
// Now drafts an accountability form after successful allocation
function allocateAsset(obj) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return { result: 'Error: System busy — try again.', formId: null }; }
  try {
    if (!obj.barcode)   return { result: 'Error: Barcode is required.', formId: null };
    if (!obj.empId)     return { result: 'Error: Employee ID is required.', formId: null };
    if (!obj.staffName) return { result: 'Error: Staff name is required.', formId: null };

    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, obj.barcode);
    if (rowIdx < 1) return { result: 'Error: Asset not found: ' + obj.barcode, formId: null };

    const curRow = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
    const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
    const curAS  = String(curRow[C.ASSET_STATUS - 1] || '').toLowerCase();

    if (curLC === 'borrow')    return { result: 'Error: Asset is currently borrowed. Return it first.', formId: null };
    if (curLC === 'dispose' || curLC === 'disposal')
                               return { result: 'Error: Disposed assets cannot be allocated.', formId: null };
    if (curLC === 'transfer')  return { result: 'Error: Asset is in an active transfer.', formId: null };
    if (curAS === 'borrowitem')
      return { result: 'Error: This asset is in the Borrow Pool. Change its status first.', formId: null };

    if (curLC === 'allocated') {
      const prevStaff = String(curRow[C.STAFF - 1] || '');
      _log('DEALLOCATE', obj.barcode,
        'Implicit dealloc from: ' + prevStaff + ' → re-allocate to: ' + obj.staffName,
        obj.allocatedBy || obj.empId || '');
    }

    const nowStr   = new Date().toLocaleString('en-PH');
    const normDiv  = _normDiv(obj.division  || '');
    const normDist = _normDist(obj.district || '');

    const updates = [
      [C.LIFECYCLE,    'Allocated'],
      [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],
      [C.EMP_ID,       obj.empId           || ''],
      [C.STAFF,        _sanitize(obj.staffName, 100)],
      [C.DESIGNATION,  obj.designation     || ''],
      [C.DEPARTMENT,   obj.department      || ''],
      [C.BASE_OFFICE,  obj.baseOffice      || ''],
      [C.DIVISION,     normDiv],
      [C.DISTRICT,     normDist],
      [C.AREA,         obj.area            || ''],
      [C.BRANCH,       _sanitize(obj.branch, 150)],
      [C.EFF_DATE,     obj.effDate         || nowStr],
      [C.REMARKS,      _sanitize(obj.remarks, 500)],
    ];
    if (obj.condition) updates.push([C.CONDITION, obj.condition]);
    if (obj.assetLocation) updates.push([C.ASSET_LOCATION, obj.assetLocation]);
    _setRow(sh, rowIdx, updates);

    // Write to Allocated log
    _allocLogSheet().appendRow([
      obj.barcode,
      obj.type      || String(curRow[C.TYPE      - 1] || ''),
      obj.brand     || String(curRow[C.BRAND     - 1] || ''),
      obj.serial    || String(curRow[C.SERIAL    - 1] || ''),
      obj.empId, _sanitize(obj.staffName, 100),
      obj.designation  || '', obj.department || '', obj.baseOffice || '',
      normDiv, normDist, obj.area || '', _sanitize(obj.branch, 150),
      obj.effDate || nowStr,
      obj.condition || String(curRow[C.CONDITION - 1] || 'Good'),
      obj.assetLocation || '',
      _sanitize(obj.remarks, 500), nowStr, obj.allocatedBy || ''
    ]);

    _log('ALLOCATE', obj.barcode,
      _sanitize(obj.staffName, 100) + ' | ' + (_sanitize(obj.branch, 150) || normDiv || ''),
      obj.allocatedBy || obj.empId || '');

    // Draft accountability form for the newly allocated asset
    const assetForForm = [{
      Barcode:     obj.barcode,
      Type:        obj.type        || String(curRow[C.TYPE      - 1] || ''),
      Brand:       obj.brand       || String(curRow[C.BRAND     - 1] || ''),
      Serial:      obj.serial      || String(curRow[C.SERIAL    - 1] || ''),
      Specs:       obj.specs       || String(curRow[C.SPECS     - 1] || ''),
      Condition:   obj.condition   || String(curRow[C.CONDITION - 1] || 'Good'),
      Staff:       _sanitize(obj.staffName, 100),
      EmpID:       obj.empId,
      Designation: obj.designation || '',
      Department:  obj.department  || '',
      BaseOffice:  obj.baseOffice  || '',
      Branch:      _sanitize(obj.branch, 150),
      Division:    normDiv,
      District:    normDist,
      Area:        obj.area        || ''
    }];

    const drafterId = obj.allocatedBy || obj.empId || '';
    const formId = draftAccountabilityForm(
      obj.empId,
      assetForForm,
      'Allocation',
      '',
      drafterId
    );

    // draftAccountabilityForm already writes APPROVAL_STATUS, FORM_ID, DRAFTED_BY, DRAFTED_AT
    // to the asset row internally. No duplicate writes needed here.

    return {
      result: 'Asset allocated to ' + _sanitize(obj.staffName, 100),
      formId: formId.startsWith('Error') ? null : formId
    };
  } catch(e) { return { result: 'Error: ' + e.message, formId: null }; }
  finally    { lock.releaseLock(); }
}

// ─── DEALLOCATE ───────────────────────────────────────────────────────────────
function deallocateAsset(barcode, remarks) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;

    const curRow = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
    const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
    if (curLC === 'borrow')  return 'Error: Return the borrow record first.';
    if (curLC === 'dispose' || curLC === 'disposal')
                             return 'Error: Disposed assets cannot be returned to spare.';

    const prevStaff = String(curRow[C.STAFF - 1] || '');
    const nowStr    = new Date().toLocaleString('en-PH');

    const updates = [
      [C.LIFECYCLE,   'Active'],   [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL,'Unassigned'],
      [C.EMP_ID,      ''],         [C.STAFF,         ''],
      [C.DESIGNATION, ''],         [C.EFF_DATE,      '']
    ];
    if (remarks) updates.push([C.REMARKS, _sanitize(remarks, 500)]);
    _setRow(sh, rowIdx, updates);

    _spareSheet().appendRow([
      String(curRow[C.EMP_ID      - 1] || ''),   // Col 1: Accountable Person ID
      String(curRow[C.STAFF       - 1] || ''),   // Col 2: Accountable Person
      'Spare',                                    // Col 3: Assignment
      String(curRow[C.DESIGNATION - 1] || ''),   // Col 4: Designation
      String(curRow[C.DEPARTMENT  - 1] || ''),   // Col 5: Department
      barcode,                                    // Col 6: Barcode
      String(curRow[C.TYPE          - 1] || ''), // Col 7: Category
      String(curRow[C.BRAND         - 1] || ''), // Col 8: Brand
      String(curRow[C.SERIAL        - 1] || ''), // Col 9: Serial No.
      String(curRow[C.CONDITION     - 1] || 'Good'), // Col 10: Condition
      String(curRow[C.PURCH_DATE    - 1] || ''), // Col 11: Purchase Date
      String(curRow[C.WARRANTY_VAL  - 1] || ''), // Col 12: Warranty Validity
      String(curRow[C.SUPPLIER      - 1] || ''), // Col 13: Supplier
      String(curRow[C.BASE_OFFICE   - 1] || ''), // Col 14: Base Office
      String(curRow[C.DIVISION      - 1] || ''), // Col 15: Division
      String(curRow[C.DISTRICT      - 1] || ''), // Col 16: District
      String(curRow[C.AREA          - 1] || ''), // Col 17: Area
      String(curRow[C.BRANCH        - 1] || ''), // Col 18: Branch
      String(curRow[C.ASSET_LOCATION- 1] || ''), // Col 19: Asset Location
      String(curRow[C.ENTRY_NAME - 1] || ''),    // Col 20: Original Enrollee (from Asset Entry)
      nowStr                                      // Col 21: Returned At
    ]);

    _log('DEALLOCATE', barcode,
      'From: ' + prevStaff + ' → Spare Pool. ' + (_sanitize(remarks, 500) || ''), '');
    return 'Asset returned to spare pool';
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

// ─── UPDATE ASSET DETAILS ─────────────────────────────────────────────────────
function updateAssetDetails(barcode, updates) {
  try {
    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;

    const fields = [];
    if (updates.brand     !== undefined) fields.push([C.BRAND,     updates.brand]);
    if (updates.condition !== undefined) fields.push([C.CONDITION, updates.condition]);
    if (updates.purchDate !== undefined) fields.push([C.PURCH_DATE,updates.purchDate]);
    if (updates.specs     !== undefined) fields.push([C.SPECS,     updates.specs]);
    if (updates.remarks   !== undefined) fields.push([C.REMARKS,   updates.remarks]);
    if (updates.supplier  !== undefined) fields.push([C.SUPPLIER,  updates.supplier]);

    if (updates.serial !== undefined) {
      const currentSerial = String(sh.getRange(rowIdx, C.SERIAL).getValue() || '').trim();
      if (updates.serial && updates.serial !== currentSerial) {
        const last = sh.getLastRow();
        if (last >= AE_DATA_START) {
          const serials = sh.getRange(
            AE_DATA_START, C.SERIAL, last - AE_DATA_START + 1, 1).getValues();
          const dupIdx = serials.findIndex(
            (r, i) => String(r[0]).trim() === updates.serial && (i + AE_DATA_START) !== rowIdx);
          if (dupIdx >= 0) return 'Error: Serial No. already used by another asset.';
        }
      }
      fields.push([C.SERIAL, updates.serial]);
    }

    _setRow(sh, rowIdx, fields);
    _log('EDIT', barcode, 'Updated: ' + Object.keys(updates).join(', '), '');
    return 'Asset details updated.';
  } catch(e) { return 'Error: ' + e.message; }
}

// ─── SINGLE TRANSFER ─────────────────────────────────────────────────────────
function saveTransfer(t) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return { result: 'Error: System busy — try again.', fromFormId: null, toFormId: null }; }

  try {
    if (!t.barcode)      return { result: 'Error: Barcode is required.', fromFormId: null, toFormId: null };
    if (!t.toEmpId)      return { result: 'Error: Destination Employee ID is required.', fromFormId: null, toFormId: null };
    if (!t.toStaff)      return { result: 'Error: Destination Staff Name is required.', fromFormId: null, toFormId: null };
    if (!t.effDate)      return { result: 'Error: Transfer Date is required.', fromFormId: null, toFormId: null };
    if (!t.transferType) return { result: 'Error: Transfer Type is required.', fromFormId: null, toFormId: null };

    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, t.barcode);
    if (rowIdx < 1) return { result: 'Error: Asset not found: ' + t.barcode, fromFormId: null, toFormId: null };

    const curRow = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
    const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
    if (curLC === 'dispose' || curLC === 'disposal')
      return { result: 'Error: Cannot transfer a disposed asset.', fromFormId: null, toFormId: null };
    if (curLC === 'borrow')
      return { result: 'Error: Return the asset first before transferring.', fromFormId: null, toFormId: null };

    const nowStr     = new Date().toLocaleString('en-PH');
    const normToDiv  = _normDiv(t.toDiv   || '');
    const normToDist = _normDist(t.toDist || '');

    _xferSheet().appendRow([
      t.barcode, t.transferType,
      _sanitize(t.fromStaff, 100), t.fromEmpId, t.fromDesig,
      t.fromDept || '', t.fromBaseOffice || '',
      t.fromDiv, t.fromDist, t.fromArea, _sanitize(t.fromBranch, 150),
      String(curRow[C.CONDITION - 1] || ''), // from condition (current state)
      String(curRow[C.ASSET_LOCATION - 1] || ''),
      _sanitize(t.fromRemarks, 500),
      _sanitize(t.toStaff, 100), t.toEmpId, t.toDesig,
      t.toDept || '', t.toBaseOffice || '',
      normToDiv, normToDist, t.toArea, _sanitize(t.toBranch, 150),
      t.toCondition || '', t.toAssetLocation || '',
      _sanitize(t.toRemarks, 500),
      t.effDate, t.status || 'Completed', nowStr
    ]);

    const curApprovalStatus = String(curRow[CAE.APPROVAL_STATUS - 1] || '').trim();
    const xferUpdates = [
      [C.LIFECYCLE,    'Allocated'],
      [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],
      [C.EMP_ID,       t.toEmpId      || ''],
      [C.STAFF,        _sanitize(t.toStaff, 100)],
      [C.DESIGNATION,  t.toDesig      || ''],
      [C.DEPARTMENT,   t.toDept       || ''],
      [C.BASE_OFFICE,  t.toBaseOffice || ''],
      [C.DIVISION,     normToDiv],
      [C.DISTRICT,     normToDist],
      [C.AREA,         t.toArea       || ''],
      [C.BRANCH,       _sanitize(t.toBranch, 150)],
      [C.EFF_DATE,     t.effDate],
      // Only reset approval status if asset is not already Confirmed
      ...(curApprovalStatus !== 'Confirmed' ? [
        [CAE.APPROVAL_STATUS,   'Draft'],
        [CAE.REJECTION_COMMENT, '']
      ] : [])
    ];
    if (t.toCondition)     xferUpdates.push([C.CONDITION,     t.toCondition]);
    if (t.toAssetLocation) xferUpdates.push([C.ASSET_LOCATION, t.toAssetLocation]);
    _setRow(sh, rowIdx, xferUpdates);

    const assetObj = {
      Barcode:   t.barcode,
      Type:      String(curRow[C.TYPE      - 1] || ''),
      Brand:     String(curRow[C.BRAND     - 1] || ''),
      Serial:    String(curRow[C.SERIAL    - 1] || ''),
      Specs:     String(curRow[C.SPECS     - 1] || ''),
      Condition: String(curRow[C.CONDITION - 1] || '')
    };

    const fromAsset = Object.assign({}, assetObj, {
      Staff:       t.fromStaff  || '', EmpID:    t.fromEmpId  || '',
      Designation: t.fromDesig  || '', Division: t.fromDiv    || '',
      District:    t.fromDist   || '', Branch:   t.fromBranch || ''
    });
    const toAsset = Object.assign({}, assetObj, {
      Staff:       t.toStaff    || '', EmpID:    t.toEmpId  || '',
      Designation: t.toDesig   || '', Division: normToDiv,
      District:    normToDist,         Branch:   t.toBranch || '',
      Condition:   t.toCondition || assetObj.Condition  // capture post-transfer condition
    });

    const drafterId  = t.fromEmpId || '';
    const fromFormId = draftAccountabilityForm(t.fromEmpId || '', [fromAsset], 'Transfer-From', '', drafterId);
    const toFormId   = draftAccountabilityForm(t.toEmpId   || '', [toAsset],   'Transfer-To',   fromFormId.startsWith('Error') ? '' : fromFormId, drafterId);

    if (!fromFormId.startsWith('Error') && !toFormId.startsWith('Error')) {
      sh.getRange(rowIdx, CAE.APPROVAL_STATUS).setValue('Draft');
      sh.getRange(rowIdx, CAE.REJECTION_COMMENT).setValue('');
      const afSh     = _afSheet();
      const fromRowI = _findAFRow(fromFormId);
      if (fromRowI > 0) afSh.getRange(fromRowI, AF.LINKED_FORM_ID).setValue(toFormId);
      // Link formId to asset row
      sh.getRange(rowIdx, CAE.FORM_ID).setValue(toFormId);
      sh.getRange(rowIdx, CAE.DRAFTED_BY).setValue(drafterId);
      sh.getRange(rowIdx, CAE.DRAFTED_AT).setValue(nowStr);
    }

    _log('TRANSFER', t.barcode,
      (_sanitize(t.fromStaff, 100) || '—') + ' → ' + _sanitize(t.toStaff, 100),
      t.fromEmpId || '');

    return {
      result:     'Transfer saved',
      fromFormId: fromFormId.startsWith('Error') ? null : fromFormId,
      toFormId:   toFormId.startsWith('Error')   ? null : toFormId
    };
  } catch(e) {
    return { result: 'Error: ' + e.message, fromFormId: null, toFormId: null };
  } finally {
    lock.releaseLock();
  }
}

// ─── BULK TRANSFER ────────────────────────────────────────────────────────────
// All barcodes come from 1 from-person, going to 1 to-person.
// Creates 1 Transfer-From form + 1 Transfer-To form covering all assets.
function saveBulkTransfer(barcodes, t) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(e) { return { result: 'Error: System busy — try again.', fromFormId: null, toFormId: null }; }

  try {
    if (!barcodes || !barcodes.length) return { result: 'Error: No assets selected.', fromFormId: null, toFormId: null };
    if (!t.toEmpId)      return { result: 'Error: Destination Employee ID is required.', fromFormId: null, toFormId: null };
    if (!t.toStaff)      return { result: 'Error: Destination Staff Name is required.', fromFormId: null, toFormId: null };
    if (!t.effDate)      return { result: 'Error: Transfer Date is required.', fromFormId: null, toFormId: null };
    if (!t.transferType) return { result: 'Error: Transfer Type is required.', fromFormId: null, toFormId: null };

    const sh        = _entrySheet();
    const nowStr    = new Date().toLocaleString('en-PH');
    const normToDiv = _normDiv(t.toDiv   || '');
    const normToDist= _normDist(t.toDist || '');
    const fromAssets = [], toAssets = [];
    const failed = [];

    barcodes.forEach(bc => {
      const rowIdx = _findRow(sh, bc);
      if (rowIdx < 1) { failed.push(bc + ' (not found)'); return; }

      const curRow = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
      const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
      if (curLC === 'dispose' || curLC === 'disposal') { failed.push(bc + ' (disposed)'); return; }
      if (curLC === 'borrow')  { failed.push(bc + ' (on borrow)'); return; }

      const assetObj = {
        Barcode:   bc,
        Type:      String(curRow[C.TYPE      - 1] || ''),
        Brand:     String(curRow[C.BRAND     - 1] || ''),
        Serial:    String(curRow[C.SERIAL    - 1] || ''),
        Specs:     String(curRow[C.SPECS     - 1] || ''),
        Condition: String(curRow[C.CONDITION - 1] || '')
      };

      fromAssets.push(Object.assign({}, assetObj, {
        Staff:       t.fromStaff  || '', EmpID:    t.fromEmpId  || '',
        Designation: t.fromDesig  || '', Division: t.fromDiv    || '',
        District:    t.fromDist   || '', Branch:   t.fromBranch || ''
      }));

      toAssets.push(Object.assign({}, assetObj, {
        Staff:       t.toStaff || '', EmpID:    t.toEmpId  || '',
        Designation: t.toDesig || '', Division: normToDiv,
        District:    normToDist,      Branch:   t.toBranch || ''
      }));

      // Write individual transfer record to Transfers sheet
      _xferSheet().appendRow([
        bc,                                       // col 1:  Barcode
        t.transferType,                           // col 2:  TransferType
        _sanitize(t.fromStaff, 100),              // col 3:  FromStaff
        t.fromEmpId,                              // col 4:  FromEmpID
        t.fromDesig   || '',                      // col 5:  FromDesig
        t.fromDiv     || '',                      // col 6:  FromDept  (reuse div as dept placeholder)
        t.fromBaseOffice || t.fromBranch || '',   // col 7:  FromBaseOffice
        t.fromDiv     || '',                      // col 8:  FromDiv
        t.fromDist    || '',                      // col 9:  FromDist
        t.fromArea    || '',                      // col 10: FromArea
        _sanitize(t.fromBranch, 150),             // col 11: FromBranch
        '',                                       // col 12: FromCondition
        '',                                       // col 13: FromAssetLoc
        _sanitize(t.fromRemarks, 500),            // col 14: FromRemarks
        _sanitize(t.toStaff, 100),                // col 15: ToStaff
        t.toEmpId,                                // col 16: ToEmpID
        t.toDesig     || '',                      // col 17: ToDesig
        t.toDept      || '',                      // col 18: ToDept
        t.toBaseOffice || t.toBranch || '',       // col 19: ToBaseOffice
        normToDiv,                                // col 20: ToDiv
        normToDist,                               // col 21: ToDist
        t.toArea      || '',                      // col 22: ToArea
        _sanitize(t.toBranch, 150),               // col 23: ToBranch
        '',                                       // col 24: ToCondition
        '',                                       // col 25: ToAssetLoc
        _sanitize(t.toRemarks, 500),              // col 26: ToRemarks
        t.effDate,                                // col 27: EffDate
        'Completed',                              // col 28: Status
        nowStr                                    // col 29: Timestamp
      ]);

      // Update asset row
      const bulkCurApproval = String(curRow[CAE.APPROVAL_STATUS - 1] || '').trim();
      const bulkSetRowUpdates = [
        [C.LIFECYCLE,    'Allocated'],  [C.ASSET_STATUS, 'Active'],
        [C.STATUS_LABEL, 'Assigned'],   [C.EMP_ID,       t.toEmpId   || ''],
        [C.STAFF,        _sanitize(t.toStaff, 100)],
        [C.DESIGNATION,  t.toDesig   || ''],
        [C.DIVISION,     normToDiv],    [C.DISTRICT,     normToDist],
        [C.AREA,         t.toArea    || ''],
        [C.BRANCH,       _sanitize(t.toBranch, 150)],
        [C.EFF_DATE,     t.effDate],
        ...(bulkCurApproval !== 'Confirmed' ? [
          [CAE.APPROVAL_STATUS,   'Draft'],
          [CAE.REJECTION_COMMENT, '']
        ] : [])
      ];
      _setRow(sh, rowIdx, bulkSetRowUpdates);
    });

    if (!fromAssets.length) {
      return { result: 'Error: No valid assets to transfer. Skipped: ' + failed.join(', '), fromFormId: null, toFormId: null };
    }

    // One shared form pair for all assets
    const drafterId  = t.fromEmpId || '';
    // Pre-generate both IDs sequentially while still under lock to prevent
    // race between sheet read and append in draftAccountabilityForm
    const fromFormId = _generateFormIDUnsafe();
    // Append From form row directly to reserve the ID slot before generating To ID
    const _nowForIds = new Date().toLocaleString('en-PH');
    const _fromContext = _getContextType(t.fromEmpId || '');
    const _fromAssetsJson = _buildAssetsSnapshot(fromAssets);
    const _fromRef = fromAssets.length ? fromAssets[0] : {};
    const _fromRow = new Array(AF_TOTAL_COLS).fill('');
    _fromRow[AF.FORM_ID - 1]      = fromFormId;
    _fromRow[AF.FORM_TYPE - 1]    = 'Transfer-From';
    _fromRow[AF.CONTEXT_TYPE - 1] = _fromContext;
    _fromRow[AF.EMP_ID - 1]       = t.fromEmpId || '';
    _fromRow[AF.STAFF_NAME - 1]   = _sanitize(_fromRef.Staff || '', 100);
    _fromRow[AF.DESIGNATION - 1]  = _fromRef.Designation || '';
    _fromRow[AF.DEPARTMENT - 1]   = _fromRef.Department  || '';
    _fromRow[AF.BRANCH - 1]       = _sanitize(_fromRef.Branch || '', 150);
    _fromRow[AF.DIVISION - 1]     = _fromRef.Division || '';
    _fromRow[AF.DISTRICT - 1]     = _fromRef.District || '';
    _fromRow[AF.ASSETS_JSON - 1]  = _fromAssetsJson;
    _fromRow[AF.STATUS - 1]       = 'Draft';
    _fromRow[AF.DRAFTED_BY - 1]   = drafterId;
    _fromRow[AF.DRAFTED_AT - 1]   = _nowForIds;
    _afSheet().appendRow(_fromRow);
    _log('DRAFT_FORM', fromFormId, 'Transfer-From | ' + (t.fromEmpId || '') + ' | ' + fromAssets.length + ' assets', drafterId);

    const toFormId = _generateFormIDUnsafe();
    const _toContext = _getContextType(t.toEmpId || '');
    const _toAssetsJson = _buildAssetsSnapshot(toAssets);
    const _toRef = toAssets.length ? toAssets[0] : {};
    const _toRow = new Array(AF_TOTAL_COLS).fill('');
    _toRow[AF.FORM_ID - 1]        = toFormId;
    _toRow[AF.FORM_TYPE - 1]      = 'Transfer-To';
    _toRow[AF.LINKED_FORM_ID - 1] = fromFormId;
    _toRow[AF.CONTEXT_TYPE - 1]   = _toContext;
    _toRow[AF.EMP_ID - 1]         = t.toEmpId || '';
    _toRow[AF.STAFF_NAME - 1]     = _sanitize(_toRef.Staff || '', 100);
    _toRow[AF.DESIGNATION - 1]    = _toRef.Designation || '';
    _toRow[AF.DEPARTMENT - 1]     = _toRef.Department  || '';
    _toRow[AF.BRANCH - 1]         = _sanitize(_toRef.Branch || '', 150);
    _toRow[AF.DIVISION - 1]       = _toRef.Division || '';
    _toRow[AF.DISTRICT - 1]       = _toRef.District || '';
    _toRow[AF.ASSETS_JSON - 1]    = _toAssetsJson;
    _toRow[AF.STATUS - 1]         = 'Draft';
    _toRow[AF.DRAFTED_BY - 1]     = drafterId;
    _toRow[AF.DRAFTED_AT - 1]     = _nowForIds;
    _afSheet().appendRow(_toRow);
    _log('DRAFT_FORM', toFormId, 'Transfer-To | ' + (t.toEmpId || '') + ' | ' + toAssets.length + ' assets', drafterId);

    if (!fromFormId.startsWith('Error') && !toFormId.startsWith('Error')) {
      const afSh     = _afSheet();
      const fromRowI = _findAFRow(fromFormId);
      if (fromRowI > 0) afSh.getRange(fromRowI, AF.LINKED_FORM_ID).setValue(toFormId);

// Link fromFormId to releasing-party assets, toFormId to receiving-party assets
    _updateAssetApprovalStatus(fromFormId, 'Draft', '');

    barcodes.forEach(function(bc) {
      const rowIdx = _findRow(sh, bc);
      if (rowIdx < 1) return;
      // After transfer the asset row belongs to the new owner; link toFormId
      sh.getRange(rowIdx, CAE.FORM_ID).setValue(toFormId);
      sh.getRange(rowIdx, CAE.DRAFTED_BY).setValue(drafterId);
      sh.getRange(rowIdx, CAE.DRAFTED_AT).setValue(nowStr);
    });
    }

    // Log one entry per successfully transferred barcode so Activity Log scoping works
    fromAssets.forEach(function(a) {
      _log('BULK_TRANSFER', a.Barcode,
        (_sanitize(t.fromStaff, 100) || '—') + ' → ' + _sanitize(t.toStaff, 100),
        t.fromEmpId || '');
    });

    let msg = 'Bulk transfer saved: ' + fromAssets.length + ' asset(s).';
    if (failed.length) msg += ' Skipped: ' + failed.join(', ');

    return {
      result:     msg,
      fromFormId: fromFormId.startsWith('Error') ? null : fromFormId,
      toFormId:   toFormId.startsWith('Error')   ? null : toFormId
    };
  } catch(e) {
    return { result: 'Error: ' + e.message, fromFormId: null, toFormId: null };
  } finally {
    lock.releaseLock();
  }
}

function getTransferData() {
  try {
    const sh   = _xferSheet();
    const last = sh.getLastRow();
    if (last < EVT_DATA_START) return [];
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 29).getValues()
      .filter(r => r[0])
      .map(r => r.map(v => String(v || '')));
  } catch(e) { return []; }
}

// ─── BORROWS ──────────────────────────────────────────────────────────────────
function saveBorrow(b) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const sh     = _entrySheet();
    const bSh    = _borrowSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const rowIdx = _findRow(sh, b.barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + b.barcode;

    const curRow  = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
    const curLC   = String(curRow[C.LIFECYCLE    - 1] || '').toLowerCase();
    const curAS   = String(curRow[C.ASSET_STATUS - 1] || '').toLowerCase();
    const isBItem = curAS === 'borrowitem';

    const borrowable = ['active', 'allocated', 'borrowitem', 'spare'];
    if (!borrowable.includes(curLC)) {
      const labels = {
        borrow:'already on borrow', returned:'Returned — re-allocate first',
        transfer:'in active Transfer', dispose:'Disposed', disposal:'Disposed'
      };
      return 'Error: Cannot borrow. Status: ' + (labels[curLC] || curLC);
    }

    bSh.appendRow([
      b.barcode, _sanitize(b.borrowerName, 100), b.empId || '', b.designation || '',
      b.baseOffice || b.branch || '',          // Col 5: Base Office
      b.division || '', b.district || '', _sanitize(b.branch, 150),
      b.borrowDate, b.expectedReturn, '',
      'Borrow', _sanitize(b.remarks, 500), nowStr
    ]);

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Borrow'],
      [C.ASSET_STATUS, isBItem ? 'BorrowItem' : 'Active'],
      [C.STATUS_LABEL, 'Assigned']
    ]);

    _log('BORROW', b.barcode,
      _sanitize(b.borrowerName, 100) + ' | due: ' + b.expectedReturn,
      b.empId || '');
    return 'Borrow saved';
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

function getBorrowData() {
  try {
    const sh   = _borrowSheet();
    const last = sh.getLastRow();
    if (last < EVT_DATA_START) return [];
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 14).getValues()
      .filter(r => r[0])
      .map(r => ({
        barcode:        String(r[0]  || ''),
        borrowerName:   String(r[1]  || ''),
        empId:          String(r[2]  || ''),
        designation:    String(r[3]  || ''),
        baseOffice:     String(r[4]  || ''),   // Col 5: Base Office (new)
        division:       String(r[5]  || ''),   // Col 6
        district:       String(r[6]  || ''),   // Col 7
        branch:         String(r[7]  || ''),   // Col 8
        borrowDate:     String(r[8]  || ''),   // Col 9
        expectedReturn: String(r[9]  || ''),   // Col 10
        actualReturn:   String(r[10] || ''),   // Col 11
        status:         String(r[11] || 'Borrow'), // Col 12
        remarks:        String(r[12] || ''),   // Col 13
        timestamp:      String(r[13] || '')    // Col 14
      }));
  } catch(e) { return []; }
}

function returnAsset(barcode, returnDate) {
  try {
    const sh      = _entrySheet();
    const bSh     = _borrowSheet();
    const retDate = returnDate || new Date().toLocaleDateString('en-PH');
    const last    = bSh.getLastRow();

    if (last >= EVT_DATA_START) {
      const data = bSh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 14).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === String(barcode) && String(data[i][11]) === 'Borrow') {
          const sheetRow = i + EVT_DATA_START;
          bSh.getRange(sheetRow, 11).setValue(retDate);   // ActualReturn at col 11
          bSh.getRange(sheetRow, 12).setValue('Returned'); // Status at col 12
          break;
        }
      }
    }

    const rowIdx = _findRow(sh, barcode);
    if (rowIdx > 0) {
      const curRow   = sh.getRange(rowIdx, 1, 1, 37).getValues()[0];
      const curAS    = String(curRow[C.ASSET_STATUS - 1] || '');
      const curStaff = String(curRow[C.STAFF - 1] || '').trim();
      const curEmpId = String(curRow[C.EMP_ID - 1] || '').trim();

      const isBItem  = curAS.toLowerCase() === 'borrowitem';
      const hasOwner = !isBItem && curStaff && curStaff !== 'Unassigned' && curEmpId;

      const restoredLC  = isBItem ? 'BorrowItem' : (hasOwner ? 'Allocated' : 'Active');
      const restoredLbl = (isBItem || !hasOwner)  ? 'Unassigned' : 'Assigned';
      const restoredAS  = isBItem ? 'BorrowItem'  : 'Active';

      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    restoredLC],
        [C.STATUS_LABEL, restoredLbl],
        [C.ASSET_STATUS, restoredAS]
      ]);
    }

    _log('RETURN', barcode, retDate, '');
    return 'Asset returned';
  } catch(e) { return 'Error: ' + e.message; }
}

// ─── DISPOSALS ────────────────────────────────────────────────────────────────
function saveDisposal(d) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const sh     = _entrySheet();
    const dSh    = _disposeSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const rowIdx = _findRow(sh, d.barcode);

    let serial = d.serial || '';
    let assetType = d.assetType || '';
    let assetLocation = d.location || '';

    if (rowIdx > 0) {
      const curRow = sh.getRange(rowIdx, 1, 1, 31).getValues()[0];
      const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
      if (curLC === 'borrow')   return 'Error: Cannot dispose a borrowed asset.';
      if (curLC === 'transfer') return 'Error: Cannot dispose an asset in active transfer.';
      if (!serial)        serial        = String(curRow[C.SERIAL        - 1] || '');
      if (!assetType)     assetType     = String(curRow[C.TYPE          - 1] || '');
      if (!assetLocation) assetLocation = String(curRow[C.ASSET_LOCATION- 1] || '');
    }

    dSh.appendRow([
      d.barcode,
      serial,                           // Col 2: Serial Number
      _sanitize(d.reason, 200),         // Col 3: Reason
      _sanitize(d.disposedBy, 100),     // Col 4: DisposedBy
      d.disposeDate,                    // Col 5: DisposeDate
      _sanitize(d.remarks, 500),        // Col 6: Remarks
      nowStr,                           // Col 7: Timestamp
      assetLocation,                    // Col 8: Location
      assetType,                        // Col 9: Equipment Type
      ''                                // Col 10: Description
    ]);

    if (rowIdx > 0) {
      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    'Dispose'],
        [C.ASSET_STATUS, 'Disposal'],
        [C.STATUS_LABEL, 'Disposed']
      ]);
    }

    _log('DISPOSE', d.barcode,
      _sanitize(d.reason, 200) + ' | ' + _sanitize(d.disposedBy, 100),
      d.disposedBy || '');
    return 'Disposal recorded';
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

function getDisposalData() {
  try {
    const sh   = _disposeSheet();
    const last = sh.getLastRow();
    if (last < EVT_DATA_START) return [];
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 10).getValues()
      .filter(r => r[0])
      .map(r => r.map(v => String(v || '')));
  } catch(e) { return []; }
}

// ─── DROPDOWN DATA ────────────────────────────────────────────────────────────
function getDropdownData() {
  try {
    const sh = _ss().getSheetByName(SH_DROPDOWN);
    if (!sh) return { categories:[], brands:{}, models:{}, suppliers:[], laptopSpecs:[] };
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1)
      return { categories:[], brands:{}, models:{}, suppliers:[], laptopSpecs:[] };

    const data  = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const row1  = data[0], row2 = data[1];
    const CAT_MAP = {
      'LAPTOP':'Laptop','LAPTOP ADAPTOR':'Laptop Adaptor','CPU':'CPU',
      'MONITOR':'Monitor','KEYBOARD':'Keyboard','MOUSE':'Mouse',
      'PRINTER':'Printer','SCANNER':'Scanner','SCANSNAP':'Scansnap',
      'SCANNER SCANSNAP':'Scanner','UPS':'UPS','EXTERNAL DRIVE':'External Drive',
      'CAMERA':'Camera','SPEAKER':'Speaker'
    };
    const result = {
      categories:[], brands:{}, models:{}, suppliers:[], laptopSpecs:[], laptopSpecValues:{}
    };
    let catPositions = [], supplierCol = -1, laptopSpecCol = -1, fieldSectionCol = -1;
    let usedMergedScannerHeader = false;

    for (let c = 0; c < row1.length; c++) {
      const raw = String(row1[c] || '').trim();
      if (!raw) continue;
      const up = raw.toUpperCase();
      if (up === 'SUPPLIERS' || up === 'SUPPLIER') { supplierCol = c; continue; }
      if (up === 'LAPTOP SPECS' || up === 'LAPTOP SPEC') { laptopSpecCol = c; continue; }
      if (up === 'FIELD' || up.startsWith('DIVISION') || up.startsWith('DISTRICT')) {
        if (fieldSectionCol < 0) fieldSectionCol = c; continue;
      }
      if (up === 'SCANNER SCANSNAP') {
        catPositions.push({ name:'Scanner', col:c }); usedMergedScannerHeader = true; continue;
      }
      catPositions.push({ name: CAT_MAP[up] || raw, col: c });
    }

    catPositions.forEach((cat, idx) => {
      const bounds = [
        catPositions[idx + 1]?.col,
        supplierCol   > -1 ? supplierCol   : null,
        laptopSpecCol > -1 ? laptopSpecCol : null,
        lastCol
      ].filter(v => v != null && v > cat.col);
      const endCol = Math.min(...bounds);
      result.categories.push(cat.name);
      result.brands[cat.name] = [];
      for (let c = cat.col; c < endCol; c++) {
        const brand = String(row2[c] || '').trim();
        if (!brand) continue;
        result.brands[cat.name].push(brand);
        const models = [];
        for (let r = 2; r < lastRow; r++) {
          const m = String(data[r][c] || '').trim();
          if (m) models.push(m);
        }
        result.models[cat.name + '|' + brand] = [...new Set(models)];
      }
    });

    if (usedMergedScannerHeader && result.brands['Scanner'] && !result.brands['Scansnap']) {
      result.categories.push('Scansnap');
      result.brands['Scansnap'] = [...result.brands['Scanner']];
      Object.keys(result.models).forEach(key => {
        if (key.startsWith('Scanner|'))
          result.models['Scansnap|' + key.slice('Scanner|'.length)] = [...result.models[key]];
      });
    }

    if (supplierCol > -1)
      for (let r = 1; r < lastRow; r++) {
        const s = String(data[r][supplierCol] || '').trim();
        if (s && !result.suppliers.includes(s)) result.suppliers.push(s);
      }

    if (laptopSpecCol > -1) {
      const specEnd = fieldSectionCol > -1 ? fieldSectionCol : lastCol;
      for (let c = laptopSpecCol; c < specEnd; c++) {
        const spec = String(row2[c] || '').trim();
        if (!spec) continue;
        result.laptopSpecs.push(spec);
        const vals = [];
        for (let r = 2; r < lastRow; r++) {
          const v = String(data[r][c] || '').trim();
          if (v) vals.push(v);
        }
        result.laptopSpecValues[spec] = vals;
      }
    }
    result.departments = getDepartmentList();
    return result;
  } catch(e) {
    return { categories:[], brands:{}, models:{}, suppliers:[], laptopSpecs:[], error:e.message };
  }
}

function getHeadOfficeDepts() {
  try {
    const sh    = _ss().getSheetByName(SH_DROPDOWN);
    if (!sh || sh.getLastRow() < 2) return [];
    const lastCol = sh.getLastColumn(), lastRow = sh.getLastRow();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    let deptCol = -1;
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c] || '').trim().toUpperCase();
      if (h === 'DEPARTMENTS' || h === 'HEAD OFFICE' || h === 'HO DEPTS') {
        deptCol = c; break;
      }
    }
    if (deptCol < 0) return [];
    const depts = [];
    for (let r = 1; r < lastRow; r++) {
      const v = String(sh.getRange(r + 1, deptCol + 1).getValue() || '').trim();
      if (v) depts.push(v);
    }
    return depts;
  } catch(e) { return []; }
}

function getDepartmentList() {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    const depts = new Set();
    if (last >= AE_DATA_START) {
      const vals = sh.getRange(AE_DATA_START, C.DEPARTMENT,
        last - AE_DATA_START + 1, 1).getValues();
      vals.forEach(r => { const v = String(r[0]||'').trim(); if(v) depts.add(v); });
    }
    // Merge with HO depts from dropdown sheet
    getHeadOfficeDepts().forEach(d => depts.add(d));
    return [...depts].sort();
  } catch(e) { return []; }
}

// ─── LOCATION DATA ────────────────────────────────────────────────────────────
function getLocationData(empId) {
  try {
    let scopeData;
    if (empId) {
      let roleTier = 'fe';
      try {
        const uSh = _ss().getSheetByName(SH_USERS);
        if (uSh && uSh.getLastRow() >= 2) {
          const uData = uSh.getRange(2, 1, uSh.getLastRow() - 1, 3).getValues();
          const uRow  = uData.find(r => String(r[2] || '').trim().toLowerCase()
                                         === String(empId).trim().toLowerCase());
          if (uRow) roleTier = _mapRoleTier(String(uRow[0] || '').trim());
        }
      } catch(e) {}
      scopeData = _parseOrgStructure(empId, roleTier);
    } else {
      scopeData = _buildFullDivDistrictMap();
      scopeData.seniorDistrictScope = scopeData.userDistricts;
    }

    const hoDepts = getHeadOfficeDepts();
    if (hoDepts.length > 0) {
      scopeData.divDistrictMap['Head Office'] = hoDepts;
    }

    return {
      divDistrictMap:      scopeData.divDistrictMap  || {},
      userDivisions:       scopeData.userDivisions   || [],
      userDistricts:       scopeData.userDistricts   || [],
      seniorDistrictScope: scopeData.seniorDistrictScope || [],
      headOfficeDepts:     hoDepts
    };
  } catch(e) {
    return {
      divDistrictMap:{}, userDivisions:[], userDistricts:[],
      seniorDistrictScope:[], headOfficeDepts:[], error:e.message
    };
  }
}

function getEngineerLocationData() { return getLocationData(null); }

function getEngineersByLocation(district, branch) {
  try {
    const sh     = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return { fe:'', senior:'' };
    const last   = sh.getLastRow();
    const data   = sh.getRange(2, 1, last - 1, 12).getValues();
    const normDist = String(district || '').trim().toLowerCase();
    const result = { fe:'', senior:'' };

    data.forEach(r => {
      const rowDist = String(r[11] || '').trim().toLowerCase();
      if (!normDist || rowDist !== normDist) return;
      if (!result.fe)     result.fe     = String(r[4] || '').trim();
      if (!result.senior) result.senior = String(r[7] || '').trim();
    });
    return result;
  } catch(e) { return { fe:'', senior:'' }; }
}

// ─── ORG LOOKUP ───────────────────────────────────────────────────────────────
function _buildOrgLookup() {
  const lookup = {};
  try {
    const sh = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return lookup;
    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, 12).getValues();

    data.forEach(r => {
      const div  = _normDiv(String(r[10] || '').trim());
      const dist = _normDist(String(r[11] || '').trim());
      if (!div || !dist) return;
      if (!lookup[dist.toLowerCase()])
        lookup[dist.toLowerCase()] = { division:div, district:dist, area:'', branch:'' };
      if (!lookup[div.toLowerCase()])
        lookup[div.toLowerCase()] = { division:div, district:'', area:'', branch:'' };
    });

    const foSh = _ss().getSheetByName('Org Structure-FO');
    if (foSh && foSh.getLastRow() >= 4) {
      const foLast = foSh.getLastRow();
      const foCol  = Math.max(foSh.getLastColumn(), 6);
      const foHdr  = foSh.getRange(3, 1, 1, foCol).getValues()[0];
      let   divC = -1, distC = -1, areaC = -1, branchC = -1;
      foHdr.forEach((h, i) => {
        const hUp = String(h || '').trim().toUpperCase();
        if (hUp.includes('DIVISION') && divC    < 0) divC    = i;
        if (hUp.includes('DISTRICT') && distC   < 0) distC   = i;
        if (hUp.includes('AREA')     && areaC   < 0) areaC   = i;
        if (hUp.includes('BRANCH')   && branchC < 0) branchC = i;
      });
      if (divC >= 0 && distC >= 0) {
        const foData = foSh.getRange(4, 1, foLast - 3, foCol).getValues();
        let cDiv = '', cDist = '', cArea = '';
        foData.forEach(r => {
          const rDiv    = divC    >= 0 ? String(r[divC]    || '').trim() : '';
          const rDist   = distC   >= 0 ? String(r[distC]   || '').trim() : '';
          const rArea   = areaC   >= 0 ? String(r[areaC]   || '').trim() : '';
          const rBranch = branchC >= 0 ? String(r[branchC] || '').trim() : '';
          if (rDiv)  cDiv  = _normDiv(rDiv);
          if (rDist) cDist = _normDist(rDist);
          if (rArea) cArea = rArea;
          const loc = { division:cDiv, district:cDist, area:cArea, branch:rBranch };
          if (rBranch)              lookup[rBranch.toLowerCase()] = loc;
          if (cDist && !lookup[cDist.toLowerCase()])
            lookup[cDist.toLowerCase()] = { division:cDiv, district:cDist, area:'', branch:'' };
          if (cDiv && !lookup[cDiv.toLowerCase()])
            lookup[cDiv.toLowerCase()] = { division:cDiv, district:'', area:'', branch:'' };
        });
      }
    }
    return lookup;
  } catch(e) {
    Logger.log('_buildOrgLookup error: ' + e.message);
    return lookup;
  }
}

// ─── ACTIVITY LOG ─────────────────────────────────────────────────────────────
function _log(action, barcode, details, performer) {
  try {
    _logSheet().appendRow([
      new Date().toLocaleString('en-PH'),
      action, barcode || '', details || '', performer || ''
    ]);
  } catch(e) {
    Logger.log('[_log FAILED] action=' + action + ' barcode=' + barcode + ' err=' + e.message);
  }
}

function getActivityLogs(page, pageSize) {
  try {
    const sh   = _logSheet();
    const last = sh.getLastRow();
    if (last < EVT_DATA_START) return { rows:[], total:0 };

    const all = sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 5).getValues()
      .filter(r => r[0])
      .reverse()
      .map(r => ({
        timestamp: String(r[0] || ''), action:    String(r[1] || ''),
        barcode:   String(r[2] || ''), details:   String(r[3] || ''),
        performer: String(r[4] || '')
      }));

    const total = all.length;
    const ps    = (pageSize && pageSize > 0) ? Number(pageSize) : 100;
    const pg    = (page     && page     > 0) ? Number(page)     : 1;
    return {
      rows:  all.slice((pg - 1) * ps, pg * ps),
      total, page:pg, pageSize:ps,
      totalPages: Math.ceil(total / ps) || 1
    };
  } catch(e) { return { rows:[], total:0 }; }
}

// ─── BATCH LOAD ───────────────────────────────────────────────────────────────
function getInitialData() {
  const assets    = getAllAssets().data || [];
  const borrows   = getBorrowData()     || [];
  const transfers = getTransferData()   || [];

  const activeBorrowMap = {};
  borrows.filter(b => b.status === 'Borrow').forEach(b => {
    if (!activeBorrowMap[b.barcode]) activeBorrowMap[b.barcode] = b;
  });
  assets.forEach(a => {
    if (a.status === 'borrowed') {
      const b = activeBorrowMap[a.Barcode];
      if (b) {
        a.BorName    = b.borrowerName;
        a.BorEmpID   = b.empId;
        a.BorDesig   = b.designation;
        a.BorDiv     = b.division;
        a.BorDist    = b.district;
        a.BorBranch  = b.branch;
        a.BorDate    = b.borrowDate;
        a.ExpReturn  = b.expectedReturn;
        a.ActReturn  = b.actualReturn;
        a.BorRemarks = b.remarks;
      }
    }
  });

  const xferMap = {};
  transfers.forEach(r => {
    if (!xferMap[r[0]] || r[20] > (xferMap[r[0]][20] || ''))
      xferMap[r[0]] = r;
  });
  assets.forEach(a => {
    if (a.status === 'transfer' && xferMap[a.Barcode]) {
      const r = xferMap[a.Barcode];
      a.XferType  = r[1]  || '';
      a.ToStaff   = r[10] || '';
      a.ToEmpID   = r[11] || '';
      a.ToDiv     = r[13] || '';
      a.ToBranch  = r[16] || '';
      a.XferDate  = r[18] || '';
    }
  });

  return {
    assets,
    borrows,
    transfers,
    disposals: getDisposalData(),
    logs:      getActivityLogs(1, 200).rows,
    orgLookup: _buildOrgLookup()
  };
}

function syncAll()           { return getAllAssets(); }
function getSpreadsheetUrl() { return _ss().getUrl(); }

// ─── STAFF MOVEMENT ───────────────────────────────────────────────────────────
function moveStaff(empId, newDiv, newDist, newArea, newBranch, assetAction) {
  try {
    if (!empId) return 'Error: Employee ID is required.';
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return 'Error: No assets found.';

    const count   = last - AE_DATA_START + 1;
    const allData = sh.getRange(AE_DATA_START, 1, count, 37).getValues();
    const nowStr  = new Date().toLocaleString('en-PH');
    const id      = String(empId).trim().toLowerCase();
    const normDiv  = _normDiv(newDiv   || '');
    const normDist = _normDist(newDist || '');

    let updated = 0, skipped = 0;
    const rowsToWrite = [];

    for (let i = 0; i < allData.length; i++) {
      const row      = allData[i];
      const rowEmpId = String(row[C.EMP_ID - 1] || '').trim().toLowerCase();
      if (rowEmpId !== id) continue;

      const lc = String(row[C.LIFECYCLE - 1] || '').toLowerCase();
      if (lc === 'borrow')                       { skipped++; continue; }
      if (lc === 'dispose' || lc === 'disposal') continue;

      const newRow = [...row];
      if (assetAction === 'spare') {
        newRow[C.LIFECYCLE    - 1] = 'Active';
        newRow[C.ASSET_STATUS - 1] = 'Active';
        newRow[C.STATUS_LABEL - 1] = 'Unassigned';
        newRow[C.EMP_ID       - 1] = '';
        newRow[C.STAFF        - 1] = '';
        newRow[C.DESIGNATION  - 1] = '';
        newRow[C.EFF_DATE     - 1] = '';
      }
      newRow[C.DIVISION      - 1] = normDiv;
      newRow[C.DISTRICT      - 1] = normDist;
      newRow[C.AREA          - 1] = newArea   || '';
      newRow[C.BRANCH        - 1] = newBranch || '';
      newRow[C.LAST_UPDATED  - 1] = nowStr;
      rowsToWrite.push({ rowIdx: i + AE_DATA_START, rowData: newRow });
      updated++;
    }

    rowsToWrite.forEach(({ rowIdx, rowData }) => {
      sh.getRange(rowIdx, 1, 1, 37).setValues([rowData]);
    });

    // Log one entry per affected asset so Activity Log scoping works correctly
    rowsToWrite.forEach(function(rw) {
      var bc = String(rw.rowData[C.BARCODE - 1] || '');
      _log('MOVE_STAFF', bc,
        'Action:' + assetAction + ' → ' + newDiv + '/' + newDist + '/' + newBranch,
        empId);
    });

    let msg = 'Staff movement recorded. ' + updated + ' asset(s) ' +
      (assetAction === 'spare' ? 'returned to spare' : 'moved to new location') + '.';
    if (skipped) msg += ' (' + skipped + ' skipped — on active borrow)';
    return msg;
  } catch(e) { return 'Error: ' + e.message; }
}

function moveOrgUnit(unitType, currentDiv, currentDist, currentArea, currentBranch,
                     newDiv, newDist, newArea) {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return 'Error: No assets found.';

    const count  = last - AE_DATA_START + 1;
    const data   = sh.getRange(AE_DATA_START, 1, count, 37).getValues();
    const nowStr = new Date().toLocaleString('en-PH');
    const normCurDiv  = _normDiv(currentDiv   || '');
    const normCurDist = _normDist(currentDist || '');
    const normNewDiv  = _normDiv(newDiv        || '');
    const normNewDist = _normDist(newDist      || '');
    let updated = 0;

    data.forEach((row, i) => {
      const rowDiv    = _normDiv(String(row[C.DIVISION  - 1] || '').trim());
      const rowDist   = _normDist(String(row[C.DISTRICT - 1] || '').trim());
      const rowArea   = String(row[C.AREA   - 1] || '').trim();
      const rowBranch = String(row[C.BRANCH - 1] || '').trim();
      const lc        = String(row[C.LIFECYCLE - 1] || '').toLowerCase();
      if (lc === 'dispose' || lc === 'disposal') return;

      const rowIdx = i + AE_DATA_START;
      const updates = [];

      if (unitType === 'district') {
        if (rowDist === normCurDist && rowDiv === normCurDiv) {
          updates.push([C.DIVISION, normNewDiv]);
          if (normNewDist && normNewDist !== normCurDist)
            updates.push([C.DISTRICT, normNewDist]);
        }
      } else if (unitType === 'area') {
        if (rowArea === currentArea && rowDist === normCurDist && rowDiv === normCurDiv) {
          updates.push([C.DIVISION, normNewDiv]);
          updates.push([C.DISTRICT, normNewDist]);
        }
      } else if (unitType === 'branch') {
        if (rowBranch === currentBranch && rowArea === currentArea &&
            rowDist   === normCurDist   && rowDiv  === normCurDiv) {
          updates.push([C.DIVISION, normNewDiv]);
          updates.push([C.DISTRICT, normNewDist]);
          updates.push([C.AREA, newArea || currentArea]);
        }
      }

      if (updates.length) {
        updates.push([C.LAST_UPDATED, nowStr]);
        updates.forEach(u => sh.getRange(rowIdx, u[0]).setValue(u[1]));
        updated++;
      }
    });

    // Log one entry per affected asset row
    if (last >= AE_DATA_START) {
      var logData = sh.getRange(AE_DATA_START, C.BARCODE, last - AE_DATA_START + 1, 1).getValues();
      // We already iterated and updated; re-check to log affected barcodes
      data.forEach(function(row, i) {
        var bc = String(row[C.BARCODE - 1] || '').trim();
        if (!bc) return;
        // Only log rows that were actually updated (check updated district matches)
        var newDist_check = _normDist(String(sh.getRange(i + AE_DATA_START, C.DISTRICT).getValue() || ''));
        if (newDist_check === _normDist(newDist || '') && bc) {
          _log('MOVE_ORG', bc,
            unitType + ': ' + currentDiv + '/' + currentDist + ' → ' + newDiv + '/' + newDist,
            '');
        }
      });
    }

    return updated > 0
      ? unitType.charAt(0).toUpperCase() + unitType.slice(1) +
        ' moved. ' + updated + ' asset(s) updated.'
      : 'Move recorded — no matching assets found.';
  } catch(e) { return 'Error: ' + e.message; }
}

// ═══════════════════════════════════════════════════════════
//  ACCOUNTABILITY FORM WORKFLOW
// ═══════════════════════════════════════════════════════════

const SH_AF  = 'AccountabilityForms';
const SH_FS  = 'FormSnapshots';
const SH_RL  = 'FormRateLimits';
const SH_CFG = 'ApprovalConfig';

const AF = {
  FORM_ID:        1, FORM_TYPE:      2, LINKED_FORM_ID: 3, CONTEXT_TYPE:   4,
  EMP_ID:         5, STAFF_NAME:     6, DESIGNATION:    7, DEPARTMENT:     8,
  BRANCH:         9, DIVISION:       10, DISTRICT:      11, ASSETS_JSON:   12,
  STATUS:         13, DRAFTED_BY:    14, DRAFTED_AT:    15, SUBMITTED_AT:  16,
  SUPERVISOR_ID:  17, REVIEWED_AT:   18, REJECTION_COMMENT: 19
};
const AF_TOTAL_COLS = 19;
const AF_DATA_START = 4;

const RL = {
  FORM_ID: 1, DRAFTED_BY: 2, RESUBMIT_COUNT: 3,
  WINDOW_START: 4, COOLDOWN_UNTIL: 5, LAST_SUBMITTED: 6, STATUS: 7
};
const RL_DATA_START = 4;

const FS = {
  FORM_ID: 1, FORM_TYPE: 2, LINKED_FORM_ID: 3, CONTEXT_TYPE: 4,
  EMP_ID: 5, STAFF_NAME: 6, DESIGNATION: 7, DEPARTMENT: 8,
  BRANCH: 9, ASSETS_JSON: 10, CONFIRMED_BY: 11, CONFIRMED_AT: 12,
  SUPERSEDED_AT: 13, SUPERSEDED_BY: 14
};
const FS_DATA_START = 4;

const RL_MAX_RESUBMITS    = 5;
const RL_WINDOW_MINUTES   = 30;
const RL_COOLDOWN_MINUTES = 120;

function _afSheet() {
  return _getOrCreate(SH_AF, [
    'Form ID','Form Type','Linked Form ID','Context Type',
    'Emp ID','Staff Name','Designation','Department','Branch / Base Office',
    'Division','District','Assets JSON','Status',
    'Drafted By','Drafted At','Submitted At',
    'Supervisor ID','Reviewed At','Rejection Comment'
  ]);
}

function _fsSheet() {
  return _getOrCreate(SH_FS, [
    'Form ID','Form Type','Linked Form ID','Context Type',
    'Emp ID','Staff Name','Designation','Department','Branch / Base Office',
    'Assets JSON','Confirmed By','Confirmed At','Superseded At','Superseded By Form'
  ]);
}

function _rlSheet() {
  return _getOrCreate(SH_RL, [
    'Form ID','Drafted By','Resubmit Count',
    'Window Start','Cooldown Until','Last Submitted At','Status'
  ]);
}

function getApprovalConfig(contextType) {
  try {
    const sh   = _ss().getSheetByName(SH_CFG);
    if (!sh) return _defaultApprovalConfig(contextType);
    const last = sh.getLastRow();
    if (last < 5) return _defaultApprovalConfig(contextType);
    const data = sh.getRange(5, 1, 2, 4).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === (contextType || '').toLowerCase()) {
        return {
          contextType:      String(data[i][0]).trim(),
          processedByLabel: String(data[i][1]).trim(),
          checkedByLabel:   String(data[i][2]).trim(),
          verifiedByName:   String(data[i][3]).trim(),
          notedByName:      'Patrick Gerard G. Reyes',
          notedByTitle:     'Department Head'
        };
      }
    }
    return _defaultApprovalConfig(contextType);
  } catch(e) { return _defaultApprovalConfig(contextType); }
}

function _defaultApprovalConfig(contextType) {
  if ((contextType || '').toLowerCase() === 'ho') {
    return {
      contextType: 'HO', processedByLabel: 'Technical Support Engineer',
      checkedByLabel: 'Senior Technical Support Engineer',
      verifiedByName: 'Sandylee Dela Cruz Paris',
      notedByName: 'Patrick Gerard G. Reyes', notedByTitle: 'Department Head'
    };
  }
  return {
    contextType: 'Field', processedByLabel: 'Field Engineer',
    checkedByLabel: 'Senior Field Engineer',
    verifiedByName: 'Maricon B. Jaropillo',
    notedByName: 'Patrick Gerard G. Reyes', notedByTitle: 'Department Head'
  };
}

function getMyForms(drafterId) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return [];
    const id   = String(drafterId || '').trim().toLowerCase();
    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    return data
      .filter(r => r[AF.FORM_ID - 1] && String(r[AF.DRAFTED_BY - 1]).trim().toLowerCase() === id)
      .map(r => _mapAfRow(r));
  } catch(e) { return []; }
}

function getPendingForms(supervisorId, roleTier) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return [];
    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    let pending = data
      .filter(r => r[AF.FORM_ID - 1] && String(r[AF.STATUS - 1]).trim() === 'Pending')
      .map(r => _mapAfRow(r));
    if (roleTier === 'ho') return pending;
    const supervisedIds = _getSupervisedEmpIds(supervisorId);
    return pending.filter(f => supervisedIds.indexOf(f.draftedBy.toLowerCase()) >= 0);
  } catch(e) { return []; }
}

function getFormDetail(formId) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return null;
    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][AF.FORM_ID - 1]).trim() === String(formId).trim())
        return _mapAfRow(data[i]);
    }
    return null;
  } catch(e) { return null; }
}

function getRateLimitStatus(formId, drafterId) {
  try {
    const row = _findRLRow(formId, drafterId);
    if (!row) return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null };
    const now           = new Date();
    const cooldownUntil = row[RL.COOLDOWN_UNTIL - 1] ? new Date(row[RL.COOLDOWN_UNTIL - 1]) : null;
    if (cooldownUntil && now < cooldownUntil)
      return { allowed: false, remaining: 0, cooldownUntil: cooldownUntil.toISOString() };
    const windowStart   = row[RL.WINDOW_START - 1] ? new Date(row[RL.WINDOW_START - 1]) : null;
    const windowExpired = !windowStart || ((now - windowStart) / 60000) > RL_WINDOW_MINUTES;
    if (windowExpired) return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null };
    const count     = parseInt(row[RL.RESUBMIT_COUNT - 1] || 0, 10);
    const remaining = Math.max(0, RL_MAX_RESUBMITS - count);
    return { allowed: remaining > 0, remaining, cooldownUntil: null };
  } catch(e) { return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null }; }
}

// Lock-free internal helper — call only from within an already-locked context.
function _generateFormIDUnsafe() {
  const sh   = _afSheet();
  const last = sh.getLastRow();
  const yr   = new Date().getFullYear();
  let   max  = 0;
  if (last >= AF_DATA_START) {
    const ids     = sh.getRange(AF_DATA_START, AF.FORM_ID, last - AF_DATA_START + 1, 1).getValues();
    const pattern = new RegExp('^FORM-' + yr + '-(\\d+)$');
    ids.forEach(r => {
      const m = String(r[0] || '').match(pattern);
      if (m) { const n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
    });
  }
  return 'FORM-' + yr + '-' + String(max + 1).padStart(3, '0');
}

// Public API — acquires its own lock; safe to call from the front-end.
function generateFormID() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    return _generateFormIDUnsafe();
  } finally { lock.releaseLock(); }
}

function checkRateLimit(formId, drafterId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(6000); }
  catch(e) { return { allowed: false, message: 'System busy — try again.' }; }
  try {
    const sh  = _rlSheet();
    const now = new Date();
    const rowIdx = _findRLRowIdx(formId, drafterId);
    if (rowIdx < 0) {
      sh.appendRow([formId, drafterId, 1, now.toLocaleString('en-PH'), '', now.toLocaleString('en-PH'), 'Active']);
      return { allowed: true, remaining: RL_MAX_RESUBMITS - 1, cooldownUntil: null };
    }
    const data = sh.getRange(rowIdx, 1, 1, 7).getValues()[0];
    const cooldownRaw   = data[RL.COOLDOWN_UNTIL - 1];
    const cooldownUntil = cooldownRaw ? new Date(cooldownRaw) : null;
    if (cooldownUntil && now < cooldownUntil) {
      const mins = Math.ceil((cooldownUntil - now) / 60000);
      return { allowed: false, remaining: 0, cooldownUntil: cooldownUntil.toISOString(), message: 'Rate limit reached. Try again in ' + mins + ' minute(s).' };
    }
    const windowStart   = data[RL.WINDOW_START - 1] ? new Date(data[RL.WINDOW_START - 1]) : null;
    const windowExpired = !windowStart || ((now - windowStart) / 60000) > RL_WINDOW_MINUTES;
    let count = parseInt(data[RL.RESUBMIT_COUNT - 1] || 0, 10);
    if (windowExpired) {
      count = 0;
      sh.getRange(rowIdx, RL.WINDOW_START).setValue(now.toLocaleString('en-PH'));
      sh.getRange(rowIdx, RL.COOLDOWN_UNTIL).setValue('');
      sh.getRange(rowIdx, RL.STATUS).setValue('Active');
    }
    count++;
    if (count > RL_MAX_RESUBMITS) {
      const cooldownEnd = new Date(now.getTime() + RL_COOLDOWN_MINUTES * 60000);
      sh.getRange(rowIdx, RL.RESUBMIT_COUNT).setValue(count);
      sh.getRange(rowIdx, RL.COOLDOWN_UNTIL).setValue(cooldownEnd.toLocaleString('en-PH'));
      sh.getRange(rowIdx, RL.STATUS).setValue('Cooldown');
      return { allowed: false, remaining: 0, cooldownUntil: cooldownEnd.toISOString(), message: 'Submission limit reached. You can resubmit after 2 hours.' };
    }
    sh.getRange(rowIdx, RL.RESUBMIT_COUNT).setValue(count);
    sh.getRange(rowIdx, RL.LAST_SUBMITTED).setValue(now.toLocaleString('en-PH'));
    return { allowed: true, remaining: RL_MAX_RESUBMITS - count, cooldownUntil: null, message: 'Submitted. ' + (RL_MAX_RESUBMITS - count) + ' attempt(s) remaining.' };
  } finally { lock.releaseLock(); }
}

function _clearRateLimit(formId, drafterId) {
  try {
    const rowIdx = _findRLRowIdx(formId, drafterId);
    if (rowIdx < 0) return;
    const sh = _rlSheet();
    sh.getRange(rowIdx, RL.RESUBMIT_COUNT).setValue(0);
    sh.getRange(rowIdx, RL.COOLDOWN_UNTIL).setValue('');
    sh.getRange(rowIdx, RL.STATUS).setValue('Clear');
  } catch(e) {}
}

function draftAccountabilityForm(empId, assets, formType, linkedFormId, draftedBy) {
  try {
    if (!empId)     return 'Error: Employee ID is required.';
    if (!formType)  return 'Error: Form type is required.';
    if (!draftedBy) return 'Error: Drafter ID is required.';

    // Use the lock-free helper as instructed
    const formId = _generateFormIDUnsafe();
    if (formId.startsWith('Error')) return formId;

    const nowStr      = new Date().toLocaleString('en-PH');
    const contextType = _getContextType(empId);
    const assetsJson  = _buildAssetsSnapshot(assets || []);
    const refAsset    = (assets && assets.length) ? assets[0] : {};
    const staffName   = refAsset.Staff         || '';
    const desig       = refAsset.Designation || '';
    const dept        = refAsset.Department  || '';
    const branch      = refAsset.Branch      || refAsset.BaseOffice || '';
    const division    = refAsset.Division    || '';
    const district    = refAsset.District    || '';

    const row = new Array(AF_TOTAL_COLS).fill('');
    row[AF.FORM_ID         - 1] = formId;
    row[AF.FORM_TYPE      - 1] = formType;
    row[AF.LINKED_FORM_ID - 1] = linkedFormId || '';
    row[AF.CONTEXT_TYPE   - 1] = contextType;
    row[AF.EMP_ID         - 1] = empId;
    row[AF.STAFF_NAME     - 1] = _sanitize(staffName, 100);
    row[AF.DESIGNATION    - 1] = _sanitize(desig, 100);
    row[AF.DEPARTMENT     - 1] = _sanitize(dept, 100);
    row[AF.BRANCH         - 1] = _sanitize(branch, 150);
    row[AF.DIVISION       - 1] = division;
    row[AF.DISTRICT       - 1] = district;
    row[AF.ASSETS_JSON    - 1] = assetsJson;
    row[AF.STATUS         - 1] = 'Draft';
    row[AF.DRAFTED_BY     - 1] = draftedBy;
    row[AF.DRAFTED_AT     - 1] = nowStr;

    _afSheet().appendRow(row);

    if (assets && assets.length) {
      const sh = _entrySheet();
      assets.forEach(function(asset) {
        const rowIdx = _findRow(sh, asset.Barcode);
        if (rowIdx < 1) return;
        sh.getRange(rowIdx, CAE.APPROVAL_STATUS).setValue('Draft');
        sh.getRange(rowIdx, CAE.FORM_ID).setValue(formId);
        sh.getRange(rowIdx, CAE.DRAFTED_BY).setValue(draftedBy);
        sh.getRange(rowIdx, CAE.DRAFTED_AT).setValue(nowStr);
        sh.getRange(rowIdx, CAE.REJECTION_COMMENT).setValue('');
      });
    }

    _log('DRAFT_FORM', formId, formType + ' | ' + empId + ' | ' + (assets ? assets.length : 0) + ' assets', draftedBy);
    return formId;

  } catch(e) { 
    return 'Error: ' + e.message; 
  }
}
function _buildAssetsSnapshot(assets) {
  if (!assets || !assets.length) return '[]';
  return JSON.stringify(assets.map(function(a) {
    return { barcode:a.Barcode||'', type:a.Type||'', brand:a.Brand||'', serial:a.Serial||'', specs:a.Specs||'', condition:a.Condition||'' };
  }));
}

function _getContextType(empId) {
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return 'HO';
    const last = sh.getLastRow();
    const id   = String(empId || '').trim().toLowerCase();
    const ids  = sh.getRange(2, 4, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() === id) return 'Field';
    }
    const supIds = sh.getRange(2, 7, last - 1, 1).getValues();
    for (let i = 0; i < supIds.length; i++) {
      if (String(supIds[i][0] || '').trim().toLowerCase() === id) return 'Field';
    }
    return 'HO';
  } catch(e) { return 'HO'; }
}

function submitFormForReview(formId, drafterId) {
  try {
    const form = getFormDetail(formId);
    if (!form)                      return { ok: false, message: 'Form not found: ' + formId };
    if (form.status === 'Confirmed') return { ok: false, message: 'This form is already confirmed.' };
    if (form.status === 'Pending')   return { ok: false, message: 'This form is already pending review.' };
    if (form.draftedBy.toLowerCase() !== String(drafterId).trim().toLowerCase())
                                    return { ok: false, message: 'Only the original drafter can submit this form.' };
    if (form.status !== 'Draft' && form.status !== 'Rejected')
                                    return { ok: false, message: 'Form status "' + form.status + '" cannot be submitted.' };

    const sh     = _afSheet();
    const rowIdx = _findAFRow(formId);
    if (rowIdx < 0) return { ok: false, message: 'Form record not found in sheet.' };

    // Rate limit check runs LAST, after all validation passes, so no credits are wasted on invalid submissions
    if (form.status === 'Rejected') {
      const rlResult = checkRateLimit(formId, drafterId);
      if (!rlResult.allowed) return { ok: false, message: rlResult.message, rateLimitStatus: rlResult };
    }

    const nowStr = new Date().toLocaleString('en-PH');
    sh.getRange(rowIdx, AF.STATUS).setValue('Pending');
    sh.getRange(rowIdx, AF.SUBMITTED_AT).setValue(nowStr);
    sh.getRange(rowIdx, AF.REJECTION_COMMENT).setValue('');
    _updateAssetApprovalStatus(formId, 'Pending', '');

    _log('SUBMIT_FORM', formId, 'Submitted for review | ' + form.formType + ' | ' + form.empId, drafterId);
    return { ok: true, formId, message: 'Form submitted for supervisor review.' };
  } catch(e) { return { ok: false, message: 'Error: ' + e.message }; }
}

function confirmForm(formId, supervisorId, roleTier) {
  try {
    const form = getFormDetail(formId);
    if (!form)                     return { ok: false, message: 'Form not found.' };
    if (form.status !== 'Pending') return { ok: false, message: 'Form is not pending review.' };
    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(form.empId.toLowerCase()) < 0)
        return { ok: false, message: 'You do not supervise the accountable person on this form.' };
    }
    const sh     = _afSheet();
    const rowIdx = _findAFRow(formId);
    if (rowIdx < 0) return { ok: false, message: 'Form record not found in sheet.' };
    const nowStr = new Date().toLocaleString('en-PH');
    sh.getRange(rowIdx, AF.STATUS).setValue('Confirmed');
    sh.getRange(rowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
    sh.getRange(rowIdx, AF.REVIEWED_AT).setValue(nowStr);
    _archiveForm(formId, supervisorId, nowStr);
    _updateAssetApprovalStatus(formId, 'Confirmed', '');
    _clearRateLimit(formId, form.draftedBy);
    _log('CONFIRM_FORM', formId, form.formType + ' | ' + form.empId + ' | confirmed', supervisorId);
    return { ok: true, message: 'Form confirmed. Asset(s) are now live in inventory.' };
  } catch(e) { return { ok: false, message: 'Error: ' + e.message }; }
}

function confirmTransferPair(fromFormId, toFormId, supervisorId, roleTier) {
  try {
    const fromForm = getFormDetail(fromFormId);
    const toForm   = getFormDetail(toFormId);
    if (!fromForm) return { ok: false, message: 'From-form not found: ' + fromFormId };
    if (!toForm)   return { ok: false, message: 'To-form not found: ' + toFormId };
    if (fromForm.status !== 'Pending') return { ok: false, message: 'From-form is not pending review.' };
    if (toForm.status   !== 'Pending') return { ok: false, message: 'To-form is not pending review.' };
    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(fromForm.empId.toLowerCase()) < 0 &&
          supervised.indexOf(toForm.empId.toLowerCase()) < 0)
        return { ok: false, message: 'You do not supervise the drafter(s) of this transfer.' };
    }
    const nowStr = new Date().toLocaleString('en-PH');
    const sh     = _afSheet();
    const fromRowIdx = _findAFRow(fromFormId);
    if (fromRowIdx > 0) {
      sh.getRange(fromRowIdx, AF.STATUS).setValue('Confirmed');
      sh.getRange(fromRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(fromRowIdx, AF.REVIEWED_AT).setValue(nowStr);
    }
    const toRowIdx = _findAFRow(toFormId);
    if (toRowIdx > 0) {
      sh.getRange(toRowIdx, AF.STATUS).setValue('Confirmed');
      sh.getRange(toRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(toRowIdx, AF.REVIEWED_AT).setValue(nowStr);
    }
    _archiveForm(toFormId, supervisorId, nowStr);
    _archiveForm(fromFormId, supervisorId, nowStr);
    _supersedeForms(fromFormId, toFormId, nowStr);
    _updateAssetApprovalStatus(fromFormId, 'Confirmed', '');
    _updateAssetApprovalStatus(toFormId,   'Confirmed', '');
    _clearRateLimit(fromFormId, fromForm.draftedBy);
    _clearRateLimit(toFormId,   toForm.draftedBy);
    _log('CONFIRM_TRANSFER', fromFormId + '+' + toFormId,
      'Transfer pair confirmed | From: ' + fromForm.empId + ' → To: ' + toForm.empId, supervisorId);
    return { ok: true, message: 'Transfer confirmed. Both forms are now live.' };
  } catch(e) { return { ok: false, message: 'Error: ' + e.message }; }
}

function rejectForm(formId, supervisorId, comment, roleTier) {
  try {
    if (!comment || !comment.trim()) return { ok: false, message: 'A rejection comment is required.' };
    const form = getFormDetail(formId);
    if (!form)                     return { ok: false, message: 'Form not found.' };
    if (form.status !== 'Pending') return { ok: false, message: 'Only Pending forms can be rejected.' };
    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(form.draftedBy.toLowerCase()) < 0)
        return { ok: false, message: 'You do not supervise the drafter of this form.' };
    }
    const sh     = _afSheet();
    const rowIdx = _findAFRow(formId);
    if (rowIdx < 0) return { ok: false, message: 'Form record not found in sheet.' };
    const nowStr  = new Date().toLocaleString('en-PH');
    const trimmed = _sanitize(comment, 500);
    sh.getRange(rowIdx, AF.STATUS).setValue('Rejected');
    sh.getRange(rowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
    sh.getRange(rowIdx, AF.REVIEWED_AT).setValue(nowStr);
    sh.getRange(rowIdx, AF.REJECTION_COMMENT).setValue(trimmed);
    _updateAssetApprovalStatus(formId, 'Rejected', trimmed);
    _log('REJECT_FORM', formId, form.formType + ' | ' + form.empId + ' | ' + trimmed, supervisorId);
    return { ok: true, message: 'Form rejected. The drafter has been notified.' };
  } catch(e) { return { ok: false, message: 'Error: ' + e.message }; }
}

function rejectTransferPair(fromFormId, toFormId, supervisorId, fromComment, toComment, roleTier) {
  try {
    const comment = (fromComment || toComment || '').trim();
    if (!comment) return { ok: false, message: 'A rejection comment is required for at least one form.' };
    const fromForm = getFormDetail(fromFormId);
    const toForm   = getFormDetail(toFormId);
    if (!fromForm) return { ok: false, message: 'From-form not found.' };
    if (!toForm)   return { ok: false, message: 'To-form not found.' };
    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(fromForm.draftedBy.toLowerCase()) < 0 &&
          supervised.indexOf(toForm.draftedBy.toLowerCase()) < 0)
        return { ok: false, message: 'You do not supervise the drafter(s) of this transfer.' };
    }
    const sh     = _afSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const fromRowIdx = _findAFRow(fromFormId);
    if (fromRowIdx > 0) {
      sh.getRange(fromRowIdx, AF.STATUS).setValue('Rejected');
      sh.getRange(fromRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(fromRowIdx, AF.REVIEWED_AT).setValue(nowStr);
      sh.getRange(fromRowIdx, AF.REJECTION_COMMENT).setValue(_sanitize(fromComment || comment, 500));
    }
    const toRowIdx = _findAFRow(toFormId);
    if (toRowIdx > 0) {
      sh.getRange(toRowIdx, AF.STATUS).setValue('Rejected');
      sh.getRange(toRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(toRowIdx, AF.REVIEWED_AT).setValue(nowStr);
      sh.getRange(toRowIdx, AF.REJECTION_COMMENT).setValue(_sanitize(toComment || comment, 500));
    }
    _updateAssetApprovalStatus(fromFormId, 'Rejected', _sanitize(fromComment || comment, 500));
    _updateAssetApprovalStatus(toFormId,   'Rejected', _sanitize(toComment   || comment, 500));
    _log('REJECT_TRANSFER', fromFormId + '+' + toFormId, 'Transfer pair rejected | ' + comment, supervisorId);
    return { ok: true, message: 'Transfer rejected. Both forms returned to drafters.' };
  } catch(e) { return { ok: false, message: 'Error: ' + e.message }; }
}

function _archiveForm(formId, confirmedBy, confirmedAt) {
  try {
    const form = getFormDetail(formId);
    if (!form) return;
    _fsSheet().appendRow([
      form.formId, form.formType, form.linkedFormId || '', form.contextType,
      form.empId, form.staffName, form.designation, form.department, form.branch,
      form.assetsJson, confirmedBy, confirmedAt, '', ''
    ]);
  } catch(e) { Logger.log('_archiveForm error: ' + e.message); }
}

function _supersedeForms(oldFormId, newFormId, nowStr) {
  try {
    const sh   = _fsSheet();
    const last = sh.getLastRow();
    if (last < FS_DATA_START) return;
    const ids = sh.getRange(FS_DATA_START, FS.FORM_ID, last - FS_DATA_START + 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(oldFormId).trim()) {
        const row = i + FS_DATA_START;
        sh.getRange(row, FS.SUPERSEDED_AT).setValue(nowStr);
        sh.getRange(row, FS.SUPERSEDED_BY).setValue(newFormId);
        break;
      }
    }
  } catch(e) { Logger.log('_supersedeForms error: ' + e.message); }
}

// UPDATED processAsset — supports accEmpId/accName for accountable person separate from inputter
function processAsset(obj, isAssign) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { result: 'Error: System busy — try again.', formId: null }; }

  try {
    const sh     = _entrySheet();
    const nowStr = new Date().toLocaleString('en-PH');

    if (!isAssign) {
      if (!obj.barcode) return { result: 'Error: Barcode is required.', formId: null };
      if (_findRow(sh, obj.barcode) > 0)
        return { result: 'Error: Barcode already exists: ' + obj.barcode, formId: null };

      if (obj.serial) {
        const last2 = sh.getLastRow();
        if (last2 >= AE_DATA_START) {
          const serials = sh.getRange(AE_DATA_START, C.SERIAL, last2 - AE_DATA_START + 1, 1).getValues();
          const dup = serials.findIndex(r => String(r[0]).trim() === String(obj.serial).trim());
          if (dup >= 0) {
            const existBC = String(sh.getRange(dup + AE_DATA_START, C.BARCODE).getValue());
            return { result: 'Error: Serial No. already registered under barcode: ' + existBC, formId: null };
          }
        }
      }

      const statusChoice = obj.statusChoice || 'Spare';
      const SM = {
        'Spare':      { lc:'Active',     as:'Active',     sl:'Unassigned' },
        'Allocated':  { lc:'Allocated',  as:'Active',     sl:'Assigned'   },
        'Disposal':   { lc:'Dispose',    as:'Disposal',   sl:'Disposed'   },
        'BorrowItem': { lc:'BorrowItem', as:'BorrowItem', sl:'Unassigned' }
      };
      const sm      = SM[statusChoice] || SM['Spare'];
      const normDiv  = _normDiv(obj.division  || '');
      const normDist = _normDist(obj.district || '');

      // ── FIX: ACCOUNTABLE PERSON LOGIC ────────────────────────────────────
      // SPARE assets: EmpID/Staff must be EMPTY so _computeStatus returns
      //   'spare' (not 'allocated').  Nobody is accountable for a spare asset
      //   until it is explicitly allocated via allocateAsset().
      //
      // NON-SPARE (direct allocation at enroll): use accEmpId if provided,
      //   else fall back to the inputter.
      // ─────────────────────────────────────────────────────────────────────
      var accEmpId = '', accName = '', accDesig = '';
      if (statusChoice !== 'Spare') {
        accEmpId = obj.accEmpId || obj.entryEmpId || '';
        accName  = obj.accName  || obj.entryName  || '';
        accDesig = obj.accDesig || '';
      }

      const row = new Array(37).fill('');
      row[C.ENTRY_EMP_ID   - 1] = obj.entryEmpId  || '';
      row[C.ENTRY_NAME     - 1] = obj.entryName   || '';
      row[C.EMP_ID         - 1] = accEmpId;
      row[C.STAFF          - 1] = _sanitize(accName, 100);
      row[C.DESIGNATION    - 1] = accDesig;
      row[C.DEPARTMENT     - 1] = obj.department  || '';
      row[C.BASE_OFFICE    - 1] = obj.baseOffice  || '';
      row[C.DIVISION       - 1] = normDiv;
      row[C.DISTRICT       - 1] = normDist;
      row[C.AREA           - 1] = obj.area        || '';
      row[C.BRANCH         - 1] = _sanitize(obj.branch, 150);
      row[C.ASSIGNMENT     - 1] = obj.assignment  || (obj.fieldType === 'Central Office' ? 'Central Office' : 'Field Office');
      row[C.EFF_DATE       - 1] = obj.effDate     || '';
      row[C.BARCODE        - 1] = obj.barcode;
      row[C.TYPE           - 1] = obj.type        || '';
      row[C.BRAND          - 1] = obj.brand       || '';
      row[C.SERIAL         - 1] = obj.serial ? String(obj.serial) : '';
      row[C.SPECS          - 1] = obj.specs       || '';
      row[C.SUPPLIER       - 1] = obj.supplier    || '';
      row[C.CONDITION      - 1] = obj.condition   || 'New';
      row[C.ASSET_LOCATION - 1] = obj.location    || '';
      row[C.LIFECYCLE      - 1] = sm.lc;
      row[C.STATUS_LABEL   - 1] = sm.sl;
      row[C.ASSET_STATUS   - 1] = sm.as;
      row[C.PURCH_DATE     - 1] = obj.purchDate   || '';
      row[C.WARRANTY_TERM  - 1] = obj.wTerm       || '';
      row[C.WARRANTY_VAL   - 1] = obj.wValidity   || '';
      row[C.REMARKS        - 1] = _sanitize(obj.remarks, 500);
      row[C.NOTES          - 1] = _sanitize(obj.notes || '', 500);
      row[C.CREATED_AT     - 1] = nowStr;
      row[C.LAST_UPDATED   - 1] = nowStr;

      // ── FIX: APPROVAL STATUS ──────────────────────────────────────────────
      // Spare assets are 'Confirmed' from birth — no approval workflow until
      //   they are allocated.  Grandfathered = true skips the print gate check.
      // Non-spare assets start as 'Draft' pending supervisor confirmation.
      // ─────────────────────────────────────────────────────────────────────
      row[CAE.APPROVAL_STATUS  - 1] = statusChoice === 'Spare' ? 'Confirmed' : 'Draft';
      row[CAE.FORM_ID          - 1] = '';
      row[CAE.DRAFTED_BY       - 1] = obj.enrolledBy || obj.entryEmpId || '';
      row[CAE.DRAFTED_AT       - 1] = nowStr;
      row[CAE.REJECTION_COMMENT- 1] = '';
      row[CAE.GRANDFATHERED    - 1] = statusChoice === 'Spare' ? true : false;

      sh.appendRow(row);
      const newRowIdx = sh.getLastRow();
      sh.getRange(newRowIdx, C.SERIAL).setNumberFormat('@STRING@');
      if (obj.serial) sh.getRange(newRowIdx, C.SERIAL).setValue(String(obj.serial));

      if (statusChoice === 'Spare') {
        _spareSheet().appendRow([
          accEmpId,                          // Col 1: Accountable Person ID
          _sanitize(accName, 100),           // Col 2: Accountable Person
          'Spare',                           // Col 3: Assignment
          accDesig,                          // Col 4: Designation
          obj.department || '',              // Col 5: Department
          obj.barcode,                       // Col 6: Barcode
          obj.type,                          // Col 7: Category
          obj.brand || '',                   // Col 8: Brand
          obj.serial ? String(obj.serial) : '', // Col 9: Serial No.
          obj.condition || 'New',            // Col 10: Condition
          obj.purchDate || '',               // Col 11: Purchase Date
          obj.wValidity || '',               // Col 12: Warranty Validity
          obj.supplier || '',                // Col 13: Supplier
          obj.baseOffice || _sanitize(obj.branch, 150) || '', // Col 14: Base Office
          normDiv,                           // Col 15: Division
          normDist,                          // Col 16: District
          obj.area || '',                    // Col 17: Area
          _sanitize(obj.branch, 150),        // Col 18: Branch
          obj.location || '',                // Col 19: Asset Location
          obj.enrolledBy || obj.entryEmpId || '', // Col 20: Enrolled By
          nowStr                             // Col 21: Created By
        ]);
      }

      _log('CREATE', obj.barcode, obj.type + ' | ' + obj.brand + ' | ' + statusChoice, obj.entryEmpId || '');

      // ── FIX: ACCOUNTABILITY FORM ──────────────────────────────────────────
      // SPARE assets: NO form created here.
      //   Reason: Nobody is accountable for unallocated spare inventory.
      //   The accountability form is created automatically by allocateAsset()
      //   when the asset is later assigned to a staff member.
      //
      // NON-SPARE (direct allocation at enrollment): create form now.
      // ─────────────────────────────────────────────────────────────────────
      if (statusChoice === 'Spare') {
        return {
          result: 'Asset added to Spare Pool: ' + obj.barcode,
          formId: null    // No form yet — form will be created on allocation
        };
      }

      // Non-spare path: create the accountability form
      const assetForForm = [{
        Barcode:     obj.barcode,
        Type:        obj.type      || '',
        Brand:       obj.brand     || '',
        Serial:      obj.serial    || '',
        Specs:       obj.specs     || '',
        Condition:   obj.condition || 'New',
        Staff:       _sanitize(accName, 100),
        Designation: accDesig,
        Department:  obj.department || '',
        BaseOffice:  obj.baseOffice || '',
        Branch:      _sanitize(obj.branch, 150),
        Division:    normDiv,
        District:    normDist,
        Area:        obj.area      || ''
      }];

      const drafterId    = obj.enrolledBy || obj.entryEmpId || '';
      const empIdForForm = accEmpId || drafterId;

      const formId = draftAccountabilityForm(empIdForForm, assetForForm, 'Enrollment', '', drafterId);

      if (!formId.startsWith('Error')) {
        sh.getRange(newRowIdx, CAE.FORM_ID).setValue(formId);
      }

      return {
        result: 'Asset created: ' + obj.barcode,
        formId: formId.startsWith('Error') ? null : formId
      };
    }

    // isAssign=true path (unchanged)
    const lc    = obj.lifecycle || 'Allocated';
    const asSt  = lc === 'Transfer' ? 'Transfer' : lc === 'Dispose' ? 'Disposal' : 'Active';
    const staff = _sanitize((obj.staff || '').trim(), 100);
    const stLbl = (staff && staff !== 'Unassigned') ? 'Assigned' : 'Unassigned';
    let rowIdx  = _findRow(sh, obj.barcode);
    if (rowIdx < 1) {
      const nr = new Array(37).fill('');
      nr[C.BARCODE      - 1] = obj.barcode;
      nr[C.CREATED_AT   - 1] = nowStr;
      nr[C.LAST_UPDATED - 1] = nowStr;
      nr[CAE.GRANDFATHERED - 1] = false;
      sh.appendRow(nr);
      rowIdx = sh.getLastRow();
    }
    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,   lc],  [C.ASSET_STATUS, asSt], [C.STATUS_LABEL, stLbl],
      [C.EMP_ID,      obj.employeeId  || ''],
      [C.STAFF,       staff           || ''],
      [C.DESIGNATION, obj.designation || ''],
      [C.DEPARTMENT,  obj.department  || ''],
      [C.BASE_OFFICE, obj.baseOffice  || ''],
      [C.DIVISION,    _normDiv(obj.division   || '')],
      [C.DISTRICT,    _normDist(obj.district  || '')],
      [C.AREA,        obj.area        || ''],
      [C.BRANCH,      _sanitize(obj.branch, 150)],
      [C.EFF_DATE,    obj.effDate     || '']
    ]);
    _log('ASSIGN', obj.barcode, staff + ' | ' + lc, obj.employeeId || '');
    return { result: 'Asset assigned successfully', formId: null };

  } catch(e) { return { result: 'Error: ' + e.message, formId: null }; }
  finally    { lock.releaseLock(); }
}

function getPendingCounts(userId, roleTier) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return { myForms: 0, pendingReview: 0 };
    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    const id   = String(userId || '').trim().toLowerCase();
    const myForms = data.filter(r =>
      r[AF.FORM_ID - 1] &&
      String(r[AF.DRAFTED_BY - 1]).trim().toLowerCase() === id &&
      ['Draft','Rejected'].indexOf(String(r[AF.STATUS - 1]).trim()) >= 0
    ).length;
    let pendingReview = 0;
    if (roleTier === 'ho') {
      pendingReview = data.filter(r => r[AF.FORM_ID - 1] && String(r[AF.STATUS - 1]).trim() === 'Pending').length;
    } else if (roleTier === 'senior') {
      const supervisedIds = _getSupervisedEmpIds(userId);
      pendingReview = data.filter(r =>
        r[AF.FORM_ID - 1] &&
        String(r[AF.STATUS - 1]).trim() === 'Pending' &&
        supervisedIds.indexOf(String(r[AF.DRAFTED_BY - 1]).trim().toLowerCase()) >= 0
      ).length;
    }
    return { myForms, pendingReview, total: myForms + pendingReview };
  } catch(e) { return { myForms: 0, pendingReview: 0, total: 0 }; }
}

function getApprovalDashboardData(userId, roleTier) {
  try {
    return {
      myForms:      getMyForms(userId),
      pendingForms: canReviewForms_server(userId, roleTier) ? getPendingForms(userId, roleTier) : [],
      config:       getApprovalConfig(_getContextType(userId))
    };
  } catch(e) { return { myForms: [], pendingForms: [], config: {} }; }
}

function canReviewForms_server(userId, roleTier) { return roleTier === 'senior' || roleTier === 'ho'; }

function _findAFRow(formId) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return -1;
    const ids = sh.getRange(AF_DATA_START, AF.FORM_ID, last - AF_DATA_START + 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(formId).trim()) return i + AF_DATA_START;
    }
    return -1;
  } catch(e) { return -1; }
}

function _findRLRow(formId, drafterId) {
  try {
    const sh   = _rlSheet();
    const last = sh.getLastRow();
    if (last < RL_DATA_START) return null;
    const data = sh.getRange(RL_DATA_START, 1, last - RL_DATA_START + 1, 7).getValues();
    const fid  = String(formId || '').trim();
    const did  = String(drafterId || '').trim().toLowerCase();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === fid && String(data[i][1]).trim().toLowerCase() === did) return data[i];
    }
    return null;
  } catch(e) { return null; }
}

function _findRLRowIdx(formId, drafterId) {
  try {
    const sh   = _rlSheet();
    const last = sh.getLastRow();
    if (last < RL_DATA_START) return -1;
    const ids = sh.getRange(RL_DATA_START, 1, last - RL_DATA_START + 1, 2).getValues();
    const fid = String(formId || '').trim();
    const did = String(drafterId || '').trim().toLowerCase();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === fid && String(ids[i][1]).trim().toLowerCase() === did) return i + RL_DATA_START;
    }
    return -1;
  } catch(e) { return -1; }
}

function _mapAfRow(row) {
  return {
    formId:           String(row[AF.FORM_ID        - 1] || ''),
    formType:         String(row[AF.FORM_TYPE      - 1] || ''),
    linkedFormId:     String(row[AF.LINKED_FORM_ID - 1] || ''),
    contextType:      String(row[AF.CONTEXT_TYPE   - 1] || 'Field'),
    empId:            String(row[AF.EMP_ID         - 1] || ''),
    staffName:        String(row[AF.STAFF_NAME     - 1] || ''),
    designation:      String(row[AF.DESIGNATION    - 1] || ''),
    department:       String(row[AF.DEPARTMENT     - 1] || ''),
    branch:           String(row[AF.BRANCH         - 1] || ''),
    division:         String(row[AF.DIVISION       - 1] || ''),
    district:         String(row[AF.DISTRICT       - 1] || ''),
    assetsJson:       String(row[AF.ASSETS_JSON    - 1] || '[]'),
    status:           String(row[AF.STATUS         - 1] || 'Draft'),
    draftedBy:        String(row[AF.DRAFTED_BY     - 1] || ''),
    draftedAt:        String(row[AF.DRAFTED_AT     - 1] || ''),
    submittedAt:      String(row[AF.SUBMITTED_AT   - 1] || ''),
    supervisorId:     String(row[AF.SUPERVISOR_ID  - 1] || ''),
    reviewedAt:       String(row[AF.REVIEWED_AT    - 1] || ''),
    rejectionComment: String(row[AF.REJECTION_COMMENT - 1] || '')
  };
}

function _updateAssetApprovalStatus(formId, newStatus, rejectionComment) {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return;
    const formIds = sh.getRange(AE_DATA_START, CAE.FORM_ID, last - AE_DATA_START + 1, 1).getValues();
    const fid     = String(formId).trim();
    for (let i = 0; i < formIds.length; i++) {
      if (String(formIds[i][0]).trim() === fid) {
        const row = i + AE_DATA_START;
        sh.getRange(row, CAE.APPROVAL_STATUS).setValue(newStatus);
        sh.getRange(row, CAE.REJECTION_COMMENT).setValue(rejectionComment || '');
      }
    }
  } catch(e) { Logger.log('_updateAssetApprovalStatus error: ' + e.message); }
}

function _getSupervisedEmpIds(supervisorId) {
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return [];
    const last  = sh.getLastRow();
    const data  = sh.getRange(2, 4, last - 1, 4).getValues();
    const supId = String(supervisorId || '').trim().toLowerCase();
    const ids   = [];
    data.forEach(r => {
      if (String(r[3] || '').trim().toLowerCase() === supId) {
        const feId = String(r[0] || '').trim().toLowerCase();
        if (feId && ids.indexOf(feId) < 0) ids.push(feId);
      }
    });
    return ids;
  } catch(e) { return []; }
}

function getUserList() {
  try {
    const sh   = _ss().getSheetByName(SH_USERS);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 9)
      .getValues().filter(r => String(r[2] || '').trim())
      .map(r => r.map(v => String(v || '')));
  } catch(e) { return []; }
}