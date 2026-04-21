// ═══════════════════════════════════════════════════════════
//  ASSET MANAGEMENT SYSTEM — Code.gs  (v6.0 — Restructured)
//
//  Key changes vs v5.x:
//    • Auth reads from Users sheet (not Masterlist)
//    • Scope reads from Org Structure sheet (not Eng. List V2)
//    • Asset Entry is now the writable Copy (31 clean columns)
//    • NO inline borrow/transfer fields in Asset Entry
//    • Borrow details merged from Borrows sheet in getInitialData()
//    • Disposal PRESERVES location data (user requirement)
//    • Deallocate PRESERVES Division/District/Area/Branch
//    • Event sheets (Borrows/Disposals/Transfers/Log): data at row 4
//    • Spare sheet now writes pool entry log
// ═══════════════════════════════════════════════════════════

const SHEET_ID    = '18tuYQKH2OLLu1NqPJiA28n8n7GNN6XR_SSZXUO4XEe8';
const SH_ENTRY    = 'Asset Entry';     // The writable master registry
const SH_USERS    = 'Users';           // Auth source (row 1=headers, data row 2)
const SH_ORG      = 'Org Structure';   // FE/SFE scope (row 1=headers, data row 2)
const SH_MASTER   = 'Masterlist';      // Employee autofill only
const SH_XFER     = 'Transfers';       // Event log (data row 4)
const SH_BORROW   = 'Borrows';         // Event log (data row 4)
const SH_DISPOSE  = 'Disposals';       // Event log (data row 4)
const SH_LOG      = 'ActivityLog';     // Event log (data row 4)
const SH_DROPDOWN = 'Drop down';
const SH_ALLOC    = 'Allocated';       // Allocation log (data row 2)
const SH_SPARE    = 'Spare';           // Spare pool log (data row 4)

// Asset Entry layout: Row 1=title, Row 2=blank, Row 3=headers, Data row 4
const AE_DATA_START  = 4;
// Event sheets layout: Rows 1-2=branding, Row 3=headers, Data row 4
const EVT_DATA_START = 4;

// ─── 31-COLUMN MAP (Asset Entry / Copy of Asset Entry) ──────────────────────
const C = {
  ENTRY_EMP_ID:  1,   // A - Inputter ID
  ENTRY_NAME:    2,   // B - Inputter Name
  EMP_ID:        3,   // C - Accountable Employee ID
  STAFF:         4,   // D - Accountable Staff
  DESIGNATION:   5,   // E - Designation
  DEPARTMENT:    6,   // F - Department
  BASE_OFFICE:   7,   // G - Base Office
  DIVISION:      8,   // H - Division
  DISTRICT:      9,   // I - District
  AREA:          10,  // J - Area
  BRANCH:        11,  // K - Branch
  ASSIGNMENT:    12,  // L - Assignment (Field Office / Central Office)
  EFF_DATE:      13,  // M - Effectivity Date
  BARCODE:       14,  // N - Barcode ← PRIMARY KEY
  TYPE:          15,  // O - Category
  BRAND:         16,  // P - Brand
  SERIAL:        17,  // Q - Serial
  SPECS:         18,  // R - Specification
  SUPPLIER:      19,  // S - Supplier
  CONDITION:     20,  // T - Condition
  ASSET_LOCATION:21,  // U - Asset Location (col T in doc = physical location)
  LIFECYCLE:     22,  // V - Lifecycle Status
  STATUS_LABEL:  23,  // W - Status Label
  ASSET_STATUS:  24,  // X - Assignment Status
  PURCH_DATE:    25,  // Y - Date of Purchase
  WARRANTY_TERM: 26,  // Z - Warranty Term
  WARRANTY_VAL:  27,  // AA - Warranty Validity
  REMARKS:       28,  // AB - Remarks
  NOTES:         29,  // AC - Notes
  CREATED_AT:    30,  // AD - Timestamp (created)
  LAST_UPDATED:  31,  // AE - Last Updated
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

// ─── UTILITY ─────────────────────────────────────────────────────────────────
function _sanitize(val, maxLen) {
  maxLen = maxLen || 500;
  return String(val || '').trim().substring(0, maxLen);
}

function _normDiv(raw) {
  if (!raw) return '';
  return String(raw).trim().replace(/^DIv/i, m => 'Div');
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
  // 'active' or empty — check whether an employee is assigned
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
    }
  }
  return sh;
}

function _entrySheet() {
  return _getOrCreate(SH_ENTRY, AE_HEADERS);
}
function _xferSheet() {
  return _getOrCreate(SH_XFER, [
    'Barcode','TransferType','FromStaff','FromEmpID','FromDesig',
    'FromDiv','FromDist','FromArea','FromBranch','FromRemarks',
    'ToStaff','ToEmpID','ToDesig','ToDiv','ToDist',
    'ToArea','ToBranch','ToRemarks','EffDate','Status','Timestamp'
  ]);
}
function _borrowSheet() {
  return _getOrCreate(SH_BORROW, [
    'Barcode','BorrowerName','EmpID','Designation',
    'Division','District','Branch',
    'BorrowDate','ExpectedReturn','ActualReturn',
    'Status','Remarks','Timestamp'
  ]);
}
function _disposeSheet() {
  return _getOrCreate(SH_DISPOSE, [
    'Barcode','Reason','DisposedBy','DisposeDate','Remarks','Timestamp'
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
    'Condition','Remarks','Timestamp','Allocated By'
  ]);
}
function _spareSheet() {
  return _getOrCreate(SH_SPARE, [
    'Barcode','Category','Brand','Serial No.','Condition',
    'Purchase Date','Warranty Validity','Supplier',
    'Division','District','Area','Branch','Asset Location',
    'Enrolled By','Timestamp','Status'
  ]);
}

// ─── ROW HELPERS ─────────────────────────────────────────────────────────────
function _findRow(sheet, barcode) {
  if (!barcode) return -1;
  try {
    const finder = sheet.createTextFinder(String(barcode).trim())
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
      if (String(vals[i][0]).trim() === String(barcode).trim())
        return i + AE_DATA_START;
    }
    return -1;
  }
}

function _setRow(sheet, rowIdx, updates) {
  const range  = sheet.getRange(rowIdx, 1, 1, TOTAL_COLS);
  const rowVals = range.getValues()[0];
  updates.forEach(u => { rowVals[u[0] - 1] = (u[1] != null ? u[1] : ''); });
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
// Users sheet roles: User | Supervisor | Super User | Admin | Super Admin
function _mapRoleTier(role) {
  const r = String(role || '').trim().toLowerCase();
  const HO_ROLES = [
    'admin', 'super admin', 'superadmin', 'administrator',
    'super user', 'superuser', 'it admin', 'it administrator',
    'system admin', 'sysadmin', 'it head', 'department head',
    'head', 'manager', 'it manager', 'iictd head', 'dept head'
  ];
  const SENIOR_ROLES = [
    'supervisor', 'senior', 'senior fe', 'senior field engineer',
    'senior engineer', 'sfe', 'sfr', 'field supervisor', 'area supervisor',
    'district supervisor', 'division supervisor', 'team lead', 'team leader',
    'lead engineer'
  ];
  if (HO_ROLES.includes(r))     return 'ho';
  if (SENIOR_ROLES.includes(r)) return 'senior';
  return 'fe';
}

// Defaults: Admin = full access incl. delete. Supervisor & FE = full actions, NO delete.
// Scope (_filterByScope) controls WHAT they see. Perms control WHAT they can do.
// Remarks column (Col I of Users sheet) can override any default.
function _parseRemarks(remarks, roleTier) {
  const r = String(remarks || '').toLowerCase().trim();

  // Scope type detection
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

// ─── MASTERLIST LOOKUP (autofill only) ───────────────────────────────────────
// Masterlist: row 1=headers, data row 2
// Col 0=EmpID, Col 2=Name, Col 4=Division, Col 5=District,
// Col 6=Area, Col 7=BaseOffice, Col 11=Position
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
        division:   _normDiv(String(row[4]  || '').trim()),
        district:   _normDist(String(row[5] || '').trim()),
        area:       String(row[6]  || '').trim(),
        baseOffice: String(row[7]  || '').trim(),
        position:   String(row[11] || '').trim()
      };
    }
  } catch(e) {}
  return {};
}

// ─── ORG STRUCTURE SCOPE ──────────────────────────────────────────────────────
// Org Structure: row 1=headers, data row 2
// Col D(3)=FE_ID | Col E(4)=FE_Name | Col F(5)=FE_Desig
// Col G(6)=Sup_ID | Col H(7)=Sup_Name | Col I(8)=Sup_Desig
// Col K(10)=Division | Col L(11)=District
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

    // HO (admin/superuser): scope = everything in Org Structure
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

    // Supervisor (Senior FE): rows where Sup_ID (col G, index 6) matches
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

    // FE (User): rows where FE_ID (col D, index 3) matches
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
// Users sheet: row 1=headers, data row 2
// ColA=Role | ColB=Password | ColC=EmpID | ColD=Name | ColE=Designation
// ColF=SupID | ColG=SupName | ColH=SupDesig | ColI=Remarks
// Col A=Role  B=Password  C=ID Number  D=Emp Name  E=Designation
// Col F=Supervisor ID  G=Supervisor Name  H=Sup Desig  I=Remarks
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

    const data   = sh.getRange(AE_DATA_START, 1, last - AE_DATA_START + 1, TOTAL_COLS).getValues();
    const result = data
      .filter((row, rowIdx) => {
        const bc = String(row[C.BARCODE - 1] || '').trim();
        const badBc = !bc || bc === '-' || bc === 'N/A' || bc === 'None' || bc === '#N/A';
        if (!badBc) return true;
        // Keep rows with no barcode ONLY if they have some location/staff data
        const hasData = [C.DIVISION,C.DISTRICT,C.BRANCH,C.STAFF,C.EMP_ID,C.DEPARTMENT,C.BASE_OFFICE,C.ASSET_LOCATION]
          .some(col => String(row[col - 1] || '').trim());
        return hasData;
      })
      .map((row, rowIdx) => {
        const get    = col => String(row[col - 1] || '');
        const approvalStatus = get(CAE.APPROVAL_STATUS) || 'Confirmed';
        const grandfathered  = String(get(CAE.GRANDFATHERED)).toLowerCase() === 'true';
        const effectiveStatus = (approvalStatus === 'Confirmed' || grandfathered)
          ? _computeStatus(get(C.LIFECYCLE), get(C.ASSET_STATUS), get(C.EMP_ID))
          : 'pending-approval';

        const rawBc = String(row[C.BARCODE - 1] || '').trim();
        const badBc = !rawBc || rawBc === '-' || rawBc === 'N/A' || rawBc === 'None' || rawBc === '#N/A';
        const syntheticKey = badBc ? ('NOBC-' + (rowIdx + AE_DATA_START)) : rawBc;

        const rawLC    = get(C.LIFECYCLE);
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
          Barcode:       badBc ? syntheticKey : rawBc,
          hasRealBarcode: !badBc,
          Type:         get(C.TYPE),
          Brand:        get(C.BRAND),
          Serial:       get(C.SERIAL),
          Specs:        get(C.SPECS),
          Supplier:     get(C.SUPPLIER),
          Condition:    get(C.CONDITION) || 'Good',
          AssetLocation:get(C.ASSET_LOCATION),
          Lifecycle:    displayLC,
          AssetStatus:  get(C.ASSET_STATUS) || 'Active',
          StatusLabel:  get(C.STATUS_LABEL) || 'Unassigned',
          jsApprovalStatus: approvalStatus,
          FormID:         get(CAE.FORM_ID) || '',
          PurchDate:    get(C.PURCH_DATE),
          WarrantyTerm: get(C.WARRANTY_TERM),
          WarrantyVal:  get(C.WARRANTY_VAL),
          Remarks:      get(C.REMARKS),
          Notes:        get(C.NOTES),
          EmpID:        get(C.EMP_ID) || 'N/A',
          Staff:        get(C.STAFF)  || 'Unassigned',
          Designation:  get(C.DESIGNATION),
          Department:   get(C.DEPARTMENT),
          BaseOffice:   get(C.BASE_OFFICE),
          Assignment:        rawAssignment,
          EffectiveAssignment: effectiveAssignment,
          Division:     _normDiv(get(C.DIVISION)),
          District:     _normDist(get(C.DISTRICT)),
          Area:         get(C.AREA),
          Branch:       get(C.BRANCH),
          EffDate:      get(C.EFF_DATE),
          CreatedAt:    get(C.CREATED_AT),
          LastUpdated:  get(C.LAST_UPDATED),
          EntryEmpId:   get(C.ENTRY_EMP_ID),
          EntryName:    get(C.ENTRY_NAME),
          // Borrow fields populated by getInitialData() merge
          BorName:'', BorEmpID:'', BorDesig:'', BorDiv:'', BorDist:'',
          BorBranch:'', BorDate:'', ExpReturn:'', ActReturn:'', BorRemarks:'',
          // Transfer display fields populated by getInitialData() merge
          XferType:'', ToStaff:'', ToEmpID:'', ToDiv:'', ToBranch:'', XferDate:'',
          status: effectiveStatus
        };
      });

    return { success: true, data: result };
  } catch(e) { return { success: false, error: e.message }; }
}

function getAssetByBarcode(barcode) {
  try {
    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return null;
    const row    = sh.getRange(rowIdx, 1, 1, TOTAL_COLS).getValues()[0];
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

// ─── ASSETS: CREATE ───────────────────────────────────────────────────────────
function processAsset_legacy(obj, isAssign) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return 'Error: System busy — try again.'; }

  try {
    const sh     = _entrySheet();
    const nowStr = new Date().toLocaleString('en-PH');

    if (!isAssign) {
      // ── ENROLL NEW ASSET ────────────────────────────────────────────────
      if (!obj.barcode) return 'Error: Barcode is required.';
      if (_findRow(sh, obj.barcode) > 0) return 'Error: Barcode already exists: ' + obj.barcode;

      // Serial duplicate check
      if (obj.serial) {
        const last2 = sh.getLastRow();
        if (last2 >= AE_DATA_START) {
          const serials = sh.getRange(
            AE_DATA_START, C.SERIAL, last2 - AE_DATA_START + 1, 1).getValues();
          const dup = serials.findIndex(r => String(r[0]).trim() === String(obj.serial).trim());
          if (dup >= 0) {
            const existBC = String(sh.getRange(dup + AE_DATA_START, C.BARCODE).getValue());
            return 'Error: Serial No. already registered under barcode: ' + existBC;
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
      const isSpare = (statusChoice === 'Spare' || statusChoice === 'BorrowItem');
      const normDiv  = _normDiv(obj.division  || '');
      const normDist = _normDist(obj.district || '');

      const row = new Array(TOTAL_COLS).fill('');
      row[C.ENTRY_EMP_ID  - 1] = obj.entryEmpId   || '';
      row[C.ENTRY_NAME    - 1] = obj.entryName     || '';
      row[C.EMP_ID        - 1] = isSpare ? '' : (obj.accEmpId || '');
      row[C.STAFF         - 1] = isSpare ? '' : _sanitize(obj.accName, 100);
      row[C.DESIGNATION   - 1] = isSpare ? '' : (obj.accRole || '');
      row[C.DEPARTMENT    - 1] = obj.department    || '';
      row[C.BASE_OFFICE   - 1] = obj.baseOffice    || '';
      row[C.DIVISION      - 1] = normDiv;
      row[C.DISTRICT      - 1] = normDist;
      row[C.AREA          - 1] = obj.area          || '';
      row[C.BRANCH        - 1] = _sanitize(obj.branch, 150);
      row[C.ASSIGNMENT    - 1] = obj.assignment    || 'Field Office';
      row[C.EFF_DATE      - 1] = obj.effDate       || '';
      row[C.BARCODE       - 1] = obj.barcode;
      row[C.TYPE          - 1] = obj.type          || '';
      row[C.BRAND         - 1] = obj.brand         || '';
      row[C.SERIAL        - 1] = obj.serial ? String(obj.serial) : '';
      row[C.SPECS         - 1] = obj.specs         || '';
      row[C.SUPPLIER      - 1] = obj.supplier      || '';
      row[C.CONDITION     - 1] = obj.condition     || 'New';
      row[C.ASSET_LOCATION- 1] = obj.location      || '';
      row[C.LIFECYCLE     - 1] = sm.lc;
      row[C.STATUS_LABEL  - 1] = sm.sl;
      row[C.ASSET_STATUS  - 1] = sm.as;
      row[C.PURCH_DATE    - 1] = obj.purchDate     || '';
      row[C.WARRANTY_TERM - 1] = obj.wTerm         || '';
      row[C.WARRANTY_VAL  - 1] = obj.wValidity     || '';
      row[C.REMARKS       - 1] = _sanitize(obj.remarks, 500);
      row[C.CREATED_AT    - 1] = nowStr;
      row[C.LAST_UPDATED  - 1] = nowStr;

      sh.appendRow(row);
      const newRowIdx = sh.getLastRow();
      sh.getRange(newRowIdx, C.SERIAL).setNumberFormat('@STRING@');
      if (obj.serial) sh.getRange(newRowIdx, C.SERIAL).setValue(String(obj.serial));

      // Write to Spare log sheet if applicable
      if (statusChoice === 'Spare') {
        _spareSheet().appendRow([
          obj.barcode, obj.type, obj.brand, obj.serial || '', obj.condition || 'New',
          obj.purchDate || '', obj.wValidity || '', obj.supplier || '',
          normDiv, normDist, obj.area || '', _sanitize(obj.branch, 150),
          obj.location || '', obj.enrolledBy || obj.entryEmpId || '',
          nowStr, 'Available'
        ]);
      }

      _log('CREATE', obj.barcode,
        obj.type + ' | ' + obj.brand + ' | ' + statusChoice,
        obj.entryEmpId || '');
      return 'Asset created: ' + obj.barcode;
    }

    // ── isAssign = true (update lifecycle from elsewhere) ──────────────────
    const lc    = obj.lifecycle || 'Allocated';
    const asSt  = lc === 'Transfer' ? 'Transfer' :
                  lc === 'Dispose'  ? 'Disposal' : 'Active';
    const staff = _sanitize((obj.staff || '').trim(), 100);
    const stLbl = (staff && staff !== 'Unassigned') ? 'Assigned' : 'Unassigned';
    let rowIdx  = _findRow(sh, obj.barcode);

    if (rowIdx < 1) {
      const nr = new Array(TOTAL_COLS).fill('');
      nr[C.BARCODE     - 1] = obj.barcode;
      nr[C.CREATED_AT  - 1] = nowStr;
      nr[C.LAST_UPDATED- 1] = nowStr;
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
    return 'Asset assigned successfully';

  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

// ─── DELETE ASSETS ────────────────────────────────────────────────────────────
function deleteAssets(barcodes, callerEmpId) {
  // Security: verify caller role from Users sheet
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

// ─── ALLOCATE ─────────────────────────────────────────────────────────────────
function allocateAsset(obj) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    if (!obj.barcode)   return 'Error: Barcode is required.';
    if (!obj.empId)     return 'Error: Employee ID is required.';
    if (!obj.staffName) return 'Error: Staff name is required.';

    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, obj.barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + obj.barcode;

    const curRow = sh.getRange(rowIdx, 1, 1, TOTAL_COLS).getValues()[0];
    const curLC  = String(curRow[C.LIFECYCLE - 1] || '').toLowerCase();
    const curAS  = String(curRow[C.ASSET_STATUS - 1] || '').toLowerCase();

    if (curLC === 'borrow')    return 'Error: Asset is currently borrowed. Return it first.';
    if (curLC === 'dispose' || curLC === 'disposal')
                               return 'Error: Disposed assets cannot be allocated.';
    if (curLC === 'transfer')  return 'Error: Asset is in an active transfer.';
    if (curAS === 'borrowitem')
      return 'Error: This asset is in the Borrow Pool. Change its status first.';

    if (curLC === 'allocated') {
      const prevStaff = String(curRow[C.STAFF - 1] || '');
      _log('DEALLOCATE', obj.barcode,
        'Implicit dealloc from: ' + prevStaff + ' → re-allocate to: ' + obj.staffName,
        obj.allocatedBy || obj.empId || '');
    }

    const nowStr   = new Date().toLocaleString('en-PH');
    const normDiv  = _normDiv(obj.division  || '');
    const normDist = _normDist(obj.district || '');

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'],   [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],    [C.EMP_ID,       obj.empId           || ''],
      [C.STAFF,        _sanitize(obj.staffName, 100)],
      [C.DESIGNATION,  obj.designation || ''],
      [C.DEPARTMENT,   obj.department  || ''],
      [C.BASE_OFFICE,  obj.baseOffice  || ''],
      [C.DIVISION,     normDiv],        [C.DISTRICT,     normDist],
      [C.AREA,         obj.area        || ''],
      [C.BRANCH,       _sanitize(obj.branch, 150)],
      [C.EFF_DATE,     obj.effDate     || nowStr],
      [C.REMARKS,      _sanitize(obj.remarks, 500)]
    ]);

    // Write to Allocated log
    _allocLogSheet().appendRow([
      obj.barcode,
      obj.type      || String(curRow[C.TYPE  - 1] || ''),
      obj.brand     || String(curRow[C.BRAND - 1] || ''),
      obj.serial    || String(curRow[C.SERIAL- 1] || ''),
      obj.empId, _sanitize(obj.staffName, 100),
      obj.designation || '', obj.department || '', obj.baseOffice || '',
      normDiv, normDist, obj.area || '', _sanitize(obj.branch, 150),
      obj.effDate || nowStr,
      obj.condition || String(curRow[C.CONDITION - 1] || 'Good'),
      _sanitize(obj.remarks, 500), nowStr, obj.allocatedBy || ''
    ]);

    _log('ALLOCATE', obj.barcode,
      _sanitize(obj.staffName, 100) + ' | ' + (_sanitize(obj.branch, 150) || normDiv || ''),
      obj.allocatedBy || obj.empId || '');
    return 'Asset allocated to ' + _sanitize(obj.staffName, 100);
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

// ─── DEALLOCATE (return to spare) ─────────────────────────────────────────────
// NOTE: Division/District/Area/Branch are PRESERVED so you know where the
//       spare asset is physically located.
function deallocateAsset(barcode, remarks) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;

    const curRow = sh.getRange(rowIdx, 1, 1, TOTAL_COLS).getValues()[0];
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
      // Division/District/Area/Branch intentionally kept — shows where spare is stored
    ];
    if (remarks) updates.push([C.REMARKS, _sanitize(remarks, 500)]);
    _setRow(sh, rowIdx, updates);

    // Write to Spare log
    _spareSheet().appendRow([
      barcode,
      String(curRow[C.TYPE          - 1] || ''),
      String(curRow[C.BRAND         - 1] || ''),
      String(curRow[C.SERIAL        - 1] || ''),
      String(curRow[C.CONDITION     - 1] || 'Good'),
      String(curRow[C.PURCH_DATE    - 1] || ''),
      String(curRow[C.WARRANTY_VAL  - 1] || ''),
      String(curRow[C.SUPPLIER      - 1] || ''),
      String(curRow[C.DIVISION      - 1] || ''),
      String(curRow[C.DISTRICT      - 1] || ''),
      String(curRow[C.AREA          - 1] || ''),
      String(curRow[C.BRANCH        - 1] || ''),
      String(curRow[C.ASSET_LOCATION- 1] || ''),
      prevStaff, nowStr, 'Available'
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

// ─── TRANSFERS ────────────────────────────────────────────────────────────────
function saveTransfer_legacy(t) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    if (!t.barcode)      return 'Error: Barcode is required.';
    if (!t.toEmpId)      return 'Error: Destination Employee ID is required.';
    if (!t.toStaff)      return 'Error: Destination Staff Name is required.';
    if (!t.effDate)      return 'Error: Transfer Date is required.';
    if (!t.transferType) return 'Error: Transfer Type is required.';

    const sh     = _entrySheet();
    const rowIdx = _findRow(sh, t.barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + t.barcode;

    const curLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
    if (curLC === 'dispose' || curLC === 'disposal')
      return 'Error: Cannot transfer a disposed asset.';
    if (curLC === 'borrow')
      return 'Error: Return the asset first before transferring.';

    const nowStr   = new Date().toLocaleString('en-PH');
    const normToDiv  = _normDiv(t.toDiv   || '');
    const normToDist = _normDist(t.toDist || '');

    // Write full record to Transfers sheet (21 columns — unchanged structure)
    _xferSheet().appendRow([
      t.barcode, t.transferType,
      _sanitize(t.fromStaff, 100), t.fromEmpId, t.fromDesig,
      t.fromDiv, t.fromDist, t.fromArea, _sanitize(t.fromBranch, 150), _sanitize(t.fromRemarks, 500),
      _sanitize(t.toStaff, 100), t.toEmpId, t.toDesig,
      normToDiv, normToDist, t.toArea, _sanitize(t.toBranch, 150), _sanitize(t.toRemarks, 500),
      t.effDate, t.status || 'Completed', nowStr
    ]);

    // Update Asset Entry with new holder (no inline transfer fields in 31-col layout)
    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'],  [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],   [C.EMP_ID,       t.toEmpId   || ''],
      [C.STAFF,        _sanitize(t.toStaff, 100)],
      [C.DESIGNATION,  t.toDesig   || ''],
      [C.DIVISION,     normToDiv],    [C.DISTRICT,     normToDist],
      [C.AREA,         t.toArea    || ''],
      [C.BRANCH,       _sanitize(t.toBranch, 150)],
      [C.EFF_DATE,     t.effDate]
    ]);

    _log('TRANSFER', t.barcode,
      (_sanitize(t.fromStaff, 100) || '—') + ' → ' + _sanitize(t.toStaff, 100),
      t.fromEmpId || '');
    return 'Transfer saved';
  } catch(e) { return 'Error: ' + e.message; }
  finally    { lock.releaseLock(); }
}

function getTransferData() {
  try {
    const sh   = _xferSheet();
    const last = sh.getLastRow();
    if (last < EVT_DATA_START) return [];
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 21).getValues()
      .filter(r => r[0])
      .map(r => r.map(v => String(v || '')));
  } catch(e) { return []; }
}

// ─── BORROWS ──────────────────────────────────────────────────────────────────
// saveBorrow: ONLY updates Lifecycle in Asset Entry.
// All borrower details go to the Borrows sheet only.
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

    const curRow  = sh.getRange(rowIdx, 1, 1, TOTAL_COLS).getValues()[0];
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

    // Write to Borrows sheet (13 columns matching fixed headers)
    bSh.appendRow([
      b.barcode, _sanitize(b.borrowerName, 100), b.empId || '', b.designation || '',
      b.division || '', b.district || '', _sanitize(b.branch, 150),
      b.borrowDate, b.expectedReturn, '',
      'Borrow', _sanitize(b.remarks, 500), nowStr
    ]);

    // Update Asset Entry — only Lifecycle; keep staff/location intact
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
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 13).getValues()
      .filter(r => r[0])
      .map(r => ({
        barcode:        String(r[0]  || ''),
        borrowerName:   String(r[1]  || ''),
        empId:          String(r[2]  || ''),
        designation:    String(r[3]  || ''),
        division:       String(r[4]  || ''),
        district:       String(r[5]  || ''),
        branch:         String(r[6]  || ''),
        borrowDate:     String(r[7]  || ''),
        expectedReturn: String(r[8]  || ''),
        actualReturn:   String(r[9]  || ''),
        status:         String(r[10] || 'Borrow'),
        remarks:        String(r[11] || ''),
        timestamp:      String(r[12] || '')
      }));
  } catch(e) { return []; }
}

function returnAsset(barcode, returnDate) {
  try {
    const sh      = _entrySheet();
    const bSh     = _borrowSheet();
    const retDate = returnDate || new Date().toLocaleDateString('en-PH');
    const last    = bSh.getLastRow();

    // Update Borrows sheet — find the active borrow row
    if (last >= EVT_DATA_START) {
      const data = bSh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 13).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === String(barcode) && String(data[i][10]) === 'Borrow') {
          const sheetRow = i + EVT_DATA_START;
          bSh.getRange(sheetRow, 10).setValue(retDate);  // Col J = ActualReturn
          bSh.getRange(sheetRow, 11).setValue('Returned'); // Col K = Status
          break;
        }
      }
    }

    // Update Asset Entry — restore previous lifecycle
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx > 0) {
      const curRow   = sh.getRange(rowIdx, 1, 1, TOTAL_COLS).getValues()[0];
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
// NOTE: Location data (Division/District/Area/Branch/Staff) is PRESERVED
//       in Asset Entry so you can see WHERE the disposed asset was/is.
function saveDisposal(d) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }
  try {
    const sh     = _entrySheet();
    const dSh    = _disposeSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const rowIdx = _findRow(sh, d.barcode);

    if (rowIdx > 0) {
      const curLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
      if (curLC === 'borrow')   return 'Error: Cannot dispose a borrowed asset.';
      if (curLC === 'transfer') return 'Error: Cannot dispose an asset in active transfer.';
    }

    // Write disposal record (6 columns matching fixed headers)
    dSh.appendRow([
      d.barcode, _sanitize(d.reason, 200), _sanitize(d.disposedBy, 100),
      d.disposeDate, _sanitize(d.remarks, 500), nowStr
    ]);

    if (rowIdx > 0) {
      // Only update lifecycle/status — ALL location & staff info is preserved
      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    'Dispose'],
        [C.ASSET_STATUS, 'Disposal'],
        [C.STATUS_LABEL, 'Disposed']
        // Division, District, Area, Branch, Staff, EmpID all KEPT intact
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
    return sh.getRange(EVT_DATA_START, 1, last - EVT_DATA_START + 1, 6).getValues()
      .filter(r => r[0])
      .map(r => r.map(v => String(v || '')));
  } catch(e) { return []; }
}

// ─── USERS ────────────────────────────────────────────────────────────────────
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
        ok:       true,
        empId:    String(row[0]  || '').trim(),
        name:     String(row[2]  || '').trim(),
        division: _normDiv(String(row[4]  || '').trim()),
        district: _normDist(String(row[5] || '').trim()),
        area:     String(row[6]  || '').trim(),
        branch:   String(row[7]  || '').trim(),
        position: String(row[11] || '').trim()
      };
    }
    return { ok: false, error: 'Employee ID not found: ' + empId };
  } catch(e) { return { ok: false, error: e.message }; }
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
        result.models[cat.name + '|' + brand] = models;
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
        if (s) result.suppliers.push(s);
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

// ─── LOCATION DATA ────────────────────────────────────────────────────────────
// Now reads from Org Structure instead of Eng. List V2
function getLocationData(empId) {
  try {
    let scopeData;
    if (empId) {
      // Determine role tier from Users sheet
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

// ─── ENGINEER LOOKUP (for Accountability Form) ────────────────────────────────
// Reads from Org Structure: Col E (index 4) = FE Name, Col H (index 7) = Sup Name
// Col L (index 11) = District
function getEngineersByLocation(district, branch) {
  try {
    const sh     = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return { fe:'', senior:'' };
    const last   = sh.getLastRow();
    const data   = sh.getRange(2, 1, last - 1, 12).getValues();
    const normDist = String(district || '').trim().toLowerCase();
    const result = { fe:'', senior:'' };

    data.forEach(r => {
      const rowDist = String(r[11] || '').trim().toLowerCase(); // Col L
      if (!normDist || rowDist !== normDist) return;
      if (!result.fe)     result.fe     = String(r[4] || '').trim(); // Col E = FE Name
      if (!result.senior) result.senior = String(r[7] || '').trim(); // Col H = Sup Name
    });
    return result;
  } catch(e) { return { fe:'', senior:'' }; }
}

// ─── ORG LOOKUP (for resolving Asset Location column back to Div/Dist) ────────
function _buildOrgLookup() {
  const lookup = {};
  try {
    const sh = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return lookup;
    const last = sh.getLastRow();
    const data = sh.getRange(2, 1, last - 1, 12).getValues();

    data.forEach(r => {
      const div  = _normDiv(String(r[10] || '').trim());   // Col K
      const dist = _normDist(String(r[11] || '').trim());  // Col L
      if (!div || !dist) return;

      // District → Division mapping (first occurrence wins)
      if (!lookup[dist.toLowerCase()])
        lookup[dist.toLowerCase()] = { division:div, district:dist, area:'', branch:'' };

      // Division → mapping
      if (!lookup[div.toLowerCase()])
        lookup[div.toLowerCase()] = { division:div, district:'', area:'', branch:'' };
    });

    // Also try to read from Org Structure-FO if it exists (branch-level data)
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
  } catch(e) {}
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

  // ── Merge active borrow details into asset objects ─────────────────────────
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

  // ── Merge latest transfer details for assets in transfer state ────────────
  // Transfers sheet columns: 0=Barcode, 1=Type, 10=ToStaff, 11=ToEmpID,
  //                          13=ToDiv, 16=ToBranch, 18=EffDate, 19=Status
  const xferMap = {};
  transfers.forEach(r => {
    if (!xferMap[r[0]] || r[20] > (xferMap[r[0]][20] || ''))
      xferMap[r[0]] = r; // keep latest by timestamp
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
    const allData = sh.getRange(AE_DATA_START, 1, count, TOTAL_COLS).getValues();
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
      sh.getRange(rowIdx, 1, 1, TOTAL_COLS).setValues([rowData]);
    });

    _log('MOVE_STAFF', empId,
      'Action:' + assetAction + ' → ' + newDiv + '/' + newDist + '/' + newBranch +
      ' | ' + updated + ' updated, ' + skipped + ' skipped', empId);

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
    const data   = sh.getRange(AE_DATA_START, 1, count, TOTAL_COLS).getValues();
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

    _log('MOVE_ORG', unitType.toUpperCase(),
      currentDiv + '/' + currentDist + '/' + currentArea + '/' + currentBranch +
      ' → ' + newDiv + '/' + newDist + '/' + newArea + ' | ' + updated + ' assets', '');

    return updated > 0
      ? unitType.charAt(0).toUpperCase() + unitType.slice(1) +
        ' moved. ' + updated + ' asset(s) updated.'
      : 'Move recorded — no matching assets found.';
  } catch(e) { return 'Error: ' + e.message; }
}

// ═══════════════════════════════════════════════════════════════════════════
//  ACCOUNTABILITY FORM WORKFLOW — Parts 1 & 2
//  Append this entire block to the bottom of Code.gs
//
//  DEPENDENCIES (already in Code.gs):
//    _ss(), _sanitize(), _log(), SHEET_ID, SH_USERS, SH_ORG, SH_MASTER
//    C (column map), AE_DATA_START, EVT_DATA_START, _findRow(), _setRow()
// ═══════════════════════════════════════════════════════════════════════════


// ─────────────────────────────────────────────────────────────────────────────
//  PART 1A — SHEET NAME CONSTANTS & COLUMN MAPS
// ─────────────────────────────────────────────────────────────────────────────

const SH_AF  = 'AccountabilityForms';
const SH_FS  = 'FormSnapshots';
const SH_RL  = 'FormRateLimits';
const SH_CFG = 'ApprovalConfig';

// New Asset Entry columns (32–37) — append to existing C map
// NOTE: Do not redefine C. Reference these directly in functions below.
const CAE = {
  APPROVAL_STATUS:    32,  // Draft | Pending | Confirmed | Rejected
  FORM_ID:            33,  // FK to AccountabilityForms
  DRAFTED_BY:         34,  // Employee ID of drafter
  DRAFTED_AT:         35,  // Timestamp
  REJECTION_COMMENT:  36,  // Rejection reason shown to drafter
  GRANDFATHERED:      37   // TRUE for all pre-existing assets
};

// AccountabilityForms sheet column positions (data starts row 4)
const AF = {
  FORM_ID:        1,   // A
  FORM_TYPE:      2,   // B  Enrollment | Transfer-From | Transfer-To
  LINKED_FORM_ID: 3,   // C  Transfer pair partner FormID
  CONTEXT_TYPE:   4,   // D  Field | HO
  EMP_ID:         5,   // E
  STAFF_NAME:     6,   // F
  DESIGNATION:    7,   // G
  DEPARTMENT:     8,   // H
  BRANCH:         9,   // I  Branch / Base Office
  DIVISION:       10,  // J
  DISTRICT:       11,  // K
  ASSETS_JSON:    12,  // L  JSON snapshot of asset rows
  STATUS:         13,  // M  Draft | Pending | Confirmed | Rejected
  DRAFTED_BY:     14,  // N  Employee ID
  DRAFTED_AT:     15,  // O  Timestamp
  SUBMITTED_AT:   16,  // P  Timestamp
  SUPERVISOR_ID:  17,  // Q  Employee ID of reviewer
  REVIEWED_AT:    18,  // R  Timestamp
  REJECTION_COMMENT: 19  // S
};
const AF_TOTAL_COLS = 19;
const AF_DATA_START = 4;

// FormRateLimits sheet column positions (data starts row 4)
const RL = {
  FORM_ID:          1,  // A
  DRAFTED_BY:       2,  // B
  RESUBMIT_COUNT:   3,  // C
  WINDOW_START:     4,  // D
  COOLDOWN_UNTIL:   5,  // E
  LAST_SUBMITTED:   6,  // F
  STATUS:           7   // G  Active | Cooldown | Clear
};
const RL_DATA_START = 4;

// FormSnapshots sheet column positions (data starts row 4)
const FS = {
  FORM_ID:          1,  // A
  FORM_TYPE:        2,  // B
  LINKED_FORM_ID:   3,  // C
  CONTEXT_TYPE:     4,  // D
  EMP_ID:           5,  // E
  STAFF_NAME:       6,  // F
  DESIGNATION:      7,  // G
  DEPARTMENT:       8,  // H
  BRANCH:           9,  // I
  ASSETS_JSON:      10, // J
  CONFIRMED_BY:     11, // K
  CONFIRMED_AT:     12, // L
  SUPERSEDED_AT:    13, // M
  SUPERSEDED_BY:    14  // N
};
const FS_DATA_START = 4;

// Rate limit constants (must match ApprovalConfig sheet)
const RL_MAX_RESUBMITS    = 5;
const RL_WINDOW_MINUTES   = 30;
const RL_COOLDOWN_MINUTES = 120;


// ─────────────────────────────────────────────────────────────────────────────
//  PART 1A — SHEET HELPER FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

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


// ─────────────────────────────────────────────────────────────────────────────
//  PART 1B — READ HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns signatory config for a given context type ('Field' or 'HO').
 * Reads from ApprovalConfig sheet rows 5-6 (signatory rules section).
 */
function getApprovalConfig(contextType) {
  try {
    const sh   = _ss().getSheetByName(SH_CFG);
    if (!sh) return _defaultApprovalConfig(contextType);
    const last = sh.getLastRow();
    if (last < 5) return _defaultApprovalConfig(contextType);

    const data = sh.getRange(5, 1, 2, 4).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === contextType.toLowerCase()) {
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
  } catch(e) {
    return _defaultApprovalConfig(contextType);
  }
}

function _defaultApprovalConfig(contextType) {
  if ((contextType || '').toLowerCase() === 'ho') {
    return {
      contextType:      'HO',
      processedByLabel: 'Technical Support Engineer',
      checkedByLabel:   'Senior Technical Support Engineer',
      verifiedByName:   'Sandylee Dela Cruz Paris',
      notedByName:      'Patrick Gerard G. Reyes',
      notedByTitle:     'Department Head'
    };
  }
  return {
    contextType:      'Field',
    processedByLabel: 'Field Engineer',
    checkedByLabel:   'Senior Field Engineer',
    verifiedByName:   'Maricon B. Jaropillo',
    notedByName:      'Patrick Gerard G. Reyes',
    notedByTitle:     'Department Head'
  };
}

/**
 * Returns all forms drafted by a specific employee.
 * Used by the "My Forms" tab on the For Approval page.
 */
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
  } catch(e) {
    return [];
  }
}

/**
 * Returns all Pending forms visible to a supervisor.
 * Scoped strictly: only forms where the drafter is supervised by supervisorId
 * in the Org Structure sheet.
 */
function getPendingForms(supervisorId, roleTier) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return [];

    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    let pending = data
      .filter(r => r[AF.FORM_ID - 1] && String(r[AF.STATUS - 1]).trim() === 'Pending')
      .map(r => _mapAfRow(r));

    // HO admin sees all pending forms
    if (roleTier === 'ho') return pending;

    // Supervisor sees only forms from their directly supervised staff
    const supervisedIds = _getSupervisedEmpIds(supervisorId);
    return pending.filter(f => supervisedIds.indexOf(f.draftedBy.toLowerCase()) >= 0);
  } catch(e) {
    return [];
  }
}

/**
 * Returns a single form record by FormID.
 */
function getFormDetail(formId) {
  try {
    const sh   = _afSheet();
    const last = sh.getLastRow();
    if (last < AF_DATA_START) return null;

    const data = sh.getRange(AF_DATA_START, 1, last - AF_DATA_START + 1, AF_TOTAL_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][AF.FORM_ID - 1]).trim() === String(formId).trim()) {
        return _mapAfRow(data[i]);
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

/**
 * Returns badge counts for the For Approval page nav item.
 * Drafters: count of their Draft + Rejected forms.
 * Supervisors: count of Pending forms in their scope.
 * Admins: count of all Pending forms.
 */
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
      pendingReview = data.filter(r =>
        r[AF.FORM_ID - 1] && String(r[AF.STATUS - 1]).trim() === 'Pending'
      ).length;
    } else if (roleTier === 'senior') {
      const supervisedIds = _getSupervisedEmpIds(userId);
      pendingReview = data.filter(r =>
        r[AF.FORM_ID - 1] &&
        String(r[AF.STATUS - 1]).trim() === 'Pending' &&
        supervisedIds.indexOf(String(r[AF.DRAFTED_BY - 1]).trim().toLowerCase()) >= 0
      ).length;
    }

    return { myForms, pendingReview, total: myForms + pendingReview };
  } catch(e) {
    return { myForms: 0, pendingReview: 0, total: 0 };
  }
}

/**
 * Returns the rate limit status for a given formId + drafter.
 * Called by frontend before showing the Resubmit button.
 */
function getRateLimitStatus(formId, drafterId) {
  try {
    const row = _findRLRow(formId, drafterId);
    if (!row) return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null };

    const now          = new Date();
    const cooldownUntil= row[RL.COOLDOWN_UNTIL - 1] ? new Date(row[RL.COOLDOWN_UNTIL - 1]) : null;

    if (cooldownUntil && now < cooldownUntil) {
      return { allowed: false, remaining: 0, cooldownUntil: cooldownUntil.toISOString() };
    }

    const windowStart = row[RL.WINDOW_START - 1] ? new Date(row[RL.WINDOW_START - 1]) : null;
    const windowExpired = !windowStart ||
      ((now - windowStart) / 60000) > RL_WINDOW_MINUTES;

    if (windowExpired) {
      return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null };
    }

    const count     = parseInt(row[RL.RESUBMIT_COUNT - 1] || 0, 10);
    const remaining = Math.max(0, RL_MAX_RESUBMITS - count);
    return { allowed: remaining > 0, remaining, cooldownUntil: null };
  } catch(e) {
    return { allowed: true, remaining: RL_MAX_RESUBMITS, cooldownUntil: null };
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  PART 1C — FORM ID GENERATOR
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Generates next sequential FormID in format FORM-YYYY-NNN.
 * Uses a script lock to prevent duplicate IDs.
 */
function generateFormID() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); }
  catch(e) { return 'Error: System busy — try again.'; }

  try {
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

    const seq  = max + 1;
    return 'FORM-' + yr + '-' + String(seq).padStart(3, '0');
  } finally {
    lock.releaseLock();
  }
}


// ─────────────────────────────────────────────────────────────────────────────
//  PART 1D — RATE LIMIT ENGINE
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Checks and updates the rate limit for a form resubmission.
 * Returns { allowed, remaining, cooldownUntil, message }
 * Writes to FormRateLimits sheet as a side effect when allowed.
 */
function checkRateLimit(formId, drafterId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(6000); }
  catch(e) { return { allowed: false, message: 'System busy — try again.' }; }

  try {
    const sh  = _rlSheet();
    const now = new Date();
    const rowIdx = _findRLRowIdx(formId, drafterId);

    // No record yet — first submission is always allowed
    if (rowIdx < 0) {
      sh.appendRow([
        formId, drafterId, 1,
        now.toLocaleString('en-PH'),  // window start
        '',                            // no cooldown
        now.toLocaleString('en-PH'),  // last submitted
        'Active'
      ]);
      return { allowed: true, remaining: RL_MAX_RESUBMITS - 1, cooldownUntil: null };
    }

    const data = sh.getRange(rowIdx, 1, 1, 7).getValues()[0];
    const cooldownRaw = data[RL.COOLDOWN_UNTIL - 1];
    const cooldownUntil = cooldownRaw ? new Date(cooldownRaw) : null;

    // Blocked by cooldown
    if (cooldownUntil && now < cooldownUntil) {
      const mins = Math.ceil((cooldownUntil - now) / 60000);
      return {
        allowed:       false,
        remaining:     0,
        cooldownUntil: cooldownUntil.toISOString(),
        message:       'Rate limit reached. Try again in ' + mins + ' minute(s).'
      };
    }

    const windowStart   = data[RL.WINDOW_START - 1] ? new Date(data[RL.WINDOW_START - 1]) : null;
    const windowExpired = !windowStart ||
      ((now - windowStart) / 60000) > RL_WINDOW_MINUTES;

    let count = parseInt(data[RL.RESUBMIT_COUNT - 1] || 0, 10);

    // Reset window if expired
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
      return {
        allowed:       false,
        remaining:     0,
        cooldownUntil: cooldownEnd.toISOString(),
        message:       'Submission limit reached. You can resubmit after 2 hours.'
      };
    }

    sh.getRange(rowIdx, RL.RESUBMIT_COUNT).setValue(count);
    sh.getRange(rowIdx, RL.LAST_SUBMITTED).setValue(now.toLocaleString('en-PH'));

    return {
      allowed:       true,
      remaining:     RL_MAX_RESUBMITS - count,
      cooldownUntil: null,
      message:       'Submitted. ' + (RL_MAX_RESUBMITS - count) + ' attempt(s) remaining in this window.'
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Clears the rate limit record for a form after successful confirmation.
 * Prevents old counts from carrying over if the form is ever re-used.
 */
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


// ─────────────────────────────────────────────────────────────────────────────
//  PART 2A — DRAFT CREATION
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Creates a new Draft accountability form in AccountabilityForms sheet.
 *
 * @param {string} empId        - Accountable staff Employee ID
 * @param {Array}  assets       - Array of asset objects (from ASSETS state)
 * @param {string} formType     - 'Enrollment' | 'Transfer-From' | 'Transfer-To'
 * @param {string} linkedFormId - Partner FormID for transfer pairs ('' for enrollment)
 * @param {string} draftedBy    - Employee ID of the person creating the draft
 * @returns {string} FormID on success, or 'Error: ...' string
 */
function draftAccountabilityForm(empId, assets, formType, linkedFormId, draftedBy) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return 'Error: System busy — try again.'; }

  try {
    if (!empId)     return 'Error: Employee ID is required.';
    if (!formType)  return 'Error: Form type is required.';
    if (!draftedBy) return 'Error: Drafter ID is required.';

    const formId    = generateFormID();
    if (formId.startsWith('Error')) return formId;

    const nowStr    = new Date().toLocaleString('en-PH');
    const contextType = _getContextType(empId);
    const assetsJson  = _buildAssetsSnapshot(assets || []);

    // Resolve staff info from first asset or masterlist
    const refAsset  = (assets && assets.length) ? assets[0] : {};
    const staffName = refAsset.Staff       || '';
    const desig     = refAsset.Designation || '';
    const dept      = refAsset.Department  || '';
    const branch    = refAsset.Branch      || refAsset.BaseOffice || '';
    const division  = refAsset.Division    || '';
    const district  = refAsset.District    || '';

    const row = new Array(AF_TOTAL_COLS).fill('');
    row[AF.FORM_ID        - 1] = formId;
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

    // Link FormID back to each asset row in Asset Entry
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
  } finally {
    lock.releaseLock();
  }
}

/**
 * Serializes an array of asset objects to a compact JSON string for storage.
 * Stores only the fields needed to render the accountability form.
 */
function _buildAssetsSnapshot(assets) {
  if (!assets || !assets.length) return '[]';
  const snapshot = assets.map(function(a) {
    return {
      barcode:   a.Barcode   || '',
      type:      a.Type      || '',
      brand:     a.Brand     || '',
      serial:    a.Serial    || '',
      specs:     a.Specs     || '',
      condition: a.Condition || ''
    };
  });
  return JSON.stringify(snapshot);
}

/**
 * Determines if an employee belongs to a Field or HO context.
 * Reads from Org Structure sheet — Field if found there, HO otherwise.
 */
function _getContextType(empId) {
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return 'HO';
    const last = sh.getLastRow();
    const id   = String(empId || '').trim().toLowerCase();
    const ids  = sh.getRange(2, 4, last - 1, 1).getValues(); // col D = FE ID
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() === id) return 'Field';
    }
    // Also check supervisor column (col G)
    const supIds = sh.getRange(2, 7, last - 1, 1).getValues();
    for (let i = 0; i < supIds.length; i++) {
      if (String(supIds[i][0] || '').trim().toLowerCase() === id) return 'Field';
    }
    return 'HO';
  } catch(e) {
    return 'HO';
  }
}


// ─────────────────────────────────────────────────────────────────────────────
//  PART 2B — SUBMIT FOR REVIEW
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Moves a Draft or Rejected form to Pending status.
 * Enforces rate limiting. Only the original drafter can submit.
 *
 * @returns {Object} { ok, formId, message, rateLimitStatus }
 */
function submitFormForReview(formId, drafterId) {
  try {
    const form = getFormDetail(formId);
    if (!form)                              return { ok: false, message: 'Form not found: ' + formId };
    if (form.status === 'Confirmed')        return { ok: false, message: 'This form is already confirmed.' };
    if (form.status === 'Pending')          return { ok: false, message: 'This form is already pending review.' };
    if (form.draftedBy.toLowerCase() !== String(drafterId).trim().toLowerCase())
                                            return { ok: false, message: 'Only the original drafter can submit this form.' };
    if (form.status !== 'Draft' && form.status !== 'Rejected')
                                            return { ok: false, message: 'Form status "' + form.status + '" cannot be submitted.' };

    // Rate limit check (only applies to Rejected resubmissions; first submit is always free)
    if (form.status === 'Rejected') {
      const rlResult = checkRateLimit(formId, drafterId);
      if (!rlResult.allowed) {
        return { ok: false, message: rlResult.message, rateLimitStatus: rlResult };
      }
    }

    const sh     = _afSheet();
    const rowIdx = _findAFRow(formId);
    if (rowIdx < 0) return { ok: false, message: 'Form record not found in sheet.' };

    const nowStr = new Date().toLocaleString('en-PH');
    sh.getRange(rowIdx, AF.STATUS).setValue('Pending');
    sh.getRange(rowIdx, AF.SUBMITTED_AT).setValue(nowStr);
    sh.getRange(rowIdx, AF.REJECTION_COMMENT).setValue('');

    // Update approval status on all linked asset rows
    _updateAssetApprovalStatus(formId, 'Pending', '');

    _log('SUBMIT_FORM', formId, 'Submitted for review | ' + form.formType + ' | ' + form.empId, drafterId);
    return { ok: true, formId, message: 'Form submitted for supervisor review.' };

  } catch(e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}


// ─────────────────────────────────────────────────────────────────────────────
//  PART 2C — SUPERVISOR ACTIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Confirms a single enrollment form.
 * Validates that supervisorId supervises the drafter (or is admin).
 * Writes confirmed snapshot to FormSnapshots. Updates Asset Entry rows.
 *
 * @returns {Object} { ok, message }
 */
function confirmForm(formId, supervisorId, roleTier) {
  try {
    const form = getFormDetail(formId);
    if (!form)                       return { ok: false, message: 'Form not found.' };
    if (form.status !== 'Pending')   return { ok: false, message: 'Form is not pending review.' };

    // Scope check — admin bypasses, supervisor must own the drafter
    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(form.draftedBy.toLowerCase()) < 0) {
        return { ok: false, message: 'You do not supervise the drafter of this form.' };
      }
    }

    const sh     = _afSheet();
    const rowIdx = _findAFRow(formId);
    if (rowIdx < 0) return { ok: false, message: 'Form record not found in sheet.' };

    const nowStr = new Date().toLocaleString('en-PH');
    sh.getRange(rowIdx, AF.STATUS).setValue('Confirmed');
    sh.getRange(rowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
    sh.getRange(rowIdx, AF.REVIEWED_AT).setValue(nowStr);

    // Archive to FormSnapshots
    _archiveForm(formId, supervisorId, nowStr);

    // Update all linked asset rows to Confirmed
    _updateAssetApprovalStatus(formId, 'Confirmed', '');

    // Clear rate limit record
    _clearRateLimit(formId, form.draftedBy);

    _log('CONFIRM_FORM', formId, form.formType + ' | ' + form.empId + ' | confirmed', supervisorId);
    return { ok: true, message: 'Form confirmed. Asset(s) are now live in inventory.' };

  } catch(e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}

/**
 * Confirms both Transfer-From and Transfer-To forms atomically.
 * Both must be Pending. If either check fails, neither is confirmed.
 *
 * @returns {Object} { ok, message }
 */
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
      if (supervised.indexOf(fromForm.draftedBy.toLowerCase()) < 0 &&
          supervised.indexOf(toForm.draftedBy.toLowerCase()) < 0) {
        return { ok: false, message: 'You do not supervise the drafter(s) of this transfer.' };
      }
    }

    const nowStr = new Date().toLocaleString('en-PH');
    const sh     = _afSheet();

    // Confirm FROM form
    const fromRowIdx = _findAFRow(fromFormId);
    if (fromRowIdx > 0) {
      sh.getRange(fromRowIdx, AF.STATUS).setValue('Confirmed');
      sh.getRange(fromRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(fromRowIdx, AF.REVIEWED_AT).setValue(nowStr);
    }

    // Confirm TO form
    const toRowIdx = _findAFRow(toFormId);
    if (toRowIdx > 0) {
      sh.getRange(toRowIdx, AF.STATUS).setValue('Confirmed');
      sh.getRange(toRowIdx, AF.SUPERVISOR_ID).setValue(supervisorId);
      sh.getRange(toRowIdx, AF.REVIEWED_AT).setValue(nowStr);
    }

    // Archive both — supersede the FROM (old holder) form
    _archiveForm(toFormId, supervisorId, nowStr);
    _archiveForm(fromFormId, supervisorId, nowStr);
    _supersedeForms(fromFormId, toFormId, nowStr);

    // Update asset approval status to Confirmed
    _updateAssetApprovalStatus(fromFormId, 'Confirmed', '');
    _updateAssetApprovalStatus(toFormId,   'Confirmed', '');

    _clearRateLimit(fromFormId, fromForm.draftedBy);
    _clearRateLimit(toFormId,   toForm.draftedBy);

    _log('CONFIRM_TRANSFER', fromFormId + '+' + toFormId,
      'Transfer pair confirmed | From: ' + fromForm.empId + ' → To: ' + toForm.empId,
      supervisorId);
    return { ok: true, message: 'Transfer confirmed. Both forms are now live.' };

  } catch(e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}

/**
 * Rejects a single pending form with a mandatory comment.
 * Flips status back to Rejected. Updates asset rows.
 *
 * @returns {Object} { ok, message }
 */
function rejectForm(formId, supervisorId, comment, roleTier) {
  try {
    if (!comment || !comment.trim()) return { ok: false, message: 'A rejection comment is required.' };

    const form = getFormDetail(formId);
    if (!form)                     return { ok: false, message: 'Form not found.' };
    if (form.status !== 'Pending') return { ok: false, message: 'Only Pending forms can be rejected.' };

    if (roleTier !== 'ho') {
      const supervised = _getSupervisedEmpIds(supervisorId);
      if (supervised.indexOf(form.draftedBy.toLowerCase()) < 0) {
        return { ok: false, message: 'You do not supervise the drafter of this form.' };
      }
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

    // Push rejection comment to all linked asset rows so drafter sees it
    _updateAssetApprovalStatus(formId, 'Rejected', trimmed);

    _log('REJECT_FORM', formId, form.formType + ' | ' + form.empId + ' | ' + trimmed, supervisorId);
    return { ok: true, message: 'Form rejected. The drafter has been notified.' };

  } catch(e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}

/**
 * Rejects both forms in a transfer pair.
 * Either form can carry the primary comment; per-form comments are optional.
 * On rejection the transfer is cancelled — the asset stays with current holder.
 *
 * @returns {Object} { ok, message }
 */
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
          supervised.indexOf(toForm.draftedBy.toLowerCase()) < 0) {
        return { ok: false, message: 'You do not supervise the drafter(s) of this transfer.' };
      }
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

    _log('REJECT_TRANSFER', fromFormId + '+' + toFormId,
      'Transfer pair rejected | ' + comment, supervisorId);
    return { ok: true, message: 'Transfer rejected. Both forms returned to drafters.' };

  } catch(e) {
    return { ok: false, message: 'Error: ' + e.message };
  }
}


// ─────────────────────────────────────────────────────────────────────────────
//  PART 2D — ARCHIVE HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Copies a confirmed form row to FormSnapshots for permanent archiving.
 */
function _archiveForm(formId, confirmedBy, confirmedAt) {
  try {
    const form = getFormDetail(formId);
    if (!form) return;

    _fsSheet().appendRow([
      form.formId,
      form.formType,
      form.linkedFormId || '',
      form.contextType,
      form.empId,
      form.staffName,
      form.designation,
      form.department,
      form.branch,
      form.assetsJson,
      confirmedBy,
      confirmedAt,
      '',  // supersededAt — filled later if needed
      ''   // supersededBy — filled later if needed
    ]);
  } catch(e) {
    Logger.log('_archiveForm error: ' + e.message);
  }
}

/**
 * Marks an archived FROM-form as superseded by a TO-form.
 * Called after a transfer pair is confirmed.
 */
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
  } catch(e) {
    Logger.log('_supersedeForms error: ' + e.message);
  }
}


// ─────────────────────────────────────────────────────────────────────────────
//  PART 2E — MODIFICATIONS TO EXISTING FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * REPLACE the existing processAsset() function with this version.
 *
 * Changes from original:
 *  - After writing the new asset row, sets ApprovalStatus = 'Draft' (col 32)
 *  - Calls draftAccountabilityForm() for all staff assets if this is an enrollment
 *  - Returns { result, formId } instead of just a string
 *    so the frontend can open the form preview
 *
 * NOTE: The isAssign=true branch (legacy update path) is unchanged.
 */
function processAsset(obj, isAssign) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(e) { return { result: 'Error: System busy — try again.', formId: null }; }

  try {
    const sh     = _entrySheet();
    const nowStr = new Date().toLocaleString('en-PH');

    if (!isAssign) {
      // ── ENROLL NEW ASSET ──────────────────────────────────────────────
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
      const isSpare = (statusChoice === 'Spare' || statusChoice === 'BorrowItem');
      const normDiv  = _normDiv(obj.division  || '');
      const normDist = _normDist(obj.district || '');

      const row = new Array(37).fill('');  // 37 cols now (was 31)
      row[C.ENTRY_EMP_ID   - 1] = obj.entryEmpId   || '';
      row[C.ENTRY_NAME     - 1] = obj.entryName     || '';
      row[C.EMP_ID         - 1] = isSpare ? '' : (obj.accEmpId || '');
      row[C.STAFF          - 1] = isSpare ? '' : _sanitize(obj.accName, 100);
      row[C.DESIGNATION    - 1] = isSpare ? '' : (obj.accRole || '');
      row[C.DEPARTMENT     - 1] = obj.department    || '';
      row[C.BASE_OFFICE    - 1] = obj.baseOffice    || '';
      row[C.DIVISION       - 1] = normDiv;
      row[C.DISTRICT       - 1] = normDist;
      row[C.AREA           - 1] = obj.area          || '';
      row[C.BRANCH         - 1] = _sanitize(obj.branch, 150);
      row[C.ASSIGNMENT     - 1] = obj.assignment    || 'Field Office';
      row[C.EFF_DATE       - 1] = obj.effDate       || '';
      row[C.BARCODE        - 1] = obj.barcode;
      row[C.TYPE           - 1] = obj.type          || '';
      row[C.BRAND          - 1] = obj.brand         || '';
      row[C.SERIAL         - 1] = obj.serial ? String(obj.serial) : '';
      row[C.SPECS          - 1] = obj.specs         || '';
      row[C.SUPPLIER       - 1] = obj.supplier      || '';
      row[C.CONDITION      - 1] = obj.condition     || 'New';
      row[C.ASSET_LOCATION - 1] = obj.location      || '';
      row[C.LIFECYCLE      - 1] = sm.lc;
      row[C.STATUS_LABEL   - 1] = sm.sl;
      row[C.ASSET_STATUS   - 1] = sm.as;
      row[C.PURCH_DATE     - 1] = obj.purchDate     || '';
      row[C.WARRANTY_TERM  - 1] = obj.wTerm         || '';
      row[C.WARRANTY_VAL   - 1] = obj.wValidity     || '';
      row[C.REMARKS        - 1] = _sanitize(obj.remarks, 500);
      row[C.CREATED_AT     - 1] = nowStr;
      row[C.LAST_UPDATED   - 1] = nowStr;
      // New approval columns
      row[CAE.APPROVAL_STATUS - 1] = 'Draft';
      row[CAE.FORM_ID         - 1] = '';  // filled after draftAccountabilityForm
      row[CAE.DRAFTED_BY      - 1] = obj.enrolledBy || obj.entryEmpId || '';
      row[CAE.DRAFTED_AT      - 1] = nowStr;
      row[CAE.REJECTION_COMMENT-1] = '';
      row[CAE.GRANDFATHERED   - 1] = false;

      sh.appendRow(row);
      const newRowIdx = sh.getLastRow();
      sh.getRange(newRowIdx, C.SERIAL).setNumberFormat('@STRING@');
      if (obj.serial) sh.getRange(newRowIdx, C.SERIAL).setValue(String(obj.serial));

      // Write to Spare log if applicable
      if (statusChoice === 'Spare') {
        _spareSheet().appendRow([
          obj.barcode, obj.type, obj.brand, obj.serial || '', obj.condition || 'New',
          obj.purchDate || '', obj.wValidity || '', obj.supplier || '',
          normDiv, normDist, obj.area || '', _sanitize(obj.branch, 150),
          obj.location || '', obj.enrolledBy || obj.entryEmpId || '',
          nowStr, 'Available'
        ]);
      }

      _log('CREATE', obj.barcode, obj.type + ' | ' + obj.brand + ' | ' + statusChoice, obj.entryEmpId || '');

      // Draft the accountability form for this new asset
      // We pass a minimal asset object — the drafter will have full context on frontend
      const assetForForm = [{
        Barcode:    obj.barcode,
        Type:       obj.type      || '',
        Brand:      obj.brand     || '',
        Serial:     obj.serial    || '',
        Specs:      obj.specs     || '',
        Condition:  obj.condition || 'New',
        Staff:      obj.accName   || '',
        Designation:obj.accRole   || '',
        Department: obj.department|| '',
        BaseOffice: obj.baseOffice|| '',
        Branch:     obj.branch    || '',
        Division:   normDiv,
        District:   normDist
      }];

      const drafterId = obj.enrolledBy || obj.entryEmpId || '';
      const empIdForForm = (isSpare ? drafterId : (obj.accEmpId || drafterId));
      const formId = draftAccountabilityForm(
        empIdForForm,
        assetForForm,
        'Enrollment',
        '',
        drafterId
      );

      if (!formId.startsWith('Error')) {
        sh.getRange(newRowIdx, CAE.FORM_ID).setValue(formId);
      }

      return {
        result: 'Asset created: ' + obj.barcode,
        formId: formId.startsWith('Error') ? null : formId
      };
    }

    // ── isAssign = true — unchanged from original processAsset() ─────────
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

  } catch(e) {
    return { result: 'Error: ' + e.message, formId: null };
  } finally {
    lock.releaseLock();
  }
}

/**
 * REPLACE the existing saveTransfer() function with this version.
 *
 * Changes from original:
 *  - After writing the transfer record, creates two Draft accountability forms
 *    (Transfer-From for old holder, Transfer-To for new holder)
 *  - Returns { result, fromFormId, toFormId } so frontend can open the
 *    paired transfer review preview
 */
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

    // Write transfer record (unchanged)
    _xferSheet().appendRow([
      t.barcode, t.transferType,
      _sanitize(t.fromStaff, 100), t.fromEmpId, t.fromDesig,
      t.fromDiv, t.fromDist, t.fromArea, _sanitize(t.fromBranch, 150), _sanitize(t.fromRemarks, 500),
      _sanitize(t.toStaff, 100), t.toEmpId, t.toDesig,
      normToDiv, normToDist, t.toArea, _sanitize(t.toBranch, 150), _sanitize(t.toRemarks, 500),
      t.effDate, t.status || 'Pending', nowStr
    ]);

    // Update Asset Entry — new holder (unchanged fields)
    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'],  [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],   [C.EMP_ID,       t.toEmpId   || ''],
      [C.STAFF,        _sanitize(t.toStaff, 100)],
      [C.DESIGNATION,  t.toDesig   || ''],
      [C.DIVISION,     normToDiv],    [C.DISTRICT,     normToDist],
      [C.AREA,         t.toArea    || ''],
      [C.BRANCH,       _sanitize(t.toBranch, 150)],
      [C.EFF_DATE,     t.effDate],
      // Reset approval to Draft pending transfer form confirmation
      [CAE.APPROVAL_STATUS, 'Draft'],
      [CAE.REJECTION_COMMENT, '']
    ]);

    // Build asset snapshot for both forms
    const assetObj = {
      Barcode:     t.barcode,
      Type:        String(curRow[C.TYPE      - 1] || ''),
      Brand:       String(curRow[C.BRAND     - 1] || ''),
      Serial:      String(curRow[C.SERIAL    - 1] || ''),
      Specs:       String(curRow[C.SPECS     - 1] || ''),
      Condition:   String(curRow[C.CONDITION - 1] || '')
    };

    // FROM form — old holder turning over the asset
    const fromAsset = Object.assign({}, assetObj, {
      Staff:       t.fromStaff   || '',
      EmpID:       t.fromEmpId   || '',
      Designation: t.fromDesig   || '',
      Division:    t.fromDiv     || '',
      District:    t.fromDist    || '',
      Branch:      t.fromBranch  || ''
    });

    // TO form — new holder receiving the asset
    const toAsset = Object.assign({}, assetObj, {
      Staff:       t.toStaff    || '',
      EmpID:       t.toEmpId    || '',
      Designation: t.toDesig    || '',
      Division:    normToDiv,
      District:    normToDist,
      Branch:      t.toBranch   || ''
    });

    const drafterId = t.fromEmpId || '';

    // Create From form first to get its ID for the linked pair
    const fromFormId = draftAccountabilityForm(
      t.fromEmpId  || '',
      [fromAsset],
      'Transfer-From',
      '',           // linkedFormId — will update after To is created
      drafterId
    );

    const toFormId = draftAccountabilityForm(
      t.toEmpId    || '',
      [toAsset],
      'Transfer-To',
      fromFormId.startsWith('Error') ? '' : fromFormId,
      drafterId
    );

    // Update FROM form with its linked pair
    if (!fromFormId.startsWith('Error') && !toFormId.startsWith('Error')) {
      const afSh      = _afSheet();
      const fromRowI  = _findAFRow(fromFormId);
      if (fromRowI > 0) {
        afSh.getRange(fromRowI, AF.LINKED_FORM_ID).setValue(toFormId);
      }
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

/**
 * Add ApprovalStatus to the getAllAssets() return object and
 * exclude non-Confirmed assets from normal pool status computation.
 *
 * INSTRUCTION: In the existing getAllAssets() function, inside the .map() callback,
 * find the return object and ADD these two lines:
 *
 *   ApprovalStatus: get(CAE.APPROVAL_STATUS) || 'Confirmed',
 *   FormID:         get(CAE.FORM_ID)         || '',
 *
 * AND modify the _computeStatus call to exclude draft/pending/rejected:
 *
 *   const approvalStatus = get(CAE.APPROVAL_STATUS) || 'Confirmed';
 *   const grandfathered  = String(get(CAE.GRANDFATHERED)).toLowerCase() === 'true';
 *   // Non-confirmed assets are invisible to normal pools
 *   const effectiveStatus = (approvalStatus === 'Confirmed' || grandfathered)
 *     ? _computeStatus(get(C.LIFECYCLE), get(C.ASSET_STATUS), get(C.EMP_ID))
 *     : 'pending-approval';
 *
 * Then use effectiveStatus instead of the direct _computeStatus call.
 *
 * This function documents the change — actual edit is in getAllAssets().
 */
function _getAllAssetsApprovalPatch() {
  // Documentation only — see instructions above for exact edit location
}

/**
 * Add pendingCounts to the getInitialData() return object.
 *
 * INSTRUCTION: In the existing getInitialData() function, add:
 *
 *   pendingCounts: getPendingCounts(
 *     <pass in userId from session>,
 *     <pass in roleTier from session>
 *   )
 *
 * Since getInitialData() doesn't receive user params currently,
 * the frontend should call getPendingCounts() separately on login:
 *   google.script.run.getPendingCounts(SESSION.username, SESSION.roleTier)
 *
 * This function documents the change.
 */
function _getInitialDataApprovalPatch() {
  // Documentation only — see instructions above
}


// ─────────────────────────────────────────────────────────────────────────────
//  INTERNAL UTILITY FUNCTIONS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Finds the row index of a form in AccountabilityForms by FormID.
 * Returns -1 if not found.
 */
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
  } catch(e) {
    return -1;
  }
}

/**
 * Finds the row in FormRateLimits matching formId + drafterId.
 * Returns the data array or null.
 */
function _findRLRow(formId, drafterId) {
  try {
    const sh   = _rlSheet();
    const last = sh.getLastRow();
    if (last < RL_DATA_START) return null;
    const data = sh.getRange(RL_DATA_START, 1, last - RL_DATA_START + 1, 7).getValues();
    const fid  = String(formId || '').trim();
    const did  = String(drafterId || '').trim().toLowerCase();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === fid &&
          String(data[i][1]).trim().toLowerCase() === did) {
        return data[i];
      }
    }
    return null;
  } catch(e) {
    return null;
  }
}

/**
 * Returns the sheet row index of a rate limit record.
 */
function _findRLRowIdx(formId, drafterId) {
  try {
    const sh   = _rlSheet();
    const last = sh.getLastRow();
    if (last < RL_DATA_START) return -1;
    const ids = sh.getRange(RL_DATA_START, 1, last - RL_DATA_START + 1, 2).getValues();
    const fid = String(formId || '').trim();
    const did = String(drafterId || '').trim().toLowerCase();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === fid &&
          String(ids[i][1]).trim().toLowerCase() === did) {
        return i + RL_DATA_START;
      }
    }
    return -1;
  } catch(e) {
    return -1;
  }
}

/**
 * Maps a raw AccountabilityForms sheet row array to a clean JS object.
 */
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

/**
 * Updates ApprovalStatus and RejectionComment on all Asset Entry rows
 * linked to a given FormID.
 */
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
  } catch(e) {
    Logger.log('_updateAssetApprovalStatus error: ' + e.message);
  }
}

/**
 * Returns an array of Employee IDs directly supervised by supervisorId.
 * Reads from Org Structure sheet — col D (FE ID), col G (Supervisor ID).
 */
function _getSupervisedEmpIds(supervisorId) {
  try {
    const sh   = _ss().getSheetByName(SH_ORG);
    if (!sh || sh.getLastRow() < 2) return [];
    const last  = sh.getLastRow();
    const data  = sh.getRange(2, 4, last - 1, 4).getValues(); // cols D-G
    const supId = String(supervisorId || '').trim().toLowerCase();
    const ids   = [];

    data.forEach(r => {
      // col G (index 3 in this range) is Supervisor ID
      if (String(r[3] || '').trim().toLowerCase() === supId) {
        const feId = String(r[0] || '').trim().toLowerCase();
        if (feId && ids.indexOf(feId) < 0) ids.push(feId);
      }
    });
    return ids;
  } catch(e) {
    return [];
  }
}

/**
 * Convenience wrapper to fetch approval dashboard data in one call.
 */
function getApprovalDashboardData(userId, roleTier) {
  try {
    return {
      myForms:      getMyForms(userId),
      pendingForms: canReviewForms_server(userId, roleTier) ? getPendingForms(userId, roleTier) : [],
      config:       getApprovalConfig()
    };
  } catch(e) {
    return { myForms: [], pendingForms: [], config: {} };
  }
}

function canReviewForms_server(userId, roleTier) {
  return roleTier === 'senior' || roleTier === 'ho';
}