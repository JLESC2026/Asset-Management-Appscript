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
  if (r === 'super admin' || r === 'superadmin' || r === 'admin') return 'ho';
  if (r === 'super user'  || r === 'superuser')                    return 'ho';
  if (r === 'supervisor')                                           return 'senior';
  return 'fe';
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
function loginUser(empId, password) {
  try {
    const sh = _ss().getSheetByName(SH_USERS);
    if (!sh) return { ok: false, error: 'Users sheet not found.' };
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, error: 'No users registered.' };
    const data = sh.getRange(2, 1, last - 1, 9).getValues();

    for (let ri = 0; ri < data.length; ri++) {
      const row   = data[ri];
      const rowId = String(row[2] || '').trim(); // Col C = Employee ID
      if (rowId.toLowerCase() !== String(empId).trim().toLowerCase()) continue;

      const role    = String(row[0] || 'User').trim(); // Col A
      const pwd     = String(row[1] || '').trim();     // Col B

      const inputHash = _hashPwd(password);
      if (_isHashed(pwd)) {
        if (inputHash !== pwd) return { ok: false, error: 'Incorrect password.' };
      } else {
        if (String(password) !== pwd) return { ok: false, error: 'Incorrect password.' };
        sh.getRange(ri + 2, 2).setValue(inputHash);
      }

      const firstLogin  = (pwd === '1234' || pwd === _hashPwd('1234'));
      const roleTier    = _mapRoleTier(role);
      const isHO        = roleTier === 'ho';

      const mlData    = _getMasterlistEntry(rowId);
      const scopeData = _parseOrgStructure(rowId, roleTier);

      return {
        ok:                  true,
        username:            rowId,
        role,
        roleTier,
        isHeadOffice:        isHO,
        name:                mlData.name || String(row[3] || rowId),
        firstLogin,
        division:            scopeData.userDivisions[0]  || mlData.division || '',
        district:            scopeData.userDistricts[0]  || mlData.district || '',
        userDivisions:       scopeData.userDivisions,
        userDistricts:       scopeData.userDistricts.length > 0
                               ? scopeData.userDistricts
                               : (mlData.district ? [mlData.district] : []),
        seniorDistrictScope: scopeData.seniorDistrictScope,
        divDistrictMap:      scopeData.divDistrictMap,
        headOfficeDepts:     [],
        area:                mlData.area       || '',
        branch:              mlData.baseOffice || '',
      };
    }
    return { ok: false, error: 'Employee ID not found.' };
  } catch(e) {
    return { ok: false, error: e.message };
  }
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
      .filter(row => {
        const bc = String(row[C.BARCODE - 1] || '').trim();
        return bc && bc !== '-' && bc !== 'N/A' && bc !== 'None' && bc !== '#N/A';
      })
      .map(row => {
        const get    = col => String(row[col - 1] || '');
        const status = _computeStatus(get(C.LIFECYCLE), get(C.ASSET_STATUS), get(C.EMP_ID));

        const rawLC    = get(C.LIFECYCLE);
        const displayLC = rawLC || {
          'allocated':'Allocated','spare':'Active','borrowed':'Borrow',
          'returned':'Returned','disposal':'Dispose','transfer':'Transfer',
          'borrow-item':'BorrowItem'
        }[status] || 'Active';

        return {
          Barcode:      get(C.BARCODE),
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
          Assignment:   get(C.ASSIGNMENT),
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
          status
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
function processAsset(obj, isAssign) {
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
function saveTransfer(t) {
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