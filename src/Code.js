// ═══════════════════════════════════════════════════════════
//  ASSET MANAGEMENT SYSTEM — Code.gs  (v5.6 — HO Department Support)
//
//  Changes vs v5.5:
//    DEPT-COL — Added DEPARTMENT (col 6) and BASE_OFFICE (col 7) to
//               the column map. These were present in the sheet but
//               not mapped, causing HO assets to be invisible in the
//               Staff Assets and Equipment Record trees (they showed
//               under "(No Division)/(No District)").
//    DEPT-COL — getAllAssets() now returns Department and BaseOffice
//               fields so the frontend can use them for HO tree views.
//    DEPT-COL — processAsset() now writes Department and BaseOffice
//               when supplied.
//    DEPT-COL — allocateAsset() and deallocateAsset() preserve Dept/BO.
//
// ═══════════════════════════════════════════════════════════

const SHEET_ID    = '18tuYQKH2OLLu1NqPJiA28n8n7GNN6XR_SSZXUO4XEe8';
const SH_MASTER   = 'Masterlist';
const SH_ENTRY    = 'Asset Entry';
const SH_XFER     = 'Transfers';
const SH_BORROW   = 'Borrows';
const SH_DISPOSE  = 'Disposals';
const SH_LOG      = 'ActivityLog';
const SH_DROPDOWN = 'Drop down';
const SH_ALLOC    = 'Allocated';

const AE_DATA_START = 4;
const C = {
  ENTRY_EMP_ID:1,  ENTRY_NAME:2,
  EMP_ID:3,        STAFF:4,          DESIGNATION:5,
  DEPARTMENT:6,    BASE_OFFICE:7,    // ← Added
  DIVISION:8,      DISTRICT:9,       AREA:10,        BRANCH:11,
  // Col 12 = Assignment (not mapped)
  EFF_DATE:13,
  BARCODE:14,      TYPE:15,          BRAND:16,       SERIAL:17,    SPECS:18,
  CONDITION:19,
  // Col 20 = Asset Location (not mapped)
  LIFECYCLE:21,    STATUS_LABEL:22,  ASSET_STATUS:23,
  PURCH_DATE:24,   WARRANTY_TERM:25, WARRANTY_VAL:26, REMARKS:27,
  // Col 28 = Notes (not mapped)
  XFER_TYPE:29,    FR_STAFF:30,      FR_EMPID:31,    FR_DESIG:32,  FR_DIV:33,
  FR_DIST:34,      FR_AREA:35,       FR_BRANCH:36,   FR_REMARKS:37,
  TO_STAFF:38,     TO_EMPID:39,      TO_DESIG:40,    TO_DIV:41,    TO_DIST:42,
  TO_AREA:43,      TO_BRANCH:44,     TO_REMARKS:45,  XFER_DATE:46,
  BOR_NAME:47,     BOR_EMPID:48,     BOR_DESIG:49,   BOR_DIV:50,   BOR_BRANCH:51,
  BOR_DATE:52,     EXP_RETURN:53,    ACT_RETURN:54,  BOR_REMARKS:55,
  CREATED_AT:56,   LAST_UPDATED:57,  SUPPLIER:58,    LOCATION:59,  ENROLLED_BY:60,
  BOR_DIST:61
};
const TOTAL_COLS = Math.max(...Object.values(C)); // = 61

// AE_HEADERS: 61 entries matching the actual sheet column layout.
const AE_HEADERS = [
  'Entry Employee ID','Entered By',
  'Accountable Employee ID','Accountable Staff','Designation',
  'Department','Base Office',
  'Division','District','Area','Branch',
  'Assignment',
  'Effectivity Date',
  'Barcode','Category','Brand','Serial No.','Specifications',
  'Condition',
  'Asset Location',
  'Lifecycle Status','Status Label','Assignment Status',
  'Date of Purchase','Warranty Term','Warranty Validity','Remarks',
  'Notes',
  'Transfer Type','From Staff','From EmpID','From Designation','From Division',
  'From District','From Area','From Branch','From Remarks',
  'To Staff','To EmpID','To Designation','To Division','To District',
  'To Area','To Branch','To Remarks','Transfer Date',
  'Borrower Name','Borrower EmpID','Borrower Designation',
  'Borrow Division','Borrow Branch','Borrow Date',
  'Expected Return Date','Actual Return Date','Borrow Remarks',
  'Created At','Last Updated','Supplier','Location','Enrolled By',
  'Borrow District'   // ← col 61
];

function _sanitize(val, maxLen) {
  maxLen = maxLen || 500;
  return String(val || '').trim().substring(0, maxLen);
}

// ─── SHEET HELPERS ───────────────────────────────────────────────────────────
function _ss() { return SpreadsheetApp.openById(SHEET_ID); }

function _getOrCreate(name, headers) {
  const ss = _ss();
  let sh = ss.getSheetByName(name);
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

function _entrySheet()   { return _getOrCreate(SH_ENTRY,   AE_HEADERS); }
function _xferSheet()    { return _getOrCreate(SH_XFER,    ['Barcode','Type','FromStaff','FromEmpID','FromDesig','FromDiv','FromDist','FromArea','FromBranch','FromRemarks','ToStaff','ToEmpID','ToDesig','ToDiv','ToDist','ToArea','ToBranch','ToRemarks','EffDate','Status','Timestamp']); }
function _borrowSheet()  { return _getOrCreate(SH_BORROW,  ['Barcode','BorrowerName','EmpID','Designation','Division','District','Branch','BorrowDate','ExpectedReturn','ActualReturn','Status','Remarks','Timestamp']); }
function _disposeSheet() { return _getOrCreate(SH_DISPOSE, ['Barcode','Reason','DisposedBy','DisposeDate','Remarks','Timestamp']); }
function _logSheet()     { return _getOrCreate(SH_LOG,     ['Timestamp','Action','Barcode','Details','Performed By']); }
function _allocLogSheet(){ return _getOrCreate(SH_ALLOC,   ['Barcode','Category','Brand','Serial No.','Employee ID','Accountable Staff','Designation','Department','Base Office','Division','District','Area','Branch','Effectivity Date','Condition','Remarks','Timestamp','Allocated By']); }

// ─── ROW FINDERS / SETTERS ───────────────────────────────────────────────────
function _findRow(sheet, barcode) {
  if (!barcode) return -1;
  try {
    const finder = sheet.createTextFinder(String(barcode).trim())
                        .matchEntireCell(true)
                        .matchCase(false);
    const range = finder.findNext();
    if (!range) return -1;
    const row = range.getRow();
    // Verify it's in the data range and in the barcode column
    if (row < AE_DATA_START || range.getColumn() !== C.BARCODE) return -1;
    return row;
  } catch (e) {
    // Fallback to linear scan if TextFinder fails
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
  // Read the entire row once
  const totalCols = TOTAL_COLS;
  const range = sheet.getRange(rowIdx, 1, 1, totalCols);
  const rowValues = range.getValues()[0]; // single read

  // Apply updates in memory
  updates.forEach(u => {
    const colIdx = u[0] - 1; // Convert 1-based to 0-based
    rowValues[colIdx] = (u[1] != null ? u[1] : '');
  });

  // Apply last-updated timestamp
  rowValues[C.LAST_UPDATED - 1] = new Date().toLocaleString('en-PH');

  // Write entire row back in ONE API call
  range.setValues([rowValues]);
}

// ─── LIFECYCLE HELPERS ───────────────────────────────────────────────────────
function _computeStatus(lc, slb, actReturn, empId, borName, assetStatus) {
  const l  = String(lc  || '').trim().toLowerCase();
  const s  = String(slb || '').trim().toLowerCase();
  const as = String(assetStatus || '').trim().toLowerCase();
  const r  = String(actReturn || '').trim();
  const hasEmp = empId  && String(empId).trim()  && String(empId).trim()  !== 'n/a';
  const hasBor = borName && String(borName).trim();

  if (l === 'borrowitem' || as === 'borrowitem')      return 'borrow-item';
  if (l === 'borrow')                                  return r ? 'returned' : 'borrowed';
  if (l === 'returned')                                return 'returned';
  if (l === 'dispose'  || l === 'disposal'  ||
      as === 'disposal' || as === 'dispose' || s === 'disposed')           return 'disposal';
  if (l === 'transfer')                                return 'transfer';
  if (l === 'allocated' || s === 'assigned')           return 'allocated';

  if (!l || l === 'active') {
    if (hasBor) return r ? 'returned' : 'borrowed';
    if (hasEmp) return 'allocated';
  }
  return 'spare';
}

function _normDiv(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';
  return s.replace(/^DIv/i, m => 'Div');
}
function _normDist(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';
  const m = s.match(/^district\s+0*(\d+)$/i);
  if (m) return 'District ' + parseInt(m[1]);
  return s;
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

// ─── PASSWORD HASHING ─────────────────────────────────────────────────────────
function _hashPwd(pwd) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(pwd)
  );
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function _isHashed(str) {
  return /^[0-9a-f]{64}$/.test(String(str));
}

// ─── ROLE CLASSIFICATION ──────────────────────────────────────────────────────
function _classifyRole(roleStr) {
  const r = String(roleStr || '').trim().toLowerCase();
  if (r.includes('senior')) return 'senior';
  return 'fe';
}

// ─── SUPERVISOR DISTRICT LOOKUP ────────────────────────────────
function getDistrictsBySupervisor(empId) {
  try {
    const ss = _ss();
    const engSh = ss.getSheetByName('Eng. List') || ss.getSheetByName('Eng List');
    if (!engSh || engSh.getLastRow() < 2) return [];

    const lastRow = engSh.getLastRow();
    const data = engSh.getRange(2, 1, lastRow - 1, 10).getValues();
    const distSet = new Set();
    const id = String(empId).trim().toLowerCase();

    data.forEach(r => {
      const supId = String(r[3] || '').trim().toLowerCase();
      if (supId === id) {
        const dist = String(r[8] || '').trim();
        if (dist) distSet.add(dist);
      }
    });

    const numSort = (a, b) => {
      const na = parseInt(a.replace(/\D+/g, '')) || 0;
      const nb = parseInt(b.replace(/\D+/g, '')) || 0;
      return na !== nb ? na - nb : a.localeCompare(b);
    };
    return [...distSet].sort(numSort);
  } catch (e) {
    return [];
  }
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
function loginUser(empId, password) {
  try {
    const sh = _ss().getSheetByName(SH_MASTER);
    if (!sh) return { ok: false, error: 'Masterlist not found.' };
    const last = sh.getLastRow();
    if (last < 5) return { ok: false, error: 'No users registered.' };
    const data = sh.getRange(5, 1, last - 4, Math.min(sh.getLastColumn(), 15)).getValues();

    for (let ri = 0; ri < data.length; ri++) {
      const row  = data[ri];
      const id   = String(row[0] || '').trim();
      const pwd  = String(row[1] || '').trim();
      const name = String(row[2] || '').trim();
      const role = String(row[11] || 'User').trim();
      if (id.toLowerCase() !== String(empId).trim().toLowerCase()) continue;

      const inputHash = _hashPwd(password);
      if (_isHashed(pwd)) {
        if (inputHash !== pwd) return { ok: false, error: 'Incorrect password.' };
      } else {
        if (String(password) !== pwd) return { ok: false, error: 'Incorrect password.' };
        sh.getRange(ri + 5, 2).setValue(inputHash);
      }

      const firstLogin = (pwd === '1234' || pwd === _hashPwd('1234'));

      const mlDivision   = String(row[4] || '').trim();
      const mlDistrict   = String(row[5] || '').trim();
      const mlBaseOffice = String(row[7] || '').trim();
      const isHeadOffice = mlBaseOffice === 'Head Office';

      const roleTier = _classifyRole(role);
      const locData  = getLocationData(id);

      const division = locData.userDivisions.length > 0
        ? locData.userDivisions[0]
        : mlDivision;
      const district = locData.userDistricts.length > 0
        ? locData.userDistricts[0]
        : mlDistrict;

      let seniorDistrictScope = [];
      if (isHeadOffice || roleTier === 'senior') {
        const isHO = isHeadOffice ||
          division === 'Head Office' ||
          (locData.userDivisions.length > 0 && locData.userDivisions[0] === 'Head Office');

        if (isHO) {
          const allDepts = locData.headOfficeDepts || getHeadOfficeDepts();
          seniorDistrictScope = allDepts.length > 0
            ? allDepts
            : (locData.userDistricts.length > 0
                ? locData.userDistricts
                : (mlDistrict ? [mlDistrict] : []));
        } else {
          const supervisedDistricts = getDistrictsBySupervisor(id);
          if (supervisedDistricts.length > 0) {
            seniorDistrictScope = supervisedDistricts;
          } else {
            const userDivs = locData.userDivisions.length > 0
              ? locData.userDivisions
              : (mlDivision ? [mlDivision] : []);
            const ddMap = locData.divDistrictMap || {};
            userDivs.forEach(div => {
              (ddMap[div] || []).forEach(d => {
                if (!seniorDistrictScope.includes(d)) seniorDistrictScope.push(d);
              });
            });
            if (seniorDistrictScope.length === 0) {
              seniorDistrictScope = locData.userDistricts.length > 0
                ? locData.userDistricts
                : (mlDistrict ? [mlDistrict] : []);
            }
          }
        }
      }

      return {
        ok: true, username: id, role,
        roleTier: isHeadOffice ? 'ho' : roleTier,
        isHeadOffice,
        name, firstLogin,
        division: isHeadOffice ? 'Head Office' : division,
        district,
        userDivisions:  locData.userDivisions,
        userDistricts:  locData.userDistricts.length > 0
                          ? locData.userDistricts
                          : (mlDistrict ? [mlDistrict] : []),
        seniorDistrictScope,
        headOfficeDepts: locData.headOfficeDepts || [],
        divDistrictMap: locData.divDistrictMap || {},
        area:   String(row[6] || '').trim(),
        branch: mlBaseOffice
      };
    }
    return { ok: false, error: 'Employee ID not found.' };
  } catch (e) { return { ok: false, error: e.message }; }
}

function changePassword(empId, newPwd) {
  try {
    if (!newPwd || newPwd.length < 4) return { ok: false, error: 'Minimum 4 characters.' };
    const sh = _ss().getSheetByName(SH_MASTER);
    if (!sh) return { ok: false, error: 'Masterlist not found.' };
    const last = sh.getLastRow();
    const ids = sh.getRange(5, 1, last - 4, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() === String(empId).trim().toLowerCase()) {
        sh.getRange(i + 5, 2).setValue(_hashPwd(newPwd));
        return { ok: true };
      }
    }
    return { ok: false, error: 'Employee ID not found.' };
  } catch (e) { return { ok: false, error: e.message }; }
}

// ─── BARCODE GENERATION ───────────────────────────────────────────────────────
function generateNextBarcode(type) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: Could not generate barcode, system busy. Try again.';
  }

  try {
    const PFX = {
      'Laptop': 'LTP', 'Laptop Adaptor': 'LAD', 'CPU': 'CPU',
      'Monitor': 'MTR', 'Printer': 'PTR', 'Scanner': 'SCN',
      'Scansnap': 'SCN', 'Keyboard': 'KBD', 'Mouse': 'MSE',
      'UPS': 'UPS', 'Camera': 'CAM', 'Speaker': 'SPR',
      'External Drive': 'EXD'
    };
    const pre  = PFX[type] || 'AST';
    const yr   = new Date().getFullYear();
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    let max = 0;

    if (last >= AE_DATA_START) {
      const barcodes = sh.getRange(
        AE_DATA_START, C.BARCODE,
        last - AE_DATA_START + 1, 1
      ).getValues();

      // FIX: filter strictly by prefix AND year before extracting sequence
      const pattern = new RegExp('^' + pre + '-' + yr + '-(\\d+)$');
      barcodes.forEach(r => {
        const bc = String(r[0] || '').trim();
        const match = bc.match(pattern);
        if (match) {
          const n = parseInt(match[1], 10);
          if (!isNaN(n) && n > max) max = n;
        }
      });
    }

    let seq = max + 1;
    let candidate = pre + '-' + yr + '-' + String(seq).padStart(3, '0');
    
    // Collision guard
    while (_findRow(sh, candidate) > 0) {
      seq++;
      candidate = pre + '-' + yr + '-' + String(seq).padStart(3, '0');
    }
    return candidate;
  } finally {
    lock.releaseLock();
  }
}

// ─── ASSETS: READ ─────────────────────────────────────────────────────────────
function getAllAssets() {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return { success: true, data: [] };
    const colCount = sh.getLastColumn();
    const data = sh.getRange(AE_DATA_START, 1, last - AE_DATA_START + 1, colCount).getValues();
    const result = data
      .filter(row => row[C.BARCODE - 1])
      .map(row => {
        const get = (col) => String(row[col - 1] || '');
        const status = _computeStatus(
          get(C.LIFECYCLE), get(C.STATUS_LABEL), get(C.ACT_RETURN),
          get(C.EMP_ID), get(C.BOR_NAME), get(C.ASSET_STATUS)
        );
        const rawLC = get(C.LIFECYCLE);
        const displayLC = rawLC || {
          'allocated':  'Allocated',
          'spare':      'Active',
          'borrowed':   'Borrow',
          'returned':   'Returned',
          'disposal':   'Dispose',
          'transfer':   'Transfer',
          'borrow-item':'BorrowItem'
        }[status] || 'Active';

        const div  = _normDiv(get(C.DIVISION));
        const dist = _normDist(get(C.DISTRICT));

        const hasXfer = displayLC === 'Transfer';
        return {
          Barcode: get(C.BARCODE), Type: get(C.TYPE), Brand: get(C.BRAND),
          Serial: get(C.SERIAL), Specs: get(C.SPECS),
          Condition: get(C.CONDITION) || 'Good', Lifecycle: displayLC,
          AssetStatus: get(C.ASSET_STATUS) || 'Active',
          StatusLabel: get(C.STATUS_LABEL) || 'Unassigned',
          PurchDate: get(C.PURCH_DATE), WarrantyTerm: get(C.WARRANTY_TERM),
          WarrantyVal: get(C.WARRANTY_VAL), Remarks: get(C.REMARKS),
          EmpID: get(C.EMP_ID) || 'N/A', Staff: get(C.STAFF) || 'Unassigned',
          Designation: get(C.DESIGNATION),
          Department: get(C.DEPARTMENT),
          BaseOffice: get(C.BASE_OFFICE),
          Division: div,
          District: dist, Area: get(C.AREA), Branch: get(C.BRANCH),
          EffDate: get(C.EFF_DATE), 
          XferType:  hasXfer ? get(C.XFER_TYPE)  : '',
          ToStaff:   hasXfer ? get(C.TO_STAFF)   : '',
          ToEmpID:   hasXfer ? get(C.TO_EMPID)   : '',
          ToDiv:     hasXfer ? get(C.TO_DIV)     : '',
          ToBranch:  hasXfer ? get(C.TO_BRANCH)  : '',
          XferDate:  hasXfer ? get(C.XFER_DATE)  : '',
          BorName: get(C.BOR_NAME), BorEmpID: get(C.BOR_EMPID),
          BorDate: get(C.BOR_DATE), ExpReturn: get(C.EXP_RETURN),
          ActReturn: get(C.ACT_RETURN), BorRemarks: get(C.BOR_REMARKS),
          BorDesig: get(C.BOR_DESIG), BorDiv: get(C.BOR_DIV), BorBranch: get(C.BOR_BRANCH),
          BorDist: colCount >= 61 ? get(C.BOR_DIST) : '',
          CreatedAt: get(C.CREATED_AT), LastUpdated: get(C.LAST_UPDATED),
          Supplier: get(C.SUPPLIER), Location: get(C.LOCATION),
          EntryEmpId: get(C.ENTRY_EMP_ID), EntryName: get(C.ENTRY_NAME),
          EnrolledBy: get(C.ENROLLED_BY), status
        };
      });
    return { success: true, data: result };
  } catch (e) { return { success: false, error: e.message }; }
}

function getAssetByBarcode(barcode) {
  try {
    const sh  = _entrySheet();
    const idx = _findRow(sh, barcode);
    if (idx < 1) return null;
    const row = sh.getRange(idx, 1, 1, sh.getLastColumn()).getValues()[0];
    const get = (col) => String(row[col - 1] || '');
    return {
      barcode: get(C.BARCODE), type: get(C.TYPE), brand: get(C.BRAND),
      serial: get(C.SERIAL), specs: get(C.SPECS),
      condition: get(C.CONDITION) || 'Good', lifecycle: get(C.LIFECYCLE) || 'Active',
      statusLabel: get(C.STATUS_LABEL) || 'Unassigned',
      purchaseDate: get(C.PURCH_DATE), warrantyTerm: get(C.WARRANTY_TERM),
      warrantyValidity: get(C.WARRANTY_VAL), remarks: get(C.REMARKS),
      employeeId: get(C.EMP_ID) || 'N/A', staff: get(C.STAFF) || 'Unassigned',
      designation: get(C.DESIGNATION), department: get(C.DEPARTMENT),
      baseOffice: get(C.BASE_OFFICE), division: _normDiv(get(C.DIVISION)),
      district: _normDist(get(C.DISTRICT)), area: get(C.AREA), branch: get(C.BRANCH),
      effDate: get(C.EFF_DATE), supplier: get(C.SUPPLIER),
      borrowDate: get(C.BOR_DATE), returnDate: get(C.EXP_RETURN),
      status: _computeStatus(get(C.LIFECYCLE), get(C.STATUS_LABEL), get(C.ACT_RETURN),
                             get(C.EMP_ID), get(C.BOR_NAME), get(C.ASSET_STATUS))
    };
  } catch (e) { return null; }
}

// ─── ASSETS: CREATE ──────────────────────────────────────────────────────────
function processAsset(obj, isAssign) {
  // 1. Get the script lock to prevent race conditions
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for other processes to finish
    lock.waitLock(10000); 
  } catch (e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }

  try {
    const sh  = _entrySheet();
    const now = new Date();
    const nowStr = now.toLocaleString('en-PH');

    if (!isAssign) {
      if (!obj.barcode) return 'Error: Barcode is required.';
      if (_findRow(sh, obj.barcode) > 0) return 'Error: Barcode already exists: ' + obj.barcode;
      
      const statusChoice = obj.statusChoice || 'Spare';
      const STATUS_MAP = {
        'Spare':      { lc: 'Active',     asSt: 'Active',     stLbl: 'Unassigned' },
        'Allocated':  { lc: 'Allocated',  asSt: 'Active',     stLbl: 'Assigned'   },
        'Disposal':   { lc: 'Dispose',    asSt: 'Disposal',   stLbl: 'Disposed'   },
        'BorrowItem': { lc: 'BorrowItem', asSt: 'BorrowItem', stLbl: 'Unassigned' }
      };
      const sm = STATUS_MAP[statusChoice] || STATUS_MAP['Spare'];
      const isSpare = statusChoice === 'Spare' || statusChoice === 'BorrowItem';

      const normDiv  = _normDiv(obj.division  || '');
      const normDist = _normDist(obj.district || '');

      const row = new Array(TOTAL_COLS).fill('');
      row[C.ENTRY_EMP_ID - 1] = obj.entryEmpId  || '';
      row[C.ENTRY_NAME   - 1] = obj.entryName   || '';
      row[C.EMP_ID       - 1] = isSpare ? '' : (obj.accEmpId  || '');
      row[C.STAFF        - 1] = isSpare ? '' : (_sanitize(obj.accName, 100)   || '');
      row[C.DESIGNATION  - 1] = isSpare ? '' : (obj.accRole   || '');
      row[C.DEPARTMENT   - 1] = obj.department  || ''; 
      row[C.BASE_OFFICE  - 1] = obj.baseOffice  || ''; 
      row[C.DIVISION     - 1] = normDiv;
      row[C.DISTRICT     - 1] = normDist;
      row[C.AREA         - 1] = obj.area        || '';
      row[C.BRANCH       - 1] = _sanitize(obj.branch, 150)      || '';
      row[C.EFF_DATE     - 1] = obj.effDate     || '';
      row[C.BARCODE      - 1] = obj.barcode;
      row[C.TYPE         - 1] = obj.type        || '';
      row[C.BRAND        - 1] = obj.brand       || '';
      row[C.SERIAL       - 1] = obj.serial ? String(obj.serial) : '';
      row[C.SPECS        - 1] = obj.specs       || '';
      row[C.CONDITION    - 1] = obj.condition   || 'New';
      row[C.LIFECYCLE    - 1] = sm.lc;
      row[C.ASSET_STATUS - 1] = sm.asSt;
      row[C.STATUS_LABEL - 1] = sm.stLbl;
      row[C.PURCH_DATE   - 1] = obj.purchDate   || '';
      row[C.WARRANTY_TERM- 1] = obj.wTerm       || '';
      row[C.WARRANTY_VAL - 1] = obj.wValidity   || '';
      row[C.REMARKS      - 1] = _sanitize(obj.remarks, 500)     || '';
      row[C.SUPPLIER     - 1] = obj.supplier    || '';
      row[C.LOCATION     - 1] = obj.location    || '';
      row[C.ENROLLED_BY  - 1] = obj.enrolledBy  || obj.entryEmpId || '';
      row[C.CREATED_AT   - 1] = nowStr;
      row[C.LAST_UPDATED - 1] = nowStr;

      // FIXED: Serial Duplicate Check (Uses 'sh' and logic from recommendation)
      if (obj.serial) {
        const curLast = sh.getLastRow();
        if (curLast >= AE_DATA_START) {
          const serials = sh.getRange(
            AE_DATA_START, 
            C.SERIAL, 
            curLast - AE_DATA_START + 1, 
            1
          ).getValues();
          
          const dupIdx = serials.findIndex(
            r => String(r[0]).trim() === String(obj.serial).trim()
          );

          if (dupIdx >= 0) {
            const existingBC = String(
              sh.getRange(dupIdx + AE_DATA_START, C.BARCODE).getValue()
            );
            return 'Error: Serial No. already registered under barcode: ' + existingBC;
          }
        }
      }

      sh.appendRow(row);
      const newRowIdx = sh.getLastRow();
      sh.getRange(newRowIdx, C.SERIAL).setNumberFormat('@STRING@');
      if (obj.serial) sh.getRange(newRowIdx, C.SERIAL).setValue(String(obj.serial));
      _log('CREATE', obj.barcode, obj.type + ' | ' + obj.brand + ' | ' + statusChoice, obj.entryEmpId || '');
      return 'Asset created: ' + obj.barcode;
    }

    // Logic for isAssign = true
    const lc    = obj.lifecycle || 'Allocated';
    const asSt  = lc === 'Transfer' ? 'Transfer' : lc === 'Dispose' ? 'Disposal' : 'Active';
    const staff = _sanitize((obj.staff || '').trim(), 100);
    const stLbl = (staff && staff !== 'Unassigned') ? 'Assigned' : 'Unassigned';
    let rowIdx  = _findRow(sh, obj.barcode);

    if (rowIdx < 1) {
      const newRow = new Array(TOTAL_COLS).fill('');
      newRow[C.BARCODE     - 1] = obj.barcode;
      newRow[C.CREATED_AT  - 1] = nowStr;
      newRow[C.LAST_UPDATED- 1] = nowStr;
      sh.appendRow(newRow);
      rowIdx = sh.getLastRow();
    }

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    lc], [C.ASSET_STATUS, asSt], [C.STATUS_LABEL, stLbl],
      [C.EMP_ID,       obj.employeeId  || ''], [C.STAFF,  staff || ''],
      [C.DESIGNATION,  obj.designation || ''],
      [C.DEPARTMENT,   obj.department  || ''],
      [C.BASE_OFFICE,  obj.baseOffice  || ''],
      [C.DIVISION, _normDiv(obj.division   || '')],
      [C.DISTRICT,     _normDist(obj.district || '')],
      [C.AREA,         obj.area || ''],
      [C.BRANCH,       _sanitize(obj.branch, 150)      || ''], [C.EFF_DATE, obj.effDate || '']
    ]);
    _log('ASSIGN', obj.barcode, staff + ' | ' + lc, obj.employeeId || '');
    return 'Asset assigned successfully';

  } catch (e) { 
    return 'Error: ' + e.message; 
  } finally {
    // 4. ALWAYS release the lock so other users can proceed
    lock.releaseLock();
  }
}

function deleteAssets(barcodes, callerEmpId) {
  // 1. SECURITY: Role Verification
  const masterSh = _ss().getSheetByName(SH_MASTER);
  if (masterSh && callerEmpId) {
    const last = masterSh.getLastRow();
    if (last >= 5) {
      const data = masterSh.getRange(5, 1, last - 4, 12).getValues();
      const caller = data.find(r => 
        String(r[0] || '').trim().toLowerCase() === String(callerEmpId).trim().toLowerCase()
      );
      
      if (!caller) return 'Error: Unauthorized — identity not verified.';
      
      const role = String(caller[11] || '').toLowerCase(); // Column L (12th column)
      if (!role.includes('senior') && !role.includes('admin')) {
        return 'Error: Unauthorized — you do not have permission to delete assets.';
      }
    }
  } else {
    return 'Error: Unauthorized — caller ID required.';
  }

  // 2. CONCURRENCY: Initialize Lock
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // Wait up to 10s for other actions to finish
  } catch (e) {
    return 'Error: System is busy. Please try again.';
  }

  try {
    const sh = _entrySheet();
    const blocked = [], toDelete = [];

    // Identify which rows to delete and which to block
    barcodes.forEach(bc => {
      const r = _findRow(sh, bc);
      if (r > 0) {
        const lc = String(sh.getRange(r, C.LIFECYCLE).getValue() || '').toLowerCase();
        // Block deletion if asset is active in a specific lifecycle
        if (lc === 'borrow' || lc === 'transfer' || lc === 'borrowitem') { 
          blocked.push(bc + ' (' + lc + ')'); 
        } else { 
          toDelete.push({ bc, r }); 
        }
      }
    });
    
    if (blocked.length && !toDelete.length) {
      return 'Error: Cannot delete — active lifecycle: ' + blocked.join(', ');
    }

    // 3. EXECUTION: Delete rows from BOTTOM to TOP
    // Sorting descending (b.r - a.r) is vital so row indices don't shift
    toDelete.sort((a, b) => b.r - a.r).forEach(({ bc, r }) => {
      sh.deleteRow(r); 
      _log('DELETE', bc, 'Deleted by ' + callerEmpId, callerEmpId);
    });

    let msg = 'Deleted ' + toDelete.length + ' asset(s)';
    if (blocked.length) {
      msg += '. Skipped ' + blocked.length + ' (Active Borrow/Transfer): ' + blocked.join(', ');
    }
    return msg;

  } catch (e) { 
    return 'Error: ' + e.message; 
  } finally {
    lock.releaseLock(); // Always release the lock
  }
}

// ─── ALLOCATE ASSET ──────────────────────────────────────────────────────────
function allocateAsset(obj) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }
  try {
    if (!obj.barcode)   return 'Error: Barcode is required.';
    if (!obj.empId)     return 'Error: Employee ID is required.';
    if (!obj.staffName) return 'Error: Staff name is required.';
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, obj.barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + obj.barcode;
    const currentLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
    if (currentLC === 'borrow')    return 'Error: Asset is currently borrowed. Return it first.';
    if (currentLC === 'dispose')   return 'Error: Disposed assets cannot be allocated.';
    if (currentLC === 'transfer')  return 'Error: Asset is in an active transfer.';
    const currentAS = String(sh.getRange(rowIdx, C.ASSET_STATUS).getValue() || '').toLowerCase();
    if (currentAS === 'borrowitem') {
      return 'Error: This asset is part of the Borrow Pool and cannot be permanently allocated. '
        + 'To assign it, first change its status via the Spare Pool.';
    }
    if (currentLC === 'allocated') {
  // Log the implicit deallocation before re-allocating
  const prevStaff = String(sh.getRange(rowIdx, C.STAFF).getValue() || '');
  _log('DEALLOCATE', obj.barcode, 
    'Implicit dealloc from: ' + prevStaff + ' → re-allocate to: ' + obj.staffName, 
    obj.allocatedBy || obj.empId || '');
}
    const nowStr = new Date().toLocaleString('en-PH');
    const rowData = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    const get = c => String(rowData[c - 1] || '');

    const normDiv  = _normDiv(obj.division  || '');
    const normDist = _normDist(obj.district || '');

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'], [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],  [C.EMP_ID,       obj.empId       || ''],
      [C.STAFF,        _sanitize(obj.staffName, 100) || ''], [C.DESIGNATION,  obj.designation || ''],
      [C.DEPARTMENT,   obj.department  || ''],   // ← Added
      [C.BASE_OFFICE,  obj.baseOffice  || ''],   // ← Added
      [C.DIVISION,     normDiv],                  [C.DISTRICT,     normDist],
      [C.AREA,         obj.area      || ''],      [C.BRANCH,       _sanitize(obj.branch, 150)      || ''],
      [C.EFF_DATE,     obj.effDate   || nowStr],  [C.REMARKS,      _sanitize(obj.remarks, 500)     || '']
    ]);
    _allocLogSheet().appendRow([
      obj.barcode, obj.type || get(C.TYPE), obj.brand || get(C.BRAND),
      obj.serial || get(C.SERIAL), obj.empId, _sanitize(obj.staffName, 100),
      obj.designation || '', obj.department || '', obj.baseOffice || '',
      normDiv, normDist, obj.area || '', _sanitize(obj.branch, 150),
      obj.effDate || nowStr, obj.condition || get(C.CONDITION) || 'Good',
      _sanitize(obj.remarks, 500), nowStr, obj.allocatedBy || ''
    ]);
    _log('ALLOCATE', obj.barcode, _sanitize(obj.staffName, 100) + ' | ' + (_sanitize(obj.branch, 150) || obj.baseOffice || normDiv || ''), obj.allocatedBy || obj.empId || '');
    return 'Asset allocated to ' + _sanitize(obj.staffName, 100);
  } catch(e) {
    return 'Error: ' + e.message;
  } finally {
    lock.releaseLock();
  }
}

// ─── DEALLOCATE ASSET ────────────────────────────────────────────────────────
function deallocateAsset(barcode, remarks) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }
  try {
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;
    const currentLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
    if (currentLC === 'borrow')  return 'Error: Return the borrow record first.';
    if (currentLC === 'dispose') return 'Error: Disposed assets cannot be returned to spare.';
    const prevStaff = String(sh.getRange(rowIdx, C.STAFF).getValue() || '');
    const updates = [
      [C.LIFECYCLE,   'Active'], [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL,'Unassigned'], [C.EMP_ID, ''], [C.STAFF, ''],
      [C.DESIGNATION, ''], [C.EFF_DATE, ''],
      // Note: preserve Department/BaseOffice — the asset stays in its location
      [C.DIVISION, ''], [C.DISTRICT, ''], [C.AREA, ''], [C.BRANCH, '']
    ];
    if (remarks) updates.push([C.REMARKS, _sanitize(remarks, 500)]);
    _setRow(sh, rowIdx, updates);
    _log('DEALLOCATE', barcode, `From: ${prevStaff} → Spare Pool. ${_sanitize(remarks, 500)||''}`, '');
    return 'Asset returned to spare pool';
  } catch (e) { return 'Error: ' + e.message; } finally {
    lock.releaseLock();
  }
}

// ─── UPDATE ASSET DETAILS (UX-5: Edit Modal) ──────────────────────────────────
function updateAssetDetails(barcode, updates) {
  try {
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;

    const fields = [];
    if (updates.brand !== undefined) fields.push([C.BRAND, updates.brand]);
    if (updates.condition !== undefined) fields.push([C.CONDITION, updates.condition]);
    if (updates.purchDate !== undefined) fields.push([C.PURCH_DATE, updates.purchDate]);
    if (updates.specs !== undefined) fields.push([C.SPECS, updates.specs]);
    if (updates.remarks !== undefined) fields.push([C.REMARKS, updates.remarks]);

    // Serial requires duplicate check
    if (updates.serial !== undefined) {
      const currentSerial = String(
        sh.getRange(rowIdx, C.SERIAL).getValue() || ''
      ).trim();
      if (updates.serial !== currentSerial && updates.serial) {
        const last = sh.getLastRow();
        if (last >= AE_DATA_START) {
          const serials = sh.getRange(
            AE_DATA_START, C.SERIAL,
            last - AE_DATA_START + 1, 1
          ).getValues();
          const dupIdx = serials.findIndex(function(r) {
            return String(r[0]).trim() === updates.serial &&
                   (AE_DATA_START + serials.indexOf(r)) !== rowIdx;
          });
          if (dupIdx >= 0) {
            return 'Error: Serial No. already used by another asset.';
          }
        }
      }
      fields.push([C.SERIAL, updates.serial]);
    }

    _setRow(sh, rowIdx, fields);
    _log('EDIT', barcode, 
      'Updated: ' + Object.keys(updates).join(', '), 
      '');
    return 'Asset details updated.';
  } catch (e) { return 'Error: ' + e.message; }
}

// ─── GET ASSETS BY POOL ───────────────────────────────────────────────────────
function getSpareAssets() {
  try {
    const all = getAllAssets();
    if (!all.success) return { success: false, error: all.error };
    return {
      success: true,
      data: all.data.filter(a => {
        const lc = (a.Lifecycle || '').toLowerCase();
        const as = (a.AssetStatus || '').toLowerCase();
        return lc === 'active' && as !== 'borrowitem';
      })
    };
  } catch (e) { return { success: false, error: e.message }; }
}

function getAllocatedAssets(districtFilter) {
  try {
    const all = getAllAssets();
    if (!all.success) return { success: false, error: all.error };
    let data = all.data.filter(a => {
      const lc = (a.Lifecycle || '').toLowerCase();
      return lc === 'allocated' || lc === 'transfer';
    });
    if (districtFilter && districtFilter !== 'all' && districtFilter !== '') {
      data = data.filter(a =>
        (a.District || '').toLowerCase() === districtFilter.toLowerCase()
      );
    }
    return { success: true, data };
  } catch (e) { return { success: false, error: e.message }; }
}

function getBorrowPoolAssets() {
  try {
    const all = getAllAssets();
    if (!all.success) return { success: false, error: all.error };
    return {
      success: true,
      data: all.data.filter(a => (a.AssetStatus || '').toLowerCase() === 'borrowitem')
    };
  } catch (e) { return { success: false, error: e.message }; }
}

function getActiveBorrows() {
  try {
    const all = getAllAssets();
    if (!all.success) return { success: false, error: all.error };
    return {
      success: true,
      data: all.data.filter(a => (a.Lifecycle || '').toLowerCase() === 'borrow')
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── TRANSFERS ───────────────────────────────────────────────────────────────
function saveTransfer(t) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }
  try {
    if (!t.barcode)      return 'Error: Barcode is required.';
    if (!t.toEmpId)      return 'Error: Destination Employee ID is required.';
    if (!t.toStaff)      return 'Error: Destination Staff Name is required.';
    if (!t.effDate)      return 'Error: Transfer Date is required.';
    if (!t.transferType) return 'Error: Transfer Type is required.';
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, t.barcode);
    if (rowIdx < 1) return 'Error: Asset barcode not found: ' + t.barcode;
    const currentLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
    if (currentLC === 'dispose' || currentLC === 'disposal')
      return 'Error: Cannot transfer a disposed asset.';
    if (currentLC === 'borrow')
      return 'Error: Cannot transfer an asset currently on borrow. Return it first.';
    const xSh    = _xferSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    xSh.appendRow([
      t.barcode, t.transferType,
      _sanitize(t.fromStaff, 100), t.fromEmpId, t.fromDesig, t.fromDiv, t.fromDist, t.fromArea, _sanitize(t.fromBranch, 150), _sanitize(t.fromRemarks, 500),
      _sanitize(t.toStaff, 100), t.toEmpId, t.toDesig, t.toDiv, t.toDist, t.toArea, _sanitize(t.toBranch, 150), _sanitize(t.toRemarks, 500),
      t.effDate, t.status || 'Completed', nowStr
    ]);

    const normToDiv  = _normDiv(t.toDiv   || '');
    const normToDist = _normDist(t.toDist || '');

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'], [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],  [C.EMP_ID,       t.toEmpId   || ''],
      [C.STAFF,        _sanitize(t.toStaff, 100)   || ''],  [C.DESIGNATION,  t.toDesig   || ''],
      [C.DIVISION,     normToDiv],          [C.DISTRICT,     normToDist],
      [C.AREA,         t.toArea    || ''],  [C.BRANCH,       _sanitize(t.toBranch, 150)  || ''],
      [C.EFF_DATE,     t.effDate],          [C.XFER_TYPE,    t.transferType || 'Permanent'],
      [C.FR_STAFF,     _sanitize(t.fromStaff, 100)   || ''],[C.FR_EMPID,     t.fromEmpId   || ''],
      [C.FR_DESIG,     t.fromDesig   || ''],[C.FR_DIV,       t.fromDiv     || ''],
      [C.FR_DIST,      t.fromDist    || ''],[C.FR_AREA,      t.fromArea    || ''],
      [C.FR_BRANCH,    _sanitize(t.fromBranch, 150)  || ''],[C.FR_REMARKS,   _sanitize(t.fromRemarks, 500) || ''],
      [C.TO_STAFF,     _sanitize(t.toStaff, 100)   || ''],  [C.TO_EMPID,     t.toEmpId   || ''],
      [C.TO_DESIG,     t.toDesig   || ''],  [C.TO_DIV,       normToDiv],
      [C.TO_DIST,      normToDist],         [C.TO_AREA,      t.toArea    || ''],
      [C.TO_BRANCH,    _sanitize(t.toBranch, 150)  || ''],  [C.TO_REMARKS,   _sanitize(t.toRemarks, 500) || ''],
      [C.XFER_DATE,    t.effDate]
    ]);
    _log('TRANSFER', t.barcode, (_sanitize(t.fromStaff, 100) || '—') + ' → ' + _sanitize(t.toStaff, 100), t.fromEmpId || '');
    return 'Transfer saved';
  } catch (e) { return 'Error: ' + e.message; } finally {
    lock.releaseLock();
  }
}

function getTransferData() {
  try {
    const sh   = _xferSheet();
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 21).getValues()
      .map(r => r.map(v => String(v || '')));
  } catch (e) { return []; }
}

// ─── BORROWS ─────────────────────────────────────────────────────────────────
function saveBorrow(b) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }
  try {
    const sh     = _entrySheet();
    const bSh    = _borrowSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const rowIdx = _findRow(sh, b.barcode);
    if (rowIdx < 1) return 'Error: Asset barcode not found: ' + b.barcode;
    const rowDataRaw  = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    const currentLC   = String(rowDataRaw[C.LIFECYCLE - 1] || '').toLowerCase();
    const currentAS   = String(rowDataRaw[C.ASSET_STATUS - 1] || '');
    const isBorrowItem = currentAS.toLowerCase() === 'borrowitem';
    const borrowable  = ['active', 'allocated', 'borrowitem'];
    if (!borrowable.includes(currentLC)) {
      const lcDisplay = {
        borrow: 'already on borrow', returned: 'Returned — re-allocate first',
        transfer: 'in active Transfer', dispose: 'Disposed', disposal: 'Disposed'
      };
      return 'Error: Cannot borrow. Status: ' + (lcDisplay[currentLC] || currentLC);
    }
    bSh.appendRow([
      b.barcode, _sanitize(b.borrowerName, 100), b.empId, b.designation,
      b.division, b.district||'', _sanitize(b.branch, 150), b.borrowDate, b.expectedReturn,
      '', 'Borrow', _sanitize(b.remarks, 500), nowStr
    ]);
    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Borrow'],
      [C.ASSET_STATUS, isBorrowItem ? 'BorrowItem' : 'Active'],
      [C.STATUS_LABEL, 'Assigned'],
      [C.BOR_NAME,     _sanitize(b.borrowerName, 100)   || ''],
      [C.BOR_EMPID,    b.empId          || ''],
      [C.BOR_DESIG,    b.designation    || ''],
      [C.BOR_DIV,      b.division       || ''],
      [C.BOR_DIST,     b.district       || ''],
      [C.BOR_BRANCH,   _sanitize(b.branch, 150)         || ''],
      [C.BOR_DATE,     b.borrowDate     || ''],
      [C.EXP_RETURN,   b.expectedReturn || ''],
      [C.ACT_RETURN,   ''],
      [C.BOR_REMARKS,  _sanitize(b.remarks, 500)        || '']
    ]);
    _log('BORROW', b.barcode, _sanitize(b.borrowerName, 100) + ' | due: ' + b.expectedReturn, b.empId || '');
    return 'Borrow saved';
  } catch (e) { return 'Error: ' + e.message; } finally {
    lock.releaseLock();
  }
}

function getBorrowData() {
  try {
    const sh   = _borrowSheet();
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 13).getValues().map(r => ({
      barcode:       String(r[0]  || ''), borrowerName:   String(r[1]  || ''),
      empId:         String(r[2]  || ''), designation:    String(r[3]  || ''),
      division:      String(r[4]  || ''), district:       String(r[5]  || ''),
      branch:        String(r[6]  || ''), borrowDate:     String(r[7]  || ''),
      expectedReturn:String(r[8]  || ''), actualReturn:   String(r[9]  || ''),
      status:        String(r[10] || 'Borrow'), remarks:  String(r[11] || ''),
      timestamp:     String(r[12] || '')
    }));
  } catch (e) { return []; }
}

function returnAsset(barcode, returnDate) {
  try {
    const sh      = _entrySheet();
    const bSh     = _borrowSheet();
    const retDate = returnDate || new Date().toLocaleDateString('en-PH');
    const last    = bSh.getLastRow();
    if (last > 1) {
      const data = bSh.getRange(2, 1, last - 1, 11).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (
          String(data[i][0]) === String(barcode) &&
          String(data[i][10]) === 'Borrow'
        ) {
          bSh.getRange(i + 2, 10).setValue(retDate);
          bSh.getRange(i + 2, 11).setValue('Returned');
          break;
        }
      }
    }
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx > 0) {
      const rowData      = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
      const origStaff    = String(rowData[C.STAFF - 1] || '').trim();
      const origAS       = String(rowData[C.ASSET_STATUS - 1] || '');
      const isBorrowItem = origAS.toLowerCase() === 'borrowitem';
      const hasOwner     = !isBorrowItem && origStaff && origStaff !== 'Unassigned';
      const restoredLC   = isBorrowItem ? 'BorrowItem' : (hasOwner ? 'Allocated' : 'Active');
      const restoredLbl  = (isBorrowItem || !hasOwner) ? 'Unassigned' : 'Assigned';
      const restoredAS   = isBorrowItem ? 'BorrowItem' : 'Active';
      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    restoredLC],
        [C.STATUS_LABEL, restoredLbl],
        [C.ASSET_STATUS, restoredAS],
        [C.ACT_RETURN,   retDate],
        [C.BOR_NAME,     ''],
        [C.BOR_EMPID,    ''],
        [C.BOR_DESIG,    ''],
        [C.BOR_DIV,      ''],
        [C.BOR_DIST,     ''],
        [C.BOR_BRANCH,   ''],
        [C.BOR_DATE,     ''],
        [C.EXP_RETURN,   ''],
        [C.BOR_REMARKS,  '']
      ]);
    }
    _log('RETURN', barcode, retDate, '');
    return 'Asset returned';
  } catch (e) { return 'Error: ' + e.message; }
}

// ─── DISPOSALS ───────────────────────────────────────────────────────────────
function saveDisposal(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
  } catch(e) {
    return 'Error: System is busy. Please try again in a few seconds.';
  }
  try {
    const sh     = _entrySheet();
    const dSh    = _disposeSheet();
    const nowStr = new Date().toLocaleString('en-PH');
    const rowIdx = _findRow(sh, d.barcode);
    if (rowIdx > 0) {
      const currentLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
      if (currentLC === 'borrow')   return 'Error: Cannot dispose an asset that is currently borrowed.';
      if (currentLC === 'transfer') return 'Error: Cannot dispose an asset in active transfer.';
    }
    dSh.appendRow([d.barcode, _sanitize(d.reason, 200), _sanitize(d.disposedBy, 100), d.disposeDate, _sanitize(d.remarks, 500), nowStr]);
    if (rowIdx > 0) {
      const cur = String(sh.getRange(rowIdx, C.REMARKS).getValue() || '');
      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    'Dispose'], [C.ASSET_STATUS, 'Disposal'],
        [C.STATUS_LABEL, 'Disposed'],
        [C.REMARKS, (cur ? cur + ' | ' : '') + 'DISPOSAL: ' + _sanitize(d.reason, 200) + ' by ' + _sanitize(d.disposedBy, 100)]
      ]);
    }
    _log('DISPOSE', d.barcode, _sanitize(d.reason, 200) + ' | ' + _sanitize(d.disposedBy, 100), d.disposedBy || '');
    return 'Disposal recorded';
  } catch (e) { return 'Error: ' + e.message; } finally {
    lock.releaseLock();
  }
}

function getDisposalData() {
  try {
    const sh   = _disposeSheet();
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 6).getValues().map(r => r.map(v => String(v || '')));
  } catch (e) { return []; }
}

// ─── USERS ───────────────────────────────────────────────────────────────────
function getUserList() {
  try {
    const sh   = _ss().getSheetByName(SH_MASTER);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 5) return [];
    return sh.getRange(5, 1, last - 4, Math.min(sh.getLastColumn(), 12))
      .getValues().filter(r => String(r[0] || '').trim()).map(r => r.map(v => String(v || '')));
  } catch (e) { return []; }
}

function getEmployeeById(empId) {
  try {
    const sh = _ss().getSheetByName(SH_MASTER);
    if (!sh) return { ok: false, error: 'Masterlist not found.' };
    const last = sh.getLastRow();
    if (last < 5) return { ok: false, error: 'No employees found.' };
    const data = sh.getRange(5, 1, last - 4, Math.min(sh.getLastColumn(), 12)).getValues();
    const id   = String(empId).trim().toLowerCase();
    for (const row of data) {
      if (String(row[0] || '').trim().toLowerCase() !== id) continue;
      return {
        ok:         true,
        empId:      String(row[0]  || '').trim(),
        name:       String(row[2]  || '').trim(),
        division:   _normDiv(String(row[4]  || '').trim()),
        district:   _normDist(String(row[5] || '').trim()),
        area:       String(row[6]  || '').trim(),
        branch:     String(row[7]  || '').trim(),
        position:   String(row[11] || '').trim()
      };
    }
    return { ok: false, error: 'Employee ID not found: ' + empId };
  } catch (e) { return { ok: false, error: e.message }; }
}

// ─── DROPDOWN DATA ────────────────────────────────────────────────────────────
function getDropdownData() {
  try {
    const sh = _ss().getSheetByName(SH_DROPDOWN);
    if (!sh) return { categories: [], brands: {}, models: {}, suppliers: [], laptopSpecs: [] };
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1)
      return { categories: [], brands: {}, models: {}, suppliers: [], laptopSpecs: [] };
    const data  = sh.getRange(1, 1, lastRow, lastCol).getValues();
    const row1  = data[0], row2 = data[1];
    const CAT_MAP = {
      'LAPTOP': 'Laptop', 'LAPTOP ADAPTOR': 'Laptop Adaptor', 'CPU': 'CPU',
      'MONITOR': 'Monitor', 'KEYBOARD': 'Keyboard', 'MOUSE': 'Mouse',
      'PRINTER': 'Printer',
      'SCANNER': 'Scanner', 'SCANSNAP': 'Scansnap',
      'SCANNER SCANSNAP': 'Scanner',
      'UPS': 'UPS', 'EXTERNAL DRIVE': 'External Drive',
      'CAMERA': 'Camera', 'SPEAKER': 'Speaker'
    };
    const result = { categories: [], brands: {}, models: {}, suppliers: [], laptopSpecs: [], laptopSpecValues: {} };
    let catPositions = [], supplierCol = -1, laptopSpecCol = -1, fieldSectionCol = -1;
    let usedMergedScannerHeader = false;

    for (let c = 0; c < row1.length; c++) {
      const raw = String(row1[c] || '').trim();
      if (!raw) continue;
      const up = raw.toUpperCase();
      if (up === 'SUPPLIERS' || up === 'SUPPLIER') { supplierCol = c; continue; }
      if (up === 'LAPTOP SPECS' || up === 'LAPTOP SPEC') { laptopSpecCol = c; continue; }
      if (up === 'FIELD' || up === 'FIELD ' || up.startsWith('DIVISION') || up.startsWith('DISTRICT')) {
        if (fieldSectionCol === -1) fieldSectionCol = c;
        continue;
      }
      if (up === 'SCANNER SCANSNAP') {
        catPositions.push({ name: 'Scanner', col: c });
        usedMergedScannerHeader = true;
        continue;
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
        if (key.startsWith('Scanner|')) {
          result.models['Scansnap|' + key.slice('Scanner|'.length)] = [...result.models[key]];
        }
      });
    }

    if (supplierCol > -1) {
      for (let r = 1; r < lastRow; r++) {
        const s = String(data[r][supplierCol] || '').trim();
        if (s) result.suppliers.push(s);
      }
    }
    if (laptopSpecCol > -1) {
      const specEndCol = fieldSectionCol > -1 ? fieldSectionCol : lastCol;
      for (let c = laptopSpecCol; c < specEndCol; c++) {
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
  } catch (e) {
    return { categories: [], brands: {}, models: {}, suppliers: [], laptopSpecs: [], error: e.message };
  }
}

function getHeadOfficeDepts() {
  try {
    const ss    = _ss();
    const ddSh  = ss.getSheetByName(SH_DROPDOWN);
    if (!ddSh || ddSh.getLastRow() < 2) return [];
    const lastCol  = ddSh.getLastColumn();
    const lastRow  = ddSh.getLastRow();
    const headerRow = ddSh.getRange(1, 1, 1, lastCol).getValues()[0];
    let deptCol = -1;
    for (let c = 0; c < headerRow.length; c++) {
      const h = String(headerRow[c] || '').trim().toUpperCase();
      if (h === 'DEPARTMENTS' || h === 'HEAD OFFICE' ||
          h === 'HO DEPTS'   || h === 'HEAD OFFICE DEPTS') {
        deptCol = c; break;
      }
    }
    if (deptCol < 0) return [];
    const depts = [];
    for (let r = 1; r < lastRow; r++) {
      const v = String(ddSh.getRange(r + 1, deptCol + 1).getValue() || '').trim();
      if (v) depts.push(v);
    }
    return depts;
  } catch (e) { return []; }
}

// ─── ENGINEER LOCATION DATA ───────────────────────────────────────────────────
function getEngineerLocationData() {
  return getLocationData(null);
}

function getLocationData(empId) {
  try {
    const ss = _ss();
    const result = {
      divDistrictMap:        {},
      userDivisions:         [],
      userDistricts:         [],
      userDivisionDistricts: {}
    };

    const ddSh = ss.getSheetByName(SH_DROPDOWN);
    if (ddSh && ddSh.getLastRow() >= 2) {
      const lastCol = ddSh.getLastColumn();
      const lastRow = ddSh.getLastRow();
      let divBlockStartCol = -1;
      if (lastCol >= 1) {
        const headerRow = ddSh.getRange(1, 1, 1, lastCol).getValues()[0];
        for (let c = 0; c < headerRow.length; c++) {
          const hdr = String(headerRow[c] || '').trim().toUpperCase();
          if (hdr.startsWith('FIELD') || hdr.startsWith('DIVISION')) {
            divBlockStartCol = c + 1;
            break;
          }
        }
      }
      if (divBlockStartCol < 1) divBlockStartCol = 46;
      const blockWidth = lastCol - divBlockStartCol + 1;
      if (blockWidth > 0 && lastCol >= divBlockStartCol) {
        const maxRows = Math.max(lastRow, 2);
        const block = ddSh.getRange(1, divBlockStartCol, maxRows, blockWidth).getValues();
        const divRow   = block[1];
        const distRows = block.slice(2);
        divRow.forEach((divName, ci) => {
          const div = String(divName || '').trim();
          if (!div) return;
          const normalDiv = div.replace(/^DIv/i, 'Div');
          const districts = [];
          distRows.forEach(dRow => {
            const d = String(dRow[ci] || '').trim();
            if (d) districts.push(d);
          });
          result.divDistrictMap[normalDiv] = districts;
        });
      }
    }

    const engSh = ss.getSheetByName('Eng. List') || ss.getSheetByName('Eng List');
    if (engSh && engSh.getLastRow() > 1) {
      const lastRow = engSh.getLastRow();
      const data = engSh.getRange(2, 1, lastRow - 1, 10).getValues();
      const divSet = new Set(), distSet = new Set();

if (empId) {
  // Look up this specific user
  data.forEach(r => {
    const rowEmpId = String(r[0] || '').trim();
    if (rowEmpId.toLowerCase() === String(empId).trim().toLowerCase()) {
      const div  = String(r[7] || '').trim();
      const dist = String(r[8] || '').trim();
      if (div)  divSet.add(div);
      if (dist) distSet.add(dist);
    }
  });
  // If still empty after targeted search, 
  // fall back to Masterlist data (already in SESSION), not all engineers
  // Just return empty sets — let the caller handle it
} else {
  // No empId — load all (used for dropdown population, not scoping)
  data.forEach(r => {
    const div  = String(r[7] || '').trim();
    const dist = String(r[8] || '').trim();
    if (div)  divSet.add(div);
    if (dist) distSet.add(dist);
  });
}

      const numSort = (a, b) => {
        const na = parseInt(a.replace(/\D+/g, '')) || 0;
        const nb = parseInt(b.replace(/\D+/g, '')) || 0;
        return na !== nb ? na - nb : a.localeCompare(b);
      };
      result.userDivisions = [...divSet].sort(numSort);
      result.userDistricts = [...distSet].sort(numSort);

      result.userDivisions.forEach(div => {
        const mapped = result.divDistrictMap[div] || [];
        result.userDivisionDistricts[div] = mapped.length > 0 ? mapped : result.userDistricts;
      });

      const _hoDepts = getHeadOfficeDepts();
      result.headOfficeDepts = _hoDepts;
      if (_hoDepts.length > 0) {
        result.divDistrictMap['Head Office'] = _hoDepts;
        result.userDivisionDistricts['Head Office'] = _hoDepts;
      }
    }

    return result;
  } catch (e) {
    return { divDistrictMap: {}, userDivisions: [], userDistricts: [], userDivisionDistricts: {}, headOfficeDepts: [], error: e.message };
  }
}

// ─── ACTIVITY LOG ─────────────────────────────────────────────────────────────
function _log(action, barcode, details, performer) {
  try {
    _logSheet().appendRow([
      new Date().toLocaleString('en-PH'),
      action, barcode, details, performer || ''
    ]);
  } catch (e) {}
}

function getActivityLogs(page, pageSize) {
  try {
    const sh   = _logSheet();
    const last = sh.getLastRow();
    if (last < 2) return { rows: [], total: 0 };
    const all = sh.getRange(2, 1, last - 1, 5).getValues().reverse()
      .map(r => ({
        timestamp: String(r[0] || ''), action: String(r[1] || ''),
        barcode:   String(r[2] || ''), details: String(r[3] || ''),
        performer: String(r[4] || '')
      }));
    const total = all.length;
    const ps    = (pageSize && pageSize > 0) ? Number(pageSize) : 100;
    const pg    = (page     && page     > 0) ? Number(page)     : 1;
    const rows  = all.slice((pg - 1) * ps, pg * ps);
    return { rows, total, page: pg, pageSize: ps, totalPages: Math.ceil(total / ps) };
  } catch (e) { return { rows: [], total: 0 }; }
}

function syncAll() { return getAllAssets(); }
function getSpreadsheetUrl() { return _ss().getUrl(); }

// ─── STAFF MOVEMENT ───────────────────────────────────────────────────────────
function moveStaff(empId, newDiv, newDist, newArea, newBranch, assetAction) {
  try {
    if (!empId) return 'Error: Employee ID is required.';
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return 'Error: No assets in system.';

    const count  = last - AE_DATA_START + 1;
    // Read ALL columns once
    const allData = sh.getRange(AE_DATA_START, 1, count, TOTAL_COLS).getValues();
    const nowStr  = new Date().toLocaleString('en-PH');
    const id      = String(empId).trim().toLowerCase();
    const normDiv  = _normDiv(newDiv   || '');
    const normDist = _normDist(newDist || '');
    
    let updated = 0, skipped = 0;
    const rowsToWrite = []; // { rowIdx, rowData }

    for (let i = 0; i < allData.length; i++) {
      const row      = allData[i];
      const rowEmpId = String(row[C.EMP_ID - 1] || '').trim().toLowerCase();
      if (rowEmpId !== id) continue;

      const lc = String(row[C.LIFECYCLE - 1] || '').toLowerCase();
      if (lc === 'borrow')                      { skipped++; continue; }
      if (lc === 'dispose' || lc === 'disposal') continue;

      // Clone the row, modify in memory
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
      newRow[C.DIVISION     - 1] = normDiv;
      newRow[C.DISTRICT     - 1] = normDist;
      newRow[C.AREA         - 1] = newArea   || '';
      newRow[C.BRANCH       - 1] = newBranch || '';
      newRow[C.LAST_UPDATED - 1] = nowStr;
      
      rowsToWrite.push({ rowIdx: i + AE_DATA_START, rowData: newRow });
      updated++;
    }

    // Batch write: one setValue call per changed row (not per cell)
    rowsToWrite.forEach(({ rowIdx, rowData }) => {
      sh.getRange(rowIdx, 1, 1, TOTAL_COLS).setValues([rowData]);
    });

    _log('MOVE_STAFF', empId,
      `Action:${assetAction} → ${newDiv}/${newDist}/${newBranch} | ` +
      `${updated} updated, ${skipped} skipped`, empId);

    let msg = `Staff movement recorded. ${updated} asset(s) ` +
      `${assetAction === 'spare' ? 'returned to spare' : 'moved to new location'}.`;
    if (skipped) msg += ` (${skipped} skipped — on active borrow)`;
    return msg;
  } catch (e) { return 'Error: ' + e.message; }
}


function moveOrgUnit(unitType, currentDiv, currentDist, currentArea, currentBranch, newDiv, newDist, newArea) {
  try {
    const sh   = _entrySheet();
    const last = sh.getLastRow();
    if (last < AE_DATA_START) return 'Error: No assets found.';
    const count  = last - AE_DATA_START + 1;
    const data   = sh.getRange(AE_DATA_START, 1, count, TOTAL_COLS).getValues();
    const nowStr = new Date().toLocaleString('en-PH');
    const normCurDiv  = _normDiv(currentDiv   || '');
    const normCurDist = _normDist(currentDist || '');
    const normNewDiv  = _normDiv(newDiv   || '');
    const normNewDist = _normDist(newDist || '');
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
  // Moving a district means reassigning it to a different division
  if (rowDist === normCurDist && rowDiv === normCurDiv) {
    updates.push([C.DIVISION, normNewDiv]);
    // Only update district name if it changed (rename scenario)
    if (normNewDist && normNewDist !== normCurDist) {
      updates.push([C.DISTRICT, normNewDist]);
    }
  }
} else if (unitType === 'area') {
  if (rowArea === currentArea && 
      rowDist === normCurDist && 
      rowDiv  === normCurDiv) {
    updates.push([C.DIVISION, normNewDiv]);
    updates.push([C.DISTRICT, normNewDist]);
    // Keep area name — only its parent changes
  }
} else if (unitType === 'branch') {
  if (rowBranch === currentBranch && 
      rowArea   === currentArea   &&
      rowDist   === normCurDist   && 
      rowDiv    === normCurDiv) {
    updates.push([C.DIVISION, normNewDiv]);
    updates.push([C.DISTRICT, normNewDist]);
    updates.push([C.AREA, newArea || currentArea]); // newArea from modal
  }
}
      if (updates.length) {
        updates.push([C.LAST_UPDATED, nowStr]);
        updates.forEach(u => sh.getRange(rowIdx, u[0]).setValue(u[1]));
        updated++;
      }
    });
    _log('MOVE_ORG', unitType.toUpperCase(),
      `${currentDiv}/${currentDist}/${currentArea}/${currentBranch} → ${newDiv}/${newDist}/${newArea} | ${updated} assets`, '');
    return updated > 0
      ? `${unitType.charAt(0).toUpperCase() + unitType.slice(1)} moved. ${updated} asset(s) updated.`
      : 'Move recorded — no matching assets found (check filters or scope).';
  } catch(e) { return 'Error: ' + e.message; }
}

// ─── UPDATE ASSET DETAILS ─────────────────────────────────────────────────────
function updateAssetDetails(barcode, updates) {
  try {
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + barcode;

    const fields = [];
    if (updates.brand     !== undefined) fields.push([C.BRAND,     updates.brand]);
    if (updates.condition !== undefined) fields.push([C.CONDITION, updates.condition]);
    if (updates.purchDate !== undefined) fields.push([C.PURCH_DATE,updates.purchDate]);
    if (updates.specs     !== undefined) fields.push([C.SPECS,     updates.specs]);
    if (updates.remarks   !== undefined) fields.push([C.REMARKS,   updates.remarks]);

    // Serial needs duplicate check
    if (updates.serial !== undefined) {
      const currentSerial = String(
        sh.getRange(rowIdx, C.SERIAL).getValue() || ''
      ).trim();
      if (updates.serial !== currentSerial && updates.serial) {
        const last = sh.getLastRow();
        if (last >= AE_DATA_START) {
          const serials = sh.getRange(
            AE_DATA_START, C.SERIAL,
            last - AE_DATA_START + 1, 1
          ).getValues();
          const dupIdx = serials.findIndex(
            (r, i) => String(r[0]).trim() === updates.serial 
              && (i + AE_DATA_START) !== rowIdx
          );
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

// ─── BATCH DATA LOADING ───────────────────────────────────────────────────────
function getInitialData() {
  return {
    assets:    getAllAssets().data    || [],
    borrows:   getBorrowData()        || [],
    transfers: getTransferData()      || [],
    disposals: getDisposalData()      || [],
    logs:      getActivityLogs(1, 200).rows || []
  };
}