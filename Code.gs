// ═══════════════════════════════════════════════════════════
//  ASSET MANAGEMENT SYSTEM — Code.gs  (v5.4 — Bug-Fix Release)
//
//  Fixes applied vs v5.3:
//    BUG-C03 — Senior seniorDistrictScope now derived from
//              Eng.List supervisor column via getDistrictsBySupervisor()
//    BUG-C04 — getDropdownData() handles 'SCANNER SCANSNAP' merged header
//    BUG-H06 — SESSION.division/district prefer Eng.List (authoritative)
//    BUG-H07 — 'network engineer' removed from ADMIN_ROLES
//    BUG-H08 — 'senior network engineer' added to ADMIN_ROLES
//    BUG-M04 — returnAsset() now clears BOR_REMARKS after return
//    BUG-L05 — _normDist()/_normDiv() applied on write paths
//              (processAsset, allocateAsset)
//
//  ⚠ GOOGLE SHEETS MANUAL STEPS REQUIRED (see bottom of this file)
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

// ─── COLUMN MAP (1-based) ────────────────────────────────────────────────────
// NOTE (BUG-C02): The actual Asset Entry sheet must have exactly 56 columns.
//   Column 56 header = 'Borrow District'.
//   If it only has 55 columns, add column 56 manually — see MANUAL STEPS below.
const C = {
  ENTRY_EMP_ID:1, ENTRY_NAME:2,
  EMP_ID:3, STAFF:4, DESIGNATION:5, DIVISION:6,
  DISTRICT:7, AREA:8, BRANCH:9, EFF_DATE:10,
  BARCODE:11, TYPE:12, BRAND:13, SERIAL:14, SPECS:15,
  CONDITION:16, LIFECYCLE:17, STATUS_LABEL:18, ASSET_STATUS:19,
  PURCH_DATE:20, WARRANTY_TERM:21, WARRANTY_VAL:22, REMARKS:23,
  XFER_TYPE:24, FR_STAFF:25, FR_EMPID:26, FR_DESIG:27, FR_DIV:28,
  FR_DIST:29, FR_AREA:30, FR_BRANCH:31, FR_REMARKS:32,
  TO_STAFF:33, TO_EMPID:34, TO_DESIG:35, TO_DIV:36, TO_DIST:37,
  TO_AREA:38, TO_BRANCH:39, TO_REMARKS:40, XFER_DATE:41,
  BOR_NAME:42, BOR_EMPID:43, BOR_DESIG:44, BOR_DIV:45, BOR_BRANCH:46,
  BOR_DATE:47, EXP_RETURN:48, ACT_RETURN:49, BOR_REMARKS:50,
  CREATED_AT:51, LAST_UPDATED:52, SUPPLIER:53, LOCATION:54, ENROLLED_BY:55,
  BOR_DIST:56
};
const TOTAL_COLS = Math.max(...Object.values(C)); // = 56

// AE_HEADERS: 56 entries, one per column.
// BUG-M01 note: These match correct English names. The existing sheet may have
//   minor typos ('Assignement Status', 'Serial') — those are cosmetic and only
//   matter if the sheet is ever recreated from scratch.
const AE_HEADERS = [
  'Entry Employee ID','Entered By',
  'Accountable Employee ID','Accountable Staff','Designation','Division',
  'District','Area','Branch','Effectivity Date',
  'Barcode','Category','Brand','Serial No.','Specifications',
  'Condition','Lifecycle Status','Status Label','Assignment Status',
  'Date of Purchase','Warranty Term','Warranty Validity','Remarks',
  'Transfer Type','From Staff','From EmpID','From Designation','From Division',
  'From District','From Area','From Branch','From Remarks',
  'To Staff','To EmpID','To Designation','To Division','To District',
  'To Area','To Branch','To Remarks','Transfer Date',
  'Borrower Name','Borrower EmpID','Borrower Designation',
  'Borrow Division','Borrow Branch','Borrow Date',
  'Expected Return Date','Actual Return Date','Borrow Remarks',
  'Created At','Last Updated','Supplier','Location','Enrolled By',
  'Borrow District'   // ← col 56  (BUG-C02 fix: header now listed)
];

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
function _allocLogSheet(){ return _getOrCreate(SH_ALLOC,   ['Barcode','Category','Brand','Serial No.','Employee ID','Accountable Staff','Designation','Division','District','Area','Branch','Effectivity Date','Condition','Remarks','Timestamp','Allocated By']); }

// ─── ROW FINDERS / SETTERS ───────────────────────────────────────────────────
function _findRow(sheet, barcode) {
  const last = sheet.getLastRow();
  if (last < AE_DATA_START) return -1;
  const count = last - AE_DATA_START + 1;
  const vals = sheet.getRange(AE_DATA_START, C.BARCODE, count, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(barcode).trim()) return i + AE_DATA_START;
  }
  return -1;
}

function _setRow(sheet, rowIdx, updates) {
  updates.forEach(u => {
    sheet.getRange(rowIdx, u[0]).setValue(u[1] != null ? u[1] : '');
  });
  sheet.getRange(rowIdx, C.LAST_UPDATED).setValue(new Date().toLocaleString('en-PH'));
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
      as === 'disposal' || as === 'dispose')           return 'disposal';
  if (l === 'transfer')                                return 'transfer';
  if (l === 'allocated' || s === 'assigned')           return 'allocated';

  if (!l || l === 'active') {
    if (hasBor) return r ? 'returned' : 'borrowed';
    if (hasEmp) return 'allocated';
  }
  return 'spare';
}

// BUG-L05 FIX: normalisation helpers now also used on WRITE paths.
// Previously _normDiv/_normDist were only called during getAllAssets() reads,
// meaning saved values could be un-normalised and fail scope comparisons.
function _normDiv(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';
  return s.replace(/^DIv/i, m => 'Div');          // e.g. 'DIvision 1' → 'Division 1'
}
function _normDist(raw) {
  if (!raw) return '';
  const s = String(raw).trim();
  if (!s) return '';
  const m = s.match(/^district\s+0*(\d+)$/i);
  if (m) return 'District ' + parseInt(m[1]);     // 'district 05' → 'District 5'
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
// Two-tier role system:
//   SENIOR — Any role whose title contains the word 'senior'.
//            Division-scoped: sees all districts in their supervised division.
//   FE     — All other roles (Field Engineer, Network Engineer, etc.).
//            District-scoped: sees only their own assigned district.
//
// There is no longer a separate admin tier. All users can add, allocate,
// borrow, transfer, and dispose assets within their scope.

function _classifyRole(roleStr) {
  const r = String(roleStr || '').trim().toLowerCase();
  if (r.includes('senior')) return 'senior';
  return 'fe';
}

// ─── SUPERVISOR DISTRICT LOOKUP (BUG-C03 FIX) ────────────────────────────────
// New function: returns all districts supervised by a given employee ID.
// Queries Eng.List column 3 (supervisor ID column, 0-indexed = r[2]).
// Used to build Senior scope in loginUser() — replaces the broken
// division→district lookup that always returned empty for Seniors.
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
      // col 3 = supervisor ID = r[2] (0-indexed in array)
      const supId = String(r[2] || '').trim().toLowerCase();
      if (supId === id) {
        // col 9 = district = r[8] (0-indexed)
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

      // Masterlist values (used as fallback only — BUG-H06 FIX)
      const mlDivision = String(row[4] || '').trim();
      const mlDistrict = String(row[5] || '').trim();

      const roleTier = _classifyRole(role);
      const locData  = getLocationData(id);

      // BUG-H06 FIX: Prefer Eng.List division/district over Masterlist.
      // Masterlist can be stale or miscategorised (e.g. Arsaga: ML=Div3, EL=Div8).
      // Eng.List is the authoritative operational record.
      const division = locData.userDivisions.length > 0
        ? locData.userDivisions[0]
        : mlDivision;
      const district = locData.userDistricts.length > 0
        ? locData.userDistricts[0]
        : mlDistrict;

      // BUG-C03 FIX: Build Senior scope from supervisor relationship in Eng.List.
      // Previously built from userDivisions which was always empty for Seniors
      // (they appear in col 3 of Eng.List, not col 0).
      let seniorDistrictScope = [];
      if (roleTier === 'senior') {
        const supervisedDistricts = getDistrictsBySupervisor(id);
        if (supervisedDistricts.length > 0) {
          // Primary path: use all districts this Senior supervises
          seniorDistrictScope = supervisedDistricts;
        } else {
          // Fallback: try division→district map, then userDistricts
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

      return {
        ok: true, username: id, role, roleTier, name, firstLogin,
        division,
        district,
        userDivisions:  locData.userDivisions,
        userDistricts:  locData.userDistricts.length > 0
                          ? locData.userDistricts
                          : (mlDistrict ? [mlDistrict] : []),
        seniorDistrictScope,
        divDistrictMap: locData.divDistrictMap || {},
        area:   String(row[6] || '').trim(),
        branch: String(row[7] || '').trim()
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
  const PFX = {
    'Laptop': 'LTP', 'Laptop Adaptor': 'LAD', 'CPU': 'CPU', 'Monitor': 'MTR',
    'Printer': 'PTR', 'Scanner': 'SCN', 'Scansnap': 'SCN', 'Keyboard': 'KBD',
    'Mouse': 'MSE', 'UPS': 'UPS', 'Camera': 'CAM', 'Speaker': 'SPR',
    'External Drive': 'EXD'
  };
  const pre  = PFX[type] || 'AST';
  const yr   = new Date().getFullYear();
  const sh   = _entrySheet();
  const last = sh.getLastRow();
  let max = 0;
  if (last >= AE_DATA_START) {
    const barcodes = sh.getRange(AE_DATA_START, C.BARCODE, last - AE_DATA_START + 1, 1).getValues();
    barcodes.forEach(r => {
      const bc = String(r[0] || '');
      const parts = bc.split('-');
      if (parts.length >= 3 && parts[0] === pre) {
        const n = parseInt(parts[parts.length - 1]);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  let candidate = pre + '-' + yr + '-' + String(max + 1).padStart(3, '0');
  while (_findRow(sh, candidate) > 0) {
    max++;
    candidate = pre + '-' + yr + '-' + String(max + 1).padStart(3, '0');
  }
  return candidate;
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

        return {
          Barcode: get(C.BARCODE), Type: get(C.TYPE), Brand: get(C.BRAND),
          Serial: get(C.SERIAL), Specs: get(C.SPECS),
          Condition: get(C.CONDITION) || 'Good', Lifecycle: displayLC,
          AssetStatus: get(C.ASSET_STATUS) || 'Active',
          StatusLabel: get(C.STATUS_LABEL) || 'Unassigned',
          PurchDate: get(C.PURCH_DATE), WarrantyTerm: get(C.WARRANTY_TERM),
          WarrantyVal: get(C.WARRANTY_VAL), Remarks: get(C.REMARKS),
          EmpID: get(C.EMP_ID) || 'N/A', Staff: get(C.STAFF) || 'Unassigned',
          Designation: get(C.DESIGNATION), Division: div,
          District: dist, Area: get(C.AREA), Branch: get(C.BRANCH),
          EffDate: get(C.EFF_DATE), XferType: get(C.XFER_TYPE),
          ToStaff: get(C.TO_STAFF), ToEmpID: get(C.TO_EMPID),
          ToDiv: get(C.TO_DIV), ToBranch: get(C.TO_BRANCH), XferDate: get(C.XFER_DATE),
          BorName: get(C.BOR_NAME), BorEmpID: get(C.BOR_EMPID),
          BorDate: get(C.BOR_DATE), ExpReturn: get(C.EXP_RETURN),
          ActReturn: get(C.ACT_RETURN), BorRemarks: get(C.BOR_REMARKS),
          BorDesig: get(C.BOR_DESIG), BorDiv: get(C.BOR_DIV), BorBranch: get(C.BOR_BRANCH),
          BorDist: colCount >= 56 ? get(C.BOR_DIST) : '',  // safe read if col exists
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
      designation: get(C.DESIGNATION), division: _normDiv(get(C.DIVISION)),
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

      // BUG-L05 FIX: normalise on write so saved values are consistent
      const normDiv  = _normDiv(obj.division  || '');
      const normDist = _normDist(obj.district || '');

      const row = new Array(TOTAL_COLS).fill('');
      row[C.ENTRY_EMP_ID - 1] = obj.entryEmpId || '';
      row[C.ENTRY_NAME   - 1] = obj.entryName  || '';
      row[C.EMP_ID       - 1] = isSpare ? '' : (obj.accEmpId || '');
      row[C.STAFF        - 1] = isSpare ? '' : (obj.accName  || '');
      row[C.DESIGNATION  - 1] = isSpare ? '' : (obj.accRole  || '');
      row[C.DIVISION     - 1] = normDiv;          // BUG-L05 FIX
      row[C.DISTRICT     - 1] = normDist;         // BUG-L05 FIX
      row[C.AREA         - 1] = obj.area      || '';
      row[C.BRANCH       - 1] = obj.branch    || '';
      row[C.EFF_DATE     - 1] = obj.effDate   || '';
      row[C.BARCODE      - 1] = obj.barcode;
      row[C.TYPE         - 1] = obj.type      || '';
      row[C.BRAND        - 1] = obj.brand     || '';
      row[C.SERIAL       - 1] = obj.serial ? String(obj.serial) : '';
      row[C.SPECS        - 1] = obj.specs     || '';
      row[C.CONDITION    - 1] = obj.condition || 'New';
      row[C.LIFECYCLE    - 1] = sm.lc;
      row[C.ASSET_STATUS - 1] = sm.asSt;
      row[C.STATUS_LABEL - 1] = sm.stLbl;
      row[C.PURCH_DATE   - 1] = obj.purchDate || '';
      row[C.WARRANTY_TERM- 1] = obj.wTerm     || '';
      row[C.WARRANTY_VAL - 1] = obj.wValidity || '';
      row[C.REMARKS      - 1] = obj.remarks   || '';
      row[C.SUPPLIER     - 1] = obj.supplier  || '';
      row[C.LOCATION     - 1] = obj.location  || '';
      row[C.ENROLLED_BY  - 1] = obj.enrolledBy || obj.entryEmpId || '';
      row[C.CREATED_AT   - 1] = nowStr;
      row[C.LAST_UPDATED - 1] = nowStr;

      if (obj.serial) {
        const curLast = _entrySheet().getLastRow();
        const allData = curLast >= AE_DATA_START
          ? _entrySheet().getRange(AE_DATA_START, C.SERIAL, curLast - AE_DATA_START + 1, 1).getValues()
          : [];
        const dupRow = allData.findIndex(r => String(r[0]).trim() === String(obj.serial).trim());
        if (dupRow >= 0) {
          const existingBC = String(_entrySheet().getRange(dupRow + AE_DATA_START, C.BARCODE).getValue());
          return 'Error: Serial No. already registered under barcode: ' + existingBC;
        }
      }
      sh.appendRow(row);
      const newRowIdx = sh.getLastRow();
      sh.getRange(newRowIdx, C.SERIAL).setNumberFormat('@STRING@');
      if (obj.serial) sh.getRange(newRowIdx, C.SERIAL).setValue(String(obj.serial));
      _log('CREATE', obj.barcode, obj.type + ' | ' + obj.brand + ' | ' + statusChoice, obj.entryEmpId || '');
      return 'Asset created: ' + obj.barcode;
    }

    const lc    = obj.lifecycle || 'Allocated';
    const asSt  = lc === 'Transfer' ? 'Transfer' : lc === 'Dispose' ? 'Disposal' : 'Active';
    const staff = (obj.staff || '').trim();
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
      [C.DESIGNATION,  obj.designation || ''], [C.DIVISION, _normDiv(obj.division   || '')],  // BUG-L05 FIX
      [C.DISTRICT,     _normDist(obj.district || '')],                                          // BUG-L05 FIX
      [C.AREA,         obj.area || ''],
      [C.BRANCH,       obj.branch      || ''], [C.EFF_DATE, obj.effDate || '']
    ]);
    _log('ASSIGN', obj.barcode, staff + ' | ' + lc, obj.employeeId || '');
    return 'Asset assigned successfully';
  } catch (e) { return 'Error: ' + e.message; }
}

function deleteAssets(barcodes) {
  try {
    const sh = _entrySheet();
    const blocked = [], toDelete = [];
    barcodes.forEach(bc => {
      const r = _findRow(sh, bc);
      if (r > 0) {
        const lc = String(sh.getRange(r, C.LIFECYCLE).getValue() || '').toLowerCase();
        if (lc === 'borrow' || lc === 'transfer') { blocked.push(bc + ' (' + lc + ')'); }
        else { toDelete.push({ bc, r }); }
      }
    });
    if (blocked.length && !toDelete.length)
      return 'Error: Cannot delete — active lifecycle: ' + blocked.join(', ');
    toDelete.sort((a, b) => b.r - a.r).forEach(({ bc, r }) => {
      sh.deleteRow(r); _log('DELETE', bc, '', '');
    });
    let msg = 'Deleted ' + toDelete.length + ' asset(s)';
    if (blocked.length) msg += '. Skipped ' + blocked.length + ' with active borrow/transfer: ' + blocked.join(', ');
    return msg;
  } catch (e) { return 'Error: ' + e.message; }
}

// ─── ALLOCATE ASSET ──────────────────────────────────────────────────────────
function allocateAsset(obj) {
  try {
    if (!obj.barcode)   return 'Error: Barcode is required.';
    if (!obj.empId)     return 'Error: Employee ID is required.';
    if (!obj.staffName) return 'Error: Staff name is required.';
    const sh = _entrySheet();
    const rowIdx = _findRow(sh, obj.barcode);
    if (rowIdx < 1) return 'Error: Asset not found: ' + obj.barcode;
    const currentLC = String(sh.getRange(rowIdx, C.LIFECYCLE).getValue() || '').toLowerCase();
    if (currentLC === 'borrow')   return 'Error: Asset is currently borrowed. Return it first.';
    if (currentLC === 'dispose')  return 'Error: Disposed assets cannot be allocated.';
    if (currentLC === 'transfer') return 'Error: Asset is in an active transfer.';
    const nowStr = new Date().toLocaleString('en-PH');
    const rowData = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    const get = c => String(rowData[c - 1] || '');

    // BUG-L05 FIX: normalise district/division on write
    const normDiv  = _normDiv(obj.division  || '');
    const normDist = _normDist(obj.district || '');

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,   'Allocated'], [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL,'Assigned'],  [C.EMP_ID,       obj.empId       || ''],
      [C.STAFF,       obj.staffName || ''],   [C.DESIGNATION,  obj.designation || ''],
      [C.DIVISION,    normDiv],               [C.DISTRICT,     normDist],      // BUG-L05 FIX
      [C.AREA,        obj.area      || ''],   [C.BRANCH,       obj.branch      || ''],
      [C.EFF_DATE,    obj.effDate   || nowStr],[C.REMARKS,      obj.remarks     || '']
    ]);
    _allocLogSheet().appendRow([
      obj.barcode, obj.type || get(C.TYPE), obj.brand || get(C.BRAND),
      obj.serial || get(C.SERIAL), obj.empId, obj.staffName,
      obj.designation || '', normDiv, normDist,
      obj.area || '', obj.branch || '', obj.effDate || nowStr,
      obj.condition || get(C.CONDITION) || 'Good',
      obj.remarks || '', nowStr, obj.allocatedBy || ''
    ]);
    _log('ALLOCATE', obj.barcode, obj.staffName + ' | ' + (obj.branch || normDiv || ''), obj.allocatedBy || obj.empId || '');
    return 'Asset allocated to ' + obj.staffName;
  } catch (e) { return 'Error: ' + e.message; }
}

// ─── DEALLOCATE ASSET ────────────────────────────────────────────────────────
function deallocateAsset(barcode, remarks) {
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
      [C.DIVISION, ''], [C.DISTRICT, ''], [C.AREA, ''], [C.BRANCH, '']
    ];
    if (remarks) updates.push([C.REMARKS, remarks]);
    _setRow(sh, rowIdx, updates);
    _log('DEALLOCATE', barcode, `From: ${prevStaff} → Spare Pool. ${remarks||''}`, '');
    return 'Asset returned to spare pool';
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
      t.fromStaff, t.fromEmpId, t.fromDesig, t.fromDiv, t.fromDist, t.fromArea, t.fromBranch, t.fromRemarks,
      t.toStaff, t.toEmpId, t.toDesig, t.toDiv, t.toDist, t.toArea, t.toBranch, t.toRemarks,
      t.effDate, t.status || 'Completed', nowStr
    ]);

    // BUG-L05 FIX: normalise on write
    const normToDiv  = _normDiv(t.toDiv   || '');
    const normToDist = _normDist(t.toDist || '');

    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Allocated'], [C.ASSET_STATUS, 'Active'],
      [C.STATUS_LABEL, 'Assigned'],  [C.EMP_ID,       t.toEmpId   || ''],
      [C.STAFF,        t.toStaff   || ''],  [C.DESIGNATION,  t.toDesig   || ''],
      [C.DIVISION,     normToDiv],          [C.DISTRICT,     normToDist],          // BUG-L05 FIX
      [C.AREA,         t.toArea    || ''],  [C.BRANCH,       t.toBranch  || ''],
      [C.EFF_DATE,     t.effDate],          [C.XFER_TYPE,    t.transferType || 'Permanent'],
      [C.FR_STAFF,     t.fromStaff   || ''],[C.FR_EMPID,     t.fromEmpId   || ''],
      [C.FR_DESIG,     t.fromDesig   || ''],[C.FR_DIV,       t.fromDiv     || ''],
      [C.FR_DIST,      t.fromDist    || ''],[C.FR_AREA,      t.fromArea    || ''],
      [C.FR_BRANCH,    t.fromBranch  || ''],[C.FR_REMARKS,   t.fromRemarks || ''],
      [C.TO_STAFF,     t.toStaff   || ''],  [C.TO_EMPID,     t.toEmpId   || ''],
      [C.TO_DESIG,     t.toDesig   || ''],  [C.TO_DIV,       normToDiv],
      [C.TO_DIST,      normToDist],         [C.TO_AREA,      t.toArea    || ''],
      [C.TO_BRANCH,    t.toBranch  || ''],  [C.TO_REMARKS,   t.toRemarks || ''],
      [C.XFER_DATE,    t.effDate]
    ]);
    _log('TRANSFER', t.barcode, (t.fromStaff || '—') + ' → ' + t.toStaff, t.fromEmpId || '');
    return 'Transfer saved';
  } catch (e) { return 'Error: ' + e.message; }
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
      b.barcode, b.borrowerName, b.empId, b.designation,
      b.division, b.district||'', b.branch, b.borrowDate, b.expectedReturn,
      '', 'Borrow', b.remarks, nowStr
    ]);
    _setRow(sh, rowIdx, [
      [C.LIFECYCLE,    'Borrow'],
      [C.ASSET_STATUS, isBorrowItem ? 'BorrowItem' : 'Active'],
      [C.STATUS_LABEL, 'Assigned'],
      [C.BOR_NAME,     b.borrowerName   || ''],
      [C.BOR_EMPID,    b.empId          || ''],
      [C.BOR_DESIG,    b.designation    || ''],
      [C.BOR_DIV,      b.division       || ''],
      [C.BOR_DIST,     b.district       || ''],
      [C.BOR_BRANCH,   b.branch         || ''],
      [C.BOR_DATE,     b.borrowDate     || ''],
      [C.EXP_RETURN,   b.expectedReturn || ''],
      [C.ACT_RETURN,   ''],
      [C.BOR_REMARKS,  b.remarks        || '']
    ]);
    _log('BORROW', b.barcode, b.borrowerName + ' | due: ' + b.expectedReturn, b.empId || '');
    return 'Borrow saved';
  } catch (e) { return 'Error: ' + e.message; }
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

// BUG-M04 FIX: returnAsset() now clears ALL borrow fields including BOR_REMARKS.
// Previously BOR_REMARKS was left populated after a return, polluting the asset record.
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
        [C.LIFECYCLE,    restoredLC],  [C.STATUS_LABEL, restoredLbl],
        [C.ASSET_STATUS, restoredAS],  [C.ACT_RETURN,   retDate],
        // BUG-M04 FIX: clear ALL borrow fields including BOR_REMARKS
        [C.BOR_NAME,''], [C.BOR_EMPID,''], [C.BOR_DESIG,''],
        [C.BOR_DIV,''],  [C.BOR_DIST,''],  [C.BOR_BRANCH,''],
        [C.BOR_DATE,''], [C.EXP_RETURN,''], [C.BOR_REMARKS,'']
      ]);
    }
    _log('RETURN', barcode, retDate, '');
    return 'Asset returned';
  } catch (e) { return 'Error: ' + e.message; }
}

// ─── DISPOSALS ───────────────────────────────────────────────────────────────
function saveDisposal(d) {
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
    dSh.appendRow([d.barcode, d.reason, d.disposedBy, d.disposeDate, d.remarks, nowStr]);
    if (rowIdx > 0) {
      const cur = String(sh.getRange(rowIdx, C.REMARKS).getValue() || '');
      _setRow(sh, rowIdx, [
        [C.LIFECYCLE,    'Dispose'], [C.ASSET_STATUS, 'Disposal'],
        [C.STATUS_LABEL, 'Disposed'],
        [C.REMARKS, (cur ? cur + ' | ' : '') + 'DISPOSAL: ' + d.reason + ' by ' + d.disposedBy]
      ]);
    }
    _log('DISPOSE', d.barcode, d.reason + ' | ' + d.disposedBy, d.disposedBy || '');
    return 'Disposal recorded';
  } catch (e) { return 'Error: ' + e.message; }
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
      // Return normalised division/district for consistent form auto-fill
      return {
        ok:       true,
        empId:    String(row[0]  || '').trim(),
        name:     String(row[2]  || '').trim(),
        division: _normDiv(String(row[4]  || '').trim()),    // BUG-L05 FIX: normalise on return
        district: _normDist(String(row[5] || '').trim()),    // BUG-L05 FIX
        area:     String(row[6]  || '').trim(),
        branch:   String(row[7]  || '').trim(),
        position: String(row[11] || '').trim()
      };
    }
    return { ok: false, error: 'Employee ID not found: ' + empId };
  } catch (e) { return { ok: false, error: e.message }; }
}

// ─── DROPDOWN DATA ────────────────────────────────────────────────────────────
// BUG-C04 FIX: Added 'SCANNER SCANSNAP' to CAT_MAP and added post-processing
// to clone Scanner data to Scansnap when the merged header is detected.
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
      // BUG-C04 FIX: handle merged 'SCANNER SCANSNAP' header — map to Scanner;
      // Scansnap is cloned from Scanner data after the loop below.
      'SCANNER SCANSNAP': 'Scanner',
      'UPS': 'UPS', 'EXTERNAL DRIVE': 'External Drive',
      'CAMERA': 'Camera', 'SPEAKER': 'Speaker'
    };
    const result = { categories: [], brands: {}, models: {}, suppliers: [], laptopSpecs: [], laptopSpecValues: {} };
    let catPositions = [], supplierCol = -1, laptopSpecCol = -1, fieldSectionCol = -1;
    // Track whether we used a merged SCANNER SCANSNAP header
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
      // BUG-C04 FIX: detect merged header
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

    // BUG-C04 FIX: clone Scanner data to Scansnap when merged header was used.
    // Only clones if Scansnap doesn't already have its own column in the sheet.
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

    // ── 1. Build Division→District map from Drop down sheet ──
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

    // ── 2. Load user's assigned divisions & districts from Eng. List ──
    const engSh = ss.getSheetByName('Eng. List') || ss.getSheetByName('Eng List');
    if (engSh && engSh.getLastRow() > 1) {
      const lastRow = engSh.getLastRow();
      const data = engSh.getRange(2, 1, lastRow - 1, 10).getValues();
      const divSet = new Set(), distSet = new Set();

      if (empId) {
        data.forEach(r => {
          const rowEmpId = String(r[0] || '').trim();
          if (rowEmpId.toLowerCase() === String(empId).trim().toLowerCase()) {
            const div  = String(r[7] || '').trim();
            const dist = String(r[8] || '').trim();
            if (div)  divSet.add(div);
            if (dist) distSet.add(dist);
          }
        });
      }

      if (divSet.size === 0 && distSet.size === 0) {
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
    }

    return result;
  } catch (e) {
    return { divDistrictMap: {}, userDivisions: [], userDistricts: [], userDivisionDistricts: {}, error: e.message };
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
    const data   = sh.getRange(AE_DATA_START, 1, count, TOTAL_COLS).getValues();
    const nowStr = new Date().toLocaleString('en-PH');
    const id     = String(empId).trim().toLowerCase();
    let updated  = 0, skipped = 0;

    for (let i = 0; i < data.length; i++) {
      const row      = data[i];
      const rowEmpId = String(row[C.EMP_ID - 1] || '').trim().toLowerCase();
      if (rowEmpId !== id) continue;

      const rowIdx = i + AE_DATA_START;
      const lc     = String(row[C.LIFECYCLE - 1] || '').toLowerCase();
      if (lc === 'borrow')                   { skipped++; continue; }
      if (lc === 'dispose' || lc === 'disposal') continue;

      const normDiv  = _normDiv(newDiv   || '');
      const normDist = _normDist(newDist || '');

      if (assetAction === 'spare') {
        [[C.LIFECYCLE,'Active'],[C.ASSET_STATUS,'Active'],[C.STATUS_LABEL,'Unassigned'],
         [C.EMP_ID,''],[C.STAFF,''],[C.DESIGNATION,''],[C.EFF_DATE,''],
         [C.DIVISION,normDiv],[C.DISTRICT,normDist],
         [C.AREA,newArea||''],[C.BRANCH,newBranch||''],[C.LAST_UPDATED,nowStr]]
        .forEach(u => sh.getRange(rowIdx, u[0]).setValue(u[1]));
      } else {
        [[C.DIVISION,normDiv],[C.DISTRICT,normDist],
         [C.AREA,newArea||''],[C.BRANCH,newBranch||''],[C.LAST_UPDATED,nowStr]]
        .forEach(u => sh.getRange(rowIdx, u[0]).setValue(u[1]));
      }
      updated++;
    }

    _log('MOVE_STAFF', empId,
      `Action:${assetAction} → ${newDiv}/${newDist}/${newBranch} | ${updated} updated, ${skipped} skipped`,
      empId);

    let msg = `Staff movement recorded. ${updated} asset(s) ${assetAction === 'spare' ? 'returned to spare at new location' : 'moved to new location'}.`;
    if (skipped) msg += ` (${skipped} skipped — on active borrow)`;
    return msg;
  } catch (e) { return 'Error: ' + e.message; }
}