# IICTD Asset Management System

A web-based asset management application built on Google Apps Script (GAS) for LifeBank Foundation's Infrastructure & Information and Communications Technology Department (IICTD). The system tracks IT equipment across Head Office and Field Office locations, managing the full lifecycle of each asset from enrollment to disposal.

---

## Features

### Asset Lifecycle Management
- **Spare Pool** — Enroll new assets into an unallocated inventory pool
- **Allocation** — Assign assets to staff members with full accountability tracking
- **Transfers** — Move assets between staff (single or bulk), with automatic form generation
- **Borrow Management** — Track temporary loans with due dates, overdue alerts, and a dedicated Borrow Pool
- **Disposal** — Record decommissioned assets with reasons and audit trail

### Accountability Form Workflow
- Auto-generated accountability forms on allocation, transfer, and enrollment
- Draft → Submit → Supervisor Review → Confirm/Reject workflow
- Rate-limited resubmission to prevent spam
- PDF-printable Equipment Accountability Forms with signatory blocks
- Form snapshots archived on confirmation

### Organizational Hierarchy
- **Field:** Division → District → Area → Branch → Staff
- **Head Office:** Department → Base Office → Staff
- Scope-based data visibility — each user sees only their assigned districts/divisions
- Role tiers: Administrator (HO), Senior Field Engineer, Field Engineer

### Views & Reports
- **Dashboard** — Summary stats, recent activity, and asset-type distribution chart
- **All Assets** — Searchable, filterable, paginated master list
- **Staff Assets** — Hierarchical tree view grouped by location and staff member
- **Equipment Record** — Location-based tree view for physical inventory tracking
- **Transfer Records** — Full transfer history with status tracking
- **Activity Log** — Auditable action log for all system events

### Access Control
- Role-based permissions (add, edit, delete, allocate, transfer, borrow, dispose)
- View-only mode for read-only accounts
- Scope isolation — FE users see their district only; seniors see their division; admins see all

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Google Apps Script (V8 runtime) |
| Frontend | HTML + CSS + Vanilla JS (no frameworks) |
| Database | Google Sheets (multi-sheet workbook) |
| Deployment | Google Apps Script Web App |
| Barcode | JsBarcode (CODE128) |
| Tooling | [clasp](https://github.com/google/clasp) (local development) |

---

## Project Structure

```
src/
├── Code.js                    # Backend — all server-side logic
├── Index.html                 # Main HTML shell / entry point
├── Stylesheet.html            # Global CSS
├── Auth.html                  # Login and password change screens
├── Modals.html                # All modal dialogs (HTML)
├── Pages_Overview.html        # Dashboard and All Assets page markup
├── Pages_AssetPools.html      # Spare, Allocated, Borrow, Disposal pages
├── Pages_Records.html         # Transfers and Activity Log pages
├── Pages_ForApproval.html     # Accountability form approval page
├── JS_Core.html               # State, helpers, auth, navigation, scope logic
├── JS_Dashboard.html          # Dashboard render logic
├── JS_AllAssets.html          # All Assets table and bulk delete
├── JS_Spare.html              # Spare Pool table
├── JS_Allocated.html          # Allocated Assets table
├── JS_Borrow.html             # Borrow management (active, pool, history)
├── JS_Disposal.html           # Disposal Records table
├── JS_Records.html            # Transfers and Activity Log tables
├── JS_StaffAssets.html        # Staff Assets hierarchical tree
├── JS_EquipmentRecord.html    # Equipment Record location tree
├── JS_AccountabilityForm.html # PDF accountability form generator
├── JS_ForApproval.html        # Approval workflow UI logic
├── JS_Modals.html             # All modal open/submit/action logic
└── Logo.html                  # Base64-encoded logo asset
appsscript.json                # GAS project manifest
.clasp.json                    # clasp configuration
```

---

## Google Sheets Structure

The system reads from and writes to a single Google Sheets workbook. The Sheet ID is configured in `Code.js`.

| Sheet | Purpose |
|---|---|
| `Asset Entry` | Master asset registry (all assets, 37+ columns) |
| `Users` | User accounts, roles, and hashed passwords |
| `Org Structure` | Field engineer → district → division mapping |
| `Masterlist` | Employee directory for auto-fill lookups |
| `Transfers` | Transfer event log |
| `Borrows` | Borrow event log |
| `Disposals` | Disposal event log |
| `ActivityLog` | Full audit log of all actions |
| `Allocated` | Allocation event log |
| `Spare` | Spare pool enrollment log |
| `AccountabilityForms` | Form workflow state |
| `FormSnapshots` | Archived confirmed forms |
| `FormRateLimits` | Resubmission rate limiting |
| `ApprovalConfig` | Signatory configuration per context |
| `Drop down` | Category, brand, model, and supplier reference data |

---

## Setup & Deployment

### Prerequisites
- A Google account with access to Google Apps Script
- [Node.js](https://nodejs.org/) and [clasp](https://github.com/google/clasp) for local development

### Initial Setup

1. **Clone the repository**
   ```bash
   git clone <repo-url>
   cd asset-management-appscript
   ```

2. **Install clasp**
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

3. **Configure the Sheet ID**  
   Open `src/Code.js` and update the constant at the top:
   ```js
   const SHEET_ID = 'your-google-sheet-id-here';
   ```

4. **Push to Google Apps Script**
   ```bash
   clasp push
   ```

5. **Deploy as a Web App**  
   In the Apps Script editor: **Deploy → New deployment → Web app**
   - Execute as: `User deploying the app`
   - Access: `Anyone`

6. **Set up the Google Sheet**  
   The backend will auto-create all required sheets on first use. Populate the `Users` sheet with at least one admin account (default password: `1234`, changed on first login).

### Local Development

```bash
# Pull latest from GAS
clasp pull

# Push local changes to GAS
clasp push

# Open GAS editor in browser
clasp open
```

---

## User Roles

| Role Tier | Key Capabilities |
|---|---|
| **Administrator (HO)** | Full access — all assets, all scopes, user management, delete |
| **Senior Field Engineer** | Manage assets in their division; review and confirm accountability forms |
| **Field Engineer** | Manage assets in their assigned districts; submit forms for review |
| **View Only** | Read-only access — no create, edit, or action buttons |

Permissions are configured per-user via the `Remarks` column in the `Users` sheet (e.g., `"no delete"`, `"view only"`, `"full access"`).

---

## Barcode Format

Auto-generated barcodes follow the pattern:

```
{TYPE_PREFIX}-{YEAR}-{SEQUENCE}
```

Examples: `LTP-2025-001`, `MTR-2025-042`, `KBD-2024-015`

| Type | Prefix |
|---|---|
| Laptop | `LTP` |
| Monitor | `MTR` |
| CPU | `CPU` |
| Keyboard | `KBD` |
| Mouse | `MSE` |
| Printer | `PTR` |
| Scanner / Scansnap | `SCN` |
| UPS | `UPS` |
| Camera | `CAM` |
| Speaker | `SPR` |
| External Drive | `EXD` |
| Laptop Adaptor | `LAD` |

---

## License

Internal use — LifeBank Foundation IICTD. All rights reserved.