# QAMP Architecture Overview

## About This Document
This document describes the overall structure of the QAMP (Quaker Arts Music Program) codebase — the scripts, workbooks, Google Drive resources, and conventions that make up the system. For function-level detail, see the `*-Functions.md` files at the repo root. For sheet-level column structure, see `Sheet-Structure.md`.

---

## Repository Structure

```
qamp-project/
├── ARCHITECTURE.md                   ← this file
├── Billing-Functions.md
├── Contacts-Functions.md
├── Payments-Functions.md
├── Responses-Functions.md
├── Sheet-Structure.md
├── Teacher-Invoice-Functions.md
├── Teacher-Responses-Functions.md
├── Utility-Functions.md
├── billing/
│   ├── Code.js
│   ├── ReRegistration.html
│   ├── .clasp.json
│   └── appsscript.json
├── contacts/
│   ├── Code.js
│   ├── .clasp.json
│   └── appsscript.json
├── payments/
│   ├── Code.js
│   ├── .clasp.json
│   └── appsscript.json
├── responses/
│   ├── Code.js
│   ├── .clasp.json
│   └── appsscript.json
├── teacher-invoice/
│   ├── Code.js
│   ├── .clasp.json
│   └── appsscript.json
├── teacher-responses/
│   ├── Code.js
│   ├── .clasp.json
│   └── appsscript.json
└── utility-library/
    ├── Code.js
    ├── Config.js
    ├── .clasp.json
    ├── .claspignore
    └── appsscript.json
```

**Conventions:**
- Each script folder contains exactly one `Code.js`, one `.clasp.json`, and one `appsscript.json`.
- `utility-library` additionally contains `Config.js` (script properties / environment config) and `.claspignore`.
- Documentation `.md` files live at the repo root, not inside subfolders.
- `.md` files follow the naming convention `[ScriptName]-Functions.md`.

---

## Scripts

### Utility Library
| | |
|---|---|
| **Folder** | `utility-library/` |
| **Documentation** | `Utility-Functions.md` |
| **Attached to** | Standalone Google Apps Script library |
| **Library prefix** | `UtilityScriptLibrary` |
| **Reads workbooks** | All (via `getWorkbook()` / `getSheet()`) |
| **Writes workbooks** | All (via `getWorkbook()` / `getSheet()`) |

The foundation of the entire system. All other scripts include Utility as a library and access it via the `UtilityScriptLibrary` prefix. Contains shared configuration (`getConfig()`, `EnvironmentManager`), the `SHEET_MAP`, `TEMPLATE_MAP`, all shared helper and data functions, and the `debugLog` function that writes to the Debug sheet in the Form Responses workbook. Config stores environment-specific workbook IDs and folder IDs in Script Properties, supporting `test` and `prod` environments.

---

### Billing
| | |
|---|---|
| **Folder** | `billing/` |
| **Documentation** | `Billing-Functions.md` |
| **Attached to** | Billing workbook |
| **Reads workbooks** | Contacts, Form Responses, Payments |
| **Writes workbooks** | Billing |

The most complex script. Manages the full billing lifecycle: semester setup, billing cycle creation, document generation (invoices, welcome letters, agreement forms, media releases), payment reconciliation, and re-registration integration. Also handles teacher-student billing updates and cumulative tracking formulas. Serves `ReRegistration.html` as a web app (deployed separately via Apps Script); the web app URL is stored in Script Properties as `webAppUrl`.

---

### Contacts
| | |
|---|---|
| **Folder** | `contacts/` |
| **Documentation** | `Contacts-Functions.md` |
| **Attached to** | Contacts workbook |
| **Reads workbooks** | *(none)* |
| **Writes workbooks** | Contacts |

Manages the Contacts workbook (Students, Parents, Teachers and Admin, Instrument List sheets). Handles cascading status changes when a teacher is marked former.

---

### Responses
| | |
|---|---|
| **Folder** | `responses/` |
| **Documentation** | `Responses-Functions.md` |
| **Attached to** | Form Responses workbook |
| **Reads workbooks** | Contacts |
| **Writes workbooks** | Form Responses, Teacher Rosters (individual Drive files) |

Processes student registration form submissions. Handles teacher assignments, roster creation and management, semester registration data, and the FieldMap that maps form headers to internal field names. Creates and manages per-teacher roster spreadsheets stored in Google Drive. Sends confirmation emails to families on form submission.

---

### Teacher Invoice
| | |
|---|---|
| **Folder** | `teacher-invoice/` |
| **Documentation** | `Teacher-Invoice-Functions.md` |
| **Attached to** | Teacher Invoices workbook |
| **Reads workbooks** | Contacts, Teacher Rosters (individual Drive files) |
| **Writes workbooks** | Teacher Invoices |

Generates monthly teacher invoices based on attendance logs in individual teacher roster files. Handles invoice document generation, attendance verification, and admin reports.

---

### Teacher Responses
| | |
|---|---|
| **Folder** | `teacher-responses/` |
| **Documentation** | `Teacher-Responses-Functions.md` |
| **Attached to** | Teacher Interest workbook |
| **Reads workbooks** | Contacts |
| **Writes workbooks** | Teacher Interest, Contacts |

Processes new and returning teacher interest form responses. Handles teacher tracking and batch processing of submissions. Updates teacher status, email, and creates partial records in Contacts for returning teachers not yet fully on file.

---

### Payments
| | |
|---|---|
| **Folder** | `payments/` |
| **Documentation** | `Payments-Functions.md` |
| **Attached to** | Payment Ledger workbook |
| **Reads workbooks** | Payment Ledger |
| **Writes workbooks** | Payment Ledger, Google Drive (PDF receipts folder) |

Handles payment receipt generation. Generates PDF receipts when a checkbox is checked in the Payment Ledger, triggered via an installable `onEdit` trigger running under script-owner credentials.

---

## Google Workbooks

| Workbook | Primary Script | Also Read By |
|---|---|---|
| Contacts | Contacts | Billing, Teacher Invoice, Teacher Responses (read + write) |
| Form Responses | Responses | Billing |
| Billing | Billing | *(none)* |
| Teacher Invoices | Teacher Invoice | *(none)* |
| Teacher Interest | Teacher Responses | *(none)* |
| Payment Ledger | Payments | Billing |

---

## Google Drive Resources

All folder IDs are stored in Script Properties and accessed via `getConfig()`. The following folders are referenced:

| Key | Purpose |
|---|---|
| `rosterFolderId` | Stores per-teacher roster spreadsheets |
| `templateFolderId` | Stores document templates (letters, invoices, agreements, media releases) |
| `generatedDocumentsFolderId` | Output folder for generated registration packet documents |
| `receiptsFolderId` | Output folder for payment receipt PDFs |

---

## Document Templates

Managed via `TEMPLATE_MAP` in Utility. Templates live in `templateFolderId` in Google Drive.

| Category | Templates |
|---|---|
| Welcome letters (print) | NewFamilyLetter-print, NewAdultLetter-print, ReturningFamilyLetter-print, ReturningAdultLetter-print, SecondInvoiceLetter-print, RevisedInvoiceLetter-print, MissingDocumentLetter-print |
| Welcome letters (email) | NewFamilyLetter-email, NewAdultLetter-email, ReturningFamilyLetter-email, ReturningAdultLetter-email, SecondInvoiceLetter-email, RevisedInvoiceLetter-email, MissingDocumentLetter-email |
| Legal / Business | Invoice, Agreement, MediaConsentChild, MediaConsentAdult, TeacherInvoiceTemplate |
| Receipts | Receipt |

---

## Shared Conventions

### Environment Management
- Two environments: `test` and `prod`, managed by `EnvironmentManager` in Utility.
- All workbook IDs and folder IDs stored in Script Properties, never hardcoded.
- Default environment is `test`; must be explicitly set to `prod` to operate on production data.

### Sheet Access
- All sheet access goes through `UtilityScriptLibrary.getSheet(sheetKey)` using keys defined in `SHEET_MAP`.
- Column headers are always referenced dynamically via `getHeaderMap()` and `normalizeHeader()`. Column indexes are never hardcoded.

### Logging
- All logging uses `UtilityScriptLibrary.debugLog(functionName, eventType, message, data, errorDetails)`.
- Logs write to the Debug sheet in the Form Responses workbook.
- Event types: `INFO`, `WARNING`, `ERROR`, `SUCCESS`, `DEBUG`.

### Error Handling
- Wrapped operations use `UtilityScriptLibrary.executeWithErrorHandling()`.
- UI-facing functions display alerts on error; background functions log and rethrow.

### Modular Design
- Code is intentionally modular: small, single-purpose functions that can be reused across callers.
- Any logic used by more than one script belongs in Utility.
- Each script's `*-Functions.md` documents every function with its category, dependencies, and callers.

### Clasp Workflow
- See `Git-AppsScript-Workflow.md` for the full push/pull workflow between local files and Google Apps Script.

---

## Out-of-Repo Components

| Component | Location | Notes |
|---|---|---|
| Teacher Roster spreadsheets | Google Drive (`rosterFolderId`) | Created dynamically by Responses script; one per teacher per year |
| Document templates | Google Drive (`templateFolderId`) | Static Google Docs used as merge templates |
| Generated documents | Google Drive (`generatedDocumentsFolderId`) | Output of Billing document generation |
| PDF receipts | Google Drive (`receiptsFolderId`) | Output of Payments receipt generation |