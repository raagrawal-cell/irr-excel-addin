# IVP IRR Analytics Engine — Excel Web Add-in
## Phase 1: Scaffold + Engine Foundation

---

## What's in this package

```
irr-excel-addin/
├── manifest.xml              ← Add-in registry (sideload this to install)
├── taskpane.html             ← Navigation hub (opens all 6 modules as dialogs)
└── engine/
    ├── xirr.js               ← Newton-Raphson + Brent XIRR (pure JS, no deps)
    ├── dataReader.js         ← Excel.run() data layer (replaces SpreadsheetApp)
    └── bridge.js             ← google.script.run compatibility shim
```

**Not yet included (Phase 2 onwards):**
```
modules/                      ← Your existing HTML files (ported in Phase 2)
├── analyticsDashboard.html
├── reverseIRR.html
├── irrSimulator.html
├── managementCommentary.html
├── configLibrary.html
└── attributionSidebar.html
engine/
├── simulation.js             ← Phase 3: SC-01–SC-10 scenario templates
└── backsolve.js              ← Phase 3: Binary search reverse IRR solver
```

---

## Setup (one-time, browser-only — no CLI needed)

### Step 1: Create the GitHub repo

1. Go to github.com → New repository → name it `irr-excel-addin`
2. Set it to **Public** (required for GitHub Pages free tier)
3. Upload all files from this package using drag-and-drop or the file editor

### Step 2: Enable GitHub Pages

1. Repo → Settings → Pages
2. Source: **Deploy from a branch** → Branch: `main` → folder: `/ (root)`
3. Click Save. Your URL will be: `https://YOUR-USERNAME.github.io/irr-excel-addin`
4. Wait 2–3 minutes for the first deployment

### Step 3: Edit manifest.xml

Replace all 6 occurrences of `YOUR-GITHUB-PAGES-URL` with your actual URL:
```
https://YOUR-USERNAME.github.io/irr-excel-addin
```

Also generate a GUID (https://guidgenerator.com/) and replace `YOUR-GUID-HERE`.

Upload the updated `manifest.xml` back to GitHub.

### Step 4: Test sideloading

**Excel Desktop (Windows/Mac):**
1. Excel → Insert → Add-ins → My Add-ins → Upload My Add-in
2. Browse to your local `manifest.xml` → Upload

**Excel Online (office.live.com):**
1. Insert → Office Add-ins → Upload My Add-in
2. Select `manifest.xml` → Upload

You should see an "IVP IRR Analytics Engine" button appear in the Home tab ribbon.

---

## Workbook requirements

The workbook must have these sheet tabs (same names as in Google Sheets):

| Sheet | Purpose | Read/Write |
|-------|---------|------------|
| `Master_Data_2_CF` | Cashflow source | Read |
| `Master_Data_2_valuation` | NAV/valuation source | Read |
| `Config_Library` | Named configurations | Read/Write |
| `Sim_Config` | API key + model | Read |
| `Simulation_Log` | Simulation history | Write |
| `Mapping_Table` | Type→bucket map | Read |

**To migrate from Google Sheets:**
1. Google Sheets → File → Download → Microsoft Excel (.xlsx)
2. Open in Excel — all sheet names are preserved exactly

---

## What works in Phase 1

✅ Add-in installs and opens task pane  
✅ Reads all workbook sheets via `excelReadTab()`  
✅ Reads `Config_Library` → shows saved configs in dropdown  
✅ Selects active config → calls `buildBundle(cfg)`  
✅ Normalises CF rows with date serial conversion  
✅ Normalises Val rows (NAV absolute-valued)  
✅ Builds full hierarchy: `hierarchyValues`, `cascading`  
✅ Computes XIRR per deal (Newton-Raphson + Brent)  
✅ Fund-level XIRR from all transactions  
✅ Shows Fund IRR, MOIC, last quarter in KPI strip  
✅ All module buttons visible (disabled until bundle loads)  
✅ Config Library button always enabled (for initial setup)  
✅ `getAnalyticsAllTransactions()` — entity cashflow detail  
✅ `getAnalyticsSubPeriodTrend()` — sub-period IRR series  
✅ `saveSimulationToLog()` — append to Simulation_Log  
✅ `getSimConfig()` — reads Anthropic API key from Sim_Config  
✅ `google.script.run` shim via bridge.js  

---

## Phase 2: Port the HTML modules (next step)

For each module, these 3 edits are needed:

### Change 1: Replace template injection

**Before (Google Sheets):**
```html
<script>
var BUNDLE     = <?!= bundle ?>;
var SIM_CONFIG = <?!= simConfig ?>;
var MAX_LEVEL  = <?= maxLevel ?>;
var PATH_SEP   = '<?= pathSep ?>';
</script>
```

**After (Excel Add-in):**
```html
<script type="module">
import './engine/bridge.js';   // ← add this ONE line at module top
// Then in your init function:
async function init() {
  await window.IVP_INIT();              // wait for bundle
  var BUNDLE     = window.IVP_BUNDLE;
  var SIM_CONFIG = window.IVP_SIM_CONFIG;
  var MAX_LEVEL  = BUNDLE.maxLevel;
  var PATH_SEP   = BUNDLE.pathSep;
  // rest of init unchanged
}
document.addEventListener('DOMContentLoaded', init);
</script>
```

### Change 2: Remove google.script.run calls (they auto-work via shim)

The `bridge.js` shim creates `window.google.script.run` automatically.
**All existing `google.script.run.X()` calls work without any code changes.**

### Change 3: Replace google.script.host.close()

```javascript
// Before:
google.script.host.close();

// After (bridge.js handles this, but if you need to explicitly):
Office.context.ui.messageParent(JSON.stringify({type: 'close'}));
```

### Module porting order (recommended)

1. `configLibrary.html` — no bundle needed, proves sheet read/write works
2. `analyticsDashboard.html` — validates bundle shape + XIRR display
3. `reverseIRR.html` — requires backsolve.js (Phase 3)
4. `irrSimulator.html` — Claude API works immediately via bridge shim
5. `managementCommentary.html` — bundle display only
6. `attributionSidebar.html` — simplest module

---

## Key technical notes

### Date handling
Excel stores dates as serial numbers (e.g. `45383` = 2024-03-31).
`dataReader.js` `normaliseCell()` auto-detects and converts these to JS Date.
Range: 36526–73050 (year 2000–2099) avoids false positives with dollar amounts.

### XIRR accuracy
Same Newton-Raphson + Brent bisection as Code.gs.
ACT/365.25 day count (matches your existing results exactly).
500 iterations, 1e-9 tolerance. Returns `null` if unsolvable (shown as N/A).

### Bundle availability in dialogs
`taskpane.html` stores bundle in `sessionStorage` before opening each dialog.
`bridge.js` reads from `sessionStorage` first, falls back to `buildBundle()` from workbook.
Date objects are revived from JSON strings via `_reviveDates()`.

### Config Library
Same schema as Google Sheets: `Name | Config_JSON | Is_Default | Saved_At`
Office.settings stores the active config name for fast retrieval on next open.

---

*IVP IRR Analytics Engine · Excel Web Add-in · Phase 1*
