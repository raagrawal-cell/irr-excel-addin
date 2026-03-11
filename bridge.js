/**
 * ═══════════════════════════════════════════════════════════════════════════
 * engine/bridge.js  —  IVP IRR Analytics Engine
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * This is the ONLY file that each ported HTML module needs to add.
 * It replaces the google.script.run bridge with local async equivalents.
 *
 * USAGE: Add this single script tag to each module HTML, before any
 * module-specific scripts:
 *
 *   <script type="module" src="../engine/bridge.js"></script>
 *
 * What it does:
 *   1. Loads the BUNDLE from sessionStorage (set by taskpane.html when opening the dialog)
 *      OR re-builds it from the workbook directly (fallback for direct dialog opens)
 *   2. Exposes window.IVP_BUNDLE, window.IVP_CFG, window.IVP_SIM_CONFIG
 *   3. Replaces the <?!= bundle ?> template injection pattern:
 *      modules must call window.IVP_INIT() which resolves when data is ready
 *   4. Creates a compatibility shim google.script.run.X() → local function
 *      so that minimal changes are needed in existing module HTML files
 *   5. Loads Office.js if available (for write-back operations like saveSimulationToLog)
 *
 * ─────────────────────────────────────────────────────────────────────────
 * REPLACING TEMPLATE INJECTION IN EACH MODULE:
 *
 * Before (Google Sheets):
 *   var BUNDLE = <?!= bundle ?>;
 *   var SIM_CONFIG = <?!= simConfig ?>;
 *
 * After (Excel Add-in):
 *   // Remove the above lines and add at top of init():
 *   await window.IVP_INIT();
 *   var BUNDLE = window.IVP_BUNDLE;
 *   var SIM_CONFIG = window.IVP_SIM_CONFIG;
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */

'use strict';

import {
  buildBundle,
  getActiveConfig,
  getConfigs,
  getSimConfig,
  getAnalyticsAllTransactions,
  getAnalyticsSubPeriodTrend,
  saveSimulationToLog,
  saveConfig,
  deleteConfig,
  setDefaultConfig,
  getSheetNames,
  getTabHeaders,
  getConfigs as getConfigsAll
} from './dataReader.js';

import {
  computeEntityIRR
} from './xirr.js';

// ── Expose backsolve and simulation modules once they are built (P3, P4) ──
// These are imported lazily to allow incremental build
let _backsolve   = null;
let _simulation  = null;

async function _loadBacksolve() {
  if (_backsolve) return _backsolve;
  try {
    _backsolve = await import('./backsolve.js');
    return _backsolve;
  } catch (e) {
    console.warn('[bridge] backsolve.js not yet built:', e.message);
    return null;
  }
}

async function _loadSimulation() {
  if (_simulation) return _simulation;
  try {
    _simulation = await import('./simulation.js');
    return _simulation;
  } catch (e) {
    console.warn('[bridge] simulation.js not yet built:', e.message);
    return null;
  }
}


// ─────────────────────────────────────────────────────────────────────────
//  INITIALISATION
// ─────────────────────────────────────────────────────────────────────────

let _initPromise = null;

/**
 * IVP_INIT — Call this at the top of each module's init/onReady function.
 *
 * Resolves when window.IVP_BUNDLE and window.IVP_CFG are ready.
 * Idempotent — safe to call multiple times; only loads once.
 *
 * @returns {Promise<void>}
 */
window.IVP_INIT = function() {
  if (_initPromise) return _initPromise;
  _initPromise = _doInit();
  return _initPromise;
};

async function _doInit() {
  // ── CRITICAL: always build the shim even if data loading fails ──────────
  // Without this, any throw inside _doInit leaves window.google undefined
  // and every rpc() call produces "google is not defined".
  try {
    // ── 1. Try sessionStorage (set by taskpane.html when opening dialog) ─
    try {
      const bundleJson  = sessionStorage.getItem('IVP_BUNDLE');
      const cfgJson     = sessionStorage.getItem('IVP_CFG');
      const configsJson = sessionStorage.getItem('IVP_CONFIGS');

      if (bundleJson) {
        window.IVP_BUNDLE  = _reviveDates(JSON.parse(bundleJson));
        window.IVP_CFG     = cfgJson     ? JSON.parse(cfgJson)     : null;
        window.IVP_CONFIGS = configsJson ? JSON.parse(configsJson) : [];
        window.IVP_SIM_CONFIG = await getSimConfig().catch(() => ({ apiKey: '', apiModel: 'claude-sonnet-4-20250514', maxTokens: 1000 }));
        return; // Data loaded from session — done (finally still runs)
      }
    } catch (e) {
      console.warn('[bridge] sessionStorage read failed, rebuilding from workbook:', e.message);
    }

    // ── 2. Fallback: re-build from workbook directly ─────────────────────
    // Happens when a module HTML is opened as a dialog (separate window context)
    window.IVP_CONFIGS = await getConfigsAll().catch(() => []);
    window.IVP_CFG     = await getActiveConfig().catch(() => null);

    if (window.IVP_CFG && window.IVP_CFG.config) {
      window.IVP_BUNDLE = await buildBundle(window.IVP_CFG.config).catch(e => {
        console.error('[bridge] buildBundle failed:', e.message);
        return null;
      });
    } else {
      window.IVP_BUNDLE = null;
    }

    window.IVP_SIM_CONFIG = await getSimConfig().catch(() => ({ apiKey: '', apiModel: 'claude-sonnet-4-20250514', maxTokens: 1000 }));

  } catch (e) {
    console.error('[bridge] _doInit error (non-fatal, shim will still be built):', e);
    window.IVP_BUNDLE     = window.IVP_BUNDLE     || null;
    window.IVP_SIM_CONFIG = window.IVP_SIM_CONFIG || { apiKey: '', apiModel: 'claude-sonnet-4-20250514', maxTokens: 1000 };
  } finally {
    // ── Always build the shim — this MUST run even if everything above failed ──
    _buildGoogleScriptRunShim();
  }
}


// ─────────────────────────────────────────────────────────────────────────
//  GOOGLE.SCRIPT.RUN COMPATIBILITY SHIM
//  Maps every google.script.run.X() call to a local function.
//  This is the key to near-zero changes in existing module HTML files.
// ─────────────────────────────────────────────────────────────────────────

function _buildGoogleScriptRunShim() {
  /**
   * The shim follows the google.script.run chaining API:
   *   google.script.run
   *     .withSuccessHandler(cb)
   *     .withFailureHandler(errCb)
   *     .functionName(args...)
   *
   * Each method returns a proxy object so the chain works.
   * The actual function is called async; success/failure callbacks are invoked.
   */

  // Helper: create a runner that wraps an async fn.
  // IMPORTANT: .bind(target) is called on the returned function so that
  // `this` inside the runner === target (the shimRoot), which is where
  // .withSuccessHandler() and .withFailureHandler() stored their callbacks.
  function makeRunner(asyncFn) {
    return function runner(...args) {
      const ctx = this; // `this` === shimRoot target (via .bind(target) below)
      asyncFn(...args)
        .then(result  => { if (ctx && ctx._success) ctx._success(result); })
        .catch(err    => { if (ctx && ctx._failure) ctx._failure(err); else console.error('[bridge] Unhandled RPC error:', err); });
    };
  }

  // Proxy handler: intercepts .withSuccessHandler / .withFailureHandler / .fnName
  const handler = {
    get(target, prop) {
      if (prop === 'withSuccessHandler') {
        return (cb) => { target._success = cb; return new Proxy(target, handler); };
      }
      if (prop === 'withFailureHandler') {
        return (cb) => { target._failure = cb; return new Proxy(target, handler); };
      }
      // Any other property = a function name to call
      if (typeof prop === 'string' && !prop.startsWith('_')) {
        const fn = _getShimFunction(prop);
        if (!fn) {
          return (...args) => {
            const err = new Error(`[bridge] Function "${prop}" not yet ported. Phase 2+ required.`);
            if (target._failure) target._failure(err);
            else console.warn(err.message);
          };
        }
        return makeRunner(fn).bind(target);
      }
      return target[prop];
    }
  };

  // Build the google.script.run proxy.
  // `run` is a getter that creates a FRESH target for every new chain,
  // so concurrent rpc() calls each get their own _success/_failure slots.
  window.google = window.google || {};
  window.google.script = {
    get run() { return new Proxy({}, handler); },
    host: {
      close: () => {
        // In a dialog: send close message to taskpane
        try {
          Office.context.ui.messageParent(JSON.stringify({ type: 'close' }));
        } catch (e) {
          window.close();
        }
      },
      editor: { focus: () => {} }
    }
  };
}


// ─────────────────────────────────────────────────────────────────────────
//  SHIM FUNCTION REGISTRY
//  Maps function name → async function with same signature as GAS version.
//  All return Promises — the runner handles the callback dispatch.
// ─────────────────────────────────────────────────────────────────────────

function _getShimFunction(name) {
  const registry = {

    // ── Analytics Dashboard ─────────────────────────────────────────────
    'getAnalyticsAllTransactions': async (pathStr, quarter) => {
      if (!window.IVP_BUNDLE) throw new Error('Bundle not loaded');
      return getAnalyticsAllTransactions(pathStr, quarter, window.IVP_BUNDLE);
    },

    'getAnalyticsSubPeriodTrend': async (pathStr, quarter) => {
      if (!window.IVP_BUNDLE) throw new Error('Bundle not loaded');
      return getAnalyticsSubPeriodTrend(pathStr, quarter, window.IVP_BUNDLE);
    },

    // ── Reverse IRR / Backsolve ──────────────────────────────────────────
    'solveReverseIRR': async (params) => {
      const bs = await _loadBacksolve();
      if (!bs) throw new Error('backsolve.js not built yet (Phase 3)');
      if (!window.IVP_BUNDLE) throw new Error('Bundle not loaded');
      return bs.solveReverseIRR(params, window.IVP_BUNDLE);
    },

    // ── SimEngine ────────────────────────────────────────────────────────
    'callClaudeAPI': async (payload) => {
      // Direct fetch to Anthropic API — no CORS issue in Excel task pane / dialog
      // This is the same as the existing XHR fallback path in irrSimulator.html
      const { apiKey, model, maxTokens, system, userMessage } = payload;
      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type':          'application/json',
          'x-api-key':             apiKey,
          'anthropic-version':     '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
          model:      model || 'claude-sonnet-4-20250514',
          max_tokens: maxTokens || 1000,
          system,
          messages: [{ role: 'user', content: userMessage }]
        })
      });
      if (!response.ok) throw new Error('Claude API error: ' + response.status);
      const data = await response.json();
      return {
        text:  (data.content || []).map(b => b.text || '').join(''),
        usage: data.usage
      };
    },

    'runSimulation': async (payload) => {
      const sim = await _loadSimulation();
      if (!sim) throw new Error('simulation.js not built yet (Phase 3)');
      if (!window.IVP_BUNDLE) throw new Error('Bundle not loaded');
      return sim.runSimulation(payload, window.IVP_BUNDLE);
    },

    'saveSimulationToLog': async (payload) => {
      return saveSimulationToLog(payload);
    },

    // ── Management Commentary ────────────────────────────────────────────
    'generateCommentaryPDF': async () => {
      // Print the current dialog content as PDF via browser print dialog
      // The user saves as PDF from the system print dialog
      window.print();
      return { ok: true };
    },

    // ── Config Library ───────────────────────────────────────────────────
    'getSheetNames': async () => {
      return getSheetNames();
    },

    'getTabHeaders': async (tab, row) => {
      return getTabHeaders(tab, row || 1);
    },

    'getConfigs': async () => {
      return getConfigsAll();
    },

    'saveConfig': async (name, cfg, isDef) => {
      await saveConfig(name, cfg, isDef);
      // Notify taskpane that config was saved (so it can refresh)
      try {
        Office.context.ui.messageParent(JSON.stringify({ type: 'config_saved', name }));
      } catch (e) { /* not in dialog context */ }
      return { ok: true };
    },

    'deleteConfig': async (name) => {
      await deleteConfig(name);
      return { ok: true };
    },

    'setDefaultConfig': async (name) => {
      await setDefaultConfig(name);
      return { ok: true };
    },

    'previewFromUI': async (cfg) => {
      // Build a quick preview bundle and return a summary string
      const bundle = await buildBundle(cfg);
      if (bundle.error) throw new Error(bundle.error);
      return `Preview: ${bundle.txnCount} records · ${bundle.deals.length} deals · ${bundle.quarters.length} quarters · Fund IRR: ${bundle.fund.irr !== null ? (bundle.fund.irr * 100).toFixed(2) + '%' : 'N/A'} · MOIC: ${bundle.fund.moic ? bundle.fund.moic.toFixed(2) + 'x' : 'N/A'}`;
    },

    'saveAndCompute': async (name, cfg, isDef) => {
      // Save config then rebuild bundle
      await saveConfig(name, cfg, isDef);
      const bundle = await buildBundle(cfg);
      window.IVP_BUNDLE = bundle;
      // Refresh sessionStorage for other modules
      try { sessionStorage.setItem('IVP_BUNDLE', JSON.stringify(bundle)); } catch (e) {}
      // Notify taskpane to reload
      try {
        Office.context.ui.messageParent(JSON.stringify({ type: 'compute_complete', name }));
      } catch (e) { /* not in dialog context */ }
      return { ok: true, bundle };
    },

  };

  return registry[name] || null;
}


// ─────────────────────────────────────────────────────────────────────────
//  UTILITIES
// ─────────────────────────────────────────────────────────────────────────

/**
 * _reviveDates — Walk a JSON-parsed object and convert ISO date strings
 * back to Date objects (JSON.stringify converts Dates to strings).
 *
 * Only converts strings that look like ISO dates to avoid false positives.
 */
function _reviveDates(obj) {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj === 'string') {
    // ISO date: "2024-03-31T00:00:00.000Z" or "2024-03-31"
    if (/^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?/.test(obj)) {
      const d = new Date(obj);
      if (!isNaN(d.getTime())) return d;
    }
    return obj;
  }
  if (Array.isArray(obj)) return obj.map(_reviveDates);
  if (typeof obj === 'object') {
    const out = {};
    for (const [k, v] of Object.entries(obj)) {
      out[k] = _reviveDates(v);
    }
    return out;
  }
  return obj;
}


// ─────────────────────────────────────────────────────────────────────────
//  AUTO-INIT
//  Kick off IVP_INIT immediately so data is loading in the background
//  before the module's own DOMContentLoaded fires.
// ─────────────────────────────────────────────────────────────────────────
window.IVP_INIT();