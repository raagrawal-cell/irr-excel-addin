/**
 * ═══════════════════════════════════════════════════════════════════════════
 * engine/xirr.js  —  IVP IRR Analytics Engine
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * Pure JavaScript XIRR solver — no Apps Script APIs, no external dependencies.
 * Ported 1:1 from Code.gs calcXIRR() conventions:
 *   - ACT/365 day count (365.25 for leap-year accuracy)
 *   - Newton-Raphson primary solver (500 iterations, 1e-9 tolerance)
 *   - Brent bisection fallback (handles sign changes, guaranteed convergence)
 *
 * Exports:
 *   calcXIRR(cashflows, guess?)     → number | null
 *   calcMetrics(cashflows, terminalNAV) → { moic, dpi, rvpi, invested, distributed }
 *   computeEntityIRR(txns, asOfDate) → { irr, metrics, valid, error }
 *
 * Cashflow contract:
 *   cashflows = [ { date: Date, amount: number }, ... ]
 *   Negative amounts = capital calls (outflows)
 *   Positive amounts = distributions or terminal NAV (inflows)
 *
 * ═══════════════════════════════════════════════════════════════════════════
 */

'use strict';

const XIRR_MAX_ITER = 500;
const XIRR_TOL      = 1e-9;
const XIRR_YEAR_MS  = 365.25 * 86400000;  // ACT/365.25

// ── Year fraction (ACT/365.25) ────────────────────────────────────────────
function yearFrac(d0, d1) {
  return (d1 - d0) / XIRR_YEAR_MS;
}

// ── Net present value at rate r ───────────────────────────────────────────
function xnpv(rate, cashflows) {
  const d0 = cashflows[0].date.getTime();
  let sum = 0;
  for (const cf of cashflows) {
    const t = (cf.date.getTime() - d0) / XIRR_YEAR_MS;
    sum += cf.amount / Math.pow(1 + rate, t);
  }
  return sum;
}

// ── Derivative of NPV w.r.t. rate ─────────────────────────────────────────
function xnpvDeriv(rate, cashflows) {
  const d0 = cashflows[0].date.getTime();
  let sum = 0;
  for (const cf of cashflows) {
    const t = (cf.date.getTime() - d0) / XIRR_YEAR_MS;
    if (t === 0) continue;
    sum -= t * cf.amount / Math.pow(1 + rate, t + 1);
  }
  return sum;
}

// ── Brent bisection fallback ──────────────────────────────────────────────
// Guaranteed convergence when Newton-Raphson fails (sign changes, flat regions)
function brentXIRR(cashflows, lo, hi, maxIter = 500, tol = 1e-9) {
  let flo = xnpv(lo, cashflows);
  let fhi = xnpv(hi, cashflows);

  // If both ends are same sign, try to widen bracket
  if (flo * fhi > 0) {
    // Scan for a bracket
    for (let step = -0.1; step > -0.999; step -= 0.05) {
      flo = xnpv(step, cashflows);
      if (flo * fhi < 0) { lo = step; break; }
    }
    if (flo * fhi > 0) return null; // no bracket found
  }

  let mid, fmid;
  for (let i = 0; i < maxIter; i++) {
    mid  = (lo + hi) / 2;
    fmid = xnpv(mid, cashflows);
    if (Math.abs(fmid) < tol || (hi - lo) / 2 < tol) return mid;
    if (flo * fmid < 0) { hi = mid; fhi = fmid; }
    else                 { lo = mid; flo = fmid; }
  }
  return mid;
}

// ── Validate cashflow array for XIRR ─────────────────────────────────────
// XIRR requires at least one sign change in the cashflow series
function validateCashflows(cashflows) {
  if (!cashflows || cashflows.length < 2) return false;
  let hasPos = false, hasNeg = false;
  for (const cf of cashflows) {
    if (cf.amount > 0) hasPos = true;
    if (cf.amount < 0) hasNeg = true;
    if (hasPos && hasNeg) return true;
  }
  return false;
}

/**
 * calcXIRR — Main XIRR function
 *
 * @param {Array}  cashflows  [{date: Date, amount: number}, ...]  sorted by date asc
 * @param {number} guess      Initial rate guess (default 0.1 = 10%)
 * @returns {number|null}     Annual rate as decimal (0.15 = 15%), or null if unsolvable
 */
export function calcXIRR(cashflows, guess = 0.10) {
  if (!validateCashflows(cashflows)) return null;

  // Sort ascending by date (defensive — caller should already sort)
  const cfs = [...cashflows].sort((a, b) => a.date - b.date);

  // ── Newton-Raphson ────────────────────────────────────────────────────
  let rate = guess;
  for (let i = 0; i < XIRR_MAX_ITER; i++) {
    const f  = xnpv(rate, cfs);
    const df = xnpvDeriv(rate, cfs);

    if (Math.abs(df) < 1e-15) break; // derivative too flat → fallback

    const next = rate - f / df;

    // Guard: rate must stay above -1 (can't have -100% return)
    if (next <= -1) break;

    if (Math.abs(next - rate) < XIRR_TOL) return next;
    rate = next;
  }

  // ── Brent bisection fallback ─────────────────────────────────────────
  const brent = brentXIRR(cfs, -0.999, 100);
  return brent;
}

/**
 * calcMetrics — Compute MOIC, DPI, RVPI from cashflows + terminal NAV
 *
 * @param {Array}  cashflows     Actual transaction cashflows (excl. terminal NAV)
 * @param {number} terminalNAV   Current/exit NAV value (positive number)
 * @returns {{ moic, dpi, rvpi, invested, distributed, nav }}
 */
export function calcMetrics(cashflows, terminalNAV = 0) {
  let invested     = 0;
  let distributed  = 0;

  for (const cf of cashflows) {
    if (cf.isCashflow !== false) { // default: treat all as cashflows
      if (cf.amount < 0) invested    += Math.abs(cf.amount);
      else               distributed += cf.amount;
    }
  }

  if (invested === 0) {
    return { moic: null, dpi: null, rvpi: null, invested: 0, distributed, nav: terminalNAV };
  }

  const moic = (distributed + terminalNAV) / invested;
  const dpi  = distributed / invested;
  const rvpi = terminalNAV / invested;

  return {
    moic:        parseFloat(moic.toFixed(4)),
    dpi:         parseFloat(dpi.toFixed(4)),
    rvpi:        parseFloat(rvpi.toFixed(4)),
    invested,
    distributed,
    nav:         terminalNAV
  };
}

/**
 * computeEntityIRR — Compute IRR for a set of transactions up to asOfDate
 *
 * This is the equivalent of Code.gs computeEntityIRR().
 * Takes an array of normalised transaction objects and a terminal date,
 * extracts the terminal NAV, builds the XIRR cashflow array, and computes.
 *
 * @param {Array}  txns      Normalised transaction objects from dataReader
 * @param {Date}   asOfDate  As-of date for terminal NAV cutoff
 * @returns {{ irr: number|null, metrics: object, valid: boolean, error: string|null }}
 */
export function computeEntityIRR(txns, asOfDate) {
  try {
    if (!txns || txns.length === 0) {
      return { irr: null, metrics: null, valid: false, error: 'No transactions' };
    }

    const cutoff = asOfDate instanceof Date ? asOfDate : new Date(asOfDate);

    // Separate cashflows and valuations
    const cfTxns  = txns.filter(t => t.isCashflow  && t.date <= cutoff);
    const valTxns = txns.filter(t => !t.isCashflow && t.date <= cutoff);

    if (cfTxns.length === 0) {
      return { irr: null, metrics: null, valid: false, error: 'No cashflows in range' };
    }

    // Terminal NAV = most recent valuation on or before asOfDate
    let terminalNAV = 0;
    if (valTxns.length > 0) {
      const latest = valTxns.reduce((best, t) => t.date > best.date ? t : best, valTxns[0]);
      terminalNAV = Math.abs(latest.amount);
    }

    // Build XIRR cashflow array: actual CFs + terminal NAV as final positive CF
    const xirrCFs = cfTxns.map(t => ({ date: t.date, amount: t.amount }));

    // Add terminal NAV only if it adds something meaningful
    if (terminalNAV > 0) {
      // Use the valuation date (or asOfDate) as the terminal date
      const termDate = valTxns.length > 0
        ? valTxns.reduce((best, t) => t.date > best.date ? t : best, valTxns[0]).date
        : cutoff;
      xirrCFs.push({ date: termDate, amount: terminalNAV });
    }

    xirrCFs.sort((a, b) => a.date - b.date);

    const irr     = calcXIRR(xirrCFs);
    const metrics = calcMetrics(cfTxns, terminalNAV);

    return {
      irr:     irr !== null ? parseFloat((irr * 100).toFixed(4)) / 100 : null, // keep as decimal
      metrics,
      valid:   irr !== null,
      error:   irr === null ? 'XIRR did not converge' : null
    };

  } catch (e) {
    return { irr: null, metrics: null, valid: false, error: e.message };
  }
}

/**
 * formatIRR — Format IRR decimal as display string
 * @param {number|null} irr   Decimal (0.1523 → "15.23%")
 * @param {number}      dp    Decimal places (default 2)
 */
export function formatIRR(irr, dp = 2) {
  if (irr === null || irr === undefined || isNaN(irr)) return 'N/A';
  return (irr * 100).toFixed(dp) + '%';
}

/**
 * formatMOIC — Format MOIC as display string
 * @param {number|null} moic  (1.234 → "1.23x")
 */
export function formatMOIC(moic, dp = 2) {
  if (moic === null || moic === undefined || isNaN(moic)) return 'N/A';
  return moic.toFixed(dp) + 'x';
}
