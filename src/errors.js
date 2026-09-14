// Office.js error → user-facing notice mapping.
//
// Office's docs warn that `error.message` is developer-facing diagnostic text, not
// fit for end users, and that the stable signal is `error.code`. So we branch on
// code (with an offline check first, since a dropped connection masquerades as all
// sorts of codes) and only ever surface the raw message as a parenthetical detail
// on otherwise-unrecognized errors. Pure: no Office/Word/DOM refs — `navigator` is
// read through `globalThis` so tests can inject a shim. Every export is total: it
// must never throw, because it runs inside catch blocks on the insert pipeline.

// A normalized {code, message} view of whatever the caller caught — Office.js
// throws OfficeExtension.Error objects, but a string or undefined can slip through.
function normalize(error) {
  if (typeof error === 'string') return { code: '', message: error };
  if (error && typeof error === 'object') {
    return { code: error.code || '', message: error.message || '' };
  }
  return { code: '', message: '' };
}

// Network signatures seen across browsers/runtimes for a failed fetch. Matched
// against the error message as a fallback when navigator.onLine is unavailable or
// the runtime didn't flip it (common inside the add-in webview).
const OFFLINE_RE = /failed to fetch|networkerror|net::|err_internet|offline/i;

// True when the failure is most plausibly a connectivity problem: the browser
// reports offline, or the message carries a network signature. Never throws.
export function isLikelyOffline(error) {
  if (globalThis.navigator?.onLine === false) return true;
  const { message } = normalize(error);
  return OFFLINE_RE.test(message);
}

// Generic (unknown/unmapped) failures get their severity from the step: losing the
// stored source or the click-to-edit tag is recoverable (the equation still got
// inserted), so those warn; failing to insert or to re-locate is a hard error.
function levelForStep(step) {
  if (step === 'store' || step === 'tag') return 'warn';
  return 'error';
}

// Append the raw diagnostic in parentheses so a support request can quote it,
// without ever putting it front-and-center as the user-facing sentence.
function withDetail(msg, message) {
  return message ? `${msg} (${message})` : msg;
}

// Turn a failed insert-pipeline step + the thrown error into {level, msg}.
// step ∈ {'insert','update-locate','tag','store'}. Never throws.
export function describeInsertError(step, error) {
  const { code, message } = normalize(error);

  // Offline wins: a connectivity failure dwarfs whatever code came back.
  if (isLikelyOffline(error)) {
    return {
      level: 'error',
      msg: 'You appear to be offline. The equation renderer needs a connection — '
        + 'reconnect and try again.',
    };
  }

  // Known, stable Office.js codes get tailored, actionable messages.
  switch (code) {
    case 'AccessDenied':
      return {
        level: 'error',
        msg: 'This document is read-only or protected, so the equation can\'t be '
          + 'added. Disable protection or save an editable copy, then try again.',
      };
    case 'ApiNotAvailable':
      return {
        level: 'error',
        msg: 'This feature needs Word on Windows or Mac desktop. The current host '
          + 'doesn\'t support it.',
      };
    case 'RequestPayloadSizeLimitExceeded':
      return {
        level: 'error',
        msg: 'The equation is too large to insert. Try a smaller equation or split '
          + 'it into parts.',
      };
  }

  // Unknown code: severity by step, message that points the user somewhere, plus
  // the raw detail in parentheses for diagnostics.
  const level = levelForStep(step);
  if (step === 'store') {
    return {
      level,
      msg: withDetail(
        'The equation was inserted, but its source couldn\'t be saved, so '
          + 'click-to-edit may not work for it. You can still re-enter it manually.',
        message,
      ),
    };
  }
  if (step === 'tag') {
    return {
      level,
      msg: withDetail(
        'The equation was inserted, but click-to-edit couldn\'t be enabled for it.',
        message,
      ),
    };
  }
  return {
    level,
    msg: withDetail('Something went wrong. Please try again.', message),
  };
}

// Shown when a click-to-edit target picture is gone (deleted by the user since it
// was inserted): we can't update in place, so we fall back to a fresh insert.
export function pictureVanishedNotice() {
  return {
    level: 'info',
    msg: 'That equation is no longer in the document — inserting it as a new '
      + 'equation instead.',
  };
}
