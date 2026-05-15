/**
 * Runtime config for the add-in.
 *
 * Storage tiers (read order):
 *   1. Office.context.roamingSettings  — authoritative when running in Outlook.
 *      Per-user, per-add-in storage held by Exchange. Survives across devices,
 *      sessions, browser cache clears, and Outlook's third-party-iframe storage
 *      partitioning (which silently drops localStorage in Outlook on the web).
 *   2. localStorage  — fast in-process cache and the only store when we're
 *      running outside Outlook (e.g. dev preview at GitHub Pages).
 *
 * Writes go to both tiers; localStorage is updated synchronously, roaming is
 * saved via saveAsync. The Settings view uses Config.save({...}) to wait for
 * the round-trip before showing "Saved.".
 */
const DEFAULTS = {
  erpBase: 'https://ms.prizm-energy.com/MS',
  apiBase: 'https://ms.prizm-energy.com/MS/outlookapi/bridge',
  apiKey: '',
  version: '2.0.0',
};

const lsKey = (k) => `prizm.${k}`;
const rsKey = (k) => `prizm.${k}`;

let roaming = null;
let roamingReady = false;

function detectRoaming() {
  if (typeof window === 'undefined') return null;
  if (typeof window.Office === 'undefined') return null;
  return window.Office.context?.roamingSettings || null;
}

export const Config = {
  /**
   * Hook up roamingSettings once Office.js has finished initializing.
   * Also runs a one-time migration: any value previously written to
   * localStorage (older builds) is copied up into roamingSettings so the
   * user doesn't have to re-enter their API key.
   */
  attachRoaming() {
    roaming = detectRoaming();
    if (!roaming) return false;
    roamingReady = true;

    const keys = ['erpBase', 'apiBase', 'apiKey'];
    let migrated = false;
    keys.forEach((k) => {
      const inRoaming = roaming.get(rsKey(k));
      const inLs = localStorage.getItem(lsKey(k));
      if ((inRoaming === undefined || inRoaming === null || inRoaming === '') && inLs) {
        roaming.set(rsKey(k), inLs);
        migrated = true;
      }
    });
    if (migrated) {
      roaming.saveAsync(() => { /* fire-and-forget */ });
    }
    return true;
  },

  /** Returns true if we're using roamingSettings, false if just localStorage. */
  isRoaming() { return roamingReady; },

  get(key) {
    if (roamingReady && roaming) {
      const v = roaming.get(rsKey(key));
      if (v !== undefined && v !== null && v !== '') return v;
    }
    const v = localStorage.getItem(lsKey(key));
    if (v !== null && v !== '') return v;
    return DEFAULTS[key];
  },

  /**
   * Write a single key. Promise resolves when the roamingSettings round-trip
   * is acknowledged by Exchange (or immediately when running outside Outlook).
   */
  set(key, value) {
    const isEmpty = value === null || value === undefined || value === '';
    if (isEmpty) {
      localStorage.removeItem(lsKey(key));
    } else {
      localStorage.setItem(lsKey(key), String(value));
    }
    if (!roamingReady || !roaming) return Promise.resolve({ status: 'no-roaming' });
    if (isEmpty) roaming.remove(rsKey(key));
    else         roaming.set(rsKey(key), String(value));
    return new Promise((resolve) => {
      try {
        roaming.saveAsync((r) => resolve({ status: r?.status || 'unknown' }));
      } catch (e) {
        resolve({ status: 'error', error: e?.message || String(e) });
      }
    });
  },

  /**
   * Batch save — preferred from the Settings UI so a single Exchange round-trip
   * persists all three values.
   * @param {object} updates partial map of keys to write
   */
  save(updates) {
    Object.entries(updates).forEach(([k, v]) => {
      const isEmpty = v === null || v === undefined || v === '';
      if (isEmpty) localStorage.removeItem(lsKey(k));
      else         localStorage.setItem(lsKey(k), String(v));
      if (roamingReady && roaming) {
        if (isEmpty) roaming.remove(rsKey(k));
        else         roaming.set(rsKey(k), String(v));
      }
    });
    if (!roamingReady || !roaming) return Promise.resolve({ status: 'no-roaming' });
    return new Promise((resolve) => {
      try {
        roaming.saveAsync((r) => resolve({ status: r?.status || 'unknown' }));
      } catch (e) {
        resolve({ status: 'error', error: e?.message || String(e) });
      }
    });
  },

  all() {
    return {
      erpBase: this.get('erpBase'),
      apiBase: this.get('apiBase'),
      apiKey: this.get('apiKey'),
      version: DEFAULTS.version,
    };
  },
};
