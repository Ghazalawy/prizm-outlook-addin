/**
 * Runtime config. Override at runtime by setting localStorage keys:
 *   prizm.erpBase       e.g. "https://ms.prizm-energy.com"
 *   prizm.apiKey        per-user API key for ERP
 *
 * These are read once at app boot. Settings view can change them.
 */
// Defaults target the Hetzner production deploy (PrizmIT upstream).
// To point a single browser at dev (Bluehost / Ghazalawy fork) without
// rebuilding, use the Settings view → API base URL:
//   https://dev.prizm-energy.com/outlookapi/bridge
//   https://dev.prizm-energy.com           (ERP base for "Open in ERP" links)
const DEFAULTS = {
  erpBase: 'https://ms.prizm-energy.com/MS',
  apiBase: 'https://ms.prizm-energy.com/MS/outlookapi/bridge',
  apiKey: '',
  version: '2.0.0',
};

const lsKey = (k) => `prizm.${k}`;

export const Config = {
  get(key) {
    const v = localStorage.getItem(lsKey(key));
    return v !== null ? v : DEFAULTS[key];
  },
  set(key, value) {
    if (value === null || value === undefined || value === '') {
      localStorage.removeItem(lsKey(key));
    } else {
      localStorage.setItem(lsKey(key), String(value));
    }
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
