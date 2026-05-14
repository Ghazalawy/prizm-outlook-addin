/**
 * Runtime config. Override at runtime by setting localStorage keys:
 *   prizm.erpBase       e.g. "https://ms.prizm-energy.com"
 *   prizm.apiKey        per-user API key for ERP
 *
 * These are read once at app boot. Settings view can change them.
 */
const DEFAULTS = {
  erpBase: 'https://ms.prizm-energy.com',
  apiBase: 'https://ms.prizm-energy.com/dev/outlookapi/bridge',
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
