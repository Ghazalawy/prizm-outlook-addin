import { Config } from '../config.js';
import { Api } from '../api.js';
import { el, mount, field, banner } from '../ui.js';

export async function render() {
  const inputs = {
    erpBase: el('input', { type: 'url', value: Config.get('erpBase') }),
    apiBase: el('input', { type: 'url', value: Config.get('apiBase') }),
    apiKey:  el('input', { type: 'password', value: Config.get('apiKey') }),
  };
  const status = el('div', {});

  const save = el('button', { class: 'btn', type: 'button', onclick: async () => {
    save.disabled = true; save.textContent = 'Saving…';
    status.replaceChildren(banner('info', 'Saving…'));
    const r = await Config.save({
      erpBase: inputs.erpBase.value.trim(),
      apiBase: inputs.apiBase.value.trim(),
      apiKey:  inputs.apiKey.value,
    });
    save.disabled = false; save.textContent = 'Save';
    if (r?.status === 'succeeded' || r?.status === 'no-roaming') {
      const where = Config.isRoaming()
        ? 'saved to your Outlook profile (syncs across devices)'
        : 'saved locally in this browser';
      status.replaceChildren(banner('ok', `Settings ${where}. Test the connection next.`));
    } else {
      status.replaceChildren(banner('err',
        'Saved locally, but could not persist to Outlook roaming settings'
        + (r?.error ? `: ${r.error}` : '. The key may be lost when Outlook restarts.')));
    }
  }}, 'Save');

  const test = el('button', { class: 'btn btn--ghost', type: 'button', onclick: async () => {
    status.replaceChildren(banner('info', 'Pinging ERP...'));
    try {
      const r = await Api.ping();
      status.replaceChildren(banner('ok', `Connected: ${r?.user || 'ok'} (v${r?.version || '?'})`));
    } catch (err) {
      status.replaceChildren(banner('err', `Failed: ${err.message}`));
    }
  }}, 'Test connection');

  const clear = el('button', { class: 'btn btn--ghost', type: 'button', onclick: async () => {
    await Config.save({ erpBase: null, apiBase: null, apiKey: null });
    window.location.reload();
  }}, 'Reset');

  const firstRun = !!window.__prizmFirstRun;
  window.__prizmFirstRun = false;

  const storageNote = Config.isRoaming()
    ? 'Stored in your Outlook profile (Exchange roamingSettings) — persists across devices and re-installs.'
    : 'Stored in this browser only (no Outlook context detected).';

  mount(el('div', {},
    firstRun
      ? banner('info',
          'First-time setup — paste your personal API key from Prizm ERP, save, then test the connection. '
          + 'Open the keys page in ERP: https://ms.prizm-energy.com/MS/admin/outlookapi/keys')
      : el('div', { class: 'card' },
          el('h3', { text: 'ERP connection' }),
          el('p', { text: storageNote }),
        ),
    field('ERP base URL', inputs.erpBase, { hint: 'Used for "Open in ERP" links, e.g. https://ms.prizm-energy.com/MS' }),
    field('API base URL', inputs.apiBase, { hint: 'Bridge endpoint, e.g. https://ms.prizm-energy.com/MS/outlookapi/bridge' }),
    field('API key',      inputs.apiKey,  { hint: 'Generate at ERP → Outlook → Add-in API keys. Shown once.' }),
    el('div', { class: 'btn-row' }, save, test, clear),
    status,
  ));
}
