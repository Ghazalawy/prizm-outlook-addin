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

  const save = el('button', { class: 'btn', type: 'button', onclick: () => {
    Config.set('erpBase', inputs.erpBase.value.trim());
    Config.set('apiBase', inputs.apiBase.value.trim());
    Config.set('apiKey',  inputs.apiKey.value);
    status.replaceChildren(banner('ok', 'Saved. Click Test connection to verify, then use the back arrow.'));
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

  const clear = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => {
    Config.set('erpBase', null);
    Config.set('apiBase', null);
    Config.set('apiKey',  null);
    window.location.reload();
  }}, 'Reset');

  const firstRun = !!window.__prizmFirstRun;
  window.__prizmFirstRun = false;

  mount(el('div', {},
    firstRun
      ? banner('info',
          'First-time setup — paste your personal API key from Prizm ERP, save, then test the connection. '
          + 'Open the keys page in ERP: https://ms.prizm-energy.com/MS/admin/outlookapi/keys')
      : el('div', { class: 'card' },
          el('h3', { text: 'ERP connection' }),
          el('p', { text: 'These are stored locally per browser/Outlook profile.' }),
        ),
    field('ERP base URL', inputs.erpBase, { hint: 'Used for "Open in ERP" links, e.g. https://ms.prizm-energy.com/MS' }),
    field('API base URL', inputs.apiBase, { hint: 'Bridge endpoint, e.g. https://ms.prizm-energy.com/MS/outlookapi/bridge' }),
    field('API key',      inputs.apiKey,  { hint: 'Generate at ERP → Outlook → Add-in API keys. Shown once.' }),
    el('div', { class: 'btn-row' }, save, test, clear),
    status,
  ));
}
