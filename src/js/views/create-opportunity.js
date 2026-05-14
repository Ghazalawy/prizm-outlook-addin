import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

export async function render() {
  mount(el('div', {}, banner('info', 'Loading email context...')));
  const snap = await Office.snapshot();

  const inputs = {
    name:      el('input', { type: 'text', value: snap.subject || '' }),
    contactName: el('input', { type: 'text', value: snap.from?.name || '' }),
    contactEmail:el('input', { type: 'email', value: snap.from?.email || '' }),
    value:     el('input', { type: 'number', value: '0', min: '0' }),
    currency:  el('select', {}, ...['AED','USD','SAR','EUR','GBP'].map((c) => el('option', { value: c, text: c }))),
    closeDate: el('input', { type: 'date', value: new Date(Date.now() + 30 * 86400_000).toISOString().slice(0,10) }),
    stage:     el('select', {}, ...['Prospect','Qualified','Proposal','Negotiation','Won','Lost']
                  .map((s) => el('option', { value: s, text: s, selected: s === 'Prospect' || undefined }))),
    notes:     el('textarea', { text: snap.bodyExcerpt || '' }),
  };

  const submit = el('button', { class: 'btn btn--block', type: 'submit' }, 'Create opportunity');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  const status = el('div', {});

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      submit.disabled = true; submit.textContent = 'Creating...';
      status.replaceChildren();

      const payload = {
        name: inputs.name.value.trim(),
        contact: { name: inputs.contactName.value.trim(), email: inputs.contactEmail.value.trim() },
        value: Number(inputs.value.value || 0),
        currency: inputs.currency.value,
        closeDate: inputs.closeDate.value,
        stage: inputs.stage.value,
        notes: inputs.notes.value,
        email: { itemId: snap.itemId, internetMessageId: snap.internetMessageId, subject: snap.subject, from: snap.from },
      };
      if (!payload.name) {
        status.replaceChildren(banner('err', 'Name is required.'));
        submit.disabled = false; submit.textContent = 'Create opportunity'; return;
      }

      try {
        const r = await Api.createOpportunity(payload);
        const url = r?.url || (r?.id ? `${Config.get('erpBase')}/admin/leads/index/${r.id}` : null);
        status.replaceChildren(banner('ok', url ? `Opportunity #${r.id} created.` : 'Opportunity created.'));
        if (url) status.appendChild(el('div', {}, el('a', { href: url, target: '_blank', rel: 'noopener', text: 'Open in ERP' })));
        submit.textContent = 'Create another'; submit.disabled = false;
      } catch (err) {
        const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
        status.replaceChildren(banner('err', msg));
        submit.disabled = false; submit.textContent = 'Retry';
      }
    },
  },
    contextBanner(snap),
    field('Opportunity name', inputs.name, { required: true }),
    row(field('Contact name', inputs.contactName), field('Contact email', inputs.contactEmail)),
    row(field('Value', inputs.value), field('Currency', inputs.currency)),
    row(field('Close date', inputs.closeDate), field('Stage', inputs.stage)),
    field('Notes', inputs.notes),
    el('div', { class: 'btn-row' }, cancel, submit),
    status,
  );

  mount(form);
}
