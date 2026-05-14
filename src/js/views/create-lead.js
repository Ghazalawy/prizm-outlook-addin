import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

export async function render() {
  mount(el('div', {}, banner('info', 'Loading email context...')));
  const snap = await Office.snapshot();

  const fromName = snap.from?.name || '';
  const fromEmail = snap.from?.email || '';
  const company = fromEmail.includes('@') ? fromEmail.split('@')[1].split('.')[0] : '';

  const inputs = {
    name:     el('input', { type: 'text', value: fromName }),
    email:    el('input', { type: 'email', value: fromEmail }),
    company:  el('input', { type: 'text', value: company }),
    phone:    el('input', { type: 'tel' }),
    source:   el('select', {}, ...['Email','Web','Referral','LinkedIn','Tender','Other']
                .map((s) => el('option', { value: s, text: s, selected: s === 'Email' || undefined }))),
    status:   el('select', {}, ...['New','Contacted','Qualified','Unqualified']
                .map((s) => el('option', { value: s, text: s, selected: s === 'New' || undefined }))),
    notes:    el('textarea', { text: `Subject: ${snap.subject || ''}\n\n${snap.bodyExcerpt || ''}` }),
  };

  const submit = el('button', { class: 'btn btn--block', type: 'submit' }, 'Create lead');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  const status = el('div', {});

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      submit.disabled = true; submit.textContent = 'Creating...'; status.replaceChildren();
      const payload = {
        name: inputs.name.value.trim(),
        email: inputs.email.value.trim(),
        company: inputs.company.value.trim(),
        phone: inputs.phone.value.trim(),
        source: inputs.source.value,
        status: inputs.status.value,
        notes: inputs.notes.value,
        email_link: { itemId: snap.itemId, internetMessageId: snap.internetMessageId },
      };
      if (!payload.email && !payload.name) {
        status.replaceChildren(banner('err', 'At least name or email is required.'));
        submit.disabled = false; submit.textContent = 'Create lead'; return;
      }
      try {
        const r = await Api.createLead(payload);
        const url = r?.url || (r?.id ? `${Config.get('erpBase')}/admin/leads/index/${r.id}` : null);
        status.replaceChildren(banner('ok', url ? `Lead #${r.id} created.` : 'Lead created.'));
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
    row(field('Name', inputs.name), field('Email', inputs.email, { required: true })),
    row(field('Company', inputs.company), field('Phone', inputs.phone)),
    row(field('Source', inputs.source), field('Status', inputs.status)),
    field('Notes', inputs.notes),
    el('div', { class: 'btn-row' }, cancel, submit),
    status,
  );

  mount(form);
}
