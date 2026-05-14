import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

export async function render() {
  mount(el('div', {}, banner('info', 'Loading email context...')));
  const snap = await Office.snapshot();

  const inputs = {
    subject:  el('input', { type: 'text', value: snap.subject || '' }),
    contact:  el('input', { type: 'email', value: snap.from?.email || '' }),
    priority: el('select', {}, ...['Low','Medium','High','Urgent']
                .map((p) => el('option', { value: p, text: p, selected: p === 'Medium' || undefined }))),
    department: el('input', { type: 'text', placeholder: 'e.g. Operations' }),
    body:     el('textarea', { text: snap.bodyExcerpt || '' }),
  };

  const submit = el('button', { class: 'btn btn--block', type: 'submit' }, 'Open ticket');
  const cancel = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  const status = el('div', {});

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      submit.disabled = true; submit.textContent = 'Opening...'; status.replaceChildren();
      const payload = {
        subject: inputs.subject.value.trim(),
        contactEmail: inputs.contact.value.trim(),
        priority: inputs.priority.value,
        department: inputs.department.value.trim(),
        body: inputs.body.value,
        email: Office.envelope(snap, { attachEmailAsEml: true, attachFiles: true }),
      };
      if (!payload.subject) {
        status.replaceChildren(banner('err', 'Subject is required.'));
        submit.disabled = false; submit.textContent = 'Open ticket'; return;
      }
      try {
        const r = await Api.createTicket(payload);
        const url = r?.url || (r?.id ? `${Config.get('erpBase')}/admin/tickets/ticket/${r.id}` : null);
        status.replaceChildren(banner('ok', url ? `Ticket #${r.id} opened.` : 'Ticket opened.'));
        if (url) status.appendChild(el('div', {}, el('a', { href: url, target: '_blank', rel: 'noopener', text: 'Open in ERP' })));
        submit.textContent = 'Open another'; submit.disabled = false;
      } catch (err) {
        const msg = err instanceof ApiError ? `Failed (${err.status||'net'}): ${err.message}` : `Failed: ${err.message}`;
        status.replaceChildren(banner('err', msg));
        submit.disabled = false; submit.textContent = 'Retry';
      }
    },
  },
    contextBanner(snap),
    field('Subject', inputs.subject, { required: true }),
    row(field('Contact email', inputs.contact), field('Priority', inputs.priority)),
    field('Department', inputs.department),
    field('Description', inputs.body),
    el('div', { class: 'btn-row' }, cancel, submit),
    status,
  );

  mount(form);
}
