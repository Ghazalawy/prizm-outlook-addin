/**
 * Create Task — pre-filled from the current email.
 * Posts to POST {apiBase}/outlook/tasks.
 *
 * Pattern to follow when adding new "Create X from email" views:
 *   1. Pull a snapshot of the email via Office.snapshot()
 *   2. Build the form with ui.field/row helpers
 *   3. On submit: call the relevant Api.createX(payload)
 *   4. Show banner with result + link to the created record (if returned)
 */
import { Office } from '../office.js';
import { Api, ApiError } from '../api.js';
import { Config } from '../config.js';
import { el, mount, field, row, banner, contextBanner } from '../ui.js';

export async function render() {
  // Loading state
  mount(el('div', {}, banner('info', 'Loading email context...')));

  const snap = await Office.snapshot();
  const todayIso = new Date().toISOString().slice(0, 10);
  const dueIso = new Date(Date.now() + 7 * 86400_000).toISOString().slice(0, 10);

  // Refdata is optional — failure shouldn't block the form
  let refdata = { priorities: ['Low','Medium','High','Urgent'], staff: [], tags: [] };
  try {
    const r = await Api.refdata();
    if (r) refdata = { ...refdata, ...r };
  } catch (_) { /* offline-friendly fallback */ }

  const inputs = {
    subject:    el('input', { type: 'text', value: snap.subject || '' }),
    startDate:  el('input', { type: 'date', value: todayIso }),
    dueDate:    el('input', { type: 'date', value: dueIso }),
    priority:   el('select', {},
      ...refdata.priorities.map((p) => el('option', { value: p, text: p, selected: p === 'Medium' || undefined })),
    ),
    relatedTo:  el('select', {},
      el('option', { value: '', text: '— none —' }),
      ...['project','customer','lead','opportunity','invoice','estimate','contract','ticket','expense','proposal']
        .map((t) => el('option', { value: t, text: t })),
    ),
    relatedQuery: el('input', { type: 'search', placeholder: 'Search ERP records...' }),
    assignees:  el('select', { multiple: true, size: 4 },
      ...refdata.staff.map((s) => el('option', { value: s.id, text: s.name })),
    ),
    tags:       el('input', { type: 'text', placeholder: 'tag1, tag2' }),
    description:el('textarea', { text: snap.bodyExcerpt || '' }),
    attachEmail:el('input', { type: 'checkbox', checked: true }),
    attachFiles:el('input', { type: 'checkbox', checked: snap.attachments?.length ? true : false }),
  };

  const submitBtn = el('button', { class: 'btn btn--block', type: 'submit' }, 'Create task in Prizm ERP');
  const cancelBtn = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  const status = el('div', {});

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      status.replaceChildren();
      submitBtn.disabled = true;
      submitBtn.textContent = 'Creating...';

      const payload = {
        subject: inputs.subject.value.trim(),
        startDate: inputs.startDate.value,
        dueDate: inputs.dueDate.value,
        priority: inputs.priority.value,
        relatedTo: inputs.relatedTo.value || null,
        relatedQuery: inputs.relatedQuery.value || null,
        assignees: Array.from(inputs.assignees.selectedOptions).map((o) => o.value),
        tags: inputs.tags.value.split(',').map((t) => t.trim()).filter(Boolean),
        description: inputs.description.value,
        email: Office.envelope(snap, {
          attachEmailAsEml: inputs.attachEmail.checked,
          attachFiles:      inputs.attachFiles.checked,
        }),
      };

      if (!payload.subject) {
        status.replaceChildren(banner('err', 'Subject is required.'));
        submitBtn.disabled = false;
        submitBtn.textContent = 'Create task in Prizm ERP';
        return;
      }

      try {
        const result = await Api.createTask(payload);
        const taskUrl = result?.url || (result?.id ? `${Config.get('erpBase')}/admin/tasks/view/${result.id}` : null);
        const msg = taskUrl
          ? `Task #${result.id} created. Open: ${taskUrl}`
          : 'Task created in Prizm ERP.';
        status.replaceChildren(banner('ok', msg));
        if (taskUrl) {
          status.appendChild(el('div', { style: { marginTop: '6px' } },
            el('a', { href: taskUrl, target: '_blank', rel: 'noopener', text: 'Open task in ERP' }),
          ));
        }
        submitBtn.textContent = 'Create another';
        submitBtn.disabled = false;
      } catch (err) {
        const msg = err instanceof ApiError
          ? `Failed (${err.status || 'net'}): ${err.message}`
          : `Failed: ${err.message}`;
        status.replaceChildren(banner('err', msg));
        submitBtn.disabled = false;
        submitBtn.textContent = 'Retry';
      }
    },
  },
    contextBanner(snap),

    field('Subject', inputs.subject, { required: true }),

    row(
      field('Start Date', inputs.startDate, { required: true }),
      field('Due Date',   inputs.dueDate),
    ),

    row(
      field('Priority',   inputs.priority),
      field('Related to', inputs.relatedTo),
    ),

    field('Related record (search)', inputs.relatedQuery, { hint: 'Type project/customer/etc name to link this task to.' }),

    field('Assignees', inputs.assignees, { hint: 'Hold Ctrl to pick multiple.' }),

    field('Tags', inputs.tags, { hint: 'Comma separated.' }),

    field('Description', inputs.description),

    el('div', { class: 'field__check' },
      inputs.attachEmail, el('label', { text: 'Attach this email as .eml to the task' }),
    ),
    el('div', { class: 'field__check' },
      inputs.attachFiles, el('label', { text: `Attach ${snap.attachments?.length || 0} email attachment(s)` }),
    ),

    el('div', { class: 'btn-row' }, cancelBtn, submitBtn),
    status,
  );

  mount(form);
}
