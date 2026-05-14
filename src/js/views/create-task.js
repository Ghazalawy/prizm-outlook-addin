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
import { el, mount, field, row, banner, contextBanner, chipsPicker, asyncSearchPicker } from '../ui.js';

function gotoSettings() { window.__prizmGo('/settings'); }

// refdata.priorities returns [{id,name,color}] from Perfex's get_tasks_priorities();
// refdata.staff returns [{id,name}] from Outlookapi_model::getstaff;
// refdata.tags returns ['name', ...] (just names from tbltags).
function normalizePriority(p) {
  if (p && typeof p === 'object') return { id: p.id, name: p.name };
  return { id: p, name: String(p) };
}

export async function render() {
  // Pre-flight: API key required for any create flow.
  if (!Config.get('apiKey')) {
    mount(el('div', {},
      banner('err', 'No API key set. Open Settings to paste your key from Prizm ERP.'),
      el('div', { class: 'btn-row' },
        el('button', { class: 'btn', type: 'button', onclick: gotoSettings }, 'Open Settings'),
      ),
    ));
    return;
  }

  // Loading state
  mount(el('div', {}, banner('info', 'Loading email context...')));

  const snap = await Office.snapshot();
  const todayIso = new Date().toISOString().slice(0, 10);
  const dueIso = new Date(Date.now() + 7 * 86400_000).toISOString().slice(0, 10);

  // Refdata is optional — failure shouldn't block the form
  let refdata = { priorities: [{id:1,name:'Low'},{id:2,name:'Medium'},{id:3,name:'High'},{id:4,name:'Urgent'}], staff: [], tags: [] };
  try {
    const r = await Api.refdata();
    if (r) refdata = { ...refdata, ...r };
  } catch (_) { /* offline-friendly fallback */ }

  const priorities = (refdata.priorities || []).map(normalizePriority);
  const staffList  = (refdata.staff || []).map((s) => ({ id: s.id, name: s.name }));
  const tagList    = (refdata.tags || []).map((t) => ({ id: t, name: t }));

  // Chip pickers — selections survive across re-renders via getSelected()
  const assigneesPicker = chipsPicker({ options: staffList, placeholder: 'Search staff...' });
  const tagsPicker      = chipsPicker({ options: tagList,   placeholder: 'Search or type tag...', allowCreate: true });

  // Related-record picker — async search that switches type when "Related to" changes.
  // Empty until the user picks a related-to type.
  let relatedType = '';
  const relatedRecordPicker = asyncSearchPicker({
    placeholder: 'Select "Related to" first',
    search: async (q) => {
      if (!relatedType) return [];
      return Api.search(relatedType, q);
    },
  });
  // Wrap so we can hide the whole row when Related to is empty.
  const relatedRecordRow = el('div', { class: 'field', hidden: true });

  const inputs = {
    subject:    el('input', { type: 'text', value: snap.subject || '' }),
    startDate:  el('input', { type: 'date', value: todayIso }),
    dueDate:    el('input', { type: 'date', value: dueIso }),
    priority:   el('select', {},
      ...priorities.map((p) => el('option', {
        value: p.id,
        text: p.name,
        selected: (p.id === 2 || /medium/i.test(p.name)) || undefined,
      })),
    ),
    relatedTo:  el('select', {},
      el('option', { value: '', text: '— none —' }),
      ...['project','customer','lead','opportunity','invoice','estimate','contract','ticket','expense','proposal']
        .map((t) => el('option', { value: t, text: t })),
    ),
    description: el('textarea', { text: snap.bodyExcerpt || '' }),
    attachEmail: el('input', { type: 'checkbox', checked: true }),
    attachFiles: el('input', { type: 'checkbox', checked: snap.attachments?.length ? true : false }),
  };

  // Populate the related-record row with the picker; toggle on relatedTo change.
  function refreshRelatedRow() {
    const type = inputs.relatedTo.value;
    relatedType = type;
    relatedRecordPicker.clear();
    if (!type) {
      relatedRecordRow.hidden = true;
      relatedRecordRow.replaceChildren();
      return;
    }
    const pretty = type.charAt(0).toUpperCase() + type.slice(1);
    relatedRecordPicker.setPlaceholder(`Search ${pretty}s…`);
    relatedRecordRow.hidden = false;
    relatedRecordRow.replaceChildren(
      el('label', {}, `Pick a ${pretty}`),
      relatedRecordPicker.node,
      el('div', { class: 'field__hint', text: `Search by name. Required to link this task to a ${type}.` }),
    );
  }
  inputs.relatedTo.addEventListener('change', refreshRelatedRow);

  const submitBtn = el('button', { class: 'btn btn--block', type: 'submit' }, 'Create task in Prizm ERP');
  const cancelBtn = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  const status = el('div', {});

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      status.replaceChildren();
      submitBtn.disabled = true;
      submitBtn.textContent = 'Creating...';

      const related = relatedRecordPicker.get();
      const payload = {
        subject: inputs.subject.value.trim(),
        startDate: inputs.startDate.value,
        dueDate: inputs.dueDate.value,
        priority: Number(inputs.priority.value) || 2,
        relatedTo: inputs.relatedTo.value || null,
        relatedId: related?.id ?? null,
        relatedLabel: related?.label ?? null,
        assignees: assigneesPicker.getSelected(),
        tags: tagsPicker.getSelected(),
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
        if (err instanceof ApiError && err.status === 401) {
          status.appendChild(el('div', { class: 'btn-row' },
            el('button', { class: 'btn btn--ghost', type: 'button', onclick: gotoSettings }, 'Fix API key in Settings'),
          ));
        }
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

    relatedRecordRow,

    field('Assignees', assigneesPicker.node, { hint: 'Type to filter staff. Click to add, × to remove.' }),

    field('Tags', tagsPicker.node, { hint: 'Type to filter existing tags or pick from list.' }),

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
