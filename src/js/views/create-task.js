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

/**
 * Open Outlook's new-message composer addressed to the task assignees,
 * pre-filled with the task instructions + a hyperlinked "Follow up task
 * has been created #N" line. The user's signature is appended automatically
 * by Outlook below our HTML.
 */
function openComposeToAssignees(result, instructions, assignees, snap) {
  const mailbox = window.Office?.context?.mailbox;
  if (!mailbox || typeof mailbox.displayNewMessageForm !== 'function') return;

  const id  = result?.id;
  const url = result?.url || '';
  const safe = (instructions || '')
    .replace(/[&<>]/g, (c) => ({ '&':'&amp;','<':'&lt;','>':'&gt;' }[c]))
    .replace(/\n/g, '<br>');
  const followLine = url
    ? `<a href="${url}">Follow up task has been created #${id}</a>`
    : `Follow up task has been created #${id}`;
  const html = `
    <p>${safe || '&nbsp;'}</p>
    <p>&mdash;</p>
    <p>${followLine}</p>
  `;

  const toRecipients = assignees
    .map((a) => a.email)
    .filter((e) => !!e && /@/.test(e));
  if (!toRecipients.length) return; // nothing to send to

  // If there's an original sender we know about, CC them for context.
  const ccRecipients = [];
  if (snap?.from?.email && /@/.test(snap.from.email)) {
    ccRecipients.push(snap.from.email);
  }

  mailbox.displayNewMessageForm({
    toRecipients,
    ccRecipients,
    subject: `Task #${id}: ${snap?.subject || 'Follow-up'}`,
    htmlBody: html,
  });
}

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
  const composeMode = snap.mode === 'compose';
  const todayIso = new Date().toISOString().slice(0, 10);
  const dueIso = new Date(Date.now() + 7 * 86400_000).toISOString().slice(0, 10);

  // Refdata is optional — failure shouldn't block the form
  let refdata = { priorities: [{id:1,name:'Low'},{id:2,name:'Medium'},{id:3,name:'High'},{id:4,name:'Urgent'}], staff: [], tags: [] };
  try {
    const r = await Api.refdata();
    if (r) refdata = { ...refdata, ...r };
  } catch (_) { /* offline-friendly fallback */ }

  const priorities  = (refdata.priorities || []).map(normalizePriority);
  const staffList   = (refdata.staff || []).map((s) => ({ id: s.id, name: s.name, email: s.email || '' }));
  const staffEmails = new Map(staffList.map((s) => [s.id, s.email]));
  const staffNames  = new Map(staffList.map((s) => [s.id, s.name]));
  const staffByEmail = new Map(
    staffList.filter((s) => s.email).map((s) => [s.email.toLowerCase(), s])
  );
  const tagList     = (refdata.tags || []).map((t) => ({ id: t, name: t }));

  // In compose mode, auto-suggest assignees from the recipients being addressed
  // (To: + CC:) matched against the active staff directory by email. So if
  // the user is forwarding to ahmad@prizm-energy.com, ahmad gets pre-selected
  // as an assignee with one click to deselect.
  const preselectedAssignees = composeMode
    ? [...(snap.to || []), ...(snap.cc || [])]
        .map((r) => staffByEmail.get((r.email || '').toLowerCase()))
        .filter(Boolean)
        .map((s) => s.id)
    : [];

  // Chip pickers — selections survive across re-renders via getSelected()
  const assigneesPicker = chipsPicker({
    options: staffList,
    placeholder: 'Search staff...',
    selected: preselectedAssignees,
  });
  const tagsPicker      = chipsPicker({ options: tagList,   placeholder: 'Search or type tag...', allowCreate: true });

  // Related-record picker — async search that switches type when "Related to" changes.
  // Empty until the user picks a related-to type.
  let relatedType = '';
  const relatedRecordPicker = asyncSearchPicker({
    placeholder: 'Select "Related to" first',
    initialList: true,        // focus → show active/recent items
    initialLimit: 10,
    expandedLimit: 100,
    search: async (q, opts) => {
      if (!relatedType) return [];
      return Api.search(relatedType, q, opts?.limit || 10);
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
    // Task description — empty by default; the user writes instructions
    // for the assignees. The email body is included separately via the
    // includeEmail checkbox below. In compose mode we pre-fill it from the
    // user's draft text so the email AND the task description stay in sync.
    description: el('textarea', {
      text: composeMode ? (snap.bodyText || '').slice(0, 2000) : '',
      placeholder: composeMode
        ? 'Pre-filled from your draft. Edit if needed — this becomes the task description.'
        : 'Task instructions for the assignees…\n(empty by default — write what you want done)',
    }),
    includeEmail:    el('input', { type: 'checkbox', checked: !composeMode }),  // already in description if compose
    attachEmail:     el('input', { type: 'checkbox', checked: !composeMode }),  // unsent draft has no .eml yet
    attachFiles:     el('input', { type: 'checkbox', checked: true }),     // always on by default
    notifyAssignees: el('input', { type: 'checkbox', checked: false }),
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

  const submitBtn  = el('button', { class: 'btn',         type: 'submit' }, 'Create task in Prizm ERP');
  const cancelBtn  = el('button', { class: 'btn btn--ghost', type: 'button', onclick: () => window.__prizmGo('/home') }, 'Cancel');
  // Primary on the left, secondary on the right (user preference).
  const actionRow  = el('div',    { class: 'btn-row btn-row--split' }, submitBtn, cancelBtn);
  const status     = el('div',    {});

  function fmtWhen(iso) {
    if (!iso) return new Date().toLocaleString();
    const d = new Date(iso);
    return Number.isNaN(d.getTime()) ? iso : d.toLocaleString();
  }

  function renderSuccess(result) {
    const id = result?.id;
    const url = result?.url || (id ? `${Config.get('erpBase')}/admin/tasks/view/${id}` : null);
    const when = fmtWhen(result?.created_at);
    const who  = result?.created_by_name || result?.created_by || 'you';

    status.replaceChildren(banner('ok', `Task #${id} created on ${when} by ${who}.`));

    // Replace the action row with View task / Create another 50-50 split.
    // Primary action (View in ERP) on the LEFT to match the form layout.
    const viewBtn = el('button', {
      class: 'btn',
      type: 'button',
      onclick: () => window.open(url, '_blank', 'noopener'),
      disabled: !url || undefined,
    }, 'View task in ERP');

    const againBtn = el('button', {
      class: 'btn btn--ghost',
      type: 'button',
      onclick: () => window.__prizmGo('/create-task'), // re-renders from scratch
    }, 'Create another');

    actionRow.replaceChildren(viewBtn, againBtn);
  }

  // Friendly explainer at the top of the form when we're in compose mode.
  const composeBanner = composeMode
    ? banner('info',
        `Drafting mode — task description is pre-filled from your draft, `
        + `and ${preselectedAssignees.length} recipient(s) matching active staff `
        + `have been pre-selected as assignees. Edit anything; submit to log the task.`)
    : null;

  const form = el('form', {
    onsubmit: async (e) => {
      e.preventDefault();
      status.replaceChildren();
      submitBtn.disabled = true;
      submitBtn.textContent = 'Creating...';

      const related = relatedRecordPicker.get();

      // Description = user instructions + (optionally) the email body
      const taskInstructions = (inputs.description.value || '').trim();
      const emailBody = inputs.includeEmail.checked ? (snap.bodyExcerpt || '') : '';
      const combinedDescription = [
        taskInstructions,
        emailBody ? `\n\n---\nOriginal email:\n${emailBody}` : '',
      ].join('').trim();

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
        description: combinedDescription,
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

        // Optional: open a new compose addressed to the assignees, pre-filled
        // with the task instructions + a hyperlinked task ID. Outlook
        // auto-appends the user's signature.
        if (inputs.notifyAssignees.checked) {
          try {
            const assigneeIds = assigneesPicker.getSelected();
            const assignees = assigneeIds.map((id) => ({
              id, email: staffEmails.get(id) || '', name: staffNames.get(id) || '',
            })).filter((a) => a.email);
            openComposeToAssignees(result, taskInstructions, assignees, snap);
          } catch (_) { /* non-fatal */ }
        }

        renderSuccess(result);
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
    composeBanner,
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

    field('Description', inputs.description, { hint: composeMode
      ? 'Captured from your draft. Edit freely — this is what the task will store.'
      : 'What needs to be done. This shows on the task and (if forwarded) in the reply body.' }),

    // Only show "append email body" when in read mode — in compose mode the
    // description IS the draft body already.
    composeMode ? null : el('div', { class: 'field__check' },
      inputs.includeEmail, el('label', { text: 'Append the original email body to the task description' }),
    ),
    // .eml capture only makes sense for existing items, not unsent drafts.
    composeMode ? null : el('div', { class: 'field__check' },
      inputs.attachEmail, el('label', { text: 'Attach the email as .eml file to the task' }),
    ),
    // Email attachments are only readable in read mode — compose mode needs
    // a different API call (getAttachmentsAsync) that isn't widely supported.
    composeMode ? null : el('div', { class: 'field__check' },
      inputs.attachFiles, el('label', { text: `Attach the ${snap.attachments?.length || 0} email attachment(s) to the task` }),
    ),
    // In compose mode, the user is already sending an email — no second email needed.
    composeMode ? null : el('div', { class: 'field__check' },
      inputs.notifyAssignees,
      el('label', { text: 'Email the assignees with the task details + "Follow up task #" link (signature added by Outlook)' }),
    ),
    composeMode ? el('div', { class: 'field__hint', style: { marginTop: '6px' } },
      'Tip: after creating the task, finish your draft and hit Send in Outlook — the task ID is already in the description so anyone replying will reference it.'
    ) : null,

    actionRow,
    status,
  );

  mount(form);
}
