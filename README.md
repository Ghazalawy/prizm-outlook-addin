# Prizm Energy — Outlook Add-in (v2)

A single Outlook add-in that surfaces multiple Prizm ERP actions in one clean
submenu on every email — modeled after the way Asana, Microsoft Teams and
OneNote show up in Outlook's overflow menu.

![menu pattern](docs/menu-preview.png)

## Why this exists

The previous two add-ins (`Prizm Task Creation`, `Prizm Energy Opportunity Add-in`)
were webpack-bundled jQuery/Select2 single-button task panes. They:

- couldn't be extended without re-bundling and re-sideloading
- had one button = one form, so each new ERP action meant a new add-in
- shipped minified, unreadable JS into prod

This v2 replaces both with **one** add-in that:

- shows a **Menu** in Outlook (read + compose) listing every ERP action
- is plain modular ES modules (no webpack, no jQuery, no Select2)
- adds a new action in two files (one manifest `<Item>` + one view module)
- is hosted on **GitHub Pages** with **GitHub Actions auto-deploy** → pushing
  to `main` updates the live add-in in Outlook **without re-sideloading**

## How "live deploy without re-sideloading" works

Outlook downloads the manifest XML **once** at sideload time. After that, every
time the user opens the task pane, Outlook fetches:

- `taskpane.html` from `https://ghazalawy.github.io/prizm-outlook-addin/taskpane.html`
- all CSS and JS modules it references
- icons from `/assets/...`

The manifest URLs are **stable** — they never change. So as long as you only
change the HTML/CSS/JS, **the file Outlook fetches next time is your new file**.
The CI pipeline (`.github/workflows/deploy.yml`) stamps a unique cache-buster
into every HTML/JS/CSS reference (`?v=<commit-sha>-<timestamp>`) so the WebView
inside Outlook can't serve a stale cached copy.

You only need to re-sideload the manifest when **the manifest itself changes**,
e.g.:

- you add a new menu Item
- you change a label or icon
- you add a new permission

For pure code/style/UX changes the loop is just: `git push origin main` → wait
~30 seconds for GitHub Actions → close and reopen the Outlook task pane.

## Repo layout

```
.
├── manifest.xml                # The only thing sideloaded into Outlook
├── src/
│   ├── taskpane.html           # SPA shell, hash-routed
│   ├── commands.html           # Reserved for function commands
│   ├── styles/taskpane.css
│   ├── assets/                 # Icons + logo
│   └── js/
│       ├── app.js              # Entry: Office.ready -> register routes -> start
│       ├── router.js           # Hash router
│       ├── ui.js               # DOM helpers (el, field, row, banner...)
│       ├── office.js           # Office.js wrappers + email snapshot
│       ├── api.js              # ERP API client
│       ├── config.js           # erpBase / apiBase / apiKey
│       ├── commands.js
│       └── views/
│           ├── home.js                # Grid of all actions
│           ├── create-task.js
│           ├── create-opportunity.js
│           ├── create-lead.js
│           ├── create-ticket.js
│           ├── link-record.js         # Attach email to existing record
│           ├── lookup-sender.js
│           └── settings.js
└── .github/workflows/deploy.yml
```

## Adding a new ERP action

Two files. That's it.

1. **Manifest** — add an `<Item>` inside the existing `Items` of the menu
   (and a matching `<bt:String>` + `<bt:Url>`):

   ```xml
   <Item id="readCreateExpense">
     <Label resid="CreateExpense.Label"/>
     ...
     <Action xsi:type="ShowTaskpane">
       <SourceLocation resid="Taskpane.CreateExpense.Url"/>
     </Action>
   </Item>
   ```
   ```xml
   <bt:Url id="Taskpane.CreateExpense.Url"
           DefaultValue="https://ghazalawy.github.io/prizm-outlook-addin/taskpane.html#/create-expense"/>
   <bt:String id="CreateExpense.Label" DefaultValue="Create Expense"/>
   ```

2. **View** — `src/js/views/create-expense.js`, following the same shape as
   `create-task.js`. Register it in `src/js/app.js`:

   ```js
   registerRoute('/create-expense', () => import('./views/create-expense.js'),
                 { title: 'Create Expense' });
   ```

After step 2 your action is live in 30s. Step 1 needs a one-time re-sideload
because Outlook caches the manifest.

## ERP backend contract

The backend lives in the Perfex CRM `outlookapi` module, controller
[`Bridge.php`](https://github.com/Ghazalawy/prizm331/blob/main/modules/outlookapi/controllers/Bridge.php).

Default `apiBase` is `https://ms.prizm-energy.com/MS/outlookapi/bridge` (the
Hetzner production deploy from `PrizmIT/prizm331`). Override per browser/profile
in the add-in's Settings view to use dev:

| Env  | ERP base                                | API base                                                |
|------|-----------------------------------------|---------------------------------------------------------|
| Prod | `https://ms.prizm-energy.com/MS`        | `https://ms.prizm-energy.com/MS/outlookapi/bridge`      |
| Dev  | `https://dev.prizm-energy.com`          | `https://dev.prizm-energy.com/outlookapi/bridge`        |

| Method | Path             | Used by                                |
|--------|------------------|----------------------------------------|
| GET    | `/ping`          | Settings · Test connection             |
| GET    | `/refdata`       | Create Task (staff / priorities / tags)|
| GET    | `/search`        | Link to Record                         |
| GET    | `/lookup`        | Lookup Sender                          |
| POST   | `/tasks`         | Create Task                            |
| POST   | `/opportunities` | Create Opportunity                     |
| POST   | `/leads`         | Create Lead                            |
| POST   | `/tickets`       | Create Ticket                          |
| POST   | `/link`          | Link to Record                         |

### Auth

`Authorization: Bearer <api_key>` on every request. Each Perfex staff user
generates their own key from **Admin sidebar → Outlook → Add-in API keys**:

- Prod: <https://ms.prizm-energy.com/MS/admin/outlookapi/keys>
- Dev:  <https://dev.prizm-energy.com/admin/outlookapi/keys>

The key is shown **once** on creation; revoke or regenerate any time.

### Attachment handling

When you tick *"Attach this email as .eml"* or *"Attach N email attachment(s)"*
on any Create / Link view, the ERP fetches the bytes directly from Microsoft
Graph (using the existing app-only token from the `outlookapi` module) and
saves them under the matching Perfex uploads folder
(`uploads/tasks/<id>/`, `uploads/leads/<id>/`, `uploads/ticket_attachments/<id>/`,
`uploads/projects/<id>/`, `uploads/opportunities/<id>/`). The files appear in
the record's normal *Attachments* tab — same as if you'd uploaded them by hand.

Files never transit the Outlook → Pages → ERP path. The add-in only sends
metadata `{id, name, size, contentType, mailbox, itemId}`; the bytes are
pulled server-to-server.

All POSTs accept a JSON body with form fields plus an `email` envelope:
```json
{
  "subject": "...",
  "...form fields...": "...",
  "email": {
    "itemId":            "...",
    "internetMessageId": "...",
    "conversationId":    "...",
    "from":              {"name":"","email":""},
    "to":                [{"name":"","email":""}],
    "attachEmailAsEml":  true,
    "attachments":       [{"id":"","name":"","size":0,"contentType":""}]
  }
}
```

Auth: `Authorization: Bearer <apiKey>` if `apiKey` is set in Settings.

## Sideloading (one-time)

Microsoft disabled **"Add from URL"** for custom Outlook add-ins in early
2026 (anti-phishing). Sideload is now download-then-upload:

1. Visit the install page: **<https://ghazalawy.github.io/prizm-outlook-addin/>** → click *Download manifest*
2. In Outlook (web or desktop): *Get add-ins → My add-ins → Add a custom add-in → **Add from File...*** → pick the manifest you just downloaded
3. Confirm the consent prompt

The raw manifest is still at `https://ghazalawy.github.io/prizm-outlook-addin/manifest.xml`
for admin tools that support manifest URLs (e.g. Microsoft 365 Admin Center →
Integrated apps → Upload custom apps).

## Local development

Open `src/taskpane.html` over any static HTTP server — Outlook isn't required
to develop the UI:

```bash
cd src
python3 -m http.server 8000
# open http://localhost:8000/taskpane.html#/create-task
```

The `Office` helper falls back to stub data when not running inside Outlook.

## Deploy pipeline

`main` branch → GitHub Actions workflow:

1. Copies `src/` into `dist/`
2. Copies `manifest.xml` next to it (so the manifest is also live-hosted)
3. Stamps a `?v=<sha>-<timestamp>` cache buster into every HTML/CSS/JS reference
4. Writes `.nojekyll`
5. Publishes `dist/` to GitHub Pages

Manual run: Actions tab → "Deploy add-in to GitHub Pages" → Run workflow.

## License

Internal Prizm Energy use.
