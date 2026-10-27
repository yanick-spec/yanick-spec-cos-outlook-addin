# CoS Drafts – Outlook add-in

A "Draft with CoS" button in classic Outlook (also works in new Outlook / OWA).

- **Reading a mail:** click the button → optional instruction or quick chip (Zusagen, Absagen, Calendly, …) → Claude writes the reply in your CoS style → **Open as reply / Reply all** opens a normal Outlook reply with the text in it, signature and quoted thread underneath. You edit and send yourself.
- **Inside a reply window:** same button → **Insert into mail** puts the text at the top of the draft. Notes you've already typed are used as input. Select your notes before clicking **Insert into mail** and the draft replaces them.

Nothing is ever sent by the add-in. It has no server: the pane calls the Claude API straight from Outlook with your own API key.

## Files

| File | What it is |
|---|---|
| `manifest.xml` | What you load into Outlook (button, permissions, URLs) |
| `manifest-test.xml` | Same add-in as "CoS Drafts (Test)", loaded from the test repo (see `SECURITY.md`) |
| `taskpane.html` | The side pane (layout and styles) |
| `taskpane.js` | The side pane's logic |
| `commands.html`, `commands.js` | Required stub for the ribbon button |
| `vendor/` | Pinned copies of MSAL (Microsoft sign-in) and JSZip (reading Office files) |
| `assets/` | Icons |
| `SECURITY.md` | How changes are made and tested, and what protects the add-in |

## Setup (about 15 minutes)

### 1. Host the files (GitHub Pages)
Outlook needs to load the pane from an HTTPS URL. There are no secrets in these files, so a public repo is fine.

1. The live add-in is served from the public repo `yanick-spec/yanick-spec-cos-outlook-addin`.
2. Repo → Settings → Pages → Source: *Deploy from a branch*, branch `main`, folder `/ (root)`.
3. After about a minute, check that `https://yanick-spec.github.io/yanick-spec-cos-outlook-addin/taskpane.html` loads.

Using a different repo name or host? Search and replace `https://yanick-spec.github.io/yanick-spec-cos-outlook-addin` in `manifest.xml`.

### 2. Get an API key
console.anthropic.com → API Keys → create a key (e.g. "Outlook CoS"). Set a monthly spend limit there. One draft costs about a cent.

### 3. Sideload into Outlook
**Classic Outlook (Windows):** Home → *Get Add-ins* (or File → Manage Add-ins) → *My add-ins* → *Custom Addins* → *Add a custom add-in* → *Add from file…* → pick `manifest.xml` → Install.
If that option isn't there, open https://aka.ms/olksideload in the browser (Outlook on the web) and do the same there. It syncs to classic Outlook after a restart.

If both are blocked, eCAPITAL IT has turned off custom add-ins. They can deploy the same `manifest.xml` just for you in the M365 admin center (Settings → Integrated apps → Upload custom apps).

### 4. First run
Open any mail → **Draft with CoS** in the ribbon (group "Chief of Staff") → Settings opens automatically → paste the API key → Save. The key is stored in your mailbox's add-in settings, so it follows you to other Outlook clients.

## Customising
- **Style rules:** Settings → Style rules. Paste learned rules from the CoS Coach changelog here. *Reset style rules* restores the defaults.
- **Model:** default is `claude-opus-5-5`. A model saved earlier in Settings stays until you change it there. Any current model ID from the Anthropic docs works.
- **Quick chips:** edit the `CHIPS` list in `taskpane.js` (through a pull request, see `SECURITY.md`). Outlook picks the change up on the next open, no re-install needed.

## Good to know
- The mail text (subject, sender, recipients, body including quoted history, up to 40k characters) is sent to the Anthropic API. Check with eCAPITAL that this is OK for deal correspondence, as with the CoS itself.
- Whole conversation, attachments and calendar need the Entra app (client) ID in Settings. Without it the add-in uses only the open mail, and the pane says so. Without calendar data it leaves `[Slot wählen]` for you.
- Drafts made here don't create Notion action items, and the CoS Coach doesn't see them.
- Only change `manifest.xml` (and re-sideload) if you change URLs, buttons or permissions. Changes to the pane go live as soon as they're merged into `main`.
