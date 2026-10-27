# Security

Whatever is on `main` runs inside Outlook within about a minute, with access to the open mail, the mailbox (through Microsoft Graph, once the Entra app ID is set), the calendar and the Anthropic API key. Everything below exists to keep that code trustworthy.

## How changes go live

1. Every change is made on a branch and merged through a pull request. Nothing is uploaded straight to `main`.
2. Before merging, the branch is pushed to the **test repo** (`yanick-spec/cos-outlook-addin-test`, GitHub Pages on its `main`) and tried in Outlook with the **CoS Drafts (Test)** button from `manifest-test.xml`.
3. The pull request diff is scanned for secrets before merging.
4. Only then is the pull request merged into `main`, which is what the live button loads.

Install the test add-in once: sideload `manifest-test.xml` the same way as `manifest.xml`. It has its own ID, so it keeps its own settings: enter the API key (and Entra app ID) in its Settings too.

Push a branch to the test repo (once: `git remote add test https://github.com/yanick-spec/cos-outlook-addin-test.git`):

    git push test HEAD:main --force

## What protects the add-in

- **No secrets in the repo.** The API key and the Entra app ID are typed into Settings and stored in the mailbox's add-in settings. Commits use the GitHub no-reply address.
- **Content Security Policy** (`taskpane.html`, `commands.html`). The pane may only run code from this site and `appsforoffice.microsoft.com`, and may only connect to Microsoft (Graph, sign-in, Office) and `api.anthropic.com`. All script lives in `.js` files, so no inline script is allowed. If you add a service, add its address to `connect-src`, and only then.
- **Drafts only.** The add-in never sends mail and Claude gets no tools: it can only return text, which you review. Keep it that way. Email content is untrusted: the prompt marks it as data, and model output is escaped before it goes into the mail.
- **Escaping.** Anything from a mail, a file name or the model goes into the page through `textContent` or `esc()`, never raw into `innerHTML`.
- **Minimal Microsoft permissions.** Graph scopes are `Calendars.ReadBasic` and `Mail.Read`. Don't add write or send scopes. In the Entra app registration, allow only this add-in's redirect URI.
- **Anthropic key.** A dedicated key with a monthly spend limit, so it can be revoked on its own.

## Pinned libraries

Third-party code is copied into `vendor/` rather than loaded from a CDN, so it can't change under us. When updating one, take it from the official npm package, record the new version and checksum here, and say so in the pull request.

| File | Version | SHA-256 |
|---|---|---|
| `vendor/msal-browser.min.js` | @azure/msal-browser 5.24.0 | `09a9a805e9f6cd80808853ba176502c25760c23526c75e317350f7b68c0100cf` |
| `vendor/jszip.min.js` | jszip 3.10.1 | `f12f367798e35ee2d9993dba6167fc61ddb52fb89880f5a99fbb606335188410` |

Check on Windows: `Get-FileHash vendor\*.js -Algorithm SHA256`.

`office.js` is loaded from Microsoft's CDN as Microsoft requires, and can't be pinned.

## Data

Mail text, the conversation, attachments and calendar free/busy windows are sent to the Anthropic API to write the draft. Check with eCAPITAL that this is OK for deal correspondence.
