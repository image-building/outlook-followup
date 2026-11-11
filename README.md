# Follow-up Tracker — Outlook add-in

Works in **classic Outlook for Windows**, **new Outlook for Mac**, **Outlook for iOS/Android** and Outlook on the web, on a Microsoft 365 mailbox.

## What it does

| You do | The add-in does |
|---|---|
| Tag an email (while writing it, or afterwards from Sent Items) | Flags your email 🚩 with a due date 14 days out and puts a reminder in your calendar |
| The other person replies | Flag turns into a ✓ checkmark, reminder is removed, email moves to **Answered** |
| You send another email in that thread | Your new email gets flagged, a new 14‑day reminder is set — it's **waiting** again |
| You send a nudge while still waiting | The flag moves to your newest email and the reminder is pushed back |
| Nobody replies for 14 days | Outlook pops up a reminder at 09:00 ("Follow up: <subject>") |

Out-of-office replies, bounces and other automatic replies (English and Dutch, plus German/French) don't count as an answer.

## One-time setup (about 15 minutes)

### 1. Host the files (free, on GitHub Pages)
Outlook add-ins have to be served over HTTPS.

1. Create a new GitHub repository named `outlook-followup` (public is fine; nothing secret is in it).
2. Upload everything in this folder to it.
3. Repository **Settings → Pages** → Source: *Deploy from a branch* → `main` / root → Save.
4. Your URL will be `https://<your-github-name>.github.io/outlook-followup`.

### 2. Register the app in Microsoft Entra
1. Go to **entra.microsoft.com → Applications → App registrations → New registration**.
2. Name: `Follow-up Tracker`. Supported accounts: *Accounts in this organizational directory only*. Leave Redirect URI empty → **Register**.
3. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
4. **API permissions → Add a permission → Microsoft Graph → Delegated**: `User.Read`, `Mail.ReadWrite`, `Calendars.ReadWrite`, `MailboxSettings.ReadWrite` → then **Grant admin consent for Imagebuilding**.

### 3. Fill in your details
With Node.js installed, from this folder run:

```
node scripts/configure.js https://<your-github-name>.github.io/outlook-followup <client-id> <tenant-id>
```

It updates `manifest.xml` and `src/config.js` and prints two redirect URIs. Back in Entra: **Authentication → Add a platform → Single-page application** and add both URIs. Commit the updated files to GitHub.

(No Node? Open `manifest.xml` and `src/config.js` and replace `YOUR-HOST/outlook-followup`, the client ID and `organizations` (with your tenant ID) by hand. The redirect URIs are `brk-multihub://<your-github-name>.github.io` and `https://<your-github-name>.github.io/outlook-followup/src/auth-dialog.html`.)

### 4. Install it in Outlook
Easiest, and works for both your Mac and Windows machines at once:
- **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps** → *Office Add-in* → upload `manifest.xml` → assign it to yourself (or the people who should have it). It can take a few hours to appear.

Or for just your own mailbox, right away:
- Open **https://aka.ms/olksideload** (Outlook on the web) → **My add-ins → Add a custom add-in → Add from file** → choose `manifest.xml`. It then shows up in Outlook on Windows and Mac too.

The first time you open the add-in it asks you to sign in once.

## Daily use
- **New email:** in the compose window click **Follow-up** (in classic Windows: the ribbon; on Mac: the *Apps* / `…` menu) and tick *Track this email for follow-up after sending*.
- **Already sent:** open it in Sent Items → **Follow-up** → *Track for follow-up*.
- **See everything:** open the Follow-up pane on any email. Click the **pin** icon to keep it open — it then checks for replies automatically while you read mail.
- **iPhone:** open an email (e.g. in Sent Items) → tap the **Apps** icon (or `…`) → **Follow-up**. You can track/untrack and see the list; tagging while composing is desktop-only.
- **Settings** in the pane let you change 14 days and the 09:00 reminder time.

## Online overview (all mailboxes)
Open **https://image-building.github.io/outlook-followup/** in any browser (also Safari on iPhone — use *Share → Add to Home Screen* for an app icon). Sign in with each Microsoft 365 account (**+ Add account**). You get one combined list with a filter per mailbox; it checks for replies when opened and every 5 minutes while open. Tick boxes and ✕ work the same as in Outlook.

## Adding another Microsoft 365 organisation
1. Entra app → **Authentication → Supported accounts** → *Accounts in any organizational directory (multitenant)*.
2. As admin of the other organisation, open the admin-consent link (see below) in a private window and approve.
3. Install `manifest.xml` in that mailbox via https://aka.ms/olksideload (signed in as that mailbox).

Admin consent link: `https://login.microsoftonline.com/organizations/adminconsent?client_id=<client-id>&redirect_uri=https://image-building.github.io/outlook-followup/`

## Good to know
- **Replies are checked when the pane is open** (on open, on every email you click while it's pinned — at most every 3 minutes — and with *Sync*). Outlook add-ins can't run fully in the background. If a reply came in while the pane was closed, the flag updates the next time it opens; if a reminder fires for something already answered, just open the pane and it clears itself. A fully background version is possible with a small Power Automate flow or Azure Function if you want that later.
- **Why a calendar reminder instead of a flag reminder:** classic Outlook doesn't fire flag reminders for emails in Sent Items. A calendar reminder works on Windows, Mac and your phone. It's marked *Free*, so it doesn't block your agenda.
- Everything lives in your mailbox: the **Follow-up** category marks tracked conversations, the flag holds waiting/done, and your settings are stored in the mailbox too — so the add-in on every device and the web page always agree. Answered emails drop off the list after 90 days.
- Tagging while composing uses an Outlook category called **Follow-up**; the add-in creates it the first time.

## Files
```
manifest.xml           Outlook add-in manifest
src/taskpane.*         The Follow-up pane
src/sync.js            Microsoft Graph sync: flags, checkmarks, reminders
src/core.js            Reply/waiting logic (unit tested)
src/auth.js            Sign-in (Outlook's built-in sign-in, with a dialog fallback for older builds)
scripts/configure.js   Fills in your URL and app IDs
test/                  core.test.js (node) and e2e.js (Playwright, mocked Outlook + Graph)
```
