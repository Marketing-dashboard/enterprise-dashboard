# Enterprise Performance Dashboard

Daily tracker for model-level Spends, Leads, Triggered Leads, CPL, and TCPL.  
Data auto-refreshes from Google Sheets via Apps Script → GitHub push.

---

## One-time Setup

### 1. Create GitHub Repo
Create a public (or private with Pages enabled) repo named `enterprise-dashboard` under your GitHub account.

### 2. Push this code
```
git remote add origin https://github.com/deepanshiahuja-dotcom/enterprise-dashboard.git
git push -u origin main
```

### 3. Install the Apps Script
1. Open the Google Sheet → **Extensions → Apps Script**
2. Paste the contents of `appscript/Code.gs`
3. Save (Ctrl+S)

### 4. Store your GitHub token
In Apps Script editor, run once in the console:
```js
setGitHubToken('ghp_your_personal_access_token_here')
```
Token needs: `repo` scope (Contents read+write).

### 5. Set up triggers
Run once:
```js
createTriggers()
```
This creates:
- **onEdit** — updates dashboard whenever Raw_Pannel or Raw_triggers is edited
- **Daily 7 AM** — backup refresh

### 6. First run
Run `updateDashboard()` manually once to populate the dashboard.

---

## Admin credentials
| Email | Password |
|---|---|
| deepanshi.ahuja@girnarsoft.com | Admin@2024 |
| neeraj.khichi@girnarsoft.com | Admin@2024 |

Change passwords directly in `index.html` → `const ADMINS = [...]`

---

## Local preview
```
python server.py
```
Opens at http://localhost:3030

---

## Data sources
| Sheet tab | Purpose |
|---|---|
| Raw_Pannel | Daily spends & leads by campaign/brand/model |
| Raw_triggers | Triggered leads (Triggered + List_ID for JLR & Citroen) |
