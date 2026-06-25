# Connecting the Master List + Search History to Google Sheets

The app keeps a **running master list** of every unique rep (deduped by LinkedIn
URL) and a **log of every search** (with new-vs-duplicate counts). To make that
data survive forever — even when the Streamlit app sleeps or redeploys — it's
stored in a Google Sheet you own.

Until this is configured, the app still works: searches run normally, but the
master list and history only last for the current browser session.

You only have to do this **once**. Takes about 10 minutes.

---

## Step 1 — Create the Google Sheet

1. Go to <https://sheets.google.com> and create a new blank spreadsheet.
2. Name it something like **"Sales Rep Sourcing — Master"**.
3. Look at the URL. The long ID between `/d/` and `/edit` is your **Sheet ID**:
   ```
   https://docs.google.com/spreadsheets/d/THIS_LONG_PART_IS_THE_ID/edit
   ```
   Copy it — you'll need it in Step 4.

> You don't need to add any tabs or headers. The app creates the **Master** and
> **Searches** tabs automatically on the first run.

---

## Step 2 — Create a Google service account (the app's robot login)

1. Go to <https://console.cloud.google.com/> and sign in.
2. Create a new project (top bar → project dropdown → **New Project**). Name it
   anything, e.g. `sales-rep-sourcing`.
3. In the search bar, find and **enable** these two APIs (one at a time):
   - **Google Sheets API**
   - **Google Drive API**
4. Go to **APIs & Services → Credentials → Create Credentials → Service account**.
   - Give it a name (e.g. `sourcing-app`), click **Create and Continue**, then
     **Done** (you can skip the optional role steps).
5. Click the new service account → **Keys** tab → **Add Key → Create new key →
   JSON**. A `.json` file downloads. Keep it safe — this is a password.

---

## Step 3 — Share the Sheet with the service account

1. Open the downloaded JSON file. Find the `"client_email"` value — it looks like
   `sourcing-app@your-project.iam.gserviceaccount.com`.
2. Back in your Google Sheet, click **Share** and paste that email in. Give it
   **Editor** access and send. (No notification needed.)

This is the step people forget — if the app can't read the sheet, it's almost
always because the sheet wasn't shared with this email.

---

## Step 4 — Put the credentials into Streamlit secrets

On Streamlit Community Cloud: open your app → **⋮ Manage app → Settings →
Secrets**, and paste the block below. (For local runs, put the same thing in
`.streamlit/secrets.toml` — that file is git-ignored so it won't be committed.)

Copy the matching values straight out of the downloaded JSON file:

```toml
APIFY_TOKEN = "your-apify-token-here"   # you already have this
MASTER_SHEET_ID = "the-long-id-from-step-1"

[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\nMIIE...long...key\n-----END PRIVATE KEY-----\n"
client_email = "sourcing-app@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
universe_domain = "googleapis.com"
```

> **Important about `private_key`:** keep it on one line with the `\n` characters
> exactly as they appear in the JSON. Don't reformat it.

---

## Step 5 — Restart and confirm

Save the secrets (the app reboots automatically). Run one search. You should see:

- The yellow "⚠️ Google Sheet not connected" banner is **gone**.
- The **Previous Searches** tab shows your run.
- The **Master List** tab fills with the reps.
- Your Google Sheet now has **Master** and **Searches** tabs with the data.

That's it — every future search adds new reps to the master and logs the run.
