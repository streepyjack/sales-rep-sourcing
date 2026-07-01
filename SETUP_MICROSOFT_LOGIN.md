# Requiring Microsoft (Albireo) Login to Access the App

The app can require every user to sign in with their **Albireo Microsoft/Outlook
account** before they can see anything. Access is restricted to
`@albireoenergy.com` accounts.

This uses Microsoft Entra ID (formerly Azure AD) as the login provider. Setup is
one-time and has two halves: register the app in Microsoft (Steps 1–3), then paste
the credentials into Streamlit (Step 4).

> ⚠️ **You may need IT/admin help.** Many Albireo tenants only let administrators
> register apps or grant consent. If a button below is greyed out or you get a
> "need admin approval" message, forward these steps to your IT admin.
>
> Until this is finished, the app stays open (no login) — nothing breaks while you set it up.

---

## Step 1 — Register the app in Microsoft Entra ID

1. Go to **https://portal.azure.com** and sign in with your Albireo account.
2. In the top search bar, type **App registrations** and open it.
3. Click **+ New registration**.
4. Fill in:
   - **Name:** `Sales Rep Sourcing`
   - **Supported account types:** choose **"Accounts in this organizational
     directory only (Albireo Energy only – Single tenant)"** — this is what limits
     it to Albireo staff.
   - **Redirect URI:** set the dropdown to **Web**, and paste:
     ```
     https://sales-rep-sourcing.streamlit.app/oauth2callback
     ```
5. Click **Register**.

---

## Step 2 — Copy the IDs and create a secret

On the app's **Overview** page (where you land after registering):

1. Copy the **Application (client) ID** — save it.
2. Copy the **Directory (tenant) ID** — save it.

Then create a client secret:

3. In the left menu, click **Certificates & secrets**.
4. Under **Client secrets**, click **+ New client secret**.
5. Description: `streamlit`; Expires: **24 months** (or your policy). Click **Add**.
6. **Immediately copy the `Value`** (not the "Secret ID"). You can't see it again
   after you leave the page — save it now.

---

## Step 3 — (Usually automatic) permissions

By default the registration includes the sign-in permissions we need
(`openid`, `profile`, `email`, `User.Read`). You normally don't have to change
anything here. If sign-in later fails with a permissions error, go to
**API permissions → + Add a permission → Microsoft Graph → Delegated permissions**
and add `openid`, `profile`, `email`, `User.Read`, then **Grant admin consent**.

---

## Step 4 — Put the credentials into Streamlit

Open your app on **https://share.streamlit.io** → your app → **⋮ Settings →
Secrets**, and **add** this block (keep everything already there, like
`APIFY_TOKEN`, `MASTER_SHEET_ID`, `gcp_service_account_json`):

```toml
[auth]
redirect_uri = "https://sales-rep-sourcing.streamlit.app/oauth2callback"
cookie_secret = "PASTE_THE_COOKIE_SECRET_BELOW"
client_id = "PASTE_APPLICATION_CLIENT_ID_FROM_STEP_2"
client_secret = "PASTE_CLIENT_SECRET_VALUE_FROM_STEP_2"
server_metadata_url = "https://login.microsoftonline.com/PASTE_TENANT_ID/v2.0/.well-known/openid-configuration"
```

Notes on filling it in:
- **cookie_secret** — use this randomly generated value (already made for you):
  ```
  212e31e89c16f6850c4d6be5d1fb4df80781f5d8e929e9b256043599ac8e8727
  ```
- **client_id** — the *Application (client) ID* from Step 2.
- **client_secret** — the secret **Value** from Step 2.
- **server_metadata_url** — replace `PASTE_TENANT_ID` with the *Directory (tenant) ID*
  from Step 2 (keep the rest of the URL exactly).

Click **Save**. The app reboots (~1 minute).

---

## Step 5 — Confirm

1. Open the app. You should now see a **"Sign in with Microsoft"** button instead
   of the tool.
2. Click it, sign in with your Albireo account, approve if prompted.
3. You should land back in the app, with **"🔐 Signed in as you@albireoenergy.com"**
   and a **Log out** button at the top.
4. Try signing in with a non-Albireo account (e.g., a personal Outlook) — it should
   be **refused** with an access-restricted message.

That's it — the app now requires an Albireo login for everyone.

> To allow accounts from more than one domain, or to change the allowed domain,
> tell me and I'll adjust `ALLOWED_EMAIL_DOMAIN` in the code.
