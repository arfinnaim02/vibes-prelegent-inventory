# Deploy VIBES × Prelegent Portal Live

## A. Upload to GitHub

Create a new repository, for example:
`vibes-prelegent-inventory`

Upload these items to the repository root:
- `.streamlit/`
- `data/`
- `.gitignore`
- `app.py`
- `requirements.txt`
- `README.md`

Important: `app.py` must be at the repository root for the simplest deployment setup.

## B. Deploy on Streamlit Community Cloud

1. Sign in to Streamlit Community Cloud with GitHub.
2. Choose **Create app / Deploy an app**.
3. Select the GitHub repository.
4. Set:
   - Repository: your VIBES Prelegent repo
   - Branch: `main`
   - Main file path: `app.py`
5. In Advanced settings, use Python 3.12 unless you have a reason to choose another supported release.
6. Choose an app URL, for example `vibes-prelegent` if available.
7. Click **Deploy**.

Your live URL will be similar to:
`https://your-app-name.streamlit.app`

## C. Updating the portal later

For code changes:
- edit or replace files in GitHub and commit;
- Streamlit Community Cloud redeploys from the repository.

For one-off inventory analysis:
- use **Upload latest inventory workbook** in the sidebar.

Important: a workbook uploaded through the Streamlit UI is session/runtime data; it is not a permanent database. If you want the portal to automatically retain every new inventory file, the next upgrade should connect it to Google Sheets, Google Drive, or a database.

## D. Recommended next production upgrades

1. Password/login for VIBES board/admin users.
2. Google Sheet or database as the live inventory source.
3. Historical snapshots by date.
4. Sales/commission data from Prelegent.
5. Automatic send/return approval workflow.
6. Exportable board PDF/Excel report.
