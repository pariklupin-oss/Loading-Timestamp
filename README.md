# Loading Timestamp Web App

Mobile-first loading register based on the columns in the provided workbook:

- **Shift details:** Date, Shift, Loading Incharge, Helper Count
- **Incharge names:** the default selection uses AJAY in place of AJEET; use **＋ Add** beside the field for a new name. A new name is stored locally right away and becomes available to other connected devices after the first loading entry with that name is saved to the Google Sheet.
- **Vehicle details:** Customer, Loading Start, Loading End, Total Hours, Vehicle No, Vehicle Feet / Type, Remarks
- **Customer search:** type in the Customer field to find and select a name from the connected `CUSTOMER_MASTER` list. A new customer name can also be typed if it is not listed. The attached file's 251 unique names are in the source sheet's `CUSTOMER_MASTER` tab.
- **Reports:** date, shift and incharge filters; CSV download
- **Operator shift cards:** the overview combines repeated records for the same incharge, date and shift into one card. Tap it to see all associated vehicles, add a late vehicle even after finishing the shift, or edit the shift details. Use Reports → Details / edit to open older shifts too.
- **Edit saved vehicle:** use Edit this vehicle in shift details; customer, vehicle number/type, loading times and remarks can be updated. Loading duration recalculates automatically.
- **One shift, many vehicles:** enter date, shift, incharge and helper count once; after the first save, keep adding vehicles to that active shift. Use **Finish shift** before starting another shift.
- **Product stage-gap report:** invoice/product rows from `STAGE_TIME`, including OQC, loading, invoice and gate-out timestamps with all three gap columns
- **PDF:** print/save layout follows the supplied second screenshot. It prints Date through OQC End plus GAP-1, GAP-2 and GAP-3; the detailed loading, invoice and gate-out timestamps remain visible in the on-screen report.

## Run locally

Open `index.html` in a browser, or serve this folder over HTTP for install/offline support:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Publish with GitHub Pages

1. Create a GitHub repository and upload the contents of this folder to its root.
2. In the repository, open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, select `main` and `/ (root)`, then save.
4. Open the published Pages address on the phone and use **Add to Home Screen** if desired.

## Connect it to the Google Sheet

The app starts in local-only mode. To enable shared data, add the included `apps-script/LoadingWebAppConnector.gs` to the existing Apps Script project bound to the Google Sheet that contains `LOADING_HEADER`, `LOADING_ITEMS`, and `STAGE_TIME`.

1. Open the Google Sheet and choose **Extensions → Apps Script**.
2. Keep your existing script functions, including `copyItemMasterData()` and `copyAllOQCData()`. Replace only the contents of the `LoadingWebAppConnector` script file with `apps-script/LoadingWebAppConnector.gs`. Do not replace or delete your other functions.
3. In the existing `LOADING TIMESTAMP` spreadsheet, confirm `CUSTOMER_MASTER` has a `Customer Name` header in A1 and names below it. This app bundle's connector reads that tab and returns its names to the searchable Customer field.
4. Choose **Deploy → Manage deployments**, edit the current web app deployment, select **New version**, and deploy. Set **Execute as** to **User accessing the web app**. Restrict access to your Google Workspace if available; otherwise allow only signed-in Google account users. Operators must be signed in and have edit access to the source sheet.
5. Keep the web app URL already set in `config.js` (or update it if the deployment URL changed), then commit the app files to the GitHub repository and wait for Pages to redeploy. Open the app while signed in to Google. The Customer field will show suggestions as the operator types.

The app checks required tab names and columns before reading or writing. It writes to `LOADING_HEADER` and `LOADING_ITEMS`, using the workbook's existing field names, and reads the private `CUSTOMER_MASTER` list. The connector does not create or rename spreadsheet tabs. Do not deploy this script with public anonymous access: it can read and write loading records in the connected spreadsheet.

Until the API URL is set and a sheet load succeeds, entries remain in browser local storage on that device. Use **Backup and restore → Download backup** regularly. The app bundle contains no customer or production records. Loading hours are calculated from start and end time; when end time is earlier than start time, the app treats it as the next day.

## Security and access

GitHub Pages publishes the app files to anyone who can access the repository/site. Do not add company records, passwords, API keys, or private links to this repository. A shared multi-user register needs a protected backend and user access controls; this static version is a single-device starter app.
