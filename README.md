# CANHS Instructional Supervision

Website of the Office of the School Head, Castor Alviar National High School, for the department head designates:
Annex B IS Plans, Annex E classroom observations (COT ratings with the Annex E-1 rubric guide, filled Rating Sheet /
Inter-Observer Agreement Form, Observation Notes Form as PDF, Word or print), Annex D PMCF (one teacher per month,
with sample entries), supervision records, personnel from the PSIPOP, and the Annex A–F files.

## Files

| File | What it is |
|---|---|
| `index.html` | The whole website |
| `config.js` | Where the shared-record address goes (step 3 below) |
| `annex/` | The Annex A–F files. The site also fills the Word and Excel templates in this folder, so keep the file names as they are |
| `apps-script/Code.gs` | The code for the Google Sheet that keeps the shared record (step 2) |

## 1. Put the site on GitHub Pages

1. Sign in to GitHub and create a new repository, e.g. `canhs-instructional-supervision` (Public).
2. Click **Add file → Upload files** and drag in everything in this folder (`index.html`, `config.js`, `README.md`, and the `annex` and `apps-script` folders). Click **Commit changes**.
3. Open **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
4. After a minute the site is at `https://<your-github-name>.github.io/canhs-instructional-supervision/`.

The site works at this point, but each person's entries stay in their own browser. Steps 2–3 give everyone one shared record.

## 2. Make the shared record (Google Sheet + Apps Script)

1. In Google Drive, create a new Google Sheet, e.g. **CANHS IS Records**.
2. In the sheet: **Extensions → Apps Script**. Delete what is in `Code.gs` and paste the contents of `apps-script/Code.gs`.
3. At the top of the code, change the three codes:
   - `ADMIN` – the School Head's code (can also set department heads and teacher assignments)
   - `EDIT` – the code you give the department heads (can add, change and delete entries)
   - `VIEW` – optional read-only code; leave `''` to turn it off
4. Save. Choose the function **setup** and click **Run** once; allow the permissions it asks for. Four tabs appear in the sheet: `obs`, `coach`, `plans`, `config`.
5. Click **Deploy → New deployment**. Type: **Web app**. Execute as: **Me**. Who has access: **Anyone**. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

## 3. Connect the site to the sheet

1. On GitHub, open `config.js`, click the pencil (Edit), and paste the web app URL between the quotes:
   ```js
   window.CANHS_IS = {
     apiUrl: "https://script.google.com/macros/s/AKfy..../exec"
   };
   ```
2. Click **Commit changes**. After a minute, open the site, go to **Connection**, and enter your code.
3. Send the department heads the site address and the `EDIT` code. They enter it once under **Connection**.

## Notes

- Only people with a code can read or save records; the Apps Script code and the codes in it are not visible on the website.
- If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → Version: New version** so the same `/exec` address keeps working.
- **Connection → Download a backup** saves every record as one JSON file. The Google Sheet itself is also a backup.
- To change a code, edit it in `Code.gs` and deploy a new version; people then enter the new code under **Connection**.
- All dates are written in full, e.g. "October 14, 2026".
