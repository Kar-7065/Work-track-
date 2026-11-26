# Work Manager – setup (about 15 minutes, free)

All data is stored in YOUR Google account: a Google Sheet (records) and a Drive folder (photos, voice notes). PDF backups can be downloaded or emailed to the admin.

**Why two parts?** The *backend* runs inside Google (Part A). The *app page* must be opened from a normal web link (Part B), because Google blocks GPS location and microphone on pages shown inside its own frame. Do both parts.

## Part A – Backend (Google Sheet + Apps Script)
1. Go to sheets.google.com → new blank Sheet, name it "Work Manager Data".
2. **Extensions → Apps Script**. Replace the contents of `Code.gs` with the provided `Code.gs`.
3. Gear **Project Settings** → tick **Show appsscript.json** → replace it with the provided `appsscript.json` (sets Delhi time).
4. In the top bar choose function **setup** → **Run**. Allow permissions (if Google says "unverified app": Advanced → Go to project; it is your own script).
5. **Deploy → New deployment → Web app** → Execute as **Me**, Who has access **Anyone** → Deploy.
6. Copy the **Web app URL** (ends with `/exec`).

## Part B – The app page (what workers open)
1. Open `Index.html` in any text editor. Near the top find `const API_URL='';` and paste your URL between the quotes:
   `const API_URL='https://script.google.com/macros/s/XXXX/exec';`
2. Save the file as **`index.html`**.
3. Go to **app.netlify.com/drop** and drag the file in (free, no sign-up needed to start; sign in to keep the link permanently). GitHub Pages also works.
4. Send the link to workers. On the phone: open it → browser menu → **Add to Home Screen**. Allow **Location** and **Microphone** when asked.

## First login (change immediately in the Me tab)
| Admin | Username | Password |
|---|---|---|
| Admin 1 | admin1 | Admin1@Delhi26 |
| Admin 2 | admin2 | Admin2@Delhi26 |

## Using it
- **Workers tab**: add a worker with username + password. They can log in with username, phone number or email. Set working hours (12-hour pickers), Day/Night shift, monthly salary and overtime rate. Admin can reset any password, including the other admin's.
- **Tasks tab**: give daily work to one worker or all; it shows on that worker's "My Work" screen.
- **Workers**: Check In / Check Out (GPS location + Delhi time saved), write updates with photo and voice note, record stock used.
- **Salary tab**: add advance or salary payments per month. Balance = monthly salary + overtime pay − advance − salary paid. Workers see only their own.
- **Overtime**: any time worked beyond the worker's set working hours.
- **Backup tab**: PDF of everything, or attendance / salary / payments / work / stock, by month and worker. Download it or email it to the admin.
- Nightly automatic email backup: in the Apps Script editor run `enableDailyBackup` once.

## Good to know
- Full monthly salary is paid by default; absent days are **not** deducted automatically. Use the payment note or adjust the payment to account for absences.
- Google Drive is used for all photos and voice notes, so nothing fills a phone or server.
- After changing code later: Deploy → Manage deployments → pencil → Version: New version → Deploy (URL stays the same). If you edit `Index.html`, upload it to Netlify again.
