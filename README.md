# CELPIP Prep Course

A 12-session CELPIP-General prep course website for Ms. Maria's classes.

- **Students** sign in with an access code you give them. No account needed.
- **The Teacher Portal** (`#teacher`, password protected) is where you create batches, add students, assign their codes, see their progress, writing and error logs, and send feedback.
- **The site** is hosted on GitHub Pages. **The data** lives in a Google Sheet through Google Apps Script.

```
index.html              the whole website (one file)
apps-script/Code.gs     the backend you paste into Google Apps Script
```

## 1. Set up the Google Sheet and Apps Script (about 10 minutes)

1. Create a new Google Sheet and name it **CELPIP Prep Course Data**.
2. In the Sheet, open **Extensions → Apps Script**.
3. Delete the sample code, then paste in everything from `apps-script/Code.gs`. Save.
4. Set your teacher password:
   - Click **Project Settings** (the gear icon), then **Script properties → Add script property**.
   - Property: `TEACHER_PASSWORD`. Value: a strong password of your choice. Save.
5. Create the sheets: in the editor, choose the `setup` function in the toolbar and click **Run**. Approve the permissions when asked. Four tabs appear in your Sheet: Batches, Students, Progress and Feedback.
6. Deploy the web app:
   - Click **Deploy → New deployment**, then the gear icon → **Web app**.
   - Execute as: **Me**.
   - Who has access: **Anyone**.
   - Click **Deploy** and copy the **Web app URL**. It ends in `/exec`.

## 2. Connect the website

Open `index.html`, find this line near the top of the `<script>`, and paste in your URL:

```js
API_URL:'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE'
```

## 3. Publish on GitHub Pages

1. Upload `index.html` to your repository, either at the root or in a folder.
2. Go to **Settings → Pages**, choose the `main` branch and click **Save**.
3. Your site is live at `https://<your-username>.github.io/<repo>/`.

## 4. Run a batch

1. Open the site, scroll to the bottom of the sign-in card and click **Teacher sign in**. Use your password.
2. Click **+ New batch**, then **+ Add student**. Enter their name, email and phone. Keep the generated code (for example `CEL-7KQ2MZ`) or type your own.
3. Click **Copy message** and send it to the student by email, text or WhatsApp.
4. The student opens the link, types the code, and their dashboard opens.

## Good to know

- **Updating the backend:** after you change `Code.gs`, go to **Deploy → Manage deployments → Edit → Version: New version → Deploy**. The URL stays the same.
- **Changing a student's code:** this keeps their progress and feedback. Send them the new code.
- **Deleting a student:** this removes their progress and feedback from the Sheet.
- **Privacy:** an access code is the student's key, so keep codes private. The Sheet is only visible to your Google account.
- **Offline:** if a student goes offline, their practice is kept on their device and retried automatically.
- **Speaking practice:** students record on their phone while the timer runs, then upload the file to play it back and score themselves. Recordings are not sent to the server.
