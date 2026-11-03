# Learnovate Qatar · CELPIP Prep Course

A 12-session CELPIP-General prep course website by Learnovate Qatar.

- **Students** sign in with an access code you give them. No account needed.
- **The Teacher Portal** (`#teacher`, password protected) is where you create batches, add students, assign their codes, see their progress, writing and error logs, and send feedback.
- **The site** is hosted on GitHub Pages. **The data** lives in a Google Sheet through Google Apps Script.

```
index.html              the whole website (one file)
apps-script/Code.gs     the backend you paste into Google Apps Script
```

## 1. Set up the Google Sheet and Apps Script (about 10 minutes)

1. Create a Google Sheet and name it **CELPIP Prep Course Data**.
2. Open **Extensions → Apps Script**. Delete the sample code, paste in everything from `apps-script/Code.gs`, and click **Save**.
3. Go back to the Sheet tab and **reload the page**. A new menu called **CELPIP Course** appears, next to Help.
4. Use the menu items in this order:
   1. **CELPIP Course → 1. Set teacher password.** The first time, Google asks you to authorize: click **Continue**, choose your account, click **Advanced → Go to (project name)**, then **Allow**. Run the menu item again, type your password and click **OK**.
   2. **CELPIP Course → 2. Set up sheets.** This creates the Batches, Students, Progress, Feedback and Messages tabs.
   3. **CELPIP Course → 3. Send me a test email.** You'll receive an email, which confirms the course is allowed to email students.
   4. **CELPIP Course → Check setup.** Every line should show ✅.
5. Deploy the web app, back in Apps Script:
   - Click **Deploy → New deployment**, then the gear icon → **Web app**.
   - Execute as: **Me**.
   - Who has access: **Anyone**.
   - Click **Deploy** and copy the URL that ends in `/exec`.

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

## Messaging and email

- **Students** have a **Messages** page. They can write to you and reply to your tasks, and a red badge shows unread messages.
- **In the Teacher Portal:**
  - **Inbox** (at the top) lists new student messages and task replies.
  - **Student detail** shows the full conversation with that student, plus a box to write to them.
  - **Message batch** sends one announcement to every student in a batch.
- **Also send by email:** tick this to email the message to the student's address as well. It's sent from your Google account with Google's free daily limit, about 100 emails per day.
- **CELPIP email-writing task:** choose this type, or click **Insert a sample email task**, to send a Writing Task 1–style prompt. The student replies on the site with a 27-minute timer and a 150–200 word counter. Their reply comes back to your inbox.
- **Email access code:** when you add a student with an email address, you can email their code right away. The same button is also on each student's page.
- **Get emailed about student messages (optional):** use **CELPIP Course → Email me when students write: on/off**.

## Updating to a new version of Code.gs

1. In Apps Script, select all of the old code and replace it with the new `apps-script/Code.gs`. Save.
2. Reload the Sheet and use **CELPIP Course → Set up sheets** to add any new tabs.
3. Go to **Deploy → Manage deployments → ✏️ Edit**. Set Version to **New version**, then click **Deploy**. The URL stays the same, so the website needs no change.

## Good to know

- **Updating the backend:** after you change `Code.gs`, go to **Deploy → Manage deployments → Edit → Version: New version → Deploy**. The URL stays the same.
- **Changing a student's code:** this keeps their progress and feedback. Send them the new code.
- **Deleting a student:** this removes their progress and feedback from the Sheet.
- **Privacy:** an access code is the student's key, so keep codes private. The Sheet is only visible to your Google account.
- **Offline:** if a student goes offline, their practice is kept on their device and retried automatically.
- **Speaking practice:** students record on their phone while the timer runs, then upload the file to play it back and score themselves. Recordings are not sent to the server.
