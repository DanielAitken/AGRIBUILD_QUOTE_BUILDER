# AgriBuild Public Quote Builder

This project now supports:

- Public access to the quote form at `/`
- Form submission to `/quote`
- Automatic PDF generation of form answers
- Local storage of each submission (PDF + answers + uploads)
- Optional email delivery (Graph/SMTP) when configured

## 1. Install and run locally

```bash
npm install
npm start
```

Open `http://localhost:3000`.

## 2. Configure environment variables

Copy `.env.example` to `.env` and set:

- `PORT`

Optional email config (only if you want forwarding):

- `APP_MAILBOX` (the mailbox identity the app sends from)
- `FORWARD_TO` (who receives forwarded submissions)
- `MAIL_TO` (legacy fallback if `FORWARD_TO` is not set)
- `GRAPH_TENANT_ID`
- `GRAPH_CLIENT_ID`
- `GRAPH_CLIENT_SECRET`

Optional SMTP fallback (if Graph is not configured):

- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_SECURE`
- `SMTP_USER`
- `SMTP_PASS`
- `SMTP_FROM`

## 3. Make it public (shareable link)

Deploy to a Node host such as Render/Railway:

1. Push this project to GitHub.
2. Create a new web service from the repo.
3. Build command: `npm install`
4. Start command: `npm start`
5. Add the same environment variables from your `.env`.
6. Deploy.

After deploy, you will get a public URL (for example `https://your-app.onrender.com/`) that you can share with customers.

## 4. How submissions work

When a user submits the form:

1. The server receives fields and uploaded drawings.
2. The server creates a PDF summary of the form.
3. The server saves everything under `submissions/<reference>/`:
4. `quote-request.pdf`, `answers.txt`, `form.json`, and uploaded files in `uploads/`.
5. If email is configured, it also forwards the submission by email.
