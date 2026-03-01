# MS Access VBA + AWS WorkMail (IMAP/SMTP)

This folder contains a VBA module you can import into an MS Access database to perform common WorkMail mailbox actions:

- Search/read emails from `INBOX` via IMAP (`SEARCH`, `UID` fetch)
- Send emails via SMTP
- Reply to an email by UID (adds `In-Reply-To` and `References`)

## Files

- `WorkMailAccessModule.bas` – VBA module to import into Access.

## Setup

1. In Access, open the VBA editor (`ALT+F11`), then import `WorkMailAccessModule.bas`.
2. Ensure `curl.exe` is available on your system `PATH`.
3. Ensure AWS WorkMail user credentials are valid for IMAP/SMTP login.
4. Use your WorkMail region endpoint format:
   - IMAP: `imap.mail.<region>.awsapps.com:993`
   - SMTP: `smtp.mail.<region>.awsapps.com:465`

## Quick start

```vb
Dim cfg As WorkMailConfig
cfg = NewWorkMailConfig("user@example.com", "PASSWORD", "us-east-1", "user@example.com")

Debug.Print SearchInbox(cfg, "UNSEEN")
Debug.Print FetchEmailByUID(cfg, "123")
Debug.Print SendEmail(cfg, "recipient@example.com", "Hello", "Hi from Access")
Debug.Print ReplyToUID(cfg, "123", "Thanks!")
```

## Notes and limitations

- Credentials are passed to `curl` command line. In production, store secrets securely and avoid hard-coding passwords.
- `ReplyToUID` uses basic header extraction and plain text body.
- For HTML email and attachments, extend `BuildMessage` to emit MIME multipart payloads.
