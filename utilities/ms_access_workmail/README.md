# MS Access VBA Function for AWS WorkMail (No import required)

This version is **copy/paste friendly** for MS Access.

If importing `.bas` is problematic in your environment, open a **Standard Module** in Access and paste the code from `WorkMailAccessModule.bas` directly.

## Main function

Use one function for all operations:

```vb
WorkMailExecute(action, username, password, region, mailFrom, [uid], [toAddress], [subject], [bodyText], [searchClause], [ccAddress], [bccAddress])
```

### Supported actions

- `SEARCH` -> IMAP `SEARCH` response from INBOX
- `FETCH` -> Raw message content by UID
- `SEND` -> Sends a new email via SMTP
- `REPLY` -> Reads original by UID and sends reply with thread headers

## Quick examples

```vb
Dim result As String

' Search unread
result = WorkMailExecute("SEARCH", "user@example.com", "PASSWORD", "us-east-1", "user@example.com", , , , , "UNSEEN")
Debug.Print result

' Fetch UID
result = WorkMailExecute("FETCH", "user@example.com", "PASSWORD", "us-east-1", "user@example.com", "123")
Debug.Print result

' Send
result = WorkMailExecute("SEND", "user@example.com", "PASSWORD", "us-east-1", "user@example.com", , "recipient@example.com", "Hello", "Hi from Access")
Debug.Print result

' Reply
result = WorkMailExecute("REPLY", "user@example.com", "PASSWORD", "us-east-1", "user@example.com", "123", , , "Thanks!")
Debug.Print result
```

## Requirements

- `curl.exe` in PATH
- WorkMail IMAP/SMTP access enabled
- Region endpoints:
  - IMAP: `imap.mail.<region>.awsapps.com:993`
  - SMTP: `smtp.mail.<region>.awsapps.com:465`

## Security note

Credentials are passed to `curl` command line by this lightweight approach. For production, consider a more secure secret-handling strategy.
