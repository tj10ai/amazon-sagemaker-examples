Attribute VB_Name = "WorkMailAccess"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' MS Access VBA module for AWS WorkMail mailbox actions using curl.exe + IMAP/SMTP
' -----------------------------------------------------------------------------
' Requirements:
'   1) curl.exe available in PATH (Windows 10/11 ships with it by default)
'   2) IMAP enabled for your WorkMail organization/user
'   3) SMTP authentication enabled for your WorkMail user
'   4) TLS endpoints, for example:
'        IMAP: imap.mail.us-east-1.awsapps.com:993
'        SMTP: smtp.mail.us-east-1.awsapps.com:465
' -----------------------------------------------------------------------------

Private Const DEFAULT_IMAP_PORT As Long = 993
Private Const DEFAULT_SMTP_PORT As Long = 465

Public Type WorkMailConfig
    Username As String
    Password As String
    ImapHost As String
    ImapPort As Long
    SmtpHost As String
    SmtpPort As Long
    MailFrom As String
End Type

Public Function NewWorkMailConfig(ByVal username As String, _
                                  ByVal password As String, _
                                  ByVal region As String, _
                                  ByVal mailFrom As String) As WorkMailConfig
    Dim cfg As WorkMailConfig
    cfg.Username = username
    cfg.Password = password
    cfg.ImapHost = "imap.mail." & region & ".awsapps.com"
    cfg.ImapPort = DEFAULT_IMAP_PORT
    cfg.SmtpHost = "smtp.mail." & region & ".awsapps.com"
    cfg.SmtpPort = DEFAULT_SMTP_PORT
    cfg.MailFrom = mailFrom
    NewWorkMailConfig = cfg
End Function

Public Function SendEmail(ByRef cfg As WorkMailConfig, _
                          ByVal toAddress As String, _
                          ByVal subject As String, _
                          ByVal bodyText As String, _
                          Optional ByVal ccAddress As String = "", _
                          Optional ByVal bccAddress As String = "") As String
    Dim emlPath As String
    emlPath = TempFilePath("workmail_send_", ".eml")

    Dim eml As String
    eml = BuildMessage(cfg.MailFrom, toAddress, ccAddress, bccAddress, subject, bodyText, "", "")
    WriteTextFile emlPath, eml

    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd " & _
          " --url ""smtps://" & cfg.SmtpHost & ":" & cfg.SmtpPort & """ & _
          " --user """ & EscapeQuotes(cfg.Username & ":" & cfg.Password) & """" & _
          " --mail-from """ & EscapeQuotes(cfg.MailFrom) & """" & _
          " --mail-rcpt """ & EscapeQuotes(toAddress) & """"

    If Len(Trim$(ccAddress)) > 0 Then
        cmd = cmd & " --mail-rcpt """ & EscapeQuotes(ccAddress) & """"
    End If

    If Len(Trim$(bccAddress)) > 0 Then
        cmd = cmd & " --mail-rcpt """ & EscapeQuotes(bccAddress) & """"
    End If

    cmd = cmd & " --upload-file """ & EscapeQuotes(emlPath) & """"

    SendEmail = RunCommandAndCapture(cmd)
End Function

Public Function SearchInbox(ByRef cfg As WorkMailConfig, _
                            Optional ByVal searchClause As String = "ALL") As String
    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd " & _
          " --url ""imaps://" & cfg.ImapHost & ":" & cfg.ImapPort & "/INBOX""" & _
          " --user """ & EscapeQuotes(cfg.Username & ":" & cfg.Password) & """" & _
          " -X ""SEARCH " & EscapeQuotes(searchClause) & """"
    SearchInbox = RunCommandAndCapture(cmd)
End Function

Public Function FetchEmailByUID(ByRef cfg As WorkMailConfig, ByVal uid As String) As String
    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd " & _
          " --url ""imaps://" & cfg.ImapHost & ":" & cfg.ImapPort & "/INBOX/;UID=" & EscapeQuotes(uid) & """" & _
          " --user """ & EscapeQuotes(cfg.Username & ":" & cfg.Password) & """"
    FetchEmailByUID = RunCommandAndCapture(cmd)
End Function

Public Function ReplyToUID(ByRef cfg As WorkMailConfig, _
                           ByVal uid As String, _
                           ByVal replyBody As String) As String
    Dim original As String
    original = FetchEmailByUID(cfg, uid)

    Dim originalFrom As String
    Dim originalSubject As String
    Dim originalMessageId As String

    originalFrom = ExtractHeaderValue(original, "From")
    originalSubject = ExtractHeaderValue(original, "Subject")
    originalMessageId = ExtractHeaderValue(original, "Message-ID")

    If Len(Trim$(originalFrom)) = 0 Then
        ReplyToUID = "ERROR: Could not extract From header from UID " & uid
        Exit Function
    End If

    Dim replySubject As String
    replySubject = originalSubject
    If InStr(1, LCase$(replySubject), "re:", vbTextCompare) <> 1 Then
        replySubject = "Re: " & replySubject
    End If

    Dim emlPath As String
    emlPath = TempFilePath("workmail_reply_", ".eml")

    Dim replyMsg As String
    replyMsg = BuildMessage(cfg.MailFrom, originalFrom, "", "", replySubject, replyBody, originalMessageId, originalMessageId)
    WriteTextFile emlPath, replyMsg

    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd " & _
          " --url ""smtps://" & cfg.SmtpHost & ":" & cfg.SmtpPort & """ & _
          " --user """ & EscapeQuotes(cfg.Username & ":" & cfg.Password) & """" & _
          " --mail-from """ & EscapeQuotes(cfg.MailFrom) & """" & _
          " --mail-rcpt """ & EscapeQuotes(originalFrom) & """" & _
          " --upload-file """ & EscapeQuotes(emlPath) & """"

    ReplyToUID = RunCommandAndCapture(cmd)
End Function

Private Function BuildMessage(ByVal fromAddress As String, _
                              ByVal toAddress As String, _
                              ByVal ccAddress As String, _
                              ByVal bccAddress As String, _
                              ByVal subject As String, _
                              ByVal bodyText As String, _
                              ByVal inReplyTo As String, _
                              ByVal referencesHeader As String) As String
    Dim lines As Collection
    Set lines = New Collection

    lines.Add "From: " & fromAddress
    lines.Add "To: " & toAddress

    If Len(Trim$(ccAddress)) > 0 Then lines.Add "Cc: " & ccAddress
    If Len(Trim$(bccAddress)) > 0 Then lines.Add "Bcc: " & bccAddress

    lines.Add "Subject: " & subject
    lines.Add "MIME-Version: 1.0"
    lines.Add "Content-Type: text/plain; charset=UTF-8"
    lines.Add "Content-Transfer-Encoding: 8bit"

    If Len(Trim$(inReplyTo)) > 0 Then lines.Add "In-Reply-To: " & inReplyTo
    If Len(Trim$(referencesHeader)) > 0 Then lines.Add "References: " & referencesHeader

    lines.Add ""
    lines.Add bodyText

    Dim i As Long
    Dim msg As String
    For i = 1 To lines.Count
        msg = msg & CStr(lines.Item(i)) & vbCrLf
    Next i

    BuildMessage = msg
End Function

Private Function ExtractHeaderValue(ByVal rawMessage As String, ByVal headerName As String) As String
    Dim marker As String
    marker = vbCrLf & headerName & ":"

    Dim source As String
    source = vbCrLf & rawMessage

    Dim pos As Long
    pos = InStr(1, source, marker, vbTextCompare)
    If pos = 0 Then Exit Function

    Dim startPos As Long
    startPos = pos + Len(marker)

    Dim endPos As Long
    endPos = InStr(startPos, source, vbCrLf)
    If endPos = 0 Then endPos = Len(source) + 1

    ExtractHeaderValue = Trim$(Mid$(source, startPos, endPos - startPos))
End Function

Private Function TempFilePath(ByVal prefix As String, ByVal extension As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String
    tempFolder = Environ$("TEMP")

    TempFilePath = tempFolder & "\" & prefix & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Timer * 1000) & extension
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal contents As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2 ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText contents
    stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    stm.Close
End Sub

Private Function EscapeQuotes(ByVal value As String) As String
    EscapeQuotes = Replace(value, """", "\""")
End Function

Private Function RunCommandAndCapture(ByVal cmd As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim execObj As Object
    Set execObj = shell.Exec("cmd /c " & cmd)

    Dim output As String
    output = ""

    Do While execObj.Status = 0
        DoEvents
    Loop

    If Not execObj.StdOut.AtEndOfStream Then
        output = output & execObj.StdOut.ReadAll
    End If

    If Not execObj.StdErr.AtEndOfStream Then
        output = output & vbCrLf & execObj.StdErr.ReadAll
    End If

    RunCommandAndCapture = Trim$(output)
End Function

Public Sub ExampleUsage()
    Dim cfg As WorkMailConfig
    cfg = NewWorkMailConfig( _
        "user@example.com", _
        "APP_PASSWORD_OR_MAILBOX_PASSWORD", _
        "us-east-1", _
        "user@example.com")

    Debug.Print "=== SEARCH ==="
    Debug.Print SearchInbox(cfg, "UNSEEN")

    Debug.Print "=== FETCH UID 123 ==="
    Debug.Print FetchEmailByUID(cfg, "123")

    Debug.Print "=== SEND ==="
    Debug.Print SendEmail(cfg, "recipient@example.com", "Test from Access", "Hello from Access VBA + WorkMail")

    Debug.Print "=== REPLY UID 123 ==="
    Debug.Print ReplyToUID(cfg, "123", "Thanks, received your email.")
End Sub
