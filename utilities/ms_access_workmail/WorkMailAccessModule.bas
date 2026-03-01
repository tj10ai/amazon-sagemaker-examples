Option Compare Database
Option Explicit

' Copy/paste friendly MS Access VBA module (Standard Module)
' for AWS WorkMail over IMAP/SMTP via curl.exe.
'
' How to use in Access:
' 1) Create a new Standard Module, e.g. "modWorkMail".
' 2) Paste this file content.
' 3) Call WorkMailExecute(...) with one of: SEARCH, FETCH, SEND, REPLY.
'
' Requirements:
' - curl.exe available in PATH
' - WorkMail mailbox access over IMAP/SMTP enabled
' - Region endpoints:
'     IMAP: imap.mail.<region>.awsapps.com:993
'     SMTP: smtp.mail.<region>.awsapps.com:465

Private Const DEFAULT_IMAP_PORT As Long = 993
Private Const DEFAULT_SMTP_PORT As Long = 465

' Main function requested for MS Access use.
' Action values:
'   SEARCH -> Result = raw SEARCH response
'   FETCH  -> Result = raw message for uid
'   SEND   -> Result = curl output for send operation
'   REPLY  -> Result = curl output for reply operation
Public Function WorkMailExecute(ByVal action As String, _
                                ByVal username As String, _
                                ByVal password As String, _
                                ByVal region As String, _
                                ByVal mailFrom As String, _
                                Optional ByVal uid As String = "", _
                                Optional ByVal toAddress As String = "", _
                                Optional ByVal subject As String = "", _
                                Optional ByVal bodyText As String = "", _
                                Optional ByVal searchClause As String = "ALL", _
                                Optional ByVal ccAddress As String = "", _
                                Optional ByVal bccAddress As String = "") As String

    Dim imapHost As String
    Dim smtpHost As String
    imapHost = "imap.mail." & region & ".awsapps.com"
    smtpHost = "smtp.mail." & region & ".awsapps.com"

    Dim normalizedAction As String
    normalizedAction = UCase$(Trim$(action))

    Select Case normalizedAction
        Case "SEARCH"
            WorkMailExecute = SearchInbox(username, password, imapHost, DEFAULT_IMAP_PORT, searchClause)

        Case "FETCH"
            If Len(Trim$(uid)) = 0 Then
                WorkMailExecute = "ERROR: uid is required for FETCH"
                Exit Function
            End If
            WorkMailExecute = FetchEmailByUID(username, password, imapHost, DEFAULT_IMAP_PORT, uid)

        Case "SEND"
            If Len(Trim$(toAddress)) = 0 Then
                WorkMailExecute = "ERROR: toAddress is required for SEND"
                Exit Function
            End If
            WorkMailExecute = SendEmail(username, password, smtpHost, DEFAULT_SMTP_PORT, mailFrom, toAddress, subject, bodyText, ccAddress, bccAddress)

        Case "REPLY"
            If Len(Trim$(uid)) = 0 Then
                WorkMailExecute = "ERROR: uid is required for REPLY"
                Exit Function
            End If
            WorkMailExecute = ReplyToUID(username, password, imapHost, DEFAULT_IMAP_PORT, smtpHost, DEFAULT_SMTP_PORT, mailFrom, uid, bodyText)

        Case Else
            WorkMailExecute = "ERROR: Unknown action. Use SEARCH, FETCH, SEND, or REPLY"
    End Select
End Function

Private Function SendEmail(ByVal username As String, _
                           ByVal password As String, _
                           ByVal smtpHost As String, _
                           ByVal smtpPort As Long, _
                           ByVal mailFrom As String, _
                           ByVal toAddress As String, _
                           ByVal subject As String, _
                           ByVal bodyText As String, _
                           ByVal ccAddress As String, _
                           ByVal bccAddress As String) As String
    Dim emlPath As String
    emlPath = TempFilePath("workmail_send_", ".eml")

    Dim eml As String
    eml = BuildMessage(mailFrom, toAddress, ccAddress, bccAddress, subject, bodyText, "", "")
    WriteTextFile emlPath, eml

    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd" & _
          " --url ""smtps://" & smtpHost & ":" & smtpPort & """" & _
          " --user """ & EscapeForCmd(username & ":" & password) & """" & _
          " --mail-from """ & EscapeForCmd(mailFrom) & """" & _
          " --mail-rcpt """ & EscapeForCmd(toAddress) & """"

    If Len(Trim$(ccAddress)) > 0 Then
        cmd = cmd & " --mail-rcpt """ & EscapeForCmd(ccAddress) & """"
    End If

    If Len(Trim$(bccAddress)) > 0 Then
        cmd = cmd & " --mail-rcpt """ & EscapeForCmd(bccAddress) & """"
    End If

    cmd = cmd & " --upload-file """ & EscapeForCmd(emlPath) & """"

    SendEmail = RunCommandAndCapture(cmd)
End Function

Private Function SearchInbox(ByVal username As String, _
                             ByVal password As String, _
                             ByVal imapHost As String, _
                             ByVal imapPort As Long, _
                             ByVal searchClause As String) As String
    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd" & _
          " --url ""imaps://" & imapHost & ":" & imapPort & "/INBOX""" & _
          " --user """ & EscapeForCmd(username & ":" & password) & """" & _
          " -X ""SEARCH " & EscapeForCmd(searchClause) & """"

    SearchInbox = RunCommandAndCapture(cmd)
End Function

Private Function FetchEmailByUID(ByVal username As String, _
                                 ByVal password As String, _
                                 ByVal imapHost As String, _
                                 ByVal imapPort As Long, _
                                 ByVal uid As String) As String
    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd" & _
          " --url ""imaps://" & imapHost & ":" & imapPort & "/INBOX/;UID=" & EscapeForCmd(uid) & """" & _
          " --user """ & EscapeForCmd(username & ":" & password) & """"

    FetchEmailByUID = RunCommandAndCapture(cmd)
End Function

Private Function ReplyToUID(ByVal username As String, _
                            ByVal password As String, _
                            ByVal imapHost As String, _
                            ByVal imapPort As Long, _
                            ByVal smtpHost As String, _
                            ByVal smtpPort As Long, _
                            ByVal mailFrom As String, _
                            ByVal uid As String, _
                            ByVal replyBody As String) As String
    Dim original As String
    original = FetchEmailByUID(username, password, imapHost, imapPort, uid)

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
    If Len(Trim$(replySubject)) = 0 Then replySubject = "(No Subject)"

    If InStr(1, LCase$(replySubject), "re:", vbTextCompare) <> 1 Then
        replySubject = "Re: " & replySubject
    End If

    Dim emlPath As String
    emlPath = TempFilePath("workmail_reply_", ".eml")

    Dim replyMsg As String
    replyMsg = BuildMessage(mailFrom, originalFrom, "", "", replySubject, replyBody, originalMessageId, originalMessageId)
    WriteTextFile emlPath, replyMsg

    Dim cmd As String
    cmd = "curl --silent --show-error --ssl-reqd" & _
          " --url ""smtps://" & smtpHost & ":" & smtpPort & """" & _
          " --user """ & EscapeForCmd(username & ":" & password) & """" & _
          " --mail-from """ & EscapeForCmd(mailFrom) & """" & _
          " --mail-rcpt """ & EscapeForCmd(originalFrom) & """" & _
          " --upload-file """ & EscapeForCmd(emlPath) & """"

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
    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    TempFilePath = tempFolder & "\" & prefix & Format$(Now, "yyyymmdd_hhnnss") & "_" & CLng(Timer * 1000) & extension
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal contents As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")

    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText contents
    stm.SaveToFile filePath, 2
    stm.Close
End Sub

Private Function EscapeForCmd(ByVal value As String) As String
    EscapeForCmd = Replace(value, """", "\""")
End Function

Private Function RunCommandAndCapture(ByVal cmd As String) As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")

    Dim execObj As Object
    Set execObj = shell.Exec("cmd /c " & cmd)

    Do While execObj.Status = 0
        DoEvents
    Loop

    Dim output As String
    output = ""

    If Not execObj.StdOut.AtEndOfStream Then output = output & execObj.StdOut.ReadAll
    If Not execObj.StdErr.AtEndOfStream Then output = output & vbCrLf & execObj.StdErr.ReadAll

    RunCommandAndCapture = Trim$(output)
End Function

Public Sub ExampleUsage()
    Dim username As String
    Dim password As String
    Dim region As String
    Dim mailFrom As String

    username = "user@example.com"
    password = "APP_PASSWORD_OR_MAILBOX_PASSWORD"
    region = "us-east-1"
    mailFrom = "user@example.com"

    Debug.Print "=== SEARCH ==="
    Debug.Print WorkMailExecute("SEARCH", username, password, region, mailFrom, , , , , "UNSEEN")

    Debug.Print "=== FETCH UID 123 ==="
    Debug.Print WorkMailExecute("FETCH", username, password, region, mailFrom, "123")

    Debug.Print "=== SEND ==="
    Debug.Print WorkMailExecute("SEND", username, password, region, mailFrom, , "recipient@example.com", "Test from Access", "Hello from Access VBA + WorkMail")

    Debug.Print "=== REPLY UID 123 ==="
    Debug.Print WorkMailExecute("REPLY", username, password, region, mailFrom, "123", , , "Thanks, received your email.")
End Sub
