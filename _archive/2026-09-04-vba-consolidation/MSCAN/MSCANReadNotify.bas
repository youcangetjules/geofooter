Attribute VB_Name = "MSCANReadNotify"
'===============================================================================
' MSCANReadNotify - ReadNotify.com tracking for Aliniant TRACKED sends
' Appends ReadNotify address suffixes to SMTP recipients (manual ActiveTracker method).
' Correct form: user@domain.com.read-notify.com
'===============================================================================
Option Explicit

Private Const DEFAULT_ALINIANT_ACCOUNT As String = "julian.garrett@aliniant.com"
Private Const DEFAULT_READNOTIFY_SUFFIX As String = ".read-notify.com"
Private Const READNOTIFY_DOMAIN As String = "read-notify.com"
' ANSI + Unicode PR_SMTP_ADDRESS (Exchange address entries vary by profile).
Private Const PR_SMTP_ADDRESS_A As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
Private Const PR_SMTP_ADDRESS_W As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

Public Function GetAliniantAccount() As String
    On Error Resume Next
    Dim overridePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String

    overridePath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_readnotify.ini"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(overridePath) Then
        Set ts = fso.OpenTextFile(overridePath, 1, False)
        Do While Not ts.AtEndOfStream
            line = Trim$(ts.ReadLine)
            If Left$(LCase$(line), 8) = "account=" Then
                GetAliniantAccount = Trim$(Mid$(line, 9))
                ts.Close
                If Len(GetAliniantAccount) > 0 Then Exit Function
            End If
        Loop
        ts.Close
    End If

    GetAliniantAccount = DEFAULT_ALINIANT_ACCOUNT
End Function

Public Function IsAliniantSend(ByVal mail As Outlook.MailItem) As Boolean
    On Error Resume Next

    Dim target As String
    Dim targetDomain As String
    Dim candidate As String
    Dim candDomain As String

    target = LCase$(Trim$(GetAliniantAccount()))
    If Len(target) = 0 Then
        IsAliniantSend = False
        Exit Function
    End If
    targetDomain = DomainOfSmtp(target)

    candidate = ""
    If Not mail.SendUsingAccount Is Nothing Then
        candidate = mail.SendUsingAccount.SmtpAddress
        If Len(candidate) = 0 Then candidate = mail.SendUsingAccount.DisplayName
    End If

    If Len(candidate) = 0 Then candidate = mail.SenderEmailAddress
    If Len(candidate) = 0 Then candidate = mail.SentOnBehalfOfName

    candidate = LCase$(Trim$(ExtractSmtpFromText(candidate)))
    If Len(candidate) = 0 Then
        IsAliniantSend = False
        Exit Function
    End If

    ' Exact account match, or same SMTP domain as the configured Aliniant account
    ' (e.g. administrator@aliniant.com when target is julian.garrett@aliniant.com).
    If StrComp(candidate, target, vbTextCompare) = 0 Then
        IsAliniantSend = True
        Exit Function
    End If

    candDomain = DomainOfSmtp(candidate)
    IsAliniantSend = (Len(targetDomain) > 0 And StrComp(candDomain, targetDomain, vbTextCompare) = 0)
End Function

Public Sub ApplyReadNotifyIfRequired(ByVal mail As Outlook.MailItem, ByVal tagIndex As Long)
    On Error GoTo EH

    If mail Is Nothing Then Exit Sub
    If tagIndex < 1 Then Exit Sub
    If Not MSCANCore.RequiresTracking(tagIndex) Then
        MSCANModLogging.WriteLog "ReadNotify: skipped (classification does not require tracking)."
        Exit Sub
    End If

    If Not ApplyReadNotifyIfRequested(mail, True, DEFAULT_READNOTIFY_SUFFIX) Then
        Err.Raise vbObjectError + 6001, "MSCANReadNotify", "ReadNotify stamping failed (see log)."
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLogLevel "error", _
        "ERROR in MSCANReadNotify.ApplyReadNotifyIfRequired | #" & Err.Number & " | " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Returns True on success / intentional skip; False on failure (including message-changed).
Public Function ApplyReadNotifyIfRequested( _
    ByVal mail As Outlook.MailItem, _
    ByVal enabled As Boolean, _
    Optional ByVal suffix As String = "") As Boolean

    On Error GoTo EH

    ApplyReadNotifyIfRequested = False

    If mail Is Nothing Then Exit Function
    If Not enabled Then
        MSCANModLogging.WriteLog "ReadNotify: skipped (user disabled tracking for this send)."
        ApplyReadNotifyIfRequested = True
        Exit Function
    End If
    If Not IsAliniantSend(mail) Then
        MSCANModLogging.WriteLog "ReadNotify: skipped (not Aliniant send account)."
        ApplyReadNotifyIfRequested = True
        Exit Function
    End If

    Dim useSuffix As String
    useSuffix = Trim$(suffix)
    If Len(useSuffix) = 0 Then useSuffix = DEFAULT_READNOTIFY_SUFFIX
    If Left$(useSuffix, 1) <> "." Then useSuffix = "." & useSuffix

    ' Resolve recipients first so Exchange names become usable SMTP addresses.
    On Error Resume Next
    mail.Recipients.ResolveAll
    Err.Clear
    On Error GoTo EH

    Dim stamped As Long
    Dim n As Long
    stamped = 0

    n = StampRecipientsByType(mail, olTo, useSuffix)
    If n < 0 Then GoTo StampFailed
    stamped = stamped + n

    n = StampRecipientsByType(mail, olCC, useSuffix)
    If n < 0 Then GoTo StampFailed
    stamped = stamped + n

    n = StampRecipientsByType(mail, olBCC, useSuffix)
    If n < 0 Then GoTo StampFailed
    stamped = stamped + n

    On Error Resume Next
    mail.Recipients.ResolveAll
    Err.Clear
    On Error GoTo EH

    MSCANModLogging.WriteLog "ReadNotify: stamped " & stamped & " recipient(s) with " & useSuffix
    ApplyReadNotifyIfRequested = True
    Exit Function

StampFailed:
    MSCANModLogging.WriteLogLevel "error", "ReadNotify: stamping aborted after recipient rewrite failure."
    ApplyReadNotifyIfRequested = False
    Exit Function

EH:
    Dim desc As String
    desc = Err.Description
    If InStr(1, LCase$(desc), "message has been changed", vbTextCompare) > 0 Or _
       (InStr(1, LCase$(desc), "has been changed", vbTextCompare) > 0 And _
        InStr(1, LCase$(desc), "message", vbTextCompare) > 0) Then
        MSCANModLogging.WriteLog "CONFLICT in ReadNotify stamping: #" & Err.Number & " - " & desc
    Else
        MSCANModLogging.WriteLogLevel "error", _
            "ERROR in MSCANReadNotify.ApplyReadNotifyIfRequested | #" & Err.Number & " | " & desc
    End If
    ApplyReadNotifyIfRequested = False
End Function

Public Sub ApplyReadNotifyForMail(ByVal mail As Outlook.MailItem)
    On Error GoTo EH

    Dim tag As String
    Dim idx As Long

    tag = MSCANCore.DetectClassification(mail.Subject)
    idx = MSCANCore.GetClassificationLevel(tag)
    ApplyReadNotifyIfRequired mail, idx
    Exit Sub

EH:
    MSCANModLogging.WriteLogLevel "error", _
        "ERROR in MSCANReadNotify.ApplyReadNotifyForMail | #" & Err.Number & " | " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Returns count stamped, or -1 if a rewrite failed (callers must treat as failure).
Private Function StampRecipientsByType( _
    ByVal mail As Outlook.MailItem, _
    ByVal recipType As Outlook.OlMailRecipientType, _
    ByVal suffix As String) As Long

    On Error GoTo EH

    Dim i As Long
    Dim recip As Outlook.Recipient
    Dim smtp As String
    Dim rebuilt As String
    Dim changed As Long
    Dim keepOriginal As String
    Dim originalField As String
    Dim removed As Boolean

    rebuilt = ""
    keepOriginal = ""
    changed = 0
    removed = False
    originalField = RecipientFieldText(mail, recipType)

    For i = 1 To mail.Recipients.Count
        Set recip = mail.Recipients(i)
        If recip.Type = recipType Then
            smtp = ResolveSmtpAddress(recip)
            If IsValidSmtp(smtp) Then
                If AlreadyStamped(smtp, suffix) Then
                    rebuilt = AppendAddress(rebuilt, smtp)
                Else
                    rebuilt = AppendAddress(rebuilt, smtp & suffix)
                    changed = changed + 1
                End If
            Else
                ' No SMTP — leave the original token alone (do not invent Name.read-notify.com).
                MSCANModLogging.WriteLog "ReadNotify: skipped non-SMTP recipient '" & recip.Name & "'"
                keepOriginal = AppendAddress(keepOriginal, recip.Name)
            End If
        End If
    Next i

    If changed = 0 Then
        StampRecipientsByType = 0
        Exit Function
    End If

    ' Drop existing recipients of this type, then write the stamped SMTP list back.
    For i = mail.Recipients.Count To 1 Step -1
        If mail.Recipients(i).Type = recipType Then
            mail.Recipients.Remove i
        End If
    Next i
    removed = True

    rebuilt = AppendAddress(rebuilt, keepOriginal)
    WriteRecipientField mail, recipType, rebuilt

    StampRecipientsByType = changed
    Exit Function

EH:
    MSCANModLogging.WriteLogLevel "error", _
        "ReadNotify StampRecipientsByType(" & CStr(recipType) & "): #" & Err.Number & " - " & Err.Description
    If removed Then
        On Error Resume Next
        WriteRecipientField mail, recipType, originalField
        mail.Recipients.ResolveAll
        Err.Clear
        MSCANModLogging.WriteLog "ReadNotify: restored original recipient field after failure."
    End If
    StampRecipientsByType = -1
End Function

Private Function RecipientFieldText(ByVal mail As Outlook.MailItem, ByVal recipType As Outlook.OlMailRecipientType) As String
    On Error Resume Next
    Select Case recipType
        Case olTo: RecipientFieldText = mail.To & ""
        Case olCC: RecipientFieldText = mail.CC & ""
        Case olBCC: RecipientFieldText = mail.BCC & ""
        Case Else: RecipientFieldText = ""
    End Select
End Function

Private Sub WriteRecipientField(ByVal mail As Outlook.MailItem, ByVal recipType As Outlook.OlMailRecipientType, ByVal value As String)
    Select Case recipType
        Case olTo: mail.To = value
        Case olCC: mail.CC = value
        Case olBCC: mail.BCC = value
    End Select
End Sub

Private Function ResolveSmtpAddress(ByVal recip As Outlook.Recipient) As String
    On Error Resume Next

    Dim smtp As String
    Dim entry As Outlook.AddressEntry
    Dim exchUser As Outlook.ExchangeUser
    Dim pa As Outlook.PropertyAccessor

    smtp = ""
    Set entry = recip.AddressEntry

    If Not entry Is Nothing Then
        If StrComp(entry.Type, "SMTP", vbTextCompare) = 0 Then
            smtp = entry.Address
        ElseIf StrComp(entry.Type, "EX", vbTextCompare) = 0 Then
            Set exchUser = entry.GetExchangeUser
            If Not exchUser Is Nothing Then
                smtp = exchUser.PrimarySmtpAddress
            End If
            If Not IsValidSmtp(smtp) Then
                Set pa = entry.PropertyAccessor
                Err.Clear
                smtp = pa.GetProperty(PR_SMTP_ADDRESS_W)
                If Not IsValidSmtp(smtp) Then
                    Err.Clear
                    smtp = pa.GetProperty(PR_SMTP_ADDRESS_A)
                End If
            End If
        Else
            smtp = entry.Address
        End If
    End If

    If Not IsValidSmtp(smtp) Then smtp = ExtractSmtpFromText(recip.Address)
    If Not IsValidSmtp(smtp) Then smtp = ExtractSmtpFromText(recip.Name)

    ResolveSmtpAddress = Trim$(smtp)
End Function

Private Function ExtractSmtpFromText(ByVal raw As String) As String
    On Error Resume Next
    Dim s As String
    Dim a As Long
    Dim b As Long

    s = Trim$(raw)
    a = InStr(1, s, "<", vbBinaryCompare)
    b = InStr(1, s, ">", vbBinaryCompare)
    If a > 0 And b > a Then
        s = Trim$(Mid$(s, a + 1, b - a - 1))
    End If

    If InStr(1, s, "@", vbBinaryCompare) = 0 Then
        ExtractSmtpFromText = ""
    Else
        ExtractSmtpFromText = s
    End If
End Function

Private Function DomainOfSmtp(ByVal address As String) As String
    Dim atPos As Long
    Dim s As String
    s = LCase$(Trim$(address))
    atPos = InStrRev(s, "@")
    If atPos < 1 Or atPos >= Len(s) Then
        DomainOfSmtp = ""
    Else
        DomainOfSmtp = Mid$(s, atPos + 1)
    End If
End Function

Private Function IsValidSmtp(ByVal address As String) As Boolean
    Dim s As String
    Dim atPos As Long
    s = Trim$(address)
    atPos = InStr(1, s, "@", vbBinaryCompare)
    IsValidSmtp = (atPos > 1) And _
                  (InStr(atPos + 1, s, "@", vbBinaryCompare) = 0) And _
                  (InStr(1, s, " ", vbBinaryCompare) = 0) And _
                  (InStr(1, s, ",", vbBinaryCompare) = 0) And _
                  (InStr(1, s, ";", vbBinaryCompare) = 0)
End Function

Private Function AlreadyStamped(ByVal address As String, Optional ByVal suffix As String = "") As Boolean
    Dim s As String
    Dim useSuffix As String
    s = LCase$(Trim$(address))
    useSuffix = LCase$(Trim$(suffix))
    If Len(useSuffix) = 0 Then useSuffix = LCase$(DEFAULT_READNOTIFY_SUFFIX)
    If Left$(useSuffix, 1) <> "." Then useSuffix = "." & useSuffix

    ' Must end with the suffix (e.g. .read-notify.com) — not a mid-string hit.
    If Len(s) >= Len(useSuffix) Then
        If Right$(s, Len(useSuffix)) = useSuffix Then
            AlreadyStamped = True
            Exit Function
        End If
    End If
    AlreadyStamped = False
End Function

Private Function AppendAddress(ByVal list As String, ByVal address As String) As String
    address = Trim$(address)
    If Len(address) = 0 Then
        AppendAddress = list
    ElseIf Len(Trim$(list)) = 0 Then
        AppendAddress = address
    Else
        AppendAddress = list & "; " & address
    End If
End Function
