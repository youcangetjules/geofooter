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
' Feature-flag segments the classification dialog can compose into the suffix,
' always dot-separated and immediately before the domain. Must match
' build_read_notify_suffix in aes\classify_dialog.py: an unlisted flag would
' survive StripReadNotifyStamp and be stamped over, producing a double suffix.
Private Const READNOTIFY_FLAGS As String = "silent,noprint,translate,certified,selfdestruct,ensured"
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

    Dim fields(0 To 2) As String
    Dim changed(0 To 2) As Long
    Dim originals(0 To 2) As String
    Dim slot As Long
    Dim stamped As Long

    For slot = 0 To 2
        originals(slot) = RecipientFieldText(mail, SlotRecipientType(slot))
    Next slot

    ' Compute every field before touching the item, so a recipient that cannot
    ' be represented aborts the send-time rewrite while the item is still intact.
    If Not BuildStampedFields(mail, useSuffix, fields, changed) Then GoTo StampFailed

    stamped = changed(0) + changed(1) + changed(2)
    If stamped = 0 Then
        MSCANModLogging.WriteLog "ReadNotify: recipients already carry " & useSuffix & "; nothing to rewrite."
        ApplyReadNotifyIfRequested = True
        Exit Function
    End If

    ' Write all three fields under one handler. Restoring only the field that
    ' failed would leave the message half-stamped and still send it.
    On Error GoTo ApplyFailed
    For slot = 0 To 2
        If changed(slot) > 0 Then
            RemoveRecipientsOfType mail, SlotRecipientType(slot)
            WriteRecipientField mail, SlotRecipientType(slot), fields(slot)
        End If
    Next slot

    On Error Resume Next
    mail.Recipients.ResolveAll
    Err.Clear
    On Error GoTo EH

    MSCANModLogging.WriteLog "ReadNotify: stamped " & stamped & " recipient(s) with " & useSuffix
    ApplyReadNotifyIfRequested = True
    Exit Function

ApplyFailed:
    MSCANModLogging.WriteLogLevel "error", _
        "ReadNotify: recipient write failed (#" & Err.Number & " - " & Err.Description & _
        "); restoring all recipient fields."
    On Error Resume Next
    ' Only the fields we wrote. RecipientFieldText yields "" if it could not read
    ' a field, so restoring an untouched one could clear live recipients.
    For slot = 0 To 2
        If changed(slot) > 0 Then
            WriteRecipientField mail, SlotRecipientType(slot), originals(slot)
        End If
    Next slot
    mail.Recipients.ResolveAll
    Err.Clear
    ApplyReadNotifyIfRequested = False
    Exit Function

StampFailed:
    MSCANModLogging.WriteLogLevel "error", "ReadNotify: stamping aborted, recipients left untouched."
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

' Builds the stamped To/CC/BCC field text in a single pass over Recipients,
' without mutating the item. fields() and changed() are indexed by slot.
'
' Returns False if any recipient cannot be represented in the rebuilt text. The
' caller must then leave the recipients alone: writing the field back would drop
' that recipient silently, and losing tracking is far better than losing them.
Private Function BuildStampedFields( _
    ByVal mail As Outlook.MailItem, _
    ByVal suffix As String, _
    ByRef fields() As String, _
    ByRef changed() As Long) As Boolean

    On Error GoTo EH

    Dim i As Long
    Dim slot As Long
    Dim recip As Outlook.Recipient
    Dim smtp As String
    Dim stampedAddr As String
    Dim kept(0 To 2) As String

    BuildStampedFields = False

    For slot = 0 To 2
        fields(slot) = ""
        changed(slot) = 0
        kept(slot) = ""
    Next slot

    For i = 1 To mail.Recipients.Count
        Set recip = mail.Recipients(i)
        slot = RecipientTypeSlot(recip.Type)
        If slot >= 0 Then
            smtp = ResolveSmtpAddress(recip)
            If IsValidSmtp(smtp) Then
                ' Rebuild from the unstamped address so changing the feature
                ' flags replaces the previous suffix instead of appending a
                ' second one.
                stampedAddr = StripReadNotifyStamp(smtp) & suffix
                fields(slot) = AppendAddress(fields(slot), stampedAddr)
                If StrComp(stampedAddr, smtp, vbTextCompare) <> 0 Then
                    changed(slot) = changed(slot) + 1
                End If
            ElseIf Len(Trim$(recip.Name)) > 0 Then
                ' No SMTP - keep the original token rather than inventing
                ' Name.read-notify.com.
                MSCANModLogging.WriteLog "ReadNotify: skipped non-SMTP recipient '" & recip.Name & "'"
                kept(slot) = AppendAddress(kept(slot), recip.Name)
            Else
                MSCANModLogging.WriteLogLevel "error", _
                    "ReadNotify: recipient " & i & " has neither an SMTP address nor a name; " & _
                    "abandoning stamping so it cannot be dropped."
                Exit Function
            End If
        End If
    Next i

    For slot = 0 To 2
        fields(slot) = AppendAddress(fields(slot), kept(slot))
    Next slot

    BuildStampedFields = True
    Exit Function

EH:
    MSCANModLogging.WriteLogLevel "error", _
        "ReadNotify BuildStampedFields: #" & Err.Number & " - " & Err.Description
    BuildStampedFields = False
End Function

Private Sub RemoveRecipientsOfType(ByVal mail As Outlook.MailItem, ByVal recipType As Outlook.OlMailRecipientType)
    Dim i As Long
    For i = mail.Recipients.Count To 1 Step -1
        If mail.Recipients(i).Type = recipType Then
            mail.Recipients.Remove i
        End If
    Next i
End Sub

Private Function SlotRecipientType(ByVal slot As Long) As Outlook.OlMailRecipientType
    Select Case slot
        Case 0: SlotRecipientType = olTo
        Case 1: SlotRecipientType = olCC
        Case Else: SlotRecipientType = olBCC
    End Select
End Function

' -1 for anything that is not a To/CC/BCC recipient.
Private Function RecipientTypeSlot(ByVal recipType As Outlook.OlMailRecipientType) As Long
    Select Case recipType
        Case olTo: RecipientTypeSlot = 0
        Case olCC: RecipientTypeSlot = 1
        Case olBCC: RecipientTypeSlot = 2
        Case Else: RecipientTypeSlot = -1
    End Select
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

' Reduces a possibly-stamped address back to the real SMTP address.
'
' Testing only for the suffix about to be applied is not enough: adding a
' feature flag makes the shorter existing stamp fail the test and get a second
' suffix appended, while removing one leaves the stale flag in place because the
' shorter suffix still matches the tail. Peeling the domain and then any known
' flag segments handles both directions.
Private Function StripReadNotifyStamp(ByVal address As String) As String
    Dim s As String
    Dim tail As String
    Dim flags() As String
    Dim f As Variant
    Dim cut As Boolean

    s = Trim$(address)

    tail = "." & READNOTIFY_DOMAIN
    If Len(s) <= Len(tail) Then
        StripReadNotifyStamp = s
        Exit Function
    End If
    If StrComp(Right$(s, Len(tail)), tail, vbTextCompare) <> 0 Then
        StripReadNotifyStamp = s
        Exit Function
    End If
    s = Left$(s, Len(s) - Len(tail))

    flags = Split(READNOTIFY_FLAGS, ",")
    Do
        cut = False
        For Each f In flags
            tail = "." & Trim$(CStr(f))
            ' Strictly greater: never consume the whole address.
            If Len(s) > Len(tail) Then
                If StrComp(Right$(s, Len(tail)), tail, vbTextCompare) = 0 Then
                    s = Left$(s, Len(s) - Len(tail))
                    cut = True
                    Exit For
                End If
            End If
        Next f
    Loop While cut

    StripReadNotifyStamp = s
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
