Attribute VB_Name = "MSCANModSenderRules"
Option Explicit

' ============================================================================
' MSCANModSenderRules
' Per-sender action rules (aes_sender_rules.json, written by the aes://
' protocol handler behind the footer buttons) plus tracking-beacon blocking
' (aes_beacon_blocking.json, written by the AES Settings dialog).
'
' Rule semantics:
'   block_attachments - quarantine + remove attachments from matching senders
'   block_beacons     - neutralise tracking pixels from matching senders
'   trusted           - exempt from both block lists and from block-all beacons
' ============================================================================

Private Const BLOCKED_PIXEL_PATH As String = "C:\GeoFooter\pages\aes_blocked_pixel.png"

' ---------------------------------------------------------------------------
' Sender identity
' ---------------------------------------------------------------------------

Public Function GetMailSenderSmtp(ByVal mail As Object) As String
    On Error Resume Next
    Dim smtp As String
    smtp = ""

    ' PR_SENDER_SMTP_ADDRESS works for Exchange (EX) senders too.
    Dim pa As Object
    Set pa = mail.PropertyAccessor
    If Not pa Is Nothing Then
        smtp = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
    End If
    If Len(smtp) = 0 Or InStr(smtp, "@") = 0 Then
        smtp = mail.SenderEmailAddress
    End If
    If InStr(smtp, "@") = 0 Then smtp = ""
    GetMailSenderSmtp = LCase$(Trim$(smtp))
End Function

Public Function GetMailSenderDomain(ByVal mail As Object) As String
    Dim smtp As String
    smtp = GetMailSenderSmtp(mail)
    Dim p As Long
    p = InStr(smtp, "@")
    If p > 0 Then GetMailSenderDomain = Mid$(smtp, p + 1) Else GetMailSenderDomain = ""
End Function

' ---------------------------------------------------------------------------
' Config / rules file access
' ---------------------------------------------------------------------------

Private Function LocalGeoFooterDir() As String
    LocalGeoFooterDir = Environ$("LOCALAPPDATA") & "\GeoFooter"
End Function

Private Function ReadTextFileQuick(ByVal filePath As String) As String
    On Error GoTo EH
    ReadTextFileQuick = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Function
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False)
    If Not ts.AtEndOfStream Then ReadTextFileQuick = ts.ReadAll
    ts.Close
    Exit Function
EH:
    ReadTextFileQuick = ""
End Function

Private Function SenderRulesJson() As String
    SenderRulesJson = ReadTextFileQuick(LocalGeoFooterDir() & "\aes_sender_rules.json")
End Function

Private Function BeaconConfigJson() As String
    BeaconConfigJson = ReadTextFileQuick(LocalGeoFooterDir() & "\aes_beacon_blocking.json")
End Function

' Crude but sufficient JSON helpers for our own well-formed config files.
Private Function JsonBoolValue(ByVal jsonText As String, ByVal key As String, ByVal defaultVal As Boolean) As Boolean
    JsonBoolValue = defaultVal
    Dim p As Long
    p = InStr(1, jsonText, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    Dim tail As String
    tail = Mid$(jsonText, p + Len(key) + 2, 24)
    If InStr(1, tail, "true", vbTextCompare) > 0 And _
       (InStr(1, tail, "false", vbTextCompare) = 0 Or _
        InStr(1, tail, "true", vbTextCompare) < InStr(1, tail, "false", vbTextCompare)) Then
        JsonBoolValue = True
    ElseIf InStr(1, tail, "false", vbTextCompare) > 0 Then
        JsonBoolValue = False
    End If
End Function

Private Function JsonStringValue(ByVal jsonText As String, ByVal key As String, ByVal defaultVal As String) As String
    JsonStringValue = defaultVal
    Dim p As Long
    p = InStr(1, jsonText, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(key) + 2, jsonText, """")
    If p = 0 Then Exit Function
    Dim q As Long
    q = InStr(p + 1, jsonText, """")
    If q > p Then JsonStringValue = Mid$(jsonText, p + 1, q - p - 1)
End Function

' Raw "[ ... ]" body of a JSON list value (lowercased).
Private Function JsonListBody(ByVal jsonText As String, ByVal key As String) As String
    JsonListBody = ""
    Dim p As Long
    p = InStr(1, jsonText, """" & key & """", vbTextCompare)
    If p = 0 Then Exit Function
    Dim s As Long
    s = InStr(p, jsonText, "[")
    If s = 0 Then Exit Function
    Dim e As Long
    e = InStr(s, jsonText, "]")
    If e = 0 Then Exit Function
    JsonListBody = LCase$(Mid$(jsonText, s + 1, e - s - 1))
End Function

Private Function ListBodyContains(ByVal listBody As String, ByVal valueLower As String) As Boolean
    ListBodyContains = False
    If Len(listBody) = 0 Or Len(valueLower) = 0 Then Exit Function
    ListBodyContains = (InStr(1, listBody, """" & valueLower & """", vbTextCompare) > 0)
End Function

' True when the sender email OR its domain is on the named rules list.
Public Function SenderOnRulesList(ByVal listName As String, ByVal senderEmail As String, ByVal senderDomain As String) As Boolean
    On Error Resume Next
    SenderOnRulesList = False
    Dim body As String
    body = JsonListBody(SenderRulesJson(), listName)
    If Len(body) = 0 Then Exit Function
    If ListBodyContains(body, LCase$(senderEmail)) Then SenderOnRulesList = True: Exit Function
    If ListBodyContains(body, LCase$(senderDomain)) Then SenderOnRulesList = True
End Function

Public Function IsSenderTrusted(ByVal mail As Object) As Boolean
    IsSenderTrusted = SenderOnRulesList("trusted", GetMailSenderSmtp(mail), GetMailSenderDomain(mail))
End Function

' ---------------------------------------------------------------------------
' Attachment blocking
' ---------------------------------------------------------------------------

Public Function ShouldBlockAttachments(ByVal mail As Object) As Boolean
    On Error Resume Next
    ShouldBlockAttachments = False
    Dim smtp As String, dom As String
    smtp = GetMailSenderSmtp(mail)
    dom = GetMailSenderDomain(mail)
    If Len(smtp) = 0 And Len(dom) = 0 Then Exit Function
    If SenderOnRulesList("trusted", smtp, dom) Then Exit Function
    ShouldBlockAttachments = SenderOnRulesList("block_attachments", smtp, dom)
End Function

' Save real (non-inline) attachments to quarantine, then remove them.
' Returns the number quarantined.
Public Function QuarantineAttachments(ByVal mail As Object) As Long
    On Error GoTo EH
    QuarantineAttachments = 0
    If mail.Attachments.count = 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim qDir As String
    qDir = LocalGeoFooterDir() & "\Quarantine\" & Format(Now, "yyyymmdd_hhnnss") & "_" & _
           Format(Int((Timer - Int(Timer)) * 1000), "000")

    Dim removed As Long
    removed = 0
    Dim i As Long
    For i = mail.Attachments.count To 1 Step -1
        Dim att As Object
        Set att = mail.Attachments.item(i)
        If Not IsInlineAttachment(att) Then
            If removed = 0 Then EnsureQuarantineDir fso, qDir
            On Error Resume Next
            att.SaveAsFile qDir & "\" & SafeFileName(att.fileName)
            If Err.Number = 0 Then
                att.Delete
                If Err.Number = 0 Then removed = removed + 1
            End If
            Err.Clear
            On Error GoTo EH
        End If
        Set att = Nothing
    Next i

    If removed > 0 Then
        On Error Resume Next
        mail.Save
        On Error GoTo EH
        MSCANModLogging.WriteLog "QuarantineAttachments: removed " & removed & _
            " attachment(s) from blocked sender " & GetMailSenderSmtp(mail) & " -> " & qDir
    End If
    QuarantineAttachments = removed
    Exit Function
EH:
    MSCANModLogging.WriteLog "QuarantineAttachments error: #" & Err.Number & " - " & Err.Description
    QuarantineAttachments = 0
End Function

Private Sub EnsureQuarantineDir(ByVal fso As Object, ByVal qDir As String)
    On Error Resume Next
    Dim parent As String
    parent = LocalGeoFooterDir() & "\Quarantine"
    If Not fso.FolderExists(LocalGeoFooterDir()) Then fso.CreateFolder LocalGeoFooterDir()
    If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    If Not fso.FolderExists(qDir) Then fso.CreateFolder qDir
End Sub

Private Function IsInlineAttachment(ByVal att As Object) As Boolean
    On Error Resume Next
    IsInlineAttachment = False
    Dim cid As String
    cid = ""
    cid = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
    If Len(cid) > 0 Then IsInlineAttachment = True
End Function

Private Function SafeFileName(ByVal fileName As String) As String
    Dim bad As Variant
    Dim result As String
    result = fileName
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        result = Replace(result, CStr(bad), "_")
    Next bad
    If Len(result) = 0 Then result = "attachment.bin"
    SafeFileName = result
End Function

' ---------------------------------------------------------------------------
' Beacon blocking
' ---------------------------------------------------------------------------

Public Function ShouldBlockBeacons(ByVal mail As Object) As Boolean
    On Error Resume Next
    ShouldBlockBeacons = False

    Dim smtp As String, dom As String
    smtp = GetMailSenderSmtp(mail)
    dom = GetMailSenderDomain(mail)
    If SenderOnRulesList("trusted", smtp, dom) Then Exit Function

    ' Per-sender rule always wins, even when the global switch is off.
    If SenderOnRulesList("block_beacons", smtp, dom) Then
        ShouldBlockBeacons = True
        Exit Function
    End If

    Dim cfg As String
    cfg = BeaconConfigJson()
    If Len(cfg) = 0 Then Exit Function
    If Not JsonBoolValue(cfg, "enabled", False) Then Exit Function
    If Not JsonBoolValue(cfg, "block_all", True) Then Exit Function

    ' Whitelisted sender domains never get blocked.
    If Len(dom) > 0 Then
        If ListBodyContains(JsonListBody(cfg, "whitelist_domains"), dom) Then Exit Function
    End If
    ShouldBlockBeacons = True
End Function

Public Function BeaconBlockingMode() As String
    Dim mode As String
    mode = LCase$(JsonStringValue(BeaconConfigJson(), "mode", "defang"))
    If mode <> "strip" Then mode = "defang"
    BeaconBlockingMode = mode
End Function

' Neutralise tracking beacons in the mail's HTML body.
' Returns the number of beacon tags stripped / defanged.
Public Function NeutralizeBeaconsInMail(ByVal mail As Object) As Long
    On Error GoTo EH
    NeutralizeBeaconsInMail = 0
    If mail.BodyFormat <> olFormatHTML Then Exit Function

    Dim body As String
    body = mail.HTMLBody
    If Len(body) = 0 Then Exit Function

    Dim mode As String
    mode = BeaconBlockingMode()

    Dim blocked As Long
    blocked = 0
    body = NeutralizeTagsIn(body, "<img", mode, blocked)
    body = NeutralizeTagsIn(body, "<v:imagedata", mode, blocked)

    If blocked > 0 Then
        mail.HTMLBody = body
        On Error Resume Next
        mail.Save
        On Error GoTo EH
        MSCANModLogging.WriteLog "NeutralizeBeaconsInMail: " & blocked & " beacon(s) " & _
            IIf(mode = "strip", "stripped", "defanged") & " for sender " & GetMailSenderSmtp(mail)
    End If
    NeutralizeBeaconsInMail = blocked
    Exit Function
EH:
    MSCANModLogging.WriteLog "NeutralizeBeaconsInMail error: #" & Err.Number & " - " & Err.Description
    NeutralizeBeaconsInMail = 0
End Function

Private Function NeutralizeTagsIn(ByVal body As String, ByVal tagPrefix As String, _
                                  ByVal mode As String, ByRef blocked As Long) As String
    Dim result As String
    Dim pos As Long
    Dim tagStart As Long
    Dim tagEnd As Long
    Dim tagText As String

    result = ""
    pos = 1
    Do
        tagStart = InStr(pos, body, tagPrefix, vbTextCompare)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, body, ">")
        If tagEnd = 0 Then Exit Do

        tagText = Mid$(body, tagStart, tagEnd - tagStart + 1)
        result = result & Mid$(body, pos, tagStart - pos)

        If MSCANModule1.IsLikelyTrackingBeacon(tagText) Then
            blocked = blocked + 1
            If mode = "strip" Then
                result = result & "<!--AES beacon blocked-->"
            Else
                result = result & DefangBeaconTag(tagText)
            End If
        Else
            result = result & tagText
        End If
        pos = tagEnd + 1
    Loop
    result = result & Mid$(body, pos)
    NeutralizeTagsIn = result
End Function

' Swap the remote src for a local inert pixel so layout survives but nothing
' pings home.
Private Function DefangBeaconTag(ByVal tagText As String) As String
    Dim p As Long
    Dim q As Long
    Dim ch As String
    Dim inert As String
    inert = "file:///" & Replace(BLOCKED_PIXEL_PATH, "\", "/")

    p = InStr(1, tagText, "src=", vbTextCompare)
    If p = 0 Then
        DefangBeaconTag = "<!--AES beacon blocked-->"
        Exit Function
    End If

    p = p + 4
    ch = Mid$(tagText, p, 1)
    If ch = """" Or ch = "'" Then
        q = InStr(p + 1, tagText, ch)
        If q > p Then
            DefangBeaconTag = Left$(tagText, p) & inert & Mid$(tagText, q)
            Exit Function
        End If
    Else
        q = InStr(p, tagText, " ")
        If q = 0 Then q = InStr(p, tagText, ">")
        If q > p Then
            DefangBeaconTag = Left$(tagText, p - 1) & """" & inert & """" & Mid$(tagText, q)
            Exit Function
        End If
    End If
    DefangBeaconTag = "<!--AES beacon blocked-->"
End Function
