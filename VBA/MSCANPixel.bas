Attribute VB_Name = "MSCANPixel"
'===============================================================================
' MSCANPixel - AES read-tracking pixel on outgoing Aliniant mail
' Adds <img src="{base}/p/{id}.gif"> before </body> and writes
' %LOCALAPPDATA%\GeoFooter\pixels\{id}.json for aes\pixel_tracking.py.
' Settings: %LOCALAPPDATA%\GeoFooter\aes_pixel.ini (enabled=0|1, base_url=...)
'===============================================================================
Option Explicit

Private Const DEFAULT_BASE_URL As String = "https://track.aliniant.net"
Private Const PIXEL_ATTR As String = "data-aes-pixel"

Private Type GUID_T
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID_T) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32" (ByRef pguid As GUID_T) As Long
#End If

Private mLoaded As Boolean
Private mEnabled As Boolean
Private mBaseUrl As String

Private Function BaseDir() As String
    BaseDir = Environ$("LOCALAPPDATA") & "\GeoFooter"
End Function

' Re-read on every send so toggling the ini takes effect without a restart.
Private Sub LoadSettings()
    On Error Resume Next
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim p As Long
    Dim k As String
    Dim v As String

    mEnabled = False
    mBaseUrl = DEFAULT_BASE_URL
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(BaseDir() & "\aes_pixel.ini") Then
        Set ts = fso.OpenTextFile(BaseDir() & "\aes_pixel.ini", 1, False)
        Do While Not ts.AtEndOfStream
            line = Trim$(ts.ReadLine)
            p = InStr(1, line, "=")
            If p > 1 Then
                k = LCase$(Trim$(Left$(line, p - 1)))
                v = Trim$(Mid$(line, p + 1))
                If k = "enabled" Then
                    mEnabled = (v = "1" Or LCase$(v) = "true" Or LCase$(v) = "yes")
                ElseIf k = "base_url" And Len(v) > 0 Then
                    If Right$(v, 1) = "/" Then v = Left$(v, Len(v) - 1)
                    mBaseUrl = v
                End If
            End If
        Loop
        ts.Close
    End If
    mLoaded = True
End Sub

Public Function PixelBaseUrl() As String
    If Not mLoaded Then LoadSettings
    PixelBaseUrl = mBaseUrl
End Function

' True for our own pixel, e.g. quoted back to us inside a reply. Callers keep
' it out of the third-party beacon count; defanging it is still fine.
Public Function IsOwnPixelTag(ByVal tagText As String) As Boolean
    Dim t As String
    t = LCase$(tagText)
    If InStr(1, t, PIXEL_ATTR, vbBinaryCompare) > 0 Then
        IsOwnPixelTag = True
        Exit Function
    End If
    IsOwnPixelTag = (InStr(1, t, LCase$(PixelBaseUrl()) & "/p/", vbBinaryCompare) > 0)
End Function

' Called from HandleItemSend after classification / ReadNotify. Never blocks
' or cancels a send: any failure is logged and the mail goes out unchanged.
Public Sub ApplyOwnPixel(ByVal mail As Outlook.MailItem)
    On Error GoTo EH

    LoadSettings
    If Not mEnabled Then Exit Sub
    If mail Is Nothing Then Exit Sub
    If Not MSCANReadNotify.IsAliniantSend(mail) Then Exit Sub
    If mail.BodyFormat <> olFormatHTML Then
        MSCANModLogging.WriteLog "MSCANPixel: skipped (not an HTML message)."
        Exit Sub
    End If

    Dim t0 As Long
    t0 = MSCANHealth.BeginOp()

    Dim pixelId As String
    pixelId = NewPixelId()
    If Len(pixelId) <> 32 Then
        MSCANModLogging.WriteLog "MSCANPixel: could not create a pixel id; sending without one."
        Exit Sub
    End If

    Dim body As String
    body = StripOwnPixels(mail.HTMLBody)

    Dim tag As String
    tag = "<img src=""" & mBaseUrl & "/p/" & pixelId & ".gif"" width=""1"" height=""1"" alt="""" " & _
          "border=""0"" " & PIXEL_ATTR & "=""1"" style=""width:1px;height:1px;border:0;"">"

    Dim p As Long
    p = InStrRev(body, "</body>", -1, vbTextCompare)
    If p > 0 Then
        body = Left$(body, p - 1) & tag & Mid$(body, p)
    Else
        body = body & tag
    End If
    mail.HTMLBody = body

    WriteLedger pixelId, mail
    MSCANHealth.EndOp "ApplyOwnPixel", t0
    MSCANModLogging.WriteLog "MSCANPixel: tracking pixel " & pixelId & " added."
    Exit Sub

EH:
    MSCANModLogging.WriteLog "MSCANPixel.ApplyOwnPixel error: #" & Err.Number & " - " & Err.Description
End Sub

' Removes earlier AES pixels (quoted from a previous send) so a reply or
' forward reports only its own opens.
Private Function StripOwnPixels(ByVal body As String) As String
    Dim result As String
    Dim pos As Long
    Dim tagStart As Long
    Dim tagEnd As Long
    Dim tagText As String

    result = ""
    pos = 1
    Do
        tagStart = InStr(pos, body, "<img", vbTextCompare)
        If tagStart = 0 Then Exit Do
        tagEnd = InStr(tagStart, body, ">")
        If tagEnd = 0 Then Exit Do
        tagText = Mid$(body, tagStart, tagEnd - tagStart + 1)
        result = result & Mid$(body, pos, tagStart - pos)
        If Not IsOwnPixelTag(tagText) Then result = result & tagText
        pos = tagEnd + 1
    Loop
    StripOwnPixels = result & Mid$(body, pos)
End Function

Private Function NewPixelId() As String
    On Error GoTo EH
    Dim g As GUID_T
    Dim s As String
    Dim i As Long
    If CoCreateGuid(g) <> 0 Then Exit Function
    s = Right$("00000000" & LCase$(Hex$(g.Data1)), 8) & _
        Right$("0000" & LCase$(Hex$(g.Data2)), 4) & _
        Right$("0000" & LCase$(Hex$(g.Data3)), 4)
    For i = 0 To 7
        s = s & Right$("00" & LCase$(Hex$(g.Data4(i))), 2)
    Next i
    NewPixelId = s
    Exit Function
EH:
    NewPixelId = ""
End Function

Private Sub WriteLedger(ByVal pixelId As String, ByVal mail As Outlook.MailItem)
    On Error GoTo EH
    Dim fso As Object
    Dim folder As String
    Dim sender As String
    Dim json As String
    Dim stm As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = BaseDir() & "\pixels"
    If Not fso.FolderExists(BaseDir()) Then fso.CreateFolder BaseDir()
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder

    sender = ""
    If Not mail.SendUsingAccount Is Nothing Then sender = mail.SendUsingAccount.SmtpAddress
    If Len(sender) = 0 Then sender = MSCANReadNotify.GetAliniantAccount()

    json = "{""id"":""" & pixelId & """," & _
           """sent_local"":""" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & """," & _
           """sender"":""" & JsonEsc(sender) & """," & _
           """to"":""" & JsonEsc(mail.To) & """," & _
           """cc"":""" & JsonEsc(mail.CC) & """," & _
           """bcc"":""" & JsonEsc(mail.BCC) & """," & _
           """subject"":""" & JsonEsc(mail.Subject) & """}"

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText json
    stm.SaveToFile folder & "\" & pixelId & ".json", 2
    stm.Close
    Exit Sub
EH:
    MSCANModLogging.WriteLog "MSCANPixel.WriteLedger error: #" & Err.Number & " - " & Err.Description
End Sub

Private Function JsonEsc(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case code
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 8: out = out & "\b"
            Case 9: out = out & "\t"
            Case 10: out = out & "\n"
            Case 12: out = out & "\f"
            Case 13: out = out & "\r"
            Case Is < 32: out = out & "\u" & Right$("0000" & Hex$(code), 4)
            Case Else: out = out & ch
        End Select
    Next i
    JsonEsc = out
End Function
