Attribute VB_Name = "MSCANModLogging"
Option Explicit

Private Const LOGGING_FILE_NAME As String = "aes_logging.json"

Private mPath As String
Private mLoggingLoaded As Boolean
Private mLoggingEnabled As Boolean
Private mLevelInfo As Boolean
Private mLevelAudit As Boolean
Private mLevelWarn As Boolean
Private mLevelDebug As Boolean

Public Function logPath() As String
    If Len(mPath) = 0 Then
        Dim paths As Variant
        paths = Array( _
            Environ$("LOCALAPPDATA") & "\GeoFooter\Logs", _
            "C:\GeoFooter\Logs", _
            Environ$("TEMP") & "\GeoFooter\Logs")
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim p As Variant
        For Each p In paths
            If EnsureFolderSimple(CStr(p)) Then
                mPath = CStr(p) & "\VBA_Log.txt"
                logPath = mPath
                Exit Function
            End If
        Next p
        mPath = ""
    End If
    logPath = mPath
End Function

Public Function IsLoggingEnabled() As Boolean
    EnsureLoggingSettingsLoaded
    IsLoggingEnabled = mLoggingEnabled
End Function

Public Sub ReloadLoggingSettings()
    mLoggingLoaded = False
    EnsureLoggingSettingsLoaded
End Sub

Public Sub WriteLog(ByVal msg As String)
    WriteLogLevel "info", msg
End Sub

Public Sub WriteLogInfo(ByVal msg As String)
    WriteLogLevel "info", msg
End Sub

Public Sub WriteLogAudit(ByVal msg As String)
    WriteLogLevel "audit", msg
End Sub

Public Sub WriteLogWarn(ByVal msg As String)
    WriteLogLevel "warn", msg
End Sub

Public Sub WriteLogDebug(ByVal msg As String)
    WriteLogLevel "debug", msg
End Sub

Public Sub WriteLogLevel(ByVal level As String, ByVal msg As String)
    On Error GoTo EH

    EnsureLoggingSettingsLoaded
    If Not mLoggingEnabled Then Exit Sub
    If Not IsLevelEnabled(level) Then Exit Sub

    Dim f As Integer
    Dim lvl As String
    lvl = UCase$(Trim$(level))
    If Len(lvl) = 0 Then lvl = "INFO"

    f = FreeFile
    Open logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd HH:nn:ss"); " | "; lvl; " | "; msg
    Close #f
    Exit Sub
EH:
    Debug.Print "WriteLog failed: " & Err.Number & " - " & Err.Description
End Sub

Private Sub EnsureLoggingSettingsLoaded()
    If mLoggingLoaded Then Exit Sub

    ' Defaults: on, info/audit/warn, debug off
    mLoggingEnabled = True
    mLevelInfo = True
    mLevelAudit = True
    mLevelWarn = True
    mLevelDebug = False

    Dim path As String
    path = GetLoggingSettingsPath()
    If Len(path) = 0 Then
        mLoggingLoaded = True
        Exit Sub
    End If

    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Or Not fso.FileExists(path) Then
        mLoggingLoaded = True
        Exit Sub
    End If

    Dim ts As Object
    Dim raw As String
    Set ts = fso.OpenTextFile(path, 1, False)
    If Not ts Is Nothing Then
        raw = ts.ReadAll
        ts.Close
    End If
    On Error GoTo 0

    If Len(raw) > 0 Then
        ApplyLoggingJson raw
    End If
    mLoggingLoaded = True
End Sub

Private Sub ApplyLoggingJson(ByVal raw As String)
    On Error Resume Next

    Dim enabledChunk As String
    enabledChunk = LCase$(ExtractJsonBoolRegion(raw, "enabled"))
    If InStr(1, enabledChunk, "false", vbTextCompare) > 0 Then
        mLoggingEnabled = False
    ElseIf InStr(1, enabledChunk, "true", vbTextCompare) > 0 Then
        mLoggingEnabled = True
    End If

    mLevelInfo = JsonLevelEnabled(raw, "info", mLevelInfo)
    mLevelAudit = JsonLevelEnabled(raw, "audit", mLevelAudit)
    mLevelWarn = JsonLevelEnabled(raw, "warn", mLevelWarn)
    mLevelDebug = JsonLevelEnabled(raw, "debug", mLevelDebug)
End Sub

Private Function JsonLevelEnabled(ByVal raw As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    On Error Resume Next
    Dim region As String
    region = LCase$(ExtractJsonBoolRegion(raw, key))
    If InStr(1, region, "false", vbTextCompare) > 0 Then
        JsonLevelEnabled = False
    ElseIf InStr(1, region, "true", vbTextCompare) > 0 Then
        JsonLevelEnabled = True
    Else
        JsonLevelEnabled = defaultValue
    End If
End Function

Private Function ExtractJsonBoolRegion(ByVal raw As String, ByVal key As String) As String
    On Error Resume Next
    Dim pos As Long
    Dim startPos As Long
    pos = InStr(1, raw, """" & key & """", vbTextCompare)
    If pos = 0 Then Exit Function
    startPos = InStr(pos, raw, ":", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    ExtractJsonBoolRegion = Mid$(raw, startPos + 1, 12)
End Function

Private Function IsLevelEnabled(ByVal level As String) As Boolean
    Select Case LCase$(Trim$(level))
        Case "info": IsLevelEnabled = mLevelInfo
        Case "audit": IsLevelEnabled = mLevelAudit
        Case "warn", "warning": IsLevelEnabled = mLevelWarn
        Case "debug": IsLevelEnabled = mLevelDebug
        Case "error": IsLevelEnabled = True ' always keep errors if logging is on
        Case Else: IsLevelEnabled = mLevelInfo
    End Select
End Function

Private Function GetLoggingSettingsPath() As String
    Dim paths As Variant
    Dim p As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    paths = Array( _
        Environ$("LOCALAPPDATA") & "\GeoFooter\" & LOGGING_FILE_NAME, _
        "C:\GeoFooter\" & LOGGING_FILE_NAME, _
        Environ$("TEMP") & "\GeoFooter\" & LOGGING_FILE_NAME)
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            GetLoggingSettingsPath = CStr(p)
            Exit Function
        End If
    Next p
    ' Prefer LocalAppData path even if file does not exist yet (defaults apply).
    GetLoggingSettingsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\" & LOGGING_FILE_NAME
End Function

Private Function EnsureFolderSimple(ByVal p As String) As Boolean
    On Error GoTo EH
    If Len(p) = 0 Then Exit Function
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    EnsureFolderSimple = True
    Exit Function
EH:
    EnsureFolderSimple = False
End Function

Public Sub OpenLog()
    On Error Resume Next
    OpenTextFileInEditor logPath
End Sub

' Prefer Notepad++; fall back to notepad.exe if it is not installed.
Public Sub OpenTextFileInEditor(ByVal filePath As String)
    On Error Resume Next
    If Len(filePath) = 0 Then Exit Sub

    Dim editor As String
    editor = ResolveNotepadPlusPlus()
    If Len(editor) > 0 Then
        Shell """" & editor & """ """ & filePath & """", vbNormalFocus
        If Err.Number = 0 Then Exit Sub
        Err.Clear
    End If

    Shell "notepad.exe """ & filePath & """", vbNormalFocus
End Sub

Private Function ResolveNotepadPlusPlus() As String
    On Error Resume Next
    ResolveNotepadPlusPlus = ""

    Dim candidates As Variant
    Dim i As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    candidates = Array( _
        Environ$("ProgramFiles") & "\Notepad++\notepad++.exe", _
        Environ$("ProgramFiles(x86)") & "\Notepad++\notepad++.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Notepad++\notepad++.exe", _
        "C:\Program Files\Notepad++\notepad++.exe", _
        "C:\Program Files (x86)\Notepad++\notepad++.exe")

    For i = LBound(candidates) To UBound(candidates)
        If Len(CStr(candidates(i))) > 0 Then
            If fso.FileExists(CStr(candidates(i))) Then
                ResolveNotepadPlusPlus = CStr(candidates(i))
                Exit Function
            End If
        End If
    Next i
End Function
