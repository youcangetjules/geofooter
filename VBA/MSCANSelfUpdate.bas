Attribute VB_Name = "MSCANSelfUpdate"
'===============================================================================
' MSCANSelfUpdate - re-import AES modules from disk, from INSIDE Outlook.
'
' One-time: File > Import File > <install root>\VBA\MSCANSelfUpdate.bas
' Each update: Alt+F11 once (leave the editor open), then Alt+F8 >
'              UpdateAesFromDisk. Fully quit Outlook afterwards.
'
' On some Outlook builds Application.VBE stays Nothing until the VBA editor
' has been opened in the current session, even with AccessVBOM=1. This module
' sends Alt+F11 and retries; if VBE is still unreachable it falls back to a
' manual import checklist (File > Import does not need AccessVBOM).
'===============================================================================
Option Explicit

Private Const SELF_NAME As String = "MSCANSelfUpdate"
Private Const VBEXT_CT_DOCUMENT As Long = 100
Private Const VBE_CTL_COMPILE As Long = 578
Private Const VBE_CTL_SAVE As Long = 3
Private Const ACCESS_VBOM_KEY As String = "HKCU\Software\Microsoft\Office\16.0\Outlook\Security\AccessVBOM"

Public Sub UpdateAesFromDisk()
    Dim root As String
    Dim vbaDir As String
    Dim files As Collection
    Dim proj As Object
    Dim report As String
    Dim f As Variant
    Dim ok As Long
    Dim failed As Long
    Dim accessVal As String

    root = ResolveRoot()
    If Len(root) = 0 Then Exit Sub
    vbaDir = root & "\VBA"

    Set files = ReadModuleList(vbaDir & "\IMPORT.txt")
    If files.Count = 0 Then
        MsgBox "No modules listed in " & vbaDir & "\IMPORT.txt.", vbCritical, "AES update"
        Exit Sub
    End If

    accessVal = ReadAccessVbom()
    Set proj = GetVbaProject()
    If proj Is Nothing Then
        ' Outlook often leaves VBE unloaded until the editor is opened once.
        NudgeVbeEditor
        SleepMs 800
        Set proj = GetVbaProject()
    End If

    If proj Is Nothing Then
        ShowManualImportGuide vbaDir, files, accessVal
        Exit Sub
    End If

    RemoveStale proj

    For Each f In files
        If ReplaceModule(proj, vbaDir & "\" & CStr(f), report) Then
            ok = ok + 1
        Else
            failed = failed + 1
        End If
    Next f

    RunVbeControl VBE_CTL_COMPILE, "Compile", report
    RunVbeControl VBE_CTL_SAVE, "Save", report

    MsgBox "Imported " & ok & " module(s), " & failed & " failed." & vbCrLf & vbCrLf & _
           report & vbCrLf & _
           "AccessVBOM registry value: " & accessVal & vbCrLf & _
           "Now fully quit Outlook (tray too) and reopen.", _
           IIf(failed = 0, vbInformation, vbExclamation), "AES update"
End Sub

' Opens the VBA folder and copies the remove/import order to the clipboard.
Public Sub ShowAesImportChecklist()
    Dim root As String
    Dim vbaDir As String
    Dim files As Collection

    root = ResolveRoot()
    If Len(root) = 0 Then Exit Sub
    vbaDir = root & "\VBA"
    Set files = ReadModuleList(vbaDir & "\IMPORT.txt")
    ShowManualImportGuide vbaDir, files, ReadAccessVbom()
End Sub

Private Function GetVbaProject() As Object
    Dim vbe As Object
    Dim proj As Object
    On Error Resume Next
    Set vbe = Application.VBE
    If Not vbe Is Nothing Then
        Set proj = vbe.ActiveVBProject
        If proj Is Nothing Then Set proj = vbe.VBProjects(1)
    End If
    On Error GoTo 0
    Set GetVbaProject = proj
End Function

Private Sub NudgeVbeEditor()
    On Error Resume Next
    CreateObject("WScript.Shell").SendKeys "%{F11}"
    On Error GoTo 0
End Sub

Private Sub ShowManualImportGuide(ByVal vbaDir As String, ByVal files As Collection, ByVal accessVal As String)
    Dim f As Variant
    Dim steps As String
    Dim i As Long

    steps = "Outlook will not hand over Application.VBE to macros on this session" & vbCrLf & _
            "(AccessVBOM registry value is already " & accessVal & " - not the problem)." & vbCrLf & vbCrLf & _
            "Do this once, by hand, with the VBA editor open (Alt+F11):" & vbCrLf & vbCrLf & _
            "1. In Project Explorer, REMOVE each existing MSCAN* module" & vbCrLf & _
            "   (right-click > Remove > No). Leave ThisOutlookSession alone." & vbCrLf & _
            "2. File > Import File, in this order:" & vbCrLf

    i = 0
    For Each f In files
        i = i + 1
        steps = steps & "   " & i & ". " & CStr(f) & vbCrLf
    Next f

    steps = steps & vbCrLf & _
            "3. Debug > Compile VBAProject" & vbCrLf & _
            "4. File > Save" & vbCrLf & _
            "5. Fully quit Outlook (tray too) and reopen." & vbCrLf & vbCrLf & _
            "The VBA folder will open; the same list is on the clipboard."

    On Error Resume Next
    CopyTextToClipboard steps
    CreateObject("Shell.Application").Open vbaDir
    On Error GoTo 0

    MsgBox steps, vbExclamation, "AES update - manual import"
End Sub

Private Function ResolveRoot() As String
    Dim p As String
    p = ReadFirstLine(Environ$("LOCALAPPDATA") & "\GeoFooter\install_root.txt")
    If Len(p) = 0 Then p = Trim$(Environ$("GEOFOOTER_ROOT"))
    If Len(p) = 0 Or Not FolderExists(p & "\VBA") Then
        p = Trim$(InputBox("GeoFooter install root (folder that contains VBA\):", "AES update", p))
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Function
    If Not FolderExists(p & "\VBA") Then
        MsgBox "No VBA folder under " & p, vbCritical, "AES update"
        Exit Function
    End If
    ResolveRoot = p
End Function

' Module file names are the "VBA\MSCAN*.bas|cls" lines in IMPORT.txt, in order.
Private Function ReadModuleList(ByVal importTxt As String) As Collection
    Dim result As New Collection
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim modFile As String

    Set ReadModuleList = result
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(importTxt) Then Exit Function

    Set ts = fso.OpenTextFile(importTxt, 1)
    Do While Not ts.AtEndOfStream
        line = Trim$(ts.ReadLine)
        If LCase$(Left$(line, 9)) = "vba\mscan" Then
            modFile = Mid$(line, 5)
            If LCase$(Right$(modFile, 4)) = ".bas" Or LCase$(Right$(modFile, 4)) = ".cls" Then
                If LCase$(modFile) <> LCase$(SELF_NAME & ".bas") Then
                    On Error Resume Next
                    result.Add modFile, LCase$(modFile)
                    On Error GoTo 0
                End If
            End If
        End If
    Loop
    ts.Close
End Function

Private Function ReplaceModule(ByVal proj As Object, ByVal path As String, ByRef report As String) As Boolean
    Dim compName As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    compName = fso.GetBaseName(path)
    If Not fso.FileExists(path) Then
        report = report & "MISSING " & compName & vbCrLf
        Exit Function
    End If

    On Error Resume Next
    RemoveComponent proj, compName
    Err.Clear
    proj.VBComponents.Import path
    If Err.Number <> 0 Then
        report = report & "FAILED " & compName & ": " & Err.Description & vbCrLf
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    ReplaceModule = True
End Function

Private Sub RemoveStale(ByVal proj As Object)
    Dim n As Variant
    For Each n In Array("UserForm1", "frmClassification", "MSCANSecurityRatingForm", _
                        "MSCANmodLogging", "MSCANmodQueueManager", "MSCANmodWatchers", _
                        "MSCANclsEmailWatcher")
        RemoveComponent proj, CStr(n)
    Next n
End Sub

' VBA defers Remove until the running macro ends, so rename first; otherwise the
' fresh import would be named e.g. MSCANPaths1.
Private Sub RemoveComponent(ByVal proj As Object, ByVal compName As String)
    Dim comp As Object
    On Error Resume Next
    Set comp = proj.VBComponents(compName)
    If comp Is Nothing Then Exit Sub
    If comp.Type = VBEXT_CT_DOCUMENT Then Exit Sub
    comp.Name = Left$(compName, 20) & "_old" & Format$(Timer * 100, "0")
    proj.VBComponents.Remove comp
End Sub

Private Sub RunVbeControl(ByVal ctlId As Long, ByVal label As String, ByRef report As String)
    Dim ctl As Object
    On Error Resume Next
    Set ctl = Application.VBE.CommandBars.FindControl(Id:=ctlId)
    If ctl Is Nothing Then
        report = report & label & ": not found - do it manually in the VBA editor." & vbCrLf
        Exit Sub
    End If
    If Not ctl.Enabled Then
        report = report & label & ": nothing to do." & vbCrLf
        Exit Sub
    End If
    ctl.Execute
    If Err.Number <> 0 Then
        report = report & label & ": failed (" & Err.Description & ") - do it manually." & vbCrLf
    Else
        report = report & label & ": done." & vbCrLf
    End If
End Sub

Private Function ReadAccessVbom() As String
    Dim sh As Object
    Dim v As Variant
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    v = sh.RegRead(ACCESS_VBOM_KEY)
    If Err.Number <> 0 Then
        ReadAccessVbom = "(missing)"
    Else
        ReadAccessVbom = CStr(v)
    End If
End Function

Private Function ReadFirstLine(ByVal path As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim s As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path, 1)
    If Not ts.AtEndOfStream Then s = Trim$(ts.ReadLine)
    ts.Close
    If Len(s) >= 2 And Left$(s, 1) = """" And Right$(s, 1) = """" Then s = Mid$(s, 2, Len(s) - 2)
    ReadFirstLine = s
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(path)
End Function

Private Sub CopyTextToClipboard(ByVal text As String)
    Dim html As Object
    Dim cfg As Object
    On Error Resume Next
    Set html = CreateObject("htmlfile")
    Set cfg = html.parentWindow.clipboardData
    cfg.SetData "text", text
    If Err.Number <> 0 Then
        Err.Clear
        CreateObject("WScript.Shell").Run "cmd /c echo " & Replace(Left$(text, 200), "&", "^&") & "| clip", 0, True
    End If
End Sub

Private Sub SleepMs(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do While Timer < t + (ms / 1000!)
        DoEvents
    Loop
End Sub
