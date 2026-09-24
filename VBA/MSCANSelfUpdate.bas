Attribute VB_Name = "MSCANSelfUpdate"
'===============================================================================
' MSCANSelfUpdate - re-import AES modules from disk, from INSIDE Outlook.
'
' Outlook often returns Nothing for Application.VBE to external COM callers
' (scripts\Import_VBA_to_Outlook.ps1), even with AccessVBOM=1. In-process VBA
' can still reach the project, so this module does the sync from Alt+F8.
'
' One-time: File > Import File > <install root>\VBA\MSCANSelfUpdate.bas
' Then any time: Alt+F8 > UpdateAesFromDisk, then restart Outlook.
'
' Self-contained on purpose: it must work while the other modules are stale.
'===============================================================================
Option Explicit

Private Const SELF_NAME As String = "MSCANSelfUpdate"
Private Const VBEXT_CT_DOCUMENT As Long = 100
Private Const VBE_CTL_COMPILE As Long = 578
Private Const VBE_CTL_SAVE As Long = 3

Public Sub UpdateAesFromDisk()
    Dim root As String
    Dim vbaDir As String
    Dim files As Collection
    Dim proj As Object
    Dim report As String
    Dim f As Variant
    Dim ok As Long
    Dim failed As Long

    root = ResolveRoot()
    If Len(root) = 0 Then Exit Sub
    vbaDir = root & "\VBA"

    Set files = ReadModuleList(vbaDir & "\IMPORT.txt")
    If files.Count = 0 Then
        MsgBox "No modules listed in " & vbaDir & "\IMPORT.txt.", vbCritical, "AES update"
        Exit Sub
    End If

    On Error Resume Next
    Set proj = Application.VBE.ActiveVBProject
    If proj Is Nothing Then Set proj = Application.VBE.VBProjects(1)
    On Error GoTo 0
    If proj Is Nothing Then
        MsgBox "Cannot reach the VBA project (Application.VBE is Nothing)." & vbCrLf & _
               "Set AccessVBOM=1, restart Outlook, press Alt+F11 once, then retry.", _
               vbCritical, "AES update"
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
           "Now fully quit Outlook (tray too) and reopen.", _
           IIf(failed = 0, vbInformation, vbExclamation), "AES update"
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
