Attribute VB_Name = "MSCANPaths"
'===============================================================================
' MSCANPaths — install-root resolution (no hard-coded C:\GeoFooter).
'
' Pointer file (written by GURI Database tab / geofooter_paths.set_install_root):
'   %LOCALAPPDATA%\GeoFooter\install_root.txt
'
' Relative layout under install root:
'   VBA\, assets\, datastore\, debuglog\, crashlogs\, scripts\
'
' Per-user runtime (settings, jobs, secrets) stays under %LOCALAPPDATA%\GeoFooter.
'===============================================================================
Option Explicit

Private Const POINTER_FILE As String = "install_root.txt"
Private mInstallRootCache As String

'-------------------------------------------------------------------------------
' Public API
'-------------------------------------------------------------------------------

Public Function GetUserDataDir() As String
    Dim p As String
    p = Environ$("LOCALAPPDATA") & "\GeoFooter"
    If EnsureDir(p) Then
        GetUserDataDir = p
        Exit Function
    End If
    p = Environ$("TEMP") & "\GeoFooter"
    If EnsureDir(p) Then GetUserDataDir = p
End Function

Public Function GetInstallRoot() As String
    Dim fso As Object
    Dim p As String
    Dim line As String

    If Len(mInstallRootCache) > 0 Then
        GetInstallRoot = mInstallRootCache
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 1) Pointer written by GURI / Python
    p = ReadPointerLine()
    If Len(p) > 0 Then
        If fso.FolderExists(p) Then
            mInstallRootCache = p
            GetInstallRoot = p
            Exit Function
        End If
    End If

    ' 2) Common relative probes from this machine (suite markers)
    Dim candidates As Variant
    Dim c As Variant
    candidates = Array( _
        Environ$("GEOFOOTER_ROOT"), _
        Environ$("LOCALAPPDATA") & "\Programs\GeoFooter", _
        Environ$("USERPROFILE") & "\GeoFooter")
    For Each c In candidates
        p = Trim$(CStr(c))
        If Len(p) > 0 Then
            If fso.FolderExists(p) Then
                If fso.FileExists(p & "\VERSION") Or fso.FolderExists(p & "\VBA") Then
                    mInstallRootCache = p
                    GetInstallRoot = p
                    Exit Function
                End If
            End If
        End If
    Next c

    ' 3) Legacy default if still present
    If fso.FolderExists("C:\GeoFooter") Then
        If fso.FileExists("C:\GeoFooter\VERSION") Or fso.FolderExists("C:\GeoFooter\VBA") Then
            mInstallRootCache = "C:\GeoFooter"
            GetInstallRoot = mInstallRootCache
            Exit Function
        End If
    End If

    GetInstallRoot = ""
End Function

Public Sub ClearInstallRootCache()
    mInstallRootCache = ""
End Sub

''' Join install root with relative segments (backslash-separated in VBA).
Public Function InstallPath(ParamArray parts() As Variant) As String
    Dim root As String
    Dim i As Long
    Dim out As String

    root = GetInstallRoot()
    If Len(root) = 0 Then Exit Function
    out = root
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(CStr(parts(i)))) > 0 Then
            out = out & "\" & Trim$(CStr(parts(i)))
        End If
    Next i
    InstallPath = out
End Function

Public Function GetVbaDir() As String
    GetVbaDir = InstallPath("VBA")
End Function

Public Function GetAssetsIconsDir() As String
    GetAssetsIconsDir = InstallPath("assets", "icons")
End Function

Public Function GetDatastoreDir() As String
    Dim p As String
    p = InstallPath("datastore")
    EnsureDir p
    GetDatastoreDir = p
End Function

Public Function GetDebugLogDir() As String
    Dim p As String
    p = InstallPath("debuglog")
    EnsureDir p
    GetDebugLogDir = p
End Function

Public Function GetCrashlogsDir() As String
    Dim p As String
    p = InstallPath("crashlogs")
    EnsureDir p
    GetCrashlogsDir = p
End Function

Public Function GetVenvPythonw() As String
    GetVenvPythonw = InstallPath(".venv", "Scripts", "pythonw.exe")
End Function

Public Function GetVenvPython() As String
    GetVenvPython = InstallPath(".venv", "Scripts", "python.exe")
End Function

Public Function GetGeolocateScript() As String
    GetGeolocateScript = InstallPath("VBA", "geolocate_headers.py")
End Function

Public Function GetGuriGuiScript() As String
    GetGuriGuiScript = InstallPath("guri_gui.py")
End Function

'-------------------------------------------------------------------------------
' Internals
'-------------------------------------------------------------------------------

Private Function ReadPointerLine() As String
    Dim path As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String

    path = GetUserDataDir() & "\" & POINTER_FILE
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function

    On Error GoTo EH
    Set ts = fso.OpenTextFile(path, 1) ' ForReading
    If Not ts.AtEndOfStream Then line = Trim$(ts.ReadLine)
    ts.Close
    If Left$(line, 1) = """" And Right$(line, 1) = """" And Len(line) >= 2 Then
        line = Mid$(line, 2, Len(line) - 2)
    End If
    ReadPointerLine = line
    Exit Function
EH:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
End Function

Private Function EnsureDir(ByVal path As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then fso.CreateFolder path
    EnsureDir = fso.FolderExists(path)
End Function
