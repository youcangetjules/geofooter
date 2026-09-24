Attribute VB_Name = "MSCANToolbar"
'===============================================================================
' MSCANToolbar - AES (Aliniant Email Scanner) command bar
' Layout (matches Home AES group): large ON/OFF + Short + Full + Deep + GURI + Aura,
' then a BeginGroup so Diagnostics / Settings / View Logs stack on the right.
' Add-ins tab always shows this bar. Home "AES (Custom)" macro pins are static.
'===============================================================================
Option Explicit

Private Const TOOLBAR_NAME As String = "AES"
Private Const LEGACY_TOOLBAR_NAME As String = "Aliniant AES"
Private Const TAG_SERVICE As String = "AES_SERVICE_TOGGLE"
Private Const TAG_FULLSCAN As String = "AES_FULLSCAN"
Private Const TAG_DEEPSCAN As String = "AES_DEEPSCAN"
Private Const TAG_SHORTSCAN As String = "AES_SHORTSCAN"
Private Const TAG_GURI As String = "AES_GURI"
Private Const TAG_AURA As String = "AES_AURA"
Private Const TAG_DIAGNOSTICS As String = "AES_DIAGNOSTICS"
Private Const TAG_SETTINGS As String = "AES_SETTINGS"
Private Const TAG_LOGS As String = "AES_LOGS"

' Cached bar reference: Application.CommandBars("name") often raises #438 on
' modern Outlook, while ActiveExplorer.CommandBars still works.
Private gToolbar As CommandBar

Public Sub CreateToolbar()
    On Error GoTo EH

    Dim bar As CommandBar

    ' Drop legacy + current bars so Home ribbon cannot keep a frozen group.
    DeleteAllNamedToolbars LEGACY_TOOLBAR_NAME
    DeleteAllNamedToolbars TOOLBAR_NAME
    Set gToolbar = Nothing

    Set bar = GetOrCreateToolbar()
    If bar Is Nothing Then
        MSCANModLogging.WriteLog "MSCANToolbar.CreateToolbar: could not get/create toolbar."
        Exit Sub
    End If

    ' Large cluster (left): ON/OFF, Short Scan, Full Scan, Deep Scan
    EnsureServiceButton bar, MSCANModWatchers.IsServiceEnabled()
    EnsureButton bar, TAG_SHORTSCAN, "AES Short Scan", "ShortScanEmail", _
        "AES: append the compact two-line footer on selected email(s)", 1561, False, 1
    ApplyButtonPicture FindButtonByTag(bar, TAG_SHORTSCAN), _
        ResolveIconPath("aes_short_scan.bmp"), ResolveIconPath("aes_short_scan_mask.bmp")

    EnsureButton bar, TAG_FULLSCAN, "AES Full Scan", "FullScanEmail", _
        "AES: full security analysis on selected email(s)", 1020, False, 1
    ApplyButtonPicture FindButtonByTag(bar, TAG_FULLSCAN), _
        ResolveIconPath("aes_full_scan.bmp"), ResolveIconPath("aes_full_scan_mask.bmp")

    EnsureButton bar, TAG_DEEPSCAN, "AES Deep Scan", "DeepScanEmail", _
        "AES: manual deep analysis — opens a report; does not change the email", 1000, False, 1
    ApplyButtonPicture FindButtonByTag(bar, TAG_DEEPSCAN), _
        ResolveIconPath("aes_deep_scan.bmp"), ResolveIconPath("aes_deep_scan_mask.bmp")

    EnsureButton bar, TAG_GURI, "GURI", "ShowGuriGui", _
        "Bring GURI Database Viewer to the foreground (starts it if needed)", 1087, False, 1
    ApplyButtonPicture FindButtonByTag(bar, TAG_GURI), _
        ResolveIconPath("aes_guri.bmp"), ResolveIconPath("aes_guri_mask.bmp")

    EnsureButton bar, TAG_AURA, "Aura", "ShowAuraGui", _
        "Aura — Aliniant Universal Removal Application (data-broker opt-out tracking)", 1088, False, 1
    ApplyButtonPicture FindButtonByTag(bar, TAG_AURA), _
        ResolveIconPath("aes_aura.bmp"), ResolveIconPath("aes_aura_mask.bmp")

    ' Right stack: BeginGroup separates these from the large icons
    EnsureButton bar, TAG_DIAGNOSTICS, "Diagnostics", "RunDiagnostics", _
        "Open interactive AES diagnostics", 1958, True, 2
    EnsureButton bar, TAG_SETTINGS, "Settings", "ShowSettings", _
        "Choose which email accounts AES scans", 548, False, 2
    EnsureButton bar, TAG_LOGS, "View Logs", "ViewLogs", _
        "Open the AES VBA log in Notepad", 19, False, 2

    OrderToolbarButtons bar
    bar.Visible = True

    MSCANModLogging.WriteLog "MSCANToolbar: Toolbar '" & TOOLBAR_NAME & "' ready (" & bar.Controls.Count & " buttons)."
    LogToolbarCaptions bar
    LogToolbarActions bar
    Exit Sub

EH:
    MSCANModLogging.WriteLog "MSCANToolbar.CreateToolbar error: #" & Err.Number & " - " & Err.Description
End Sub

Private Function ToolbarStillValid(ByVal bar As CommandBar) As Boolean
    On Error Resume Next
    Dim n As Long
    If bar Is Nothing Then Exit Function
    Err.Clear
    n = bar.Controls.Count
    ToolbarStillValid = (Err.Number = 0)
End Function

Private Sub DeleteAllNamedToolbars(ByVal barName As String)
    On Error Resume Next

    Dim i As Long
    Dim n As Long
    Dim bar As CommandBar
    Dim exp As Outlook.Explorer
    Dim deleted As Long

    Err.Clear
    n = Application.CommandBars.Count
    If Err.Number = 0 Then
        For i = n To 1 Step -1
            Err.Clear
            Set bar = Application.CommandBars(i)
            If Not bar Is Nothing Then
                If StrComp(bar.Name, barName, vbTextCompare) = 0 Then
                    bar.Delete
                    deleted = deleted + 1
                End If
            End If
        Next i
    End If

    For Each exp In Application.Explorers
        Err.Clear
        n = exp.CommandBars.Count
        If Err.Number <> 0 Then GoTo NextExplorer
        For i = n To 1 Step -1
            Err.Clear
            Set bar = exp.CommandBars(i)
            If Not bar Is Nothing Then
                If StrComp(bar.Name, barName, vbTextCompare) = 0 Then
                    bar.Delete
                    deleted = deleted + 1
                End If
            End If
        Next i
NextExplorer:
    Next exp

    If deleted > 0 Then
        MSCANModLogging.WriteLogInfo "DeleteAllNamedToolbars: removed " & deleted & " '" & barName & "' bar(s)."
    End If
End Sub

' Find toolbar by name on Application or any Explorer (index walk; name lookup often #438).
Private Function FindExistingToolbar() As CommandBar
    On Error Resume Next

    Dim bar As CommandBar
    Dim exp As Outlook.Explorer
    Dim i As Long
    Dim n As Long

    Err.Clear
    n = Application.CommandBars.Count
    If Err.Number = 0 Then
        For i = 1 To n
            Err.Clear
            Set bar = Application.CommandBars(i)
            If Not bar Is Nothing Then
                If StrComp(bar.Name, TOOLBAR_NAME, vbTextCompare) = 0 Then
                    Set FindExistingToolbar = bar
                    Exit Function
                End If
            End If
        Next i
    End If

    For Each exp In Application.Explorers
        Err.Clear
        n = exp.CommandBars.Count
        If Err.Number <> 0 Then GoTo NextFindExplorer
        For i = 1 To n
            Err.Clear
            Set bar = exp.CommandBars(i)
            If Not bar Is Nothing Then
                If StrComp(bar.Name, TOOLBAR_NAME, vbTextCompare) = 0 Then
                    Set FindExistingToolbar = bar
                    Exit Function
                End If
            End If
        Next i
NextFindExplorer:
    Next exp

    Set FindExistingToolbar = Nothing
End Function

Private Function GetOrCreateToolbar() As CommandBar
    On Error Resume Next

    Dim bar As CommandBar
    Dim addErr As Long
    Dim addDesc As String

    If ToolbarStillValid(gToolbar) Then
        Set GetOrCreateToolbar = gToolbar
        Exit Function
    End If
    Set gToolbar = Nothing

    Set bar = FindExistingToolbar()
    If Not bar Is Nothing Then
        Set gToolbar = bar
        Set GetOrCreateToolbar = bar
        Exit Function
    End If

    ' Temporary:=False so Outlook keeps it on the Add-ins tab.
    Err.Clear
    Set bar = Application.CommandBars.Add(Name:=TOOLBAR_NAME, Position:=msoBarTop, MenuBar:=False, Temporary:=False)
    If Not bar Is Nothing And Err.Number = 0 Then
        Set gToolbar = bar
        Set GetOrCreateToolbar = bar
        Exit Function
    End If
    addErr = Err.Number
    addDesc = Err.Description

    If Not Application.ActiveExplorer Is Nothing Then
        Err.Clear
        Set bar = Nothing
        Set bar = Application.ActiveExplorer.CommandBars.Add( _
            Name:=TOOLBAR_NAME, Position:=msoBarTop, MenuBar:=False, Temporary:=False)
        If Not bar Is Nothing And Err.Number = 0 Then
            Set gToolbar = bar
            Set GetOrCreateToolbar = bar
            Exit Function
        End If
        MSCANModLogging.WriteLog "GetOrCreateToolbar Explorer.Add failed: #" & Err.Number & " - " & Err.Description
    ElseIf addErr <> 0 Then
        MSCANModLogging.WriteLog "GetOrCreateToolbar Add failed: #" & addErr & " - " & addDesc
    End If

    Set GetOrCreateToolbar = Nothing
End Function

Private Sub LogToolbarCaptions(ByVal bar As CommandBar)
    On Error Resume Next
    Dim ctl As CommandBarControl
    Dim parts As String
    If bar Is Nothing Then Exit Sub
    For Each ctl In bar.Controls
        If Len(parts) > 0 Then parts = parts & " | "
        parts = parts & ctl.Caption
    Next ctl
    MSCANModLogging.WriteLogInfo "Toolbar captions: " & parts
End Sub

Private Sub LogToolbarActions(ByVal bar As CommandBar)
    On Error Resume Next
    Dim ctl As CommandBarControl
    Dim parts As String
    If bar Is Nothing Then Exit Sub
    For Each ctl In bar.Controls
        If Len(parts) > 0 Then parts = parts & " | "
        parts = parts & ctl.Caption & "->" & ctl.OnAction
    Next ctl
    MSCANModLogging.WriteLogInfo "Toolbar OnAction: " & parts
End Sub

' Outlook Add-ins CommandBars resolve macros like the Macros dialog (Project1.Proc),
' not Module.Proc. Prefer project-qualified bare Public Sub names.
Private Function MacroOnAction(ByVal procName As String) As String
    On Error Resume Next

    Dim proj As String
    Dim bare As String
    Dim dot As Long

    bare = Trim$(procName)
    dot = InStrRev(bare, ".")
    If dot > 0 Then bare = Mid$(bare, dot + 1)

    Err.Clear
    proj = Application.VBE.ActiveVBProject.Name
    If Err.Number <> 0 Or Len(proj) = 0 Then proj = "Project1"

    MacroOnAction = proj & "." & bare
End Function

Private Sub EnsureButton(ByVal bar As CommandBar, ByVal tag As String, ByVal caption As String, _
    ByVal procName As String, ByVal tooltip As String, ByVal faceId As Long, _
    Optional ByVal beginGroup As Boolean = False, Optional ByVal priority As Long = 3)

    On Error Resume Next

    Dim btn As CommandBarButton
    Dim action As String
    Set btn = FindButtonByTag(bar, tag)
    If btn Is Nothing Then
        Set btn = bar.Controls.Add(Type:=msoControlButton, Temporary:=False)
        If btn Is Nothing Then
            MSCANModLogging.WriteLog "EnsureButton: Add failed for " & tag & " #" & Err.Number
            Exit Sub
        End If
    End If

    action = MacroOnAction(procName)
    btn.Tag = tag
    btn.Caption = caption
    btn.OnAction = action
    btn.TooltipText = tooltip
    btn.BeginGroup = beginGroup
    If priority > 0 Then btn.Priority = priority
    Err.Clear
    btn.Style = msoButtonIconAndCaption
    If Err.Number <> 0 Then
        Err.Clear
        btn.Style = msoButtonCaption
    End If
    If faceId > 0 Then
        Err.Clear
        btn.FaceId = faceId
    End If
End Sub

' Force left-to-right: ON/OFF, Short, Full, Deep, GURI, Aura | Diagnostics, Settings, Logs
Private Sub OrderToolbarButtons(ByVal bar As CommandBar)
    On Error Resume Next

    Dim tags As Variant
    Dim i As Long
    Dim btn As CommandBarButton

    If bar Is Nothing Then Exit Sub
    tags = Array(TAG_SERVICE, TAG_SHORTSCAN, TAG_FULLSCAN, TAG_DEEPSCAN, TAG_GURI, TAG_AURA, _
                 TAG_DIAGNOSTICS, TAG_SETTINGS, TAG_LOGS)

    For i = LBound(tags) To UBound(tags)
        Set btn = FindButtonByTag(bar, CStr(tags(i)))
        If Not btn Is Nothing Then
            btn.Move Before:=i + 1
            If CStr(tags(i)) = TAG_DIAGNOSTICS Then
                btn.BeginGroup = True
            Else
                btn.BeginGroup = False
            End If
            ' Large cluster: indices 0..5 (service through Aura)
            If i <= 5 Then
                btn.Priority = 1
            Else
                btn.Priority = 2
            End If
        End If
    Next i
End Sub

Private Function FindButtonByTag(ByVal bar As CommandBar, ByVal tag As String) As CommandBarButton
    On Error Resume Next

    Dim ctl As CommandBarControl
    For Each ctl In bar.Controls
        If ctl.Tag = tag Then
            Set FindButtonByTag = ctl
            Exit Function
        End If
    Next ctl

    Set FindButtonByTag = Nothing
End Function

Private Function ResolveIconPath(ByVal fileName As String) As String
    On Error Resume Next

    Dim paths As Variant
    Dim p As Variant
    Dim fso As Object

    paths = Array( _
        Environ$("LOCALAPPDATA") & "\GeoFooter\icons\" & fileName, _
        MSCANPaths.GetAssetsIconsDir() & "\" & fileName, _
        MSCANPaths.InstallPath("VBA", "icons") & "\" & fileName)

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveIconPath = CStr(p)
            Exit Function
        End If
    Next p
End Function

Private Sub ApplyButtonPicture(ByVal btn As CommandBarButton, ByVal iconPath As String, _
    Optional ByVal maskPath As String = "")
    On Error Resume Next
    If btn Is Nothing Then Exit Sub
    If Len(iconPath) = 0 Then Exit Sub
    If Dir$(iconPath) = "" Then Exit Sub

    Err.Clear
    btn.Style = msoButtonIconAndCaption
    btn.Picture = LoadPicture(iconPath)
    If Len(maskPath) > 0 Then
        If Dir$(maskPath) <> "" Then
            Err.Clear
            btn.Mask = LoadPicture(maskPath)
        End If
    End If
End Sub

Private Sub MigrateLegacyButtons(ByVal bar As CommandBar)
    On Error Resume Next

    Dim ctl As CommandBarControl
    Dim btn As CommandBarButton
    Dim hasTagged As Boolean

    For Each ctl In bar.Controls
        If Len(ctl.Tag) > 0 And Left$(ctl.Tag, 4) = "AES_" Then
            hasTagged = True
            Exit For
        End If
    Next ctl
    If hasTagged Then Exit Sub

    ' Old toolbar: assign tags by caption / OnAction so we update in place (no duplicates).
    For Each ctl In bar.Controls
        Set btn = ctl
        If btn Is Nothing Then GoTo NextLegacy

        If InStr(1, btn.OnAction, "ToggleService", vbTextCompare) > 0 Or _
           InStr(1, btn.Caption, "Activate", vbTextCompare) > 0 Or _
           InStr(1, btn.Caption, "Service", vbTextCompare) > 0 Then
            btn.Tag = TAG_SERVICE
        ElseIf InStr(1, btn.OnAction, "ShortScan", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Short Scan", vbTextCompare) > 0 Then
            btn.Tag = TAG_SHORTSCAN
        ElseIf InStr(1, btn.OnAction, "DeepScan", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Deep Scan", vbTextCompare) > 0 Then
            btn.Tag = TAG_DEEPSCAN
        ElseIf InStr(1, btn.OnAction, "ShowAuraGui", vbTextCompare) > 0 Or _
               StrComp(Trim$(btn.Caption), "Aura", vbTextCompare) = 0 Then
            btn.Tag = TAG_AURA
        ElseIf InStr(1, btn.OnAction, "ShowGuriGui", vbTextCompare) > 0 Or _
               InStr(1, btn.OnAction, "Guri", vbTextCompare) > 0 Or _
               StrComp(Trim$(btn.Caption), "GURI", vbTextCompare) = 0 Then
            btn.Tag = TAG_GURI
        ElseIf InStr(1, btn.OnAction, "FullScan", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Full Scan", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Complete", vbTextCompare) > 0 Then
            If Len(btn.Tag) = 0 Then btn.Tag = TAG_FULLSCAN
        ElseIf InStr(1, btn.OnAction, "Diagnostics", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Diagnostic", vbTextCompare) > 0 Then
            btn.Tag = TAG_DIAGNOSTICS
        ElseIf InStr(1, btn.OnAction, "Settings", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Settings", vbTextCompare) > 0 Then
            btn.Tag = TAG_SETTINGS
        ElseIf InStr(1, btn.OnAction, "ViewLogs", vbTextCompare) > 0 Or _
               InStr(1, btn.OnAction, "OpenLog", vbTextCompare) > 0 Or _
               InStr(1, btn.Caption, "Log", vbTextCompare) > 0 Then
            btn.Tag = TAG_LOGS
        End If
NextLegacy:
    Next ctl
End Sub

Public Sub ReScanEmail()
    FullScanEmail
End Sub

Public Sub FullScanEmail()
    On Error GoTo EH

    Dim explorer As Outlook.explorer
    Set explorer = Application.ActiveExplorer

    If explorer Is Nothing Then
        MSCANModStatus.ShowStatus "Open a mail folder and select email(s) for Full Scan."
        Exit Sub
    End If

    If explorer.Selection.Count = 0 Then
        MSCANModStatus.ShowStatus "Select one or more emails, then click AES Full Scan."
        Exit Sub
    End If

    MSCANModLogging.WriteLog "FullScanEmail: Processing " & explorer.Selection.Count & " selected item(s)."

    Dim completed As Long
    Dim hasFull As Long
    Dim skipped As Long
    Dim i As Long

    For i = 1 To explorer.Selection.Count
        Dim selectedItem As Object
        Set selectedItem = explorer.Selection.Item(i)

        If Not MSCANModule1.IsScannableItem(selectedItem) Then
            skipped = skipped + 1
            GoTo NextFullScanItem
        End If

        Dim mail As Object
        Set mail = selectedItem

        If MSCANModule1.MailHasFullFooter(mail) Then
            hasFull = hasFull + 1
            MSCANModLogging.WriteLog "FullScanEmail: Full footer already present - " & mail.Subject
        Else
            MSCANModule1.ProcessCompleteFooter mail
            completed = completed + 1
            MSCANModLogging.WriteLog "FullScanEmail: Full footer requested - " & mail.Subject
        End If

NextFullScanItem:
        Set mail = Nothing
        Set selectedItem = Nothing
    Next i

    MSCANModStatus.ShowStatus BuildFullScanStatusLine(completed, hasFull, skipped)
    Exit Sub

EH:
    MSCANModLogging.WriteLog "FullScanEmail error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "AES Full Scan failed: " & Err.Description
End Sub

Public Sub CompleteFooterEmail()
    FullScanEmail
End Sub

Public Sub DeepScanEmail()
    On Error GoTo EH

    Dim explorer As Outlook.explorer
    Set explorer = Application.ActiveExplorer

    If explorer Is Nothing Then
        MSCANModStatus.ShowStatus "Open a mail folder and select email(s) for Deep Scan."
        Exit Sub
    End If

    If explorer.Selection.Count = 0 Then
        MSCANModStatus.ShowStatus "Select one or more emails, then click AES Deep Scan."
        Exit Sub
    End If

    MSCANModLogging.WriteLog "DeepScanEmail: Processing " & explorer.Selection.Count & " selected item(s)."

    Dim completed As Long
    Dim skipped As Long
    Dim i As Long

    For i = 1 To explorer.Selection.Count
        Dim selectedItem As Object
        Set selectedItem = explorer.Selection.Item(i)

        If Not MSCANModule1.IsScannableItem(selectedItem) Then
            skipped = skipped + 1
            GoTo NextDeepScanItem
        End If

        Dim mail As Object
        Set mail = selectedItem

        MSCANModule1.ProcessDeepScan mail
        completed = completed + 1
        MSCANModLogging.WriteLog "DeepScanEmail: Deep scan requested - " & mail.Subject

NextDeepScanItem:
        Set mail = Nothing
        Set selectedItem = Nothing
    Next i

    MSCANModStatus.ShowStatus BuildDeepScanStatusLine(completed, skipped)
    Exit Sub

EH:
    MSCANModLogging.WriteLog "DeepScanEmail error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "AES Deep Scan failed: " & Err.Description
End Sub

Private Function BuildDeepScanStatusLine(ByVal completed As Long, ByVal skipped As Long) As String
    If completed = 0 And skipped > 0 Then
        BuildDeepScanStatusLine = "Selection does not contain email messages."
        Exit Function
    End If

    Dim msg As String
    msg = ""

    If completed > 0 Then
        msg = "AES Deep Scan started for " & completed & " email(s) — report opens when ready"
    End If

    If skipped > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & skipped & " non-email item(s) skipped"
    End If

    If Len(msg) = 0 Then msg = "Nothing to process."
    BuildDeepScanStatusLine = msg
End Function

Public Sub ShortScanEmail()
    On Error GoTo EH

    Dim explorer As Outlook.explorer
    Set explorer = Application.ActiveExplorer

    If explorer Is Nothing Then
        MSCANModStatus.ShowStatus "Open a mail folder and select email(s) for AES Short Scan."
        Exit Sub
    End If

    If explorer.Selection.Count = 0 Then
        MSCANModStatus.ShowStatus "Select one or more emails, then click AES Short Scan."
        Exit Sub
    End If

    MSCANModLogging.WriteLog "ShortScanEmail: Processing " & explorer.Selection.Count & " selected item(s)."

    Dim completed As Long
    Dim rescanned As Long
    Dim skipped As Long
    Dim i As Long

    For i = 1 To explorer.Selection.Count
        Dim selectedItem As Object
        Set selectedItem = explorer.Selection.Item(i)

        If Not MSCANModule1.IsScannableItem(selectedItem) Then
            skipped = skipped + 1
            GoTo NextShortScanItem
        End If

        Dim mail As Object
        Set mail = selectedItem

        Dim isRescan As Boolean
        isRescan = MSCANModule1.MailHasFooter(mail)

        ' Clicking Short Scan is an explicit request, so an existing result is
        ' rescanned and replaced rather than left alone.
        ' ProcessEmail is async — footer is applied when Python finishes.
        MSCANModule1.ProcessEmail mail, True

        If isRescan Then
            rescanned = rescanned + 1
            MSCANModLogging.WriteLog "ShortScanEmail: Rescan started, existing footer will be replaced - " & mail.Subject
        Else
            completed = completed + 1
            MSCANModLogging.WriteLog "ShortScanEmail: Compact scan started - " & mail.Subject
        End If

NextShortScanItem:
        Set mail = Nothing
        Set selectedItem = Nothing
    Next i

    MSCANModStatus.ShowStatus BuildShortScanStatusLine(completed, rescanned, skipped)
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ShortScanEmail error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "AES Short Scan failed: " & Err.Description
End Sub

Private Function BuildShortScanStatusLine(ByVal completed As Long, ByVal rescanned As Long, ByVal skipped As Long) As String
    If completed = 0 And rescanned = 0 And skipped > 0 Then
        BuildShortScanStatusLine = "Selection does not contain email messages."
        Exit Function
    End If

    Dim msg As String
    msg = ""

    If completed > 0 Then
        msg = "AES Short Scan started for " & completed & " email(s) — footer applies when ready"
    End If

    If rescanned > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & "rescanning " & rescanned & " email(s) — existing result will be replaced"
    End If

    If skipped > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & skipped & " non-email item(s) skipped"
    End If

    If Len(msg) = 0 Then msg = "Nothing to process."
    BuildShortScanStatusLine = msg
End Function

Private Function BuildFullScanStatusLine(ByVal completed As Long, ByVal hasFull As Long, ByVal skipped As Long) As String
    If completed = 0 And hasFull = 0 And skipped > 0 Then
        BuildFullScanStatusLine = "Selection does not contain email messages."
        Exit Function
    End If

    If completed = 0 And hasFull > 0 And skipped = 0 Then
        BuildFullScanStatusLine = "Selected email(s) already have the full AES footer."
        Exit Function
    End If

    Dim msg As String
    msg = ""

    If completed > 0 Then
        msg = "Full Scan applied to " & completed & " email(s)"
    End If

    If hasFull > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & hasFull & " already have full footer"
    End If

    If skipped > 0 Then
        If Len(msg) > 0 Then msg = msg & "; "
        msg = msg & skipped & " non-email item(s) skipped"
    End If

    If Len(msg) = 0 Then msg = "Nothing to process."
    BuildFullScanStatusLine = msg
End Function

Public Sub RunDiagnostics()
    On Error GoTo EH
    MSCANModLogging.WriteLog "RunDiagnostics: User opened interactive diagnostics."
    MSCANDiagnostics.ShowDiagnosticsDialog
    Exit Sub

EH:
    MSCANModLogging.WriteLog "RunDiagnostics error: #" & Err.Number & " - " & Err.Description
    MsgBox "Diagnostics failed: " & Err.Description, vbExclamation, "AES"
End Sub

Public Sub ShowSettings()
    On Error GoTo EH
    MSCANModLogging.WriteLog "ShowSettings: User opened AES settings."
    MSCANSettings.ShowSettingsDialog
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ShowSettings error: #" & Err.Number & " - " & Err.Description
    MsgBox "Settings failed: " & Err.Description, vbExclamation, "AES"
End Sub

' Large AES ribbon / toolbar: bring GURI GUI to the foreground (or start it).
Public Sub ShowGuriGui()
    On Error GoTo EH

    Dim fso As Object
    Dim shell As Object
    Dim py As String
    Dim script As String
    Dim cmd As String
    Dim paths As Variant
    Dim p As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    py = ""
    paths = Array( _
        MSCANPaths.GetVenvPythonw(), _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\pythonw.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\pythonw.exe", _
        MSCANPaths.GetVenvPython())
    For Each p In paths
        If Len(CStr(p)) > 0 Then
            If fso.FileExists(CStr(p)) Then
                py = CStr(p)
                Exit For
            End If
        End If
    Next p

    script = ""
    paths = Array( _
        MSCANPaths.GetGuriGuiScript(), _
        Environ$("LOCALAPPDATA") & "\GeoFooter\guri_gui.py")
    For Each p In paths
        If Len(CStr(p)) > 0 Then
            If fso.FileExists(CStr(p)) Then
                script = CStr(p)
                Exit For
            End If
        End If
    Next p

    If Len(py) = 0 Or Len(script) = 0 Then
        MSCANModLogging.WriteLog "ShowGuriGui: python or guri_gui.py not found (py=" & py & " script=" & script & ")."
        MSCANModStatus.ShowStatus "GURI GUI not found — set Install root in GURI Database tab"
        Exit Sub
    End If

    ' --raise: single-instance IPC brings an existing window forward; else starts GURI.
    cmd = """" & py & """ """ & script & """ --raise"
    shell.Run cmd, 0, False
    MSCANModLogging.WriteLog "ShowGuriGui: launched " & cmd
    MSCANModStatus.ShowStatus "Opening GURI…"
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ShowGuriGui error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "Could not open GURI: " & Err.Description
End Sub

' Large AES ribbon / toolbar: open Aura (data-broker removal) inside GURI.
Public Sub ShowAuraGui()
    On Error GoTo EH

    Dim fso As Object
    Dim shell As Object
    Dim py As String
    Dim script As String
    Dim cmd As String
    Dim paths As Variant
    Dim p As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")

    py = ""
    paths = Array( _
        MSCANPaths.GetVenvPythonw(), _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\pythonw.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\pythonw.exe", _
        MSCANPaths.GetVenvPython())
    For Each p In paths
        If Len(CStr(p)) > 0 Then
            If fso.FileExists(CStr(p)) Then
                py = CStr(p)
                Exit For
            End If
        End If
    Next p

    script = ""
    paths = Array( _
        MSCANPaths.GetGuriGuiScript(), _
        Environ$("LOCALAPPDATA") & "\GeoFooter\guri_gui.py")
    For Each p In paths
        If Len(CStr(p)) > 0 Then
            If fso.FileExists(CStr(p)) Then
                script = CStr(p)
                Exit For
            End If
        End If
    Next p

    If Len(py) = 0 Or Len(script) = 0 Then
        MSCANModLogging.WriteLog "ShowAuraGui: python or guri_gui.py not found (py=" & py & " script=" & script & ")."
        MSCANModStatus.ShowStatus "Aura GUI not found — set Install root in GURI Database tab"
        Exit Sub
    End If

    ' --aura selects the Aura tab; --raise brings an existing instance forward.
    cmd = """" & py & """ """ & script & """ --raise --aura"
    shell.Run cmd, 0, False
    MSCANModLogging.WriteLog "ShowAuraGui: launched " & cmd
    MSCANModStatus.ShowStatus "Opening Aura…"
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ShowAuraGui error: #" & Err.Number & " - " & Err.Description
    MSCANModStatus.ShowStatus "Could not open Aura: " & Err.Description
End Sub

Public Sub ViewLogs()
    On Error GoTo EH
    MSCANModLogging.WriteLog "ViewLogs: User opened VBA log."
    MSCANModLogging.OpenLog
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ViewLogs error: #" & Err.Number & " - " & Err.Description
    MsgBox "Could not open log: " & Err.Description, vbExclamation, "AES"
End Sub

Public Sub ToggleService()
    On Error GoTo EH

    MSCANModLogging.WriteLogInfo "ToggleService: clicked (Add-ins / toolbar)."

    If MSCANModWatchers.IsServiceEnabled() Then
        MSCANModWatchers.DisableService
        MSCANModStatus.ShowStatus "AES scanning is now OFF."
    Else
        MSCANModWatchers.EnableService
        MSCANModStatus.ShowStatus "AES scanning is ON. Active watchers: " & MSCANModWatchers.GetWatcherCount()
    End If

    MSCANAppBootstrap.RefreshServiceUI
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ToggleService error: #" & Err.Number & " - " & Err.Description
    MsgBox "Could not change service state: " & Err.Description, vbExclamation, "AES"
End Sub

Public Sub UpdateServiceButtonCaption()
    ' Outlook freezes captions on an existing Home-ribbon CommandBar group.
    ' Only a full toolbar wipe + recreate updates what you see.
    On Error Resume Next
    CreateToolbar
    WriteServiceStateFile
End Sub

' State for AesRibbonHost COM add-in (live Home ribbon ON/OFF icons).
Public Sub WriteServiceStateFile()
    On Error Resume Next

    Dim path As String
    Dim fso As Object
    Dim ts As Object
    Dim parent As String
    Dim enabled As String

    path = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_service_state.json"
    Set fso = CreateObject("Scripting.FileSystemObject")
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    End If

    If MSCANModWatchers.IsServiceEnabled() Then
        enabled = "true"
    Else
        enabled = "false"
    End If

    Set ts = fso.CreateTextFile(path, True, False)
    Dim busy As String
    Dim pending As Long
    pending = 0
    On Error Resume Next
    pending = MSCANModule1.PendingAsyncJobCount()
    On Error GoTo 0
    If pending > 0 Then busy = "true" Else busy = "false"

    ts.Write "{""enabled"": " & enabled & _
             ", ""watchers"": " & CLng(MSCANModWatchers.GetWatcherCount()) & _
             ", ""busy"": " & busy & _
             ", ""pending"": " & CLng(pending) & "}"
    ts.Close
End Sub

Public Sub ApplyServiceBusyAppearance()
    ' Lightweight update (no toolbar wipe) when scan busy state changes.
    On Error Resume Next

    Dim bar As CommandBar
    Dim btn As CommandBarButton
    Dim isOn As Boolean
    Dim busy As Boolean
    Dim pending As Long

    Set bar = GetOrCreateToolbar()
    If bar Is Nothing Then Exit Sub
    Set btn = FindButtonByTag(bar, TAG_SERVICE)
    If btn Is Nothing Then Exit Sub

    isOn = MSCANModWatchers.IsServiceEnabled()
    pending = MSCANModule1.PendingAsyncJobCount()
    busy = (pending > 0)

    If busy Then
        btn.Caption = "AES PROC"
        btn.TooltipText = "AES is processing " & pending & " scan(s). Footer/report applies when ready."
        ApplyButtonPicture btn, ResolveIconPath("aes_service_busy.bmp"), ResolveIconPath("aes_service_busy_mask.bmp")
    ElseIf isOn And MSCANModWatchers.GetWatcherCount() = 0 Then
        btn.Caption = "AES ON (0 inboxes!)"
        btn.TooltipText = "AES is ON but NO inboxes are being watched — all accounts are disabled. Open AES Settings and enable at least one account."
        ApplyButtonPicture btn, ResolveIconPath("aes_service_off.bmp"), ResolveIconPath("aes_service_off_mask.bmp")
    ElseIf isOn Then
        btn.Caption = "AES ON"
        btn.TooltipText = "AES scanning is ON (" & MSCANModWatchers.GetWatcherCount() & " inbox(es)). Click to turn OFF."
        ApplyButtonPicture btn, ResolveIconPath("aes_service_on.bmp"), ResolveIconPath("aes_service_on_mask.bmp")
    Else
        btn.Caption = "AES OFF"
        btn.TooltipText = "AES scanning is OFF. Click to turn ON."
        ApplyButtonPicture btn, ResolveIconPath("aes_service_off.bmp"), ResolveIconPath("aes_service_off_mask.bmp")
    End If
End Sub

Private Sub EnsureServiceButton(ByVal bar As CommandBar, ByVal isOn As Boolean)
    On Error Resume Next

    Dim btn As CommandBarButton
    Dim watcherCount As Long
    Dim busy As Boolean
    Dim pending As Long

    If bar Is Nothing Then Exit Sub

    Set btn = FindButtonByTag(bar, TAG_SERVICE)
    If btn Is Nothing Then
        Set btn = bar.Controls.Add(Type:=msoControlButton, Temporary:=False)
    End If
    If btn Is Nothing Then
        MSCANModLogging.WriteLogWarn "EnsureServiceButton: could not add control (#" & Err.Number & ")."
        Exit Sub
    End If

    watcherCount = MSCANModWatchers.GetWatcherCount()
    pending = MSCANModule1.PendingAsyncJobCount()
    busy = (pending > 0)

    btn.Tag = TAG_SERVICE
    btn.OnAction = MacroOnAction("ToggleService")
    btn.BeginGroup = False
    btn.Priority = 1

    If busy Then
        btn.Caption = "AES PROC"
        btn.TooltipText = "AES is processing " & pending & " scan(s). Footer/report applies when ready."
        btn.FaceId = 1087
    ElseIf isOn And watcherCount = 0 Then
        ' Service is ON but nothing is actually being watched — make that visible.
        btn.Caption = "AES ON (0 inboxes!)"
        btn.TooltipText = "AES is ON but NO inboxes are being watched — all accounts are disabled. Open AES Settings and enable at least one account."
        btn.FaceId = 1088  ' warning-style fallback icon
    ElseIf isOn Then
        btn.Caption = "AES ON"
        btn.TooltipText = "AES scanning is ON (" & watcherCount & " inbox(es)). Click to turn OFF."
        btn.FaceId = 1087  ' fallback checkmark if BMP missing
    Else
        btn.Caption = "AES OFF"
        btn.TooltipText = "AES scanning is OFF. Click to turn ON."
        btn.FaceId = 478   ' fallback cross if BMP missing
    End If

    Err.Clear
    btn.Style = msoButtonIconAndCaption
    If Err.Number <> 0 Then
        Err.Clear
        btn.Style = msoButtonCaption
    End If

    ' Yellow (busy) / green tick (ON) / red cross (OFF)
    If busy Then
        ApplyButtonPicture btn, ResolveIconPath("aes_service_busy.bmp"), ResolveIconPath("aes_service_busy_mask.bmp")
    ElseIf isOn Then
        ApplyButtonPicture btn, ResolveIconPath("aes_service_on.bmp"), ResolveIconPath("aes_service_on_mask.bmp")
    Else
        ApplyButtonPicture btn, ResolveIconPath("aes_service_off.bmp"), ResolveIconPath("aes_service_off_mask.bmp")
    End If

    MSCANModLogging.WriteLogInfo "EnsureServiceButton: '" & btn.Caption & "' (service=" & _
        IIf(isOn, "ON", "OFF") & ", busy=" & IIf(busy, "yes", "no") & ", watchers=" & watcherCount & ")"
End Sub

Public Sub RemoveToolbar()
    On Error Resume Next
    DeleteAllNamedToolbars TOOLBAR_NAME
    DeleteAllNamedToolbars LEGACY_TOOLBAR_NAME
    Set gToolbar = Nothing
End Sub
