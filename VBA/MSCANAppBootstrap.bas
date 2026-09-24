Attribute VB_Name = "MSCANAppBootstrap"
'===============================================================================
' MSCANAppBootstrap - Startup, ribbon, and deferred watcher init
' Import via File > Import File. Logic lives here so ThisOutlookSession stays thin.
'===============================================================================
Option Explicit

' Give Outlook time to paint the inbox before AES touches stores/watchers.
Private Const DEFERRED_STARTUP_DELAY As String = "0:00:45"
Private Const TOOLBAR_DELAY As String = "0:00:50"

Private gRibbon As IRibbonUI
Private gStartupScheduled As Boolean

Public Sub HandleApplicationStartup()
    On Error GoTo EH

    MSCANEventHandlers.HandleStartup
    MSCANModWatchers.PrepareForStartup
    MSCANModLogging.WriteLog "Application_Startup: AES (Aliniant Email Scanner) starting."

    gStartupScheduled = False
    ScheduleDeferredStartup
    Exit Sub

EH:
    MSCANModLogging.WriteLog "Application_Startup error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub ScheduleDeferredStartup()
    On Error GoTo EH

    If gStartupScheduled Then Exit Sub

    ' Outlook.Application has no OnTime (Excel-only). Start after a short
    ' DoEvents yield so the UI can paint, then attach watchers.
    gStartupScheduled = True
    MSCANModLogging.WriteLog "Application_Startup: Starting AES (Outlook has no OnTime; no " & DEFERRED_STARTUP_DELAY & " delay)."
    DoEvents
    DeferredStartup
    DoEvents
    DeferredToolbar
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ScheduleDeferredStartup error: #" & Err.Number & " - " & Err.Description
    DeferredStartup
    DeferredToolbar
End Sub

Public Sub DeferredStartup()
    On Error GoTo EH

    gStartupScheduled = False

    ' Quiet period first so any early ItemAdd events are ignored.
    MSCANModQueueManager.InitializeQueueManager
    MSCANModQueueManager.BeginStartupQuietPeriod
    MSCANModWatchers.InitializeWatchersStaggered
    Exit Sub

EH:
    MSCANModLogging.WriteLog "DeferredStartup error: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub DeferredToolbar()
    On Error GoTo EH
    MSCANToolbar.CreateToolbar
    ' Clear stale busy flag from prior session (in-memory jobs do not survive restart).
    MSCANToolbar.WriteServiceStateFile
    MSCANToolbar.ApplyServiceBusyAppearance
    RefreshServiceUI
    Exit Sub
EH:
    MSCANModLogging.WriteLog "DeferredToolbar error: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub HandleApplicationQuit()
    On Error Resume Next

    MSCANModLogging.WriteLog "Application_Quit: AES add-in shutting down."
    MSCANModWatchers.CleanupWatchers
    MSCANToolbar.RemoveToolbar
    MSCANEventHandlers.HandleQuit
End Sub

Public Sub OnRibbonLoad(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
    MSCANModLogging.WriteLog "OnRibbonLoad: Ribbon loaded."
End Sub

Public Sub GetWatcherStatusLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    If Not MSCANModWatchers.IsServiceEnabled() Then
        returnedVal = "AES OFF - use Activate on AES tab"
    ElseIf MSCANModWatchers.AreWatchersActive() Then
        returnedVal = "Scanning ON (" & MSCANModWatchers.GetWatcherCount() & " inboxes)"
    Else
        returnedVal = "AES ON - no watchers (check Settings)"
    End If
End Sub

Public Sub GetWatcherStatusImage(ByVal control As IRibbonControl, ByRef returnedVal)
    GetServiceToggleImage control, returnedVal
End Sub

Public Sub OnWatcherButtonClick(ByVal control As IRibbonControl)
    If Not MSCANModWatchers.IsServiceEnabled() Then
        MSCANModStatus.ShowStatus "AES scanning is OFF. Click Activate on the AES tab."
    Else
        MSCANModWatchers.CheckWatchers
        MSCANModStatus.ShowStatus "AES watchers active: " & MSCANModWatchers.GetWatcherCount() & " inbox(es)."
    End If
    RefreshServiceUI
End Sub

Public Sub OnFullScanClick(ByVal control As IRibbonControl)
    MSCANToolbar.FullScanEmail
End Sub

Public Sub OnDeepScanClick(ByVal control As IRibbonControl)
    MSCANToolbar.DeepScanEmail
End Sub

Public Sub OnShortScanClick(ByVal control As IRibbonControl)
    MSCANToolbar.ShortScanEmail
End Sub

Public Sub OnCompleteFooterClick(ByVal control As IRibbonControl)
    MSCANToolbar.FullScanEmail
End Sub

Public Sub OnDiagnosticsClick(ByVal control As IRibbonControl)
    MSCANToolbar.RunDiagnostics
End Sub

Public Sub OnSettingsClick(ByVal control As IRibbonControl)
    MSCANToolbar.ShowSettings
End Sub

Public Sub OnViewLogsClick(ByVal control As IRibbonControl)
    MSCANToolbar.ViewLogs
End Sub

Public Sub OnServiceToggleClick(ByVal control As IRibbonControl)
    MSCANToolbar.ToggleService
End Sub

Public Sub GetServiceToggleLabel(ByVal control As IRibbonControl, ByRef returnedVal)
    If MSCANModule1.IsScanBusy() Then
        returnedVal = "AES PROC"
    ElseIf MSCANModWatchers.IsServiceEnabled() Then
        returnedVal = "AES ON"
    Else
        returnedVal = "AES OFF"
    End If
End Sub

Public Sub GetServiceToggleImage(ByVal control As IRibbonControl, ByRef returnedVal)
    ' Prefer custom BMP when present; otherwise leave empty (use getImageMso).
    On Error GoTo Fallback
    Dim iconPath As String
    iconPath = GetServiceIconPath()
    If Len(iconPath) > 0 Then
        Set returnedVal = LoadPicture(iconPath)
        Exit Sub
    End If
Fallback:
    Set returnedVal = Nothing
End Sub

Public Sub GetServiceToggleImageMso(ByVal control As IRibbonControl, ByRef returnedVal)
    ' Built-in Office icons when custom BMP not used.
    If MSCANModule1.IsScanBusy() Then
        returnedVal = "HappyFace"
    ElseIf MSCANModWatchers.IsServiceEnabled() Then
        returnedVal = "AcceptInvitation"
    Else
        returnedVal = "DeclineInvitation"
    End If
End Sub

Public Sub GetServiceToggleScreentip(ByVal control As IRibbonControl, ByRef returnedVal)
    If MSCANModule1.IsScanBusy() Then
        returnedVal = "AES is processing scan(s)"
    ElseIf MSCANModWatchers.IsServiceEnabled() Then
        returnedVal = "AES scanning is ON"
    Else
        returnedVal = "AES scanning is OFF"
    End If
End Sub

Public Sub GetServiceToggleSupertip(ByVal control As IRibbonControl, ByRef returnedVal)
    If MSCANModule1.IsScanBusy() Then
        returnedVal = "AES is analyzing header(s) / building footer or deep-scan report. Returns to green/red when finished."
    ElseIf MSCANModWatchers.IsServiceEnabled() Then
        returnedVal = "Automatic inbox scanning is active on " & MSCANModWatchers.GetWatcherCount() & " inbox(es). Click to turn OFF."
    Else
        returnedVal = "Automatic inbox scanning is stopped. Click to turn AES scanning ON."
    End If
End Sub

Public Sub RefreshServiceUI()
    On Error Resume Next

    MSCANToolbar.UpdateServiceButtonCaption
    MSCANToolbar.WriteServiceStateFile

    If MSCANModWatchers.IsServiceEnabled() Then
        MSCANModStatus.ShowPersistentServiceStatus True, MSCANModWatchers.GetWatcherCount()
    Else
        MSCANModStatus.ShowPersistentServiceStatus False
    End If

    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "AESActivateDeactivate"
        gRibbon.InvalidateControl "AESActivateDeactivateTab"
        gRibbon.InvalidateControl "WatcherStatus"
        gRibbon.Invalidate
    Else
        MSCANModLogging.WriteLogDebug "RefreshServiceUI: gRibbon is Nothing (customUI onLoad not hooked)."
    End If

    ' COM Ribbon host (Home large icons) — invalidate after VBA toggle
    NotifyAesRibbonHost
End Sub

' Lightweight refresh when async scans start/finish (yellow busy icon).
Public Sub RefreshScanBusyUI()
    On Error Resume Next

    MSCANToolbar.WriteServiceStateFile
    MSCANToolbar.ApplyServiceBusyAppearance

    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "AESActivateDeactivate"
        gRibbon.InvalidateControl "AESActivateDeactivateTab"
    End If

    NotifyAesRibbonHost
End Sub

Public Sub NotifyAesRibbonHost()
    On Error Resume Next

    Dim addin As Object
    Set addin = Application.COMAddIns.Item("Aliniant.AesRibbonHost")
    If addin Is Nothing Then Exit Sub
    If addin.Object Is Nothing Then Exit Sub
    addin.Object.InvalidateAesRibbon
End Sub

Public Function GetServiceIconPath() As String
    Dim fileName As String
    If MSCANModule1.IsScanBusy() Then
        fileName = "aes_service_busy.bmp"
    ElseIf MSCANModWatchers.IsServiceEnabled() Then
        fileName = "aes_service_on.bmp"
    Else
        fileName = "aes_service_off.bmp"
    End If

    Dim paths As Variant
    Dim p As Variant
    Dim fso As Object

    paths = Array( _
        Environ$("LOCALAPPDATA") & "\GeoFooter\icons\" & fileName, _
        MSCANPaths.GetAssetsIconsDir() & "\" & fileName)

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            GetServiceIconPath = CStr(p)
            Exit Function
        End If
    Next p
End Function

Public Sub ResetWatcher()
    MSCANToolbar.ReScanEmail
End Sub
