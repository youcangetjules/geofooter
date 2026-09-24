Attribute VB_Name = "MSCANModStatus"
'===============================================================================
' MSCANModStatus - Show AES messages in the Outlook status bar
' Note: Outlook.Explorer.Caption is read-only in this Outlook build — do not assign it.
' ON/OFF is shown via status-bar text and the AES ribbon, not the window title.
'===============================================================================
Option Explicit

Private Const STATUS_PREFIX As String = "AES: "
Private Const RESTORE_DELAY_SEC As Long = 5

Private gHasPersistent As Boolean
Private gPersistEnabled As Boolean
Private gPersistWatcherCount As Long
Private gPersistStatusText As String
Private gRestoreDue As Date
Private gRestoreNudgePending As Boolean

' Transient status (footer ready, scan progress, etc.).
' After a short delay, re-applies the last persistent ON/OFF line when one exists.
Public Sub ShowStatus(ByVal msg As String)
    On Error Resume Next

    Dim displayText As String
    displayText = FormatStatusText(msg)
    ApplyStatusToAllExplorers displayText
    MSCANModLogging.WriteLog "Status: " & msg

    If gHasPersistent Then
        gRestoreDue = Now + TimeSerial(0, 0, RESTORE_DELAY_SEC)
        ScheduleStatusRestoreNudge
    End If
End Sub

' Persistent service line in the status bar (and ribbon via separate invalidate).
Public Sub ShowPersistentServiceStatus(ByVal enabled As Boolean, Optional ByVal watcherCount As Long = 0)
    On Error Resume Next

    Dim msg As String
    If Not enabled Then
        msg = "scanning OFF"
        watcherCount = 0
    ElseIf watcherCount <= 0 Then
        msg = "ON - no active watchers"
    Else
        msg = "scanning ON - " & watcherCount & " inbox(es)"
    End If

    gHasPersistent = True
    gPersistEnabled = enabled
    gPersistWatcherCount = watcherCount
    gPersistStatusText = FormatStatusText(msg)
    gRestoreDue = 0
    gRestoreNudgePending = False

    ApplyStatusToAllExplorers gPersistStatusText
    MSCANModLogging.WriteLog "Status: " & msg
End Sub

' Re-apply persisted ON/OFF when a delayed nudge fires (or any ItemLoad nudge).
Public Sub RestorePersistentStatusIfDue()
    On Error Resume Next

    If Not gHasPersistent Then Exit Sub
    If gRestoreDue = 0 Then Exit Sub

    If Now < gRestoreDue Then
        ' Early nudge — reschedule for the remaining time.
        gRestoreNudgePending = False
        ScheduleStatusRestoreNudge
        Exit Sub
    End If

    gRestoreDue = 0
    gRestoreNudgePending = False
    ApplyStatusToAllExplorers gPersistStatusText
End Sub

' Kept for callers; Explorer.Caption cannot be written on this Outlook build.
Public Sub UpdateExplorerCaptions(ByVal isOn As Boolean)
    ' no-op — Caption is read-only
End Sub

Public Sub ClearStatus()
    On Error Resume Next

    Dim exp As Outlook.Explorer
    For Each exp In Application.Explorers
        ClearStatusOnExplorer exp
    Next exp

    gHasPersistent = False
    gPersistStatusText = ""
    gRestoreDue = 0
    gRestoreNudgePending = False
End Sub

'-------------------------------------------------------------------------------
' Internals
'-------------------------------------------------------------------------------

Private Sub ApplyStatusToAllExplorers(ByVal text As String)
    On Error Resume Next

    Dim exp As Outlook.Explorer
    For Each exp In Application.Explorers
        SetStatusOnExplorer exp, text
    Next exp
End Sub

Private Function FormatStatusText(ByVal msg As String) As String
    Dim t As String
    t = Trim$(msg & "")

    ' Strip repeated AES / AES: prefixes from callers so we never show "AES: AES …".
    Do While Len(t) > 0
        If StrComp(Left$(t, 5), "AES: ", vbTextCompare) = 0 Then
            t = Trim$(Mid$(t, 6))
        ElseIf StrComp(Left$(t, 4), "AES ", vbTextCompare) = 0 Then
            t = Trim$(Mid$(t, 5))
        ElseIf StrComp(t, "AES", vbTextCompare) = 0 Then
            t = ""
            Exit Do
        Else
            Exit Do
        End If
    Loop

    If Len(t) = 0 Then
        FormatStatusText = "AES"
    Else
        FormatStatusText = STATUS_PREFIX & t
    End If
End Function

Private Sub SetStatusOnExplorer(ByVal exp As Outlook.Explorer, ByVal text As String)
    On Error Resume Next

    If exp Is Nothing Then Exit Sub

    Dim bar As Office.CommandBar
    Dim ctl As Office.CommandBarControl
    Dim i As Long

    Set bar = FindStatusBar(exp)
    If bar Is Nothing Then Exit Sub

    For Each ctl In bar.Controls
        If ctl.Type = msoControlLabel Then
            ctl.Caption = text
            Exit Sub
        End If
    Next ctl

    ' Prefer a button/label-like control over blindly writing Controls(1)
    ' (which can be a gripper or unrelated chrome on some Outlook builds).
    For i = 1 To bar.Controls.Count
        Set ctl = bar.Controls(i)
        If ctl.Type = msoControlLabel Or ctl.Type = msoControlButton Then
            ctl.Caption = text
            Exit Sub
        End If
    Next i
End Sub

Private Function FindStatusBar(ByVal exp As Outlook.Explorer) As Office.CommandBar
    On Error Resume Next

    Dim bar As Office.CommandBar
    Dim nameHint As String

    Set FindStatusBar = Nothing
    If exp Is Nothing Then Exit Function

    Set bar = exp.CommandBars("Status Bar")
    If Not bar Is Nothing Then
        Set FindStatusBar = bar
        Exit Function
    End If

    Set bar = exp.CommandBars("StatusBar")
    If Not bar Is Nothing Then
        Set FindStatusBar = bar
        Exit Function
    End If

    For Each bar In exp.CommandBars
        nameHint = LCase$(bar.Name & "")
        If InStr(1, nameHint, "status", vbTextCompare) > 0 Then
            Set FindStatusBar = bar
            Exit Function
        End If
    Next bar
End Function

Private Sub ClearStatusOnExplorer(ByVal exp As Outlook.Explorer)
    On Error Resume Next
    SetStatusOnExplorer exp, ""
End Sub

Private Sub ScheduleStatusRestoreNudge()
    On Error GoTo EH

    Dim delayMs As Long
    Dim remainSec As Long
    Dim vbsPath As String
    Dim fso As Object
    Dim ts As Object
    Dim folderPath As String
    Dim shell As Object

    If gRestoreDue = 0 Then Exit Sub
    If gRestoreNudgePending Then Exit Sub

    remainSec = DateDiff("s", Now, gRestoreDue)
    If remainSec < 1 Then remainSec = 1
    If remainSec > 30 Then remainSec = 30
    delayMs = remainSec * 1000&

    vbsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_status_restore.vbs"
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(vbsPath)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    Set ts = fso.CreateTextFile(vbsPath, True, False)
    ts.WriteLine "WScript.Sleep " & CStr(delayMs)
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine "Dim ol, itm"
    ts.WriteLine "Set ol = GetObject(, ""Outlook.Application"")"
    ts.WriteLine "If ol Is Nothing Then WScript.Quit 1"
    ts.WriteLine "Set itm = ol.GetNamespace(""MAPI"").GetDefaultFolder(6).Items.GetFirst"
    ts.WriteLine "If itm Is Nothing Or Err.Number <> 0 Then"
    ts.WriteLine "  Err.Clear"
    ts.WriteLine "  Set itm = ol.GetNamespace(""MAPI"").GetDefaultFolder(5).Items.GetFirst"
    ts.WriteLine "End If"
    ts.Close

    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe //Nologo //B " & Chr$(34) & vbsPath & Chr$(34), 0, False
    gRestoreNudgePending = True
    Exit Sub

EH:
    gRestoreNudgePending = False
End Sub
