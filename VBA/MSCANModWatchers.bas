Attribute VB_Name = "MSCANModWatchers"
Option Explicit

' Outlook.Application does not support Application.OnTime (Excel-only).
' Staggered attach previously used OnTime and aborted after the first inbox (#438).
' Between accounts we pause ~200ms (MSCANIdle.WaitMs, never a DoEvents spin).

Private Const ATTACH_YIELD_SEC As Double = 0.2
Private Const STATE_FILE_NAME As String = "aes_service_state.json"

Private gWatchers As Collection
Private gServiceEnabled As Boolean
Private gPendingAccounts As Collection

' Called at Application_Startup before deferred watcher attach.
' Restores ON/OFF from aes_service_state.json (default ON if missing/unreadable).
Public Sub PrepareForStartup()
    On Error Resume Next
    gServiceEnabled = ReadPersistedServiceEnabled(True)
    MSCANModLogging.WriteLog "PrepareForStartup: service enabled=" & CStr(gServiceEnabled) & _
        " (from " & ServiceStatePath() & ")"
End Sub

Public Function IsServiceEnabled() As Boolean
    IsServiceEnabled = gServiceEnabled
End Function

Public Sub EnableService()
    gServiceEnabled = True
    MSCANModLogging.WriteLog "AES service enabled."
    InitializeWatchersStaggered
    PersistServiceEnabled True
End Sub

Public Sub DisableService()
    gServiceEnabled = False
    MSCANModLogging.WriteLog "AES service disabled."
    CleanupWatchers
    PersistServiceEnabled False
End Sub

Public Sub SetServiceEnabled(ByVal enabled As Boolean)
    If enabled Then
        EnableService
    Else
        DisableService
    End If
End Sub

Public Function GetWatcherCount() As Long
    If gWatchers Is Nothing Then
        GetWatcherCount = 0
    Else
        GetWatcherCount = gWatchers.Count
    End If
End Function

Public Function AreWatchersActive() As Boolean
    AreWatchersActive = (GetWatcherCount() > 0)
End Function

' Legacy name - same path as staggered so queue/watcher init cannot diverge.
Public Sub InitializeWatchers()
    InitializeWatchersStaggered
End Sub

' Attach enabled inboxes with a short yield between each.
Public Sub InitializeWatchersStaggered()
    On Error GoTo EH

    If Not gServiceEnabled Then
        MSCANModLogging.WriteLog "InitializeWatchersStaggered: Skipped because service is disabled."
        Exit Sub
    End If

    CleanupWatchers
    MSCANModQueueManager.InitializeQueueManager
    Set gWatchers = New Collection
    Set gPendingAccounts = New Collection

    Dim acc As Outlook.Account
    Dim label As String
    For Each acc In Application.Session.Accounts
        label = ""
        On Error Resume Next
        label = acc.DisplayName
        If MSCANSettings.IsAccountScanEnabled(acc) Then
            gPendingAccounts.Add acc
            Err.Clear
            On Error GoTo EH
        Else
            MSCANModLogging.WriteLog "InitializeWatchersStaggered: Skipping disabled account " & label
            Err.Clear
            On Error GoTo EH
        End If
    Next acc

    MSCANModLogging.WriteLog "InitializeWatchersStaggered: " & gPendingAccounts.Count & " account(s) queued for attach."
    AttachAllPendingWatchers
    Exit Sub

EH:
    MSCANModLogging.WriteLog "InitializeWatchersStaggered error: #" & Err.Number & " - " & Err.Description
End Sub

' Public for any callers that still invoke the old one-at-a-time name.
Public Sub AttachNextPendingWatcher()
    AttachAllPendingWatchers
End Sub

Private Sub AttachAllPendingWatchers()
    On Error GoTo EH

    If Not gServiceEnabled Then Exit Sub
    If gPendingAccounts Is Nothing Then Exit Sub

    Dim acc As Outlook.Account
    Dim first As Boolean
    first = True

    Do While gPendingAccounts.Count > 0
        If Not first Then YieldBriefly ATTACH_YIELD_SEC
        first = False

        Set acc = gPendingAccounts.Item(1)
        gPendingAccounts.Remove 1
        AttachWatcherForAccount acc
    Loop

    MSCANModLogging.WriteLog "InitializeWatchersStaggered: Complete. " & GetWatcherCount() & " watchers active."
    Exit Sub

EH:
    MSCANModLogging.WriteLog "AttachAllPendingWatchers error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub AttachWatcherForAccount(ByVal acc As Outlook.Account)
    Dim inboxFolder As Outlook.Folder
    Dim watcher As MSCANEmailWatcher
    Dim label As String
    Dim store As Outlook.Store
    Dim entryKey As String
    Dim initOk As Boolean

    On Error GoTo EH

    If acc Is Nothing Then Exit Sub
    If Not MSCANSettings.IsAccountScanEnabled(acc) Then Exit Sub

    label = acc.DisplayName

    Set store = Nothing
    Set inboxFolder = Nothing
    On Error Resume Next
    Set store = acc.DeliveryStore
    If Not store Is Nothing Then
        Set inboxFolder = store.GetDefaultFolder(olFolderInbox)
    End If
    Err.Clear
    On Error GoTo EH

    If inboxFolder Is Nothing Then
        MSCANModLogging.WriteLog "AttachWatcher: No inbox for account '" & label & "'"
        Exit Sub
    End If

    entryKey = ""
    On Error Resume Next
    entryKey = inboxFolder.EntryID
    Err.Clear
    On Error GoTo EH

    If gWatchers Is Nothing Then Set gWatchers = New Collection

    Set watcher = New MSCANEmailWatcher
    initOk = watcher.Init(inboxFolder.Items, label)
    If Not initOk Or Not watcher.IsWatching Then
        MSCANModLogging.WriteLog "AttachWatcher: Init failed for '" & label & "' - not counted."
        Set watcher = Nothing
        Exit Sub
    End If

    On Error Resume Next
    If Len(entryKey) > 0 Then
        gWatchers.Add watcher, entryKey
    Else
        Err.Number = 1 ' force unkeyed path
    End If
    If Err.Number <> 0 Then
        Err.Clear
        gWatchers.Add watcher
        MSCANModLogging.WriteLog "Watcher initialized for: " & label & " (unkeyed)"
    Else
        MSCANModLogging.WriteLog "Watcher initialized for: " & label
    End If
    Err.Clear
    On Error GoTo EH

    Set inboxFolder = Nothing
    Set watcher = Nothing
    Set store = Nothing
    Exit Sub

EH:
    MSCANModLogging.WriteLog "AttachWatcherForAccount error (" & label & "): #" & Err.Number & " - " & Err.Description
    Set watcher = Nothing
    Set inboxFolder = Nothing
    Set store = Nothing
End Sub

Public Sub CleanupWatchers()
    On Error Resume Next
    Set gPendingAccounts = Nothing
    Set gWatchers = Nothing
End Sub

Public Sub CheckWatchers()
    MSCANModLogging.WriteLog "CheckWatchers: Reinitializing watchers"
    If Not gServiceEnabled Then
        MSCANModLogging.WriteLog "CheckWatchers: Service is disabled; enable service first."
        Exit Sub
    End If
    InitializeWatchersStaggered
End Sub

'-------------------------------------------------------------------------------
' Persistence + yield helpers
'-------------------------------------------------------------------------------

Private Function ServiceStatePath() As String
    ServiceStatePath = Environ$("LOCALAPPDATA") & "\GeoFooter\" & STATE_FILE_NAME
End Function

Private Function ReadPersistedServiceEnabled(ByVal defaultValue As Boolean) As Boolean
    On Error GoTo EH

    Dim path As String
    Dim fso As Object
    Dim ts As Object
    Dim raw As String
    Dim pos As Long
    Dim chunk As String

    ReadPersistedServiceEnabled = defaultValue
    path = ServiceStatePath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Or Not fso.FileExists(path) Then Exit Function

    Set ts = fso.OpenTextFile(path, 1, False)
    If ts Is Nothing Then Exit Function
    raw = LCase$(ts.ReadAll)
    ts.Close

    pos = InStr(1, raw, """enabled""", vbTextCompare)
    If pos = 0 Then Exit Function
    chunk = Mid$(raw, pos, 24)
    If InStr(1, chunk, "false", vbTextCompare) > 0 Then
        ReadPersistedServiceEnabled = False
    ElseIf InStr(1, chunk, "true", vbTextCompare) > 0 Then
        ReadPersistedServiceEnabled = True
    End If
    Exit Function

EH:
    ReadPersistedServiceEnabled = defaultValue
End Function

' Lightweight write so OFF survives even if toolbar refresh is skipped.
Private Sub PersistServiceEnabled(ByVal enabled As Boolean)
    On Error Resume Next
    ' Prefer full state write (watchers/busy/pending) when toolbar module is present.
    MSCANToolbar.WriteServiceStateFile
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    Dim path As String
    Dim fso As Object
    Dim ts As Object
    Dim parent As String
    Dim en As String

    path = ServiceStatePath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then fso.CreateFolder parent
    End If
    If enabled Then en = "true" Else en = "false"
    Set ts = fso.CreateTextFile(path, True, False)
    ts.Write "{""enabled"": " & en & ", ""watchers"": " & CLng(GetWatcherCount()) & _
             ", ""busy"": false, ""pending"": 0}"
    ts.Close
End Sub

Private Sub YieldBriefly(ByVal seconds As Double)
    On Error Resume Next
    If seconds <= 0 Then Exit Sub
    MSCANIdle.WaitMs CLng(seconds * 1000#)
End Sub
