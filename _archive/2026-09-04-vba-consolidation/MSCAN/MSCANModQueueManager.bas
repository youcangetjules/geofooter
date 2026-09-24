Attribute VB_Name = "MSCANModQueueManager"
Option Explicit

' After watcher attach, defer *processing* (not queuing) until Outlook finishes syncing.
Private Const STARTUP_QUEUE_QUIET As String = "0:01:30"
Private Const STARTUP_QUIET_MS As Long = 90000
' Gap between queue ticks so Send / UI stay responsive (Outlook has no OnTime).
Private Const QUEUE_TICK_DELAY_MS As Long = 1500
' Async delay before looking for a post-send inbox copy.
Private Const POST_SEND_DRAIN_MS As Long = 12000

Private m_MailQueue As Collection
Private m_IsProcessing As Boolean
Private m_ProcessScheduled As Boolean
Private m_PostSendSubjects As Collection
Private m_PostSendDrainPending As Boolean
Private m_QueueTickPending As Boolean
Private m_QueueTickScheduledAt As Date
Private m_QuietResumePending As Boolean
Private m_EarliestProcessAt As Date
Private m_NotReadyRetries As Object ' Scripting.Dictionary EntryID -> count
Private m_LastCatchUpAt As Date
Private m_LastQueueActivityAt As Date
Private Const QUEUE_TICK_STALE_SEC As Long = 8

Public Sub InitializeQueueManager()
    On Error GoTo EH
    Set m_MailQueue = New Collection
    m_IsProcessing = False
    m_ProcessScheduled = False
    m_PostSendDrainPending = False
    m_QueueTickPending = False
    m_QuietResumePending = False
    Set m_NotReadyRetries = CreateObject("Scripting.Dictionary")
    m_NotReadyRetries.CompareMode = vbTextCompare
    If m_EarliestProcessAt = 0 Then m_EarliestProcessAt = Now
    MSCANModLogging.WriteLog "MSCANModQueueManager initialized."
    Exit Sub
EH:
    MSCANModLogging.WriteLog "InitializeQueueManager error: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub BeginStartupQuietPeriod()
    ' Call after watchers attach so sync-driven ItemAdd floods do not
    ' monopolize Outlook while the inbox is still loading.
    ' Mail is still queued during quiet; processing resumes when quiet ends.
    ' Drop stale backlog from a prior session — it only blocks new mail.
    If Not m_MailQueue Is Nothing Then
        If m_MailQueue.Count > 0 Then
            MSCANModLogging.WriteLog "BeginStartupQuietPeriod: clearing " & m_MailQueue.Count & " stale queued item(s)."
            Set m_MailQueue = New Collection
        End If
    End If
    m_EarliestProcessAt = Now + TimeValue(STARTUP_QUEUE_QUIET)
    MSCANModLogging.WriteLog "BeginStartupQuietPeriod: queue only until " & Format$(m_EarliestProcessAt, "hh:nn:ss") & "; then drain."
    LaunchDelayedQuietResume
End Sub

Public Function IsStartupQuiet() As Boolean
    If m_EarliestProcessAt = 0 Then
        IsStartupQuiet = False
    Else
        IsStartupQuiet = (Now < m_EarliestProcessAt)
    End If
End Function

Public Sub QueueMailForProcessing(ByVal item As Object, Optional ByVal highPriority As Boolean = False)
    On Error GoTo EH

    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub
    If m_MailQueue Is Nothing Then InitializeQueueManager
    If item Is Nothing Or Not MSCANModule1.IsScannableItem(item) Then Exit Sub

    Dim mail As Object
    Set mail = item

    If Not MSCANSettings.IsMailItemAccountScanEnabled(mail) Then
        MSCANModLogging.WriteLog "QueueMailForProcessing: skipped (account scan off) - " & Left$(CStr(mail.Subject), 60)
        Exit Sub
    End If

    Dim uniqueKey As String
    uniqueKey = mail.EntryID
    If Len(uniqueKey) = 0 Then Exit Sub

    If Not ItemInQueue(uniqueKey) Then
        If highPriority Then
            PrependMailToQueue mail, uniqueKey
        Else
            m_MailQueue.Add mail, uniqueKey
        End If
        m_LastQueueActivityAt = Now
        If IsStartupQuiet() Then
            MSCANModLogging.WriteLog "Queued (quiet" & IIf(highPriority, ", priority", "") & "): " & Left$(mail.Subject, 80) & " | Queue size: " & m_MailQueue.Count
        Else
            MSCANModLogging.WriteLog "Queued" & IIf(highPriority, " (priority)", "") & ": " & Left$(mail.Subject, 80) & " | Queue size: " & m_MailQueue.Count
        End If
    ElseIf highPriority Then
        PromoteMailInQueue uniqueKey
    End If

    ' During quiet: keep the mail, do not start processing yet (resume VBS will drain).
    If IsStartupQuiet() Then Exit Sub

    ' Always defer off NewMailEx/ItemAdd — never ProcessEmail on the UI event thread.
    LaunchDelayedQueueTick

    ' New mail: drain immediately as well — VBS ItemLoad nudges often stall for minutes.
    If highPriority And Not m_IsProcessing Then
        DeferredProcessQueue
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "QueueMailForProcessing error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub PrependMailToQueue(ByVal mail As Object, ByVal uniqueKey As String)
    On Error Resume Next
    Dim newQ As Collection
    Dim i As Long
    Dim existing As Object
    Dim existingKey As String

    Set newQ = New Collection
    newQ.Add mail, uniqueKey

    For i = 1 To m_MailQueue.Count
        Set existing = m_MailQueue(i)
        existingKey = existing.EntryID
        If Len(existingKey) > 0 Then
            If StrComp(existingKey, uniqueKey, vbBinaryCompare) <> 0 Then
                newQ.Add existing, existingKey
            End If
        End If
    Next i

    Set m_MailQueue = newQ
End Sub

Private Sub PromoteMailInQueue(ByVal uniqueKey As String)
    On Error Resume Next
    If Not ItemInQueue(uniqueKey) Then Exit Sub
    Dim mail As Object
    Set mail = m_MailQueue(uniqueKey)
    m_MailQueue.Remove uniqueKey
    PrependMailToQueue mail, uniqueKey
End Sub

Public Sub ScheduleDeferredProcessQueue()
    On Error GoTo EH

    If m_ProcessScheduled Then Exit Sub
    If m_IsProcessing Then
        LaunchDelayedQueueTick
        Exit Sub
    End If

    If IsStartupQuiet() Then
        MSCANModLogging.WriteLog "ScheduleDeferredProcessQueue: quiet until " & Format$(m_EarliestProcessAt, "hh:nn:ss") & "; will drain after quiet."
        LaunchDelayedQuietResume
        Exit Sub
    End If

    m_ProcessScheduled = True
    MSCANModLogging.WriteLog "ScheduleDeferredProcessQueue: Processing one queue tick."
    DeferredProcessQueue
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ScheduleDeferredProcessQueue error: #" & Err.Number & " - " & Err.Description
    m_ProcessScheduled = False
End Sub

Public Sub DeferredProcessQueue()
    On Error GoTo EH

    m_ProcessScheduled = False

    If IsStartupQuiet() Then
        MSCANModLogging.WriteLog "DeferredProcessQueue: still in quiet period until " & Format$(m_EarliestProcessAt, "hh:nn:ss")
        LaunchDelayedQuietResume
        Exit Sub
    End If

    ' One item per tick — keeps Send responsive; schedule another tick if more remain.
    ProcessQueue

    If Not m_MailQueue Is Nothing Then
        If m_MailQueue.Count > 0 And Not m_IsProcessing And Not IsStartupQuiet() Then
            LaunchDelayedQueueTick
        End If
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "DeferredProcessQueue error: #" & Err.Number & " - " & Err.Description
    m_ProcessScheduled = False
End Sub

Public Sub ProcessQueue()
    On Error GoTo EH
    Static inCall As Boolean

    If inCall Then Exit Sub
    inCall = True

    If m_MailQueue Is Nothing Then InitializeQueueManager
    If m_IsProcessing Then GoTo EXIT_SUB
    If m_MailQueue.Count = 0 Then GoTo EXIT_SUB
    If Now < m_EarliestProcessAt Then GoTo EXIT_SUB

    m_IsProcessing = True
    MSCANModLogging.WriteLog "ProcessQueue: Started with " & m_MailQueue.Count & " items."

    Dim mail As Object
    Dim startTime As Single
    Dim elapsed As Single
    Dim entryKey As String
    Dim subjectHint As String
    Dim footerSkips As Long
    Const MAX_FOOTER_SKIPS As Long = 40
    Dim startedScan As Boolean

    footerSkips = 0
    startedScan = False

    Do While m_MailQueue.Count > 0 And footerSkips < MAX_FOOTER_SKIPS And Not startedScan
        On Error Resume Next
        Set mail = m_MailQueue.Item(1)
        entryKey = ""
        subjectHint = ""
        If Not mail Is Nothing Then
            entryKey = mail.EntryID
            subjectHint = mail.Subject
        End If
        m_MailQueue.Remove 1
        On Error GoTo EH

        If Not IsItemValid(mail) Then
            MSCANModLogging.WriteLog "Skipped invalid mail reference."
            ClearNotReadyRetry entryKey
            GoTo NextQueueItem
        End If

        ' Drop already-stamped mail without reading HTMLBody (catch-up backlog).
        If MSCANModule1.IsAesScanned(mail) Then
            footerSkips = footerSkips + 1
            ClearNotReadyRetry entryKey
            GoTo NextQueueItem
        End If

        startTime = Timer
        MSCANModLogging.WriteLog "Processing mail: " & subjectHint
        MSCANModStatus.ShowStatus "Processing footer: " & Left$(subjectHint, 60) & " (" & m_MailQueue.Count & " left)"
        MSCANModule1.ProcessEmail mail
        elapsed = Timer - startTime
        MSCANModLogging.WriteLog "Processed: " & subjectHint & " (" & Format(elapsed, "0.00") & "s)"

        If ItemInQueue(entryKey) Then
            If Not RegisterNotReadyRetry(entryKey) Then
                RemoveFromQueue entryKey
                MSCANModLogging.WriteLog "Giving up on not-ready mail after retries: " & Left$(subjectHint, 80)
            End If
            startedScan = True
        ElseIf MSCANModule1.IsAesScanned(mail) Then
            ' Footer already present — stamped; keep draining backlog this tick.
            footerSkips = footerSkips + 1
            ClearNotReadyRetry entryKey
        Else
            ' Async geo job started (or export failed) — one real scan per tick.
            startedScan = True
            ClearNotReadyRetry entryKey
        End If

NextQueueItem:
        Set mail = Nothing
        DoEvents
    Loop

    m_LastQueueActivityAt = Now
    MSCANModLogging.WriteLog "ProcessQueue: Tick complete. Remaining = " & m_MailQueue.Count & _
        IIf(footerSkips > 0, " (skipped " & footerSkips & " already-scanned)", "")

EXIT_SUB:
    m_IsProcessing = False
    inCall = False
    Exit Sub

EH:
    MSCANModLogging.WriteLog "ProcessQueue error: #" & Err.Number & " - " & Err.Description
    m_IsProcessing = False
    inCall = False
End Sub

Private Function IsItemValid(ByVal mail As Object) As Boolean
    On Error Resume Next
    If mail Is Nothing Then Exit Function
    If Not MSCANModule1.IsScannableItem(mail) Then Exit Function
    Dim subject As String
    subject = mail.Subject
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    IsItemValid = True
End Function

Private Function ItemInQueue(ByVal key As String) As Boolean
    On Error Resume Next
    If Len(key) = 0 Then Exit Function
    If m_MailQueue Is Nothing Then Exit Function
    Dim temp As Object
    Set temp = m_MailQueue(key)
    ItemInQueue = (Err.Number = 0)
    Err.Clear
End Function

Private Sub RemoveFromQueue(ByVal key As String)
    On Error Resume Next
    If Len(key) = 0 Then Exit Sub
    If m_MailQueue Is Nothing Then Exit Sub
    m_MailQueue.Remove key
    Err.Clear
    ClearNotReadyRetry key
End Sub

Private Function RegisterNotReadyRetry(ByVal key As String) As Boolean
    ' Returns True if another retry is allowed.
    On Error Resume Next
    Const MAX_RETRIES As Long = 3
    If Len(key) = 0 Then
        RegisterNotReadyRetry = False
        Exit Function
    End If
    If m_NotReadyRetries Is Nothing Then
        Set m_NotReadyRetries = CreateObject("Scripting.Dictionary")
        m_NotReadyRetries.CompareMode = vbTextCompare
    End If
    Dim n As Long
    If m_NotReadyRetries.Exists(key) Then
        n = CLng(m_NotReadyRetries(key)) + 1
    Else
        n = 1
    End If
    m_NotReadyRetries(key) = n
    RegisterNotReadyRetry = (n <= MAX_RETRIES)
End Function

Private Sub ClearNotReadyRetry(ByVal key As String)
    On Error Resume Next
    If Len(key) = 0 Then Exit Sub
    If m_NotReadyRetries Is Nothing Then Exit Sub
    If m_NotReadyRetries.Exists(key) Then m_NotReadyRetries.Remove key
End Sub

Public Sub ForceResumeProcessing()
    If m_IsProcessing Then
        LaunchDelayedQueueTick
        Exit Sub
    End If
    MSCANModLogging.WriteLog "ForceResumeProcessing called."
    m_EarliestProcessAt = Now
    m_QuietResumePending = False
    DrainPostSendSubjects
    ScheduleDeferredProcessQueue
End Sub

' Called from ItemSend — must NOT do heavy work (no inbox scan / ProcessQueue).
Public Sub SchedulePostSendFooterScan(ByVal subject As String)
    On Error GoTo EH

    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub
    If Len(Trim$(subject)) = 0 Then Exit Sub

    If m_PostSendSubjects Is Nothing Then Set m_PostSendSubjects = New Collection
    m_PostSendSubjects.Add subject
    MSCANModLogging.WriteLog "SchedulePostSendFooterScan: queued '" & Left$(subject, 80) & "' (async drain; not during ItemSend)"

    LaunchDelayedPostSendDrain
    Exit Sub

EH:
    MSCANModLogging.WriteLog "SchedulePostSendFooterScan error: #" & Err.Number & " - " & Err.Description
End Sub

' Public entry point for delayed WScript / ItemAdd / ForceResume.
Public Sub DrainPostSendSubjects()
    On Error GoTo EH

    If m_PostSendSubjects Is Nothing Then Exit Sub
    If m_PostSendSubjects.Count = 0 Then Exit Sub
    If m_IsProcessing Then Exit Sub

    PostSendFooterScan
    Exit Sub

EH:
    MSCANModLogging.WriteLog "DrainPostSendSubjects error: #" & Err.Number & " - " & Err.Description
End Sub

Public Sub PostSendFooterScan()
    On Error GoTo EH

    If m_PostSendSubjects Is Nothing Or m_PostSendSubjects.Count = 0 Then Exit Sub

    MSCANModLogging.WriteLog "PostSendFooterScan: Checking " & m_PostSendSubjects.Count & " recently sent subject(s)."

    Dim i As Long
    For i = 1 To m_PostSendSubjects.Count
        FindAndQueueRecentMailBySubject CStr(m_PostSendSubjects(i))
    Next i

    Set m_PostSendSubjects = New Collection
    m_PostSendDrainPending = False

    If Not m_MailQueue Is Nothing Then
        If m_MailQueue.Count > 0 And Not m_IsProcessing Then
            ScheduleDeferredProcessQueue
        End If
    End If
    Exit Sub

EH:
    m_PostSendDrainPending = False
    MSCANModLogging.WriteLog "PostSendFooterScan error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub LaunchDelayedQueueTick()
    On Error GoTo EH

    If m_MailQueue Is Nothing Then Exit Sub
    If m_MailQueue.Count = 0 Then Exit Sub
    If IsStartupQuiet() Then
        LaunchDelayedQuietResume
        Exit Sub
    End If

    If m_QueueTickPending Then
        If m_QueueTickScheduledAt > 0 Then
            If DateDiff("s", m_QueueTickScheduledAt, Now) < QUEUE_TICK_STALE_SEC Then Exit Sub
        End If
        MSCANModLogging.WriteLog "LaunchDelayedQueueTick: stale tick latch cleared (queue=" & m_MailQueue.Count & ")."
        m_QueueTickPending = False
    End If

    m_QueueTickPending = True
    m_QueueTickScheduledAt = Now

    Dim vbsPath As String
    vbsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_queue_tick.vbs"
    If Not WriteQueueTickScript(vbsPath, QUEUE_TICK_DELAY_MS) Then
        MSCANModLogging.WriteLog "LaunchDelayedQueueTick: could not write tick script."
        m_QueueTickPending = False
        m_QueueTickScheduledAt = 0
        Exit Sub
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe //Nologo //B " & Chr$(34) & vbsPath & Chr$(34), 0, False
    MSCANModLogging.WriteLog "LaunchDelayedQueueTick: next tick in " & QUEUE_TICK_DELAY_MS & " ms (queue=" & m_MailQueue.Count & ")."
    Exit Sub

EH:
    m_QueueTickPending = False
    m_QueueTickScheduledAt = 0
    MSCANModLogging.WriteLog "LaunchDelayedQueueTick error: #" & Err.Number & " - " & Err.Description
End Sub

' If the VBS nudge never lands (Outlook idle / cached GetFirst), unstick the queue.
Private Sub MaybeDrainQueueIfStale()
    On Error Resume Next

    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub
    If m_MailQueue Is Nothing Or m_MailQueue.Count = 0 Then Exit Sub
    If m_IsProcessing Or IsStartupQuiet() Then Exit Sub

    Dim stale As Boolean
    stale = False

    If m_QueueTickPending Then
        If m_QueueTickScheduledAt = 0 Then
            stale = True
        ElseIf DateDiff("s", m_QueueTickScheduledAt, Now) >= QUEUE_TICK_STALE_SEC Then
            stale = True
        End If
    ElseIf m_LastQueueActivityAt > 0 Then
        If DateDiff("s", m_LastQueueActivityAt, Now) >= QUEUE_TICK_STALE_SEC Then
            stale = True
        End If
    End If

    If stale Then
        MSCANModLogging.WriteLog "MaybeDrainQueueIfStale: forcing drain (queue=" & m_MailQueue.Count & ")."
        m_QueueTickPending = False
        m_QueueTickScheduledAt = 0
        DeferredProcessQueue
    End If
End Sub

' Queue tick: sleep briefly, then retry inbox nudges — single GetFirst often misses ItemLoad.
Private Function WriteQueueTickScript(ByVal vbsPath As String, ByVal sleepMs As Long) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Dim ts As Object
    Dim folderPath As String
    Dim logPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(vbsPath)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    logPath = Environ$("LOCALAPPDATA") & "\GeoFooter\Logs\VBA_Log.txt"

    Set ts = fso.CreateTextFile(vbsPath, True, False)
    ts.WriteLine "WScript.Sleep " & CStr(sleepMs)
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine "Dim ol, items, itm, n, dummy, fsoLog, logPath, tsLog"
    ts.WriteLine "Set ol = GetObject(, ""Outlook.Application"")"
    ts.WriteLine "If ol Is Nothing Then WScript.Quit 1"
    ts.WriteLine "For n = 1 To 12"
    ts.WriteLine "  Err.Clear"
    ts.WriteLine "  Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(6).Items"
    ts.WriteLine "  If items Is Nothing Or Err.Number <> 0 Then"
    ts.WriteLine "    Err.Clear"
    ts.WriteLine "    Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(5).Items"
    ts.WriteLine "  End If"
    ts.WriteLine "  If Not items Is Nothing Then"
    ts.WriteLine "    items.Sort ""[ReceivedTime]"", True"
    ts.WriteLine "    If (n Mod 2) = 1 Then"
    ts.WriteLine "      Set itm = items.GetFirst"
    ts.WriteLine "    Else"
    ts.WriteLine "      Set itm = items.GetLast"
    ts.WriteLine "    End If"
    ts.WriteLine "    If Not itm Is Nothing Then dummy = itm.EntryID"
    ts.WriteLine "  End If"
    ts.WriteLine "  WScript.Sleep 500"
    ts.WriteLine "Next"
    ts.WriteLine "If Err.Number <> 0 Then"
    ts.WriteLine "  logPath = CreateObject(""WScript.Shell"").ExpandEnvironmentStrings(""%LOCALAPPDATA%\GeoFooter\Logs\VBA_Log.txt"")"
    ts.WriteLine "  Set fsoLog = CreateObject(""Scripting.FileSystemObject"")"
    ts.WriteLine "  Set tsLog = fsoLog.OpenTextFile(logPath, 8, True)"
    ts.WriteLine "  tsLog.WriteLine Now & "" | VBS nudge failed (AesQueueTick): #"" & Err.Number & "" - "" & Err.Description"
    ts.WriteLine "  tsLog.Close"
    ts.WriteLine "End If"
    ts.Close
    WriteQueueTickScript = True
    Exit Function

EH:
    WriteQueueTickScript = False
End Function

' Called from MSCANModule1.NudgeAsyncWork on every Application_ItemLoad.
' Delayed VBS pings land here indirectly (they nudge Outlook by reading an
' item). Casual browsing ItemLoads must stay cheap — only latch-driven nudges
' may CatchUp / ProcessQueue (HTMLBody + geo work freeze the UI).
Public Sub NudgeQueueWork()
    On Error GoTo EH

    MaybeDrainQueueIfStale

    Dim hadQueueTick As Boolean
    Dim hadPostSend As Boolean
    Dim hadQuietResume As Boolean
    hadQueueTick = m_QueueTickPending
    hadPostSend = m_PostSendDrainPending
    hadQuietResume = m_QuietResumePending And (Not IsStartupQuiet())

    m_QueueTickPending = False
    m_PostSendDrainPending = False
    If Not IsStartupQuiet() Then m_QuietResumePending = False

    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub

    ' User clicked/opened mail: do not scan inboxes or process the queue here.
    If Not (hadQueueTick Or hadPostSend Or hadQuietResume) Then Exit Sub

    If hadPostSend Or hadQuietResume Or hadQueueTick Then
        DrainPostSendSubjects
    End If
    ' Catch-up after startup quiet only — not on every queue tick (that refloods the queue).
    If hadQuietResume Then
        CatchUpMissedInboxMail
    End If
    If m_MailQueue Is Nothing Then Exit Sub
    If m_MailQueue.Count = 0 Then Exit Sub
    If m_IsProcessing Then Exit Sub
    If IsStartupQuiet() Then Exit Sub
    DeferredProcessQueue
    Exit Sub
EH:
    MSCANModLogging.WriteLog "NudgeQueueWork error: #" & Err.Number & " - " & Err.Description
End Sub

' Throttled inbox sweep: queue recent unscanned items that watchers missed.
Public Sub CatchUpMissedInboxMail()
    On Error GoTo EH

    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub
    If IsStartupQuiet() Then Exit Sub
    If m_LastCatchUpAt > 0 Then
        If DateDiff("s", m_LastCatchUpAt, Now) < 120 Then Exit Sub
    End If
    m_LastCatchUpAt = Now

    If m_MailQueue Is Nothing Then InitializeQueueManager

    Dim acc As Outlook.Account
    Dim inboxFolder As Outlook.folder
    Dim inboxItems As Outlook.Items
    Dim itm As Object
    Dim idx As Long
    Dim checkCount As Long
    Dim queued As Long
    Dim received As Date

    queued = 0
    For Each acc In Application.Session.Accounts
        On Error Resume Next
        Set inboxFolder = Nothing
        If Not MSCANSettings.IsAccountScanEnabled(acc) Then GoTo NextCatchUpAccount
        If Not acc.DeliveryStore Is Nothing Then
            Set inboxFolder = acc.DeliveryStore.GetDefaultFolder(olFolderInbox)
        End If
        If inboxFolder Is Nothing Then GoTo NextCatchUpAccount

        Set inboxItems = inboxFolder.Items
        inboxItems.Sort "[ReceivedTime]", True
        checkCount = 2
        If inboxItems.Count < checkCount Then checkCount = inboxItems.Count

        For idx = 1 To checkCount
            Set itm = inboxItems.Item(idx)
            If Not MSCANModule1.IsScannableItem(itm) Then GoTo NextCatchUpItem
            ' Already stamped with AES footer — never re-queue (no HTMLBody needed).
            If MSCANModule1.IsAesScanned(itm) Then GoTo NextCatchUpItem

            received = 0
            Err.Clear
            received = itm.ReceivedTime
            If Err.Number <> 0 Or received = 0 Then
                Err.Clear
                If Not itm.UnRead Then GoTo NextCatchUpItem
            ElseIf DateDiff("n", received, Now) > 20 Then
                ' Only recover very recent misses — older mail is backlog noise.
                GoTo NextCatchUpItem
            ElseIf Not itm.UnRead And DateDiff("n", received, Now) > 5 Then
                GoTo NextCatchUpItem
            End If

            QueueMailForProcessing itm
            queued = queued + 1

NextCatchUpItem:
            Set itm = Nothing
            Err.Clear
        Next idx

NextCatchUpAccount:
        Err.Clear
        On Error GoTo EH
    Next acc

    If queued > 0 Then
        MSCANModLogging.WriteLog "CatchUpMissedInboxMail: queued " & queued & " missed item(s)."
        If Not m_IsProcessing And Not IsStartupQuiet() Then ScheduleDeferredProcessQueue
    End If
    Exit Sub

EH:
    MSCANModLogging.WriteLog "CatchUpMissedInboxMail error: #" & Err.Number & " - " & Err.Description
End Sub

' Entry point from aes_queue_tick.vbs
Public Sub OnQueueTick()
    On Error GoTo EH
    m_QueueTickPending = False
    m_QueueTickScheduledAt = 0
    If Not MSCANModWatchers.IsServiceEnabled() Then Exit Sub
    If m_MailQueue Is Nothing Then Exit Sub
    If m_MailQueue.Count = 0 Then Exit Sub
    ScheduleDeferredProcessQueue
    Exit Sub
EH:
    m_QueueTickPending = False
    MSCANModLogging.WriteLog "OnQueueTick error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub LaunchDelayedQuietResume()
    On Error GoTo EH

    If m_QuietResumePending Then Exit Sub
    m_QuietResumePending = True

    Dim vbsPath As String
    Dim delayMs As Long
    Dim remainingSec As Long

    delayMs = STARTUP_QUIET_MS
    If m_EarliestProcessAt > Now Then
        remainingSec = DateDiff("s", Now, m_EarliestProcessAt)
        If remainingSec < 1 Then remainingSec = 1
        delayMs = remainingSec * 1000&
        If delayMs < 1000 Then delayMs = 1000
        If delayMs > STARTUP_QUIET_MS + 5000 Then delayMs = STARTUP_QUIET_MS
    End If

    vbsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_quiet_resume.vbs"
    If Not WriteSleepRunScript(vbsPath, delayMs, "AesForceResume") Then
        MSCANModLogging.WriteLog "LaunchDelayedQuietResume: could not write resume script."
        m_QuietResumePending = False
        Exit Sub
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe //Nologo //B " & Chr$(34) & vbsPath & Chr$(34), 0, False
    MSCANModLogging.WriteLog "LaunchDelayedQuietResume: will drain queue in " & delayMs & " ms."
    Exit Sub

EH:
    m_QuietResumePending = False
    MSCANModLogging.WriteLog "LaunchDelayedQuietResume error: #" & Err.Number & " - " & Err.Description
End Sub

' Writes a VBS that sleeps, then nudges Outlook by reading one inbox item.
' That fires Application_ItemLoad inside Outlook -> MSCANModule1.NudgeAsyncWork,
' which drains the queue / reconciles async jobs. (This Outlook build exposes
' neither Application.Run nor ThisOutlookSession publics to external COM, so
' direct method calls are impossible; purposeName is only for the log line.)
Private Function WriteSleepRunScript(ByVal vbsPath As String, ByVal sleepMs As Long, ByVal purposeName As String) As Boolean
    On Error GoTo EH

    Dim fso As Object
    Dim ts As Object
    Dim folderPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(vbsPath)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    Set ts = fso.CreateTextFile(vbsPath, True, False)
    ts.WriteLine "WScript.Sleep " & CStr(sleepMs)
    ts.WriteLine "On Error Resume Next"
    ts.WriteLine "Dim ol, items, itm, n, dummy, fsoLog, logPath, tsLog"
    ts.WriteLine "Set ol = GetObject(, ""Outlook.Application"")"
    ts.WriteLine "If ol Is Nothing Then WScript.Quit 1"
    ts.WriteLine "For n = 1 To 6"
    ts.WriteLine "  Err.Clear"
    ts.WriteLine "  Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(6).Items"
    ts.WriteLine "  If items Is Nothing Or Err.Number <> 0 Then"
    ts.WriteLine "    Err.Clear"
    ts.WriteLine "    Set items = ol.GetNamespace(""MAPI"").GetDefaultFolder(5).Items"
    ts.WriteLine "  End If"
    ts.WriteLine "  If Not items Is Nothing Then"
    ts.WriteLine "    items.Sort ""[ReceivedTime]"", True"
    ts.WriteLine "    If (n Mod 2) = 1 Then"
    ts.WriteLine "      Set itm = items.GetFirst"
    ts.WriteLine "    Else"
    ts.WriteLine "      Set itm = items.GetLast"
    ts.WriteLine "    End If"
    ts.WriteLine "    If Not itm Is Nothing Then dummy = itm.EntryID"
    ts.WriteLine "  End If"
    ts.WriteLine "  WScript.Sleep 400"
    ts.WriteLine "Next"
    ts.WriteLine "If Err.Number <> 0 Then"
    ts.WriteLine "  logPath = CreateObject(""WScript.Shell"").ExpandEnvironmentStrings(""%LOCALAPPDATA%\GeoFooter\Logs\VBA_Log.txt"")"
    ts.WriteLine "  Set fsoLog = CreateObject(""Scripting.FileSystemObject"")"
    ts.WriteLine "  Set tsLog = fsoLog.OpenTextFile(logPath, 8, True)"
    ts.WriteLine "  tsLog.WriteLine Now & "" | VBS nudge failed (" & purposeName & "): #"" & Err.Number & "" - "" & Err.Description"
    ts.WriteLine "  tsLog.Close"
    ts.WriteLine "End If"
    ts.Close
    WriteSleepRunScript = True
    Exit Function

EH:
    WriteSleepRunScript = False
End Function

Private Sub LaunchDelayedPostSendDrain()
    On Error GoTo EH

    If m_PostSendDrainPending Then Exit Sub
    m_PostSendDrainPending = True

    Dim vbsPath As String
    vbsPath = Environ$("LOCALAPPDATA") & "\GeoFooter\aes_postsend_drain.vbs"
    If Not WriteSleepRunScript(vbsPath, POST_SEND_DRAIN_MS, "AesDrainPostSend") Then
        MSCANModLogging.WriteLog "LaunchDelayedPostSendDrain: could not write drain script; will rely on ItemAdd."
        m_PostSendDrainPending = False
        Exit Sub
    End If

    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    ' Async — must not block ItemSend.
    shell.Run "wscript.exe //Nologo //B " & Chr$(34) & vbsPath & Chr$(34), 0, False
    MSCANModLogging.WriteLog "LaunchDelayedPostSendDrain: scheduled async drain (" & POST_SEND_DRAIN_MS & " ms)."
    Exit Sub

EH:
    m_PostSendDrainPending = False
    MSCANModLogging.WriteLog "LaunchDelayedPostSendDrain error: #" & Err.Number & " - " & Err.Description
End Sub

Private Sub FindAndQueueRecentMailBySubject(ByVal targetSubject As String)
    On Error GoTo EH

    Dim acc As Outlook.Account
    Dim inboxFolder As Outlook.folder
    Dim inboxItems As Outlook.Items
    Dim mail As Object
    Dim idx As Long
    Dim checkCount As Long

    For Each acc In Application.Session.Accounts
        On Error Resume Next
        Set inboxFolder = Nothing
        If Not acc.DeliveryStore Is Nothing Then
            Set inboxFolder = acc.DeliveryStore.GetDefaultFolder(olFolderInbox)
        End If
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextAccount
        End If
        On Error GoTo EH

        If inboxFolder Is Nothing Then GoTo NextAccount

        If Not MSCANSettings.IsAccountScanEnabled(acc) Then GoTo NextAccount

        Set inboxItems = inboxFolder.Items
        inboxItems.Sort "[ReceivedTime]", True

        checkCount = 10
        If inboxItems.Count < checkCount Then checkCount = inboxItems.Count

        For idx = 1 To checkCount
            If MSCANModule1.IsScannableItem(inboxItems(idx)) Then
                Set mail = inboxItems(idx)
                If StrComp(mail.Subject, targetSubject, vbTextCompare) = 0 Then
                    If Not MSCANModule1.MailHasFooter(mail) Then
                        MSCANModLogging.WriteLog "PostSendFooterScan: Queuing inbox copy on " & acc.DisplayName & " - " & Left$(targetSubject, 80)
                        QueueMailForProcessing mail
                    End If
                    Exit Sub
                End If
            End If
        Next idx

NextAccount:
    Next acc
    Exit Sub

EH:
    MSCANModLogging.WriteLog "FindAndQueueRecentMailBySubject error: #" & Err.Number & " - " & Err.Description
End Sub
