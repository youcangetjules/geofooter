Attribute VB_Name = "MSCANHealth"
Option Explicit

' Session health counters for the AES add-in.
'
' The main log is a flat stream of INFO lines, so a slow-burn fault (a commit
' conflict storm, a queue that never drains, work running while the user types)
' only shows up if you already know what to grep for. Every counter here is
' written back as a single "HEALTH" line, and anything that crosses a threshold
' raises a WARN immediately. To review a session:
'
'   grep HEALTH "%LOCALAPPDATA%\GeoFooter\Logs\VBA_Log.txt"
'
' or run scripts\aes_log_health.py for a parsed summary.

#If VBA7 Then
    Private Declare PtrSafe Function ApiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#Else
    Private Declare Function ApiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
#End If

' An operation holding the Outlook UI thread this long is long enough for the
' user to see a stall while typing.
Private Const STALL_WARN_MS As Long = 1200
' Consecutive failed commits that mean the item is never going to settle.
Private Const CONFLICT_STORM As Long = 3
' A backlog this deep will churn the reading pane for several minutes.
Private Const QUEUE_BACKLOG_WARN As Long = 25
Private Const SNAPSHOT_EVERY_MS As Long = 300000

Private m_SessionStart As Date
Private m_Started As Boolean

Private m_CommitAttempts As Long
Private m_CommitConflicts As Long
Private m_ConflictRun As Long
Private m_ConflictStormLogged As Boolean

Private m_FootersApplied As Long
Private m_FootersAbandoned As Long
Private m_DuplicateFooters As Long
Private m_MissingFooters As Long
Private m_WorstFooterCount As Long
Private m_MarkFailures As Long

Private m_ComposeDeferrals As Long
Private m_OnScreenDeferrals As Long
Private m_QueuePeak As Long
Private m_BacklogWarned As Boolean

Private m_Stalls As Long
Private m_WorstStallMs As Long
Private m_WorstStallOp As String
Private m_UiBusyMs As Double

Private m_LastSnapshotTick As Long

Public Sub StartSession()
    m_SessionStart = Now
    m_Started = True
    m_CommitAttempts = 0
    m_CommitConflicts = 0
    m_ConflictRun = 0
    m_ConflictStormLogged = False
    m_FootersApplied = 0
    m_FootersAbandoned = 0
    m_DuplicateFooters = 0
    m_MissingFooters = 0
    m_WorstFooterCount = 0
    m_MarkFailures = 0
    m_ComposeDeferrals = 0
    m_OnScreenDeferrals = 0
    m_QueuePeak = 0
    m_BacklogWarned = False
    m_Stalls = 0
    m_WorstStallMs = 0
    m_WorstStallOp = ""
    m_UiBusyMs = 0
    m_LastSnapshotTick = ApiGetTickCount()
    MSCANModLogging.WriteLog "HEALTH | session counters reset"
End Sub

'------------------------------------------------------------------------
' UI thread timing - the "Outlook freezes while I type" detector
'------------------------------------------------------------------------

' Returns a tick to hand back to EndOp. Always pair the two.
Public Function BeginOp() As Long
    BeginOp = ApiGetTickCount()
End Function

Public Sub EndOp(ByVal opName As String, ByVal startTick As Long)
    On Error Resume Next
    Dim elapsed As Long
    elapsed = ElapsedMs(startTick)
    m_UiBusyMs = m_UiBusyMs + elapsed

    If elapsed < STALL_WARN_MS Then Exit Sub

    m_Stalls = m_Stalls + 1
    If elapsed > m_WorstStallMs Then
        m_WorstStallMs = elapsed
        m_WorstStallOp = opName
    End If
    MSCANModLogging.WriteLogWarn "HEALTH | UI stall: " & opName & " held the Outlook thread for " & _
        CStr(elapsed) & "ms (typing can drop keystrokes above " & CStr(STALL_WARN_MS) & "ms)"
    MaybeSnapshot
End Sub

' GetTickCount wraps roughly every 49 days; treat a wrap as no elapsed time
' rather than reporting a negative or absurd duration.
Private Function ElapsedMs(ByVal startTick As Long) As Long
    Dim nowTick As Long
    nowTick = ApiGetTickCount()
    If nowTick < startTick Then
        ElapsedMs = 0
    Else
        ElapsedMs = nowTick - startTick
    End If
End Function

'------------------------------------------------------------------------
' Footer commit health - the "reading pane flickers" detector
'------------------------------------------------------------------------

Public Sub NoteCommitAttempt()
    m_CommitAttempts = m_CommitAttempts + 1
End Sub

Public Sub NoteCommitConflict(ByVal subjectHint As String)
    On Error Resume Next
    m_CommitConflicts = m_CommitConflicts + 1
    m_ConflictRun = m_ConflictRun + 1

    If m_ConflictRun >= CONFLICT_STORM And Not m_ConflictStormLogged Then
        m_ConflictStormLogged = True
        MSCANModLogging.WriteLogWarn "HEALTH | commit conflict storm: " & CStr(m_ConflictRun) & _
            " consecutive 'message has been changed' failures (last: " & Left$(subjectHint, 60) & _
            "). Each retry repaints the mail - this is what looks like flickering."
        MaybeSnapshot
    End If
End Sub

Public Sub NoteCommitSucceeded()
    m_ConflictRun = 0
    m_ConflictStormLogged = False
End Sub

Public Sub NoteFooterApplied()
    m_FootersApplied = m_FootersApplied + 1
    m_ConflictRun = 0
    m_ConflictStormLogged = False
End Sub

Public Sub NoteFooterAbandoned()
    m_FootersAbandoned = m_FootersAbandoned + 1
End Sub

' The scanned mark failed to save. The footer is in the body but the message
' looks unscanned, so it will be picked up and rewritten again.
Public Sub NoteMarkFailed()
    On Error Resume Next
    m_MarkFailures = m_MarkFailures + 1
    If m_MarkFailures = 5 Then
        MSCANModLogging.WriteLogWarn "HEALTH | the scanned mark has failed to save 5 times; " & _
            "those messages will be scanned again on every pass."
    End If
End Sub

' Every message must carry exactly one scan block. Anything else means old
' footers are surviving removal (markers stripped in transit) or a message is
' being footered twice.
Public Sub NoteFooterCount(ByVal count As Long)
    On Error Resume Next
    If count = 1 Then Exit Sub

    If count > 1 Then
        m_DuplicateFooters = m_DuplicateFooters + 1
        If count > m_WorstFooterCount Then m_WorstFooterCount = count
    Else
        m_MissingFooters = m_MissingFooters + 1
    End If
    MaybeSnapshot
End Sub

'------------------------------------------------------------------------
' Queue and compose health
'------------------------------------------------------------------------

Public Sub NoteQueueDepth(ByVal depth As Long)
    On Error Resume Next
    If depth > m_QueuePeak Then m_QueuePeak = depth

    If depth >= QUEUE_BACKLOG_WARN And Not m_BacklogWarned Then
        m_BacklogWarned = True
        MSCANModLogging.WriteLogWarn "HEALTH | queue backlog: " & CStr(depth) & _
            " items pending. Draining this many will keep rewriting mail for several minutes."
    ElseIf depth < QUEUE_BACKLOG_WARN Then
        m_BacklogWarned = False
    End If
    MaybeSnapshot
End Sub

Public Sub NoteComposeDeferral()
    m_ComposeDeferrals = m_ComposeDeferrals + 1
End Sub

' A footer held back because its message was selected or open. Expected in
' small numbers; a large count means mail is sitting unfootered because the
' user keeps it on screen.
Public Sub NoteOnScreenDeferral()
    On Error Resume Next
    m_OnScreenDeferrals = m_OnScreenDeferrals + 1
    If m_OnScreenDeferrals = 50 Then
        MSCANModLogging.WriteLogWarn "HEALTH | 50 footers deferred because the message was on screen; " & _
            "mail may be staying unfootered."
    End If
End Sub

'------------------------------------------------------------------------
' Snapshot
'------------------------------------------------------------------------

Private Sub MaybeSnapshot()
    On Error Resume Next
    If ElapsedMs(m_LastSnapshotTick) < SNAPSHOT_EVERY_MS Then Exit Sub
    LogSnapshot "periodic"
End Sub

Public Sub LogSnapshot(ByVal reason As String)
    On Error Resume Next
    If Not m_Started Then StartSession
    m_LastSnapshotTick = ApiGetTickCount()

    Dim conflictPct As Long
    If m_CommitAttempts > 0 Then
        conflictPct = CLng((CDbl(m_CommitConflicts) / CDbl(m_CommitAttempts)) * 100#)
    End If

    MSCANModLogging.WriteLog "HEALTH | " & reason & _
        " since=" & Format$(m_SessionStart, "HH:nn:ss") & _
        " commits=" & CStr(m_CommitAttempts) & _
        " conflicts=" & CStr(m_CommitConflicts) & "(" & CStr(conflictPct) & "%)" & _
        " applied=" & CStr(m_FootersApplied) & _
        " abandoned=" & CStr(m_FootersAbandoned) & _
        " dupFooters=" & CStr(m_DuplicateFooters) & _
        IIf(m_WorstFooterCount > 1, "(worst " & CStr(m_WorstFooterCount) & ")", "") & _
        " missingFooters=" & CStr(m_MissingFooters) & _
        " markFailures=" & CStr(m_MarkFailures) & _
        " queuePeak=" & CStr(m_QueuePeak) & _
        " composeDeferrals=" & CStr(m_ComposeDeferrals) & _
        " onScreenDeferrals=" & CStr(m_OnScreenDeferrals) & _
        " uiStalls=" & CStr(m_Stalls) & _
        " worstStall=" & CStr(m_WorstStallMs) & "ms" & _
        IIf(Len(m_WorstStallOp) > 0, "(" & m_WorstStallOp & ")", "") & _
        " uiBusy=" & CStr(CLng(m_UiBusyMs / 1000#)) & "s"
End Sub

' One-line summary for the diagnostics dialog / ribbon.
Public Function HealthSummary() As String
    HealthSummary = "commits " & CStr(m_CommitAttempts) & ", conflicts " & CStr(m_CommitConflicts) & _
        ", applied " & CStr(m_FootersApplied) & ", abandoned " & CStr(m_FootersAbandoned) & _
        ", duplicate footers " & CStr(m_DuplicateFooters) & _
        ", queue peak " & CStr(m_QueuePeak) & ", UI stalls " & CStr(m_Stalls)
End Function
