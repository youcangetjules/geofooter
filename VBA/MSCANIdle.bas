Attribute VB_Name = "MSCANIdle"
Option Explicit
'===============================================================================
' MSCANIdle - waiting and compose detection.
'
' Two rules keep Outlook typeable while AES works:
'
' 1. Never spin on DoEvents. A "Do While Timer < t: DoEvents: Loop" re-enters
'    Outlook's message pump at 100% CPU. Keystrokes typed into a compose window
'    are pumped through VBA instead of the editor, so characters are dropped and
'    the window appears to freeze. WaitMs sleeps the thread instead: no message
'    pump, no re-entrancy, no lost input.
'
' 2. Never run scan / footer work while the user is writing. IsUserComposing
'    covers both a compose Inspector and an inline reply in the reading pane.
'===============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' Cache the compose answer briefly: the retry loops ask on every pass.
Private Const COMPOSE_CACHE_MS As Long = 750
Private m_ComposeCheckedAt As Double
Private m_ComposeCached As Boolean
Private m_ComposeHasCache As Boolean

' Sleep without pumping messages. Long waits are sliced so a wedged call
' cannot hold the UI thread for more than SLICE_MS at a time.
Public Sub WaitMs(ByVal ms As Long)
    On Error Resume Next
    Const SLICE_MS As Long = 50
    Dim remaining As Long

    If ms <= 0 Then Exit Sub
    remaining = ms
    Do While remaining > 0
        If remaining > SLICE_MS Then
            ApiSleep SLICE_MS
            remaining = remaining - SLICE_MS
        Else
            ApiSleep remaining
            remaining = 0
        End If
    Loop
End Sub

' True while a message is being written: a compose Inspector (new mail, reply,
' forward) or an inline reply in the reading pane. A received mail opened for
' reading reports Sent = True and does not count.
Public Function IsUserComposing() As Boolean
    On Error Resume Next

    Dim nowMs As Double
    nowMs = Timer * 1000#
    If m_ComposeHasCache Then
        ' Timer resets at midnight - treat a backwards jump as expiry.
        If nowMs >= m_ComposeCheckedAt And (nowMs - m_ComposeCheckedAt) < COMPOSE_CACHE_MS Then
            IsUserComposing = m_ComposeCached
            Exit Function
        End If
    End If

    Dim answer As Boolean
    answer = AnyComposeWindowOpen()

    m_ComposeCached = answer
    m_ComposeCheckedAt = nowMs
    m_ComposeHasCache = True
    IsUserComposing = answer
End Function

Private Function AnyComposeWindowOpen() As Boolean
    On Error Resume Next

    Dim insp As Object
    Dim item As Object
    Dim inline As Object

    For Each insp In Application.Inspectors
        Set item = Nothing
        Set item = insp.CurrentItem
        If Not item Is Nothing Then
            If IsUnsentItem(item) Then
                AnyComposeWindowOpen = True
                Exit Function
            End If
        End If
    Next insp

    ' Inline reply in the reading pane creates no Inspector.
    Set inline = Nothing
    Set inline = Application.ActiveExplorer.ActiveInlineResponse
    If Not inline Is Nothing Then
        AnyComposeWindowOpen = True
    End If
End Function

Private Function IsUnsentItem(ByVal item As Object) As Boolean
    On Error Resume Next
    Dim wasSent As Boolean
    wasSent = True
    Err.Clear
    wasSent = CBool(item.Sent)
    If Err.Number <> 0 Then
        ' Non-mail items (appointments, contacts) have no Sent - not a compose.
        Err.Clear
        IsUnsentItem = False
        Exit Function
    End If
    IsUnsentItem = Not wasSent
End Function
