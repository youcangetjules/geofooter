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

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowTextW Lib "user32" (ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
#Else
    Private Declare Sub ApiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
#End If

' Cache the compose answer briefly: the retry loops ask on every pass.
Private Const COMPOSE_CACHE_MS As Long = 750
' A draft only blocks AES while the user is actually typing into it. After
' this long without keyboard or mouse input the writer has paused, and a
' footer write (~1s) cannot land between keystrokes.
Private Const COMPOSE_TYPING_GAP_MS As Long = 4000
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
' forward) or an inline reply in the reading pane is the foreground window and
' the user has typed or clicked within COMPOSE_TYPING_GAP_MS. A received mail
' opened for reading reports Sent = True and does not count.
'
' An open draft on its own is not enough: every footer, queue tick and job
' timeout waits on this, so an idle draft must not hold them for hours.
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
    If MsSinceLastInput() >= COMPOSE_TYPING_GAP_MS Then
        answer = False
    Else
        answer = ComposeWindowInForeground()
    End If

    m_ComposeCached = answer
    m_ComposeCheckedAt = nowMs
    m_ComposeHasCache = True
    IsUserComposing = answer
End Function

' Matches the foreground window against Outlook's compose windows by caption.
' Typing elsewhere (another app, or a received mail) does not count.
Private Function ComposeWindowInForeground() As Boolean
    On Error Resume Next

    Dim fg As String
    fg = ForegroundCaption()

    Dim insp As Object
    Dim item As Object
    Dim inline As Object

    For Each insp In Application.Inspectors
        Set item = Nothing
        Set item = insp.CurrentItem
        If Not item Is Nothing Then
            If IsUnsentItem(item) Then
                ' Unreadable foreground caption: assume this draft has focus
                ' rather than risk stealing keystrokes.
                If Len(fg) = 0 Then
                    ComposeWindowInForeground = True
                    Exit Function
                End If
                If StrComp(CStr(insp.Caption), fg, vbTextCompare) = 0 Then
                    ComposeWindowInForeground = True
                    Exit Function
                End If
            End If
        End If
    Next insp

    ' Inline reply in the reading pane creates no Inspector. Check every
    ' Explorer, not just the active one - with two windows open, typing a reply
    ' in the background one was not counted as composing.
    Dim exp As Object
    Dim hasInline As Boolean
    For Each exp In Application.Explorers
        Set inline = Nothing
        Err.Clear
        Set inline = exp.ActiveInlineResponse
        ' Some folder views raise here; that is "cannot tell", which only
        ' counts when this Explorer is the window being typed into.
        hasInline = (Err.Number <> 0) Or Not (inline Is Nothing)
        Err.Clear
        If hasInline Then
            If Len(fg) = 0 Then
                ComposeWindowInForeground = True
                Exit Function
            End If
            If StrComp(CStr(exp.Caption), fg, vbTextCompare) = 0 Then
                ComposeWindowInForeground = True
                Exit Function
            End If
        End If
    Next exp
End Function

Private Function ForegroundCaption() As String
    On Error Resume Next
    Dim buf As String
    Dim n As Long
    buf = String$(512, vbNullChar)
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    hwnd = GetForegroundWindow()
    If hwnd = 0 Then Exit Function
    n = GetWindowTextW(hwnd, StrPtr(buf), 512)
    If n > 0 Then ForegroundCaption = Left$(buf, n)
End Function

' Milliseconds since the last keyboard or mouse input anywhere in the session.
' GetTickCount wraps every ~49.7 days and both values are unsigned DWORDs held
' in signed Longs, so the subtraction is done in Double.
Private Function MsSinceLastInput() As Double
    On Error Resume Next
    Const TWO_32 As Double = 4294967296#
    Dim lii As LASTINPUTINFO
    lii.cbSize = Len(lii)
    If GetLastInputInfo(lii) = 0 Then
        MsSinceLastInput = 0
        Exit Function
    End If

    Dim nowTick As Double
    Dim lastTick As Double
    nowTick = CDbl(GetTickCount())
    lastTick = CDbl(lii.dwTime)
    If nowTick < 0 Then nowTick = nowTick + TWO_32
    If lastTick < 0 Then lastTick = lastTick + TWO_32

    Dim gap As Double
    gap = nowTick - lastTick
    If gap < 0 Then gap = gap + TWO_32
    MsSinceLastInput = gap
End Function

' True when this message is selected in the reading pane or open in a window.
'
' Writing HTMLBody forces Outlook to re-render the item, so rewriting whatever
' is on screen is exactly what the user sees as flickering. Automatic scans
' defer; explicit ribbon actions do not call this because the user asked for
' those and expects the mail to change.
Public Function IsItemOnScreen(ByVal entryId As String) As Boolean
    On Error Resume Next
    IsItemOnScreen = False
    If Len(entryId) = 0 Then Exit Function

    Dim insp As Object
    Dim item As Object
    Dim sel As Object
    Dim i As Long

    For Each insp In Application.Inspectors
        Set item = Nothing
        Set item = insp.CurrentItem
        If Not item Is Nothing Then
            If StrComp(CStr(item.EntryID), entryId, vbTextCompare) = 0 Then
                IsItemOnScreen = True
                Exit Function
            End If
        End If
    Next insp

    Set sel = Nothing
    Set sel = Application.ActiveExplorer.Selection
    If sel Is Nothing Then Exit Function

    ' A large multi-select is not worth walking on the UI thread; the user is
    ' not reading any single one of those messages.
    If sel.Count > 20 Then Exit Function

    For i = 1 To sel.Count
        Set item = Nothing
        Set item = sel.item(i)
        If Not item Is Nothing Then
            If StrComp(CStr(item.EntryID), entryId, vbTextCompare) = 0 Then
                IsItemOnScreen = True
                Exit Function
            End If
        End If
    Next i
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
