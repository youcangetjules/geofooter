VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSCANEmailWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents WatchedItems As Outlook.Items
Attribute WatchedItems.VB_VarHelpID = -1
Private mFolderLabel As String

Public Function Init(ByVal Items As Outlook.Items, Optional ByVal folderLabel As String = "") As Boolean
    On Error GoTo EH
    Init = False
    If Items Is Nothing Then
        MSCANModLogging.WriteLog "MSCANEmailWatcher.Init: Items collection is Nothing."
        Exit Function
    End If

    mFolderLabel = folderLabel
    ' Do not touch Items.Count or FolderPath — both can enumerate huge stores.
    Set WatchedItems = Items
    If WatchedItems Is Nothing Then
        MSCANModLogging.WriteLog "MSCANEmailWatcher.Init: failed to bind Items for " & folderLabel
        Exit Function
    End If
    If Len(mFolderLabel) = 0 Then mFolderLabel = "(inbox)"
    MSCANModLogging.WriteLog "MSCANEmailWatcher: Now watching: " & mFolderLabel
    Init = True
    Exit Function

EH:
    Init = False
    Set WatchedItems = Nothing
    MSCANModLogging.WriteLog "MSCANEmailWatcher.Init error: #" & Err.Number & " - " & Err.Description
End Function

Public Function IsWatching() As Boolean
    IsWatching = Not (WatchedItems Is Nothing)
End Function

Private Sub WatchedItems_ItemAdd(ByVal item As Object)
    On Error Resume Next

    If item Is Nothing Then Exit Sub
    ' MailItem + ReportItem (ReadNotify IPNRN read receipts, etc.)
    If Not MSCANModule1.IsScannableItem(item) Then
        MSCANModLogging.WriteLog "WatchedItems_ItemAdd: skipped non-scannable in " & mFolderLabel & _
            " class=" & CStr(item.MessageClass)
        Exit Sub
    End If

    ' Ignore ancient backfill; keep a wide window because Google/IMAP OST sync
    ' often fires ItemAdd minutes/hours after ReceivedTime.
    If Not IsRecentMail(item) Then
        MSCANModLogging.WriteLog "WatchedItems_ItemAdd: skipped stale in " & mFolderLabel & _
            " - " & Left$(CStr(item.Subject), 60)
        Exit Sub
    End If

    MSCANModLogging.WriteLog "WatchedItems_ItemAdd: new mail in " & mFolderLabel & " class=" & item.MessageClass
    ' Opportunistic drain of subjects queued at ItemSend (never scanned inside ItemSend).
    ' Safe during quiet — DrainPostSendSubjects exits if nothing pending / processing.
    MSCANModQueueManager.DrainPostSendSubjects
    ' Queue even during startup quiet; processing resumes when quiet ends.
    MSCANModQueueManager.QueueMailForProcessing item, True
End Sub

Private Function IsRecentMail(ByVal item As Object) As Boolean
    On Error Resume Next
    Dim received As Date
    Dim unread As Boolean
    received = 0
    unread = False
    If TypeOf item Is Outlook.MailItem Then
        received = item.ReceivedTime
        unread = item.UnRead
    Else
        ' ReportItem has no ReceivedTime — use MAPI delivery time, then CreationTime.
        received = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E060040")
        If Err.Number <> 0 Or received = 0 Then
            Err.Clear
            received = item.CreationTime
        End If
    End If
    If Err.Number <> 0 Or received = 0 Then
        Err.Clear
        IsRecentMail = True
        Exit Function
    End If
    ' 24h window; still-unread mail within 48h also accepted (missed ItemAdd recovery).
    If DateDiff("n", received, Now) <= (24 * 60) Then
        IsRecentMail = True
    ElseIf unread And DateDiff("n", received, Now) <= (48 * 60) Then
        IsRecentMail = True
    Else
        IsRecentMail = False
    End If
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    Set WatchedItems = Nothing
    MSCANModLogging.WriteLog "MSCANEmailWatcher: Stopped watching: " & mFolderLabel
End Sub
