VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmailWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module: clsEmailWatcher
' Watches an Outlook folder for new items.

Option Explicit

' Declare the Items collection WithEvents to handle its events
Private WithEvents WatchedItems As Outlook.items
Attribute WatchedItems.VB_VarHelpID = -1
Private mFolderPath As String

' Initializes the watcher to monitor a specific folder's Items collection.
Public Sub Init(ByVal items As Outlook.items)
    On Error GoTo EH
    Set WatchedItems = items
    mFolderPath = WatchedItems.parent.folderPath ' Store the path for logging
    Module1.WriteLog "clsEmailWatcher.Init: Now watching folder " & mFolderPath
    Exit Sub
EH:
    Module1.WriteLog "clsEmailWatcher.Init error: " & Err.Number & " - " & Err.Description
End Sub

Private Sub WatchedItems_ItemAdd(ByVal item As Object)
    On Error GoTo EH
    If TypeOf item Is Outlook.MailItem Then
        Module1.WriteLog "clsEmailWatcher: New MailItem received. Subject: " & item.Subject
        Module1.ProcessEmailSecurity item
    End If
    Exit Sub
EH:
    Module1.WriteLog "clsEmailWatcher.ItemAdd error: " & Err.Number & " - " & Err.Description
End Sub
