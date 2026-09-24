VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' ThisOutlookSession - thin event stubs only (Outlook requires events here).
' See IMPORT.txt: import this file, then move code into the built-in object.
' All logic is in importable modules (MSCANAppBootstrap, MSCANEventHandlers, etc.).
'===============================================================================
Option Explicit

Private Sub Application_Startup()
    On Error GoTo EH
    MSCANAppBootstrap.HandleApplicationStartup
    Exit Sub
EH:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    MSCANModLogging.WriteLogLevel "error", _
        "ThisOutlookSession.Application_Startup: #" & errNum & " - " & errDesc
End Sub

Private Sub Application_Quit()
    On Error Resume Next
    MSCANAppBootstrap.HandleApplicationQuit
End Sub

Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    On Error GoTo EH

    Dim allowed As Boolean
    allowed = MSCANEventHandlers.HandleItemSend(Item, Cancel)

    ' Prefer explicit Cancel from the handler; also honour False return.
    If Not allowed Then Cancel = True
    Exit Sub

EH:
    ' Do not Resume Next — a missing/broken handler must not silently send unmarked mail.
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    MSCANModLogging.WriteLogLevel "error", _
        "ThisOutlookSession.Application_ItemSend: #" & errNum & " - " & errDesc
    Cancel = True
    MsgBox "AES could not process this send (classification handler failed)." & vbCrLf & vbCrLf & _
           "The send was cancelled. Check the AES log (View Logs), then try again." & vbCrLf & vbCrLf & _
           "Error: " & errDesc, _
           vbExclamation, "AES"
End Sub

' More reliable than per-folder ItemAdd for some stores (Google/IMAP OST sync
' can deliver mail without firing Items.ItemAdd). EntryIDs are comma-separated.
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    On Error Resume Next
    MSCANEventHandlers.HandleNewMailEx EntryIDCollection
End Sub

' Fires whenever any item loads (user clicks, or a helper VBS "nudges" Outlook
' by reading an inbox item). Drives async-job completion and queue drain,
' because external scripts cannot call VBA directly in this Outlook build.
' Do NOT touch Item here — its properties are not accessible during ItemLoad.
Private Sub Application_ItemLoad(ByVal Item As Object)
    On Error Resume Next
    MSCANModule1.NudgeAsyncWork
End Sub

'-------------------------------------------------------------------------------
' External automation bridge (legacy — kept for Outlook builds that expose
' ThisOutlookSession publics on the Application object; this build does NOT,
' so helper scripts use the ItemLoad nudge above instead).
'-------------------------------------------------------------------------------
Public Sub AesCompleteAsyncFooter(ByVal jobId As String)
    On Error Resume Next
    MSCANModule1.CompleteAsyncFooter jobId
End Sub

Public Sub AesFailAsyncFooter(ByVal jobId As String)
    On Error Resume Next
    MSCANModule1.FailAsyncFooter jobId
End Sub

Public Sub AesQueueTick()
    On Error Resume Next
    MSCANModQueueManager.OnQueueTick
End Sub

Public Sub AesForceResume()
    On Error Resume Next
    MSCANModQueueManager.ForceResumeProcessing
End Sub

Public Sub AesDrainPostSend()
    On Error Resume Next
    MSCANModQueueManager.DrainPostSendSubjects
End Sub

Public Sub AesUpgradeToFullScan(ByVal entryId As String)
    On Error Resume Next
    MSCANModule1.UpgradeToFullScanByEntryID entryId
End Sub
