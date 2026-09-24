Attribute VB_Name = "MSCANEventHandlers"
'===============================================================================
' MSCANEventHandlers - Event Handler Logic
' Author: Julian Garrett (Rewritten)
' Purpose: Contains the logic for handling Outlook events
'===============================================================================
Option Explicit

Private Const MODULE_NAME As String = "MSCANEventHandlers"

'===============================================================================
' MODULE-LEVEL VARIABLES FOR FORM RESULTS
'===============================================================================

' These replace the UDT since VBA can't return UDTs from functions easily
Private m_FormCancelled As Boolean
Private m_FormSelectedTag As String
Private m_FormUpdatedSubject As String
Private m_FormReadNotify As Boolean
Private m_FormReadNotifySuffix As String

'===============================================================================
' NEW MAIL (Application.NewMailEx) — backup when Items.ItemAdd is missed
'===============================================================================

Public Sub HandleNewMailEx(ByVal entryIdCollection As String)
    On Error GoTo ErrorHandler

    If Not MSCANModWatchers.IsServiceEnabled() Then
        MSCANModLogging.WriteLogDebug "NewMailEx: ignored (service OFF)."
        Exit Sub
    End If
    If Len(entryIdCollection) = 0 Then Exit Sub

    Dim parts() As String
    Dim i As Long
    Dim entryId As String
    Dim itm As Object
    Dim queued As Long
    Dim skipped As Long

    parts = Split(entryIdCollection, ",")
    queued = 0
    skipped = 0
    For i = LBound(parts) To UBound(parts)
        entryId = Trim$(parts(i))
        If Len(entryId) = 0 Then GoTo NextNewMailPart

        On Error Resume Next
        Set itm = Application.Session.GetItemFromID(entryId)
        Err.Clear
        On Error GoTo ErrorHandler

        If itm Is Nothing Then
            skipped = skipped + 1
            MSCANModLogging.WriteLogDebug "NewMailEx: skip — GetItemFromID failed for " & Left$(entryId, 24)
            GoTo NextNewMailPart
        End If
        If Not MSCANModule1.IsScannableItem(itm) Then
            skipped = skipped + 1
            MSCANModLogging.WriteLogDebug "NewMailEx: skip non-scannable class=" & CStr(itm.MessageClass)
            GoTo NextNewMailPart
        End If
        ' Account filter before any body/header work — HTMLBody freezes Outlook.
        If Not MSCANSettings.IsMailItemAccountScanEnabled(itm) Then
            skipped = skipped + 1
            MSCANModLogging.WriteLog "NewMailEx: skip (account scan OFF) - " & Left$(CStr(itm.Subject), 60)
            GoTo NextNewMailPart
        End If

        MSCANModLogging.WriteLog "NewMailEx: queueing " & Left$(CStr(itm.Subject), 80) & _
            " class=" & CStr(itm.MessageClass)
        MSCANModQueueManager.NoteNewMailExEvent CStr(itm.Subject)
        MSCANModQueueManager.QueueMailForProcessing itm, True
        queued = queued + 1

NextNewMailPart:
        Set itm = Nothing
    Next i

    If queued > 0 Then
        MSCANModLogging.WriteLog "NewMailEx: queued " & queued & " item(s)" & _
            IIf(skipped > 0, ", skipped " & skipped, "") & "."
    ElseIf skipped > 0 Then
        MSCANModLogging.WriteLog "NewMailEx: nothing queued (" & skipped & " skipped)."
    End If
    Exit Sub

ErrorHandler:
    MSCANModLogging.WriteLog "HandleNewMailEx error: #" & Err.Number & " - " & Err.Description
End Sub

'===============================================================================
' ITEM SEND HANDLER
'===============================================================================

Public Function HandleItemSend(ByVal Item As Object, ByRef Cancel As Boolean) As Boolean
    '''Handles the ItemSend event - validates classification before sending
    '''
    '''Args:
    '''    Item: The Outlook item being sent
    '''    Cancel: Set to True to prevent sending
    '''
    '''Returns:
    '''    True if item should be sent, False if cancelled
    
    On Error GoTo ErrorHandler
    
    ' Only process mail items
    If Not TypeOf Item Is Outlook.mailItem Then
        HandleItemSend = True
        Exit Function
    End If
    
    Dim mailItem As Outlook.mailItem
    Set mailItem = Item
    
    MSCANCore.Log "ItemSend triggered for: " & Left$(mailItem.Subject, 50)
    LogSendDiagnostics mailItem, "ItemSend-start"
    
    ' Check if already classified
    If MSCANCore.HasClassification(mailItem.Subject) Then
        MSCANCore.Log "Email already classified, allowing send"
        ApplyReadNotifyWithConflictRetry mailItem, True, "", True
        SchedulePostSendFooterScan mailItem
        HandleItemSend = True
        Exit Function
    End If
    
    ' Show classification form
    ShowClassificationForm mailItem
    
    ' Handle the result
    If m_FormCancelled Then
        MSCANCore.Log "User cancelled classification"
        Cancel = True
        HandleItemSend = False
        Exit Function
    End If
    
    ' Apply the classification (conflict-aware)
    If Not ApplySubjectWithConflictRetry(mailItem, m_FormUpdatedSubject) Then
        Cancel = True
        HandleItemSend = False
        Exit Function
    End If

    MSCANCore.Log "Applied classification: " & m_FormSelectedTag & _
        " | ReadNotify=" & IIf(m_FormReadNotify, "Yes", "No") & _
        IIf(Len(m_FormReadNotifySuffix) > 0, " (" & m_FormReadNotifySuffix & ")", "")

    ApplyReadNotifyWithConflictRetry mailItem, m_FormReadNotify, m_FormReadNotifySuffix, False
    LogSendDiagnostics mailItem, "ItemSend-after-mutate"
    SchedulePostSendFooterScan mailItem
    HandleItemSend = True
    Exit Function
    
ErrorHandler:
    If IsMessageChangedError(Err.Number, Err.Description) Then
        LogConflict "HandleItemSend", Err.Number, Err.Description, mailItem
    Else
        MSCANCore.LogError MODULE_NAME & ".HandleItemSend", Err.Number, Err.Description
    End If
    
    ' On error, ask user what to do
    Dim response As VbMsgBoxResult
    response = MsgBox("An error occurred during classification." & vbCrLf & _
                      "Error: " & Err.Description & vbCrLf & vbCrLf & _
                      "Do you want to send the email anyway?", _
                      vbExclamation + vbYesNo, "MSCAN Error")
    
    If response = vbNo Then
        Cancel = True
        HandleItemSend = False
    Else
        SchedulePostSendFooterScan mailItem
        HandleItemSend = True
    End If
End Function

Private Sub LogSendDiagnostics(ByVal mail As Outlook.MailItem, ByVal stage As String)
    On Error Resume Next

    Dim entryId As String
    Dim modTime As String
    Dim created As String
    Dim sz As String
    Dim acct As String

    If mail Is Nothing Then Exit Sub

    entryId = Left$(mail.EntryID, 40)
    modTime = Format$(mail.LastModificationTime, "yyyy-mm-dd hh:nn:ss")
    created = Format$(mail.CreationTime, "yyyy-mm-dd hh:nn:ss")
    sz = CStr(mail.Size)
    If Not mail.SendUsingAccount Is Nothing Then
        acct = mail.SendUsingAccount.SmtpAddress
        If Len(acct) = 0 Then acct = mail.SendUsingAccount.DisplayName
    End If

    Dim line As String
    line = "SendDiag[" & stage & "]: subject='" & Left$(mail.Subject, 60) & "'" & _
           " | EntryID=" & entryId & "..." & _
           " | Created=" & created & _
           " | LastMod=" & modTime & _
           " | Size=" & sz & _
           " | Account=" & acct & _
           " | To=" & Left$(mail.To, 80) & _
           " | CC=" & Left$(mail.CC, 60) & _
           " | Aliniant=" & IIf(MSCANReadNotify.IsAliniantSend(mail), "Yes", "No")

    MSCANCore.Log line
    MSCANModLogging.WriteLog line
End Sub

Private Sub LogConflict( _
    ByVal source As String, _
    ByVal errNum As Long, _
    ByVal errDesc As String, _
    ByVal mail As Outlook.MailItem)

    On Error Resume Next

    Dim line As String
    line = "CONFLICT: message changed during AES " & source & _
           " | Err#" & errNum & " - " & errDesc & _
           " | Likely Outlook autosave/Exchange sync and/or ReadNotify ActiveTracker " & _
           "while AES also mutated the item."

    MSCANCore.Log line
    MSCANModLogging.WriteLog line
    LogSendDiagnostics mail, "conflict-" & source
End Sub

Private Function IsMessageChangedError(ByVal errNum As Long, ByVal errDesc As String) As Boolean
    Dim d As String
    d = LCase$(errDesc)
    IsMessageChangedError = (InStr(1, d, "message has been changed", vbTextCompare) > 0) Or _
                            (InStr(1, d, "has been changed", vbTextCompare) > 0 And _
                             InStr(1, d, "message", vbTextCompare) > 0)
End Function

Private Function ApplySubjectWithConflictRetry( _
    ByVal mail As Outlook.MailItem, _
    ByVal newSubject As String) As Boolean

    On Error GoTo EH

    Dim attempt As Long
    For attempt = 1 To 2
        On Error Resume Next
        Err.Clear
        mail.Subject = newSubject
        If Err.Number = 0 Then
            ApplySubjectWithConflictRetry = True
            Exit Function
        End If

        If IsMessageChangedError(Err.Number, Err.Description) Then
            LogConflict "ApplySubject(attempt " & attempt & ")", Err.Number, Err.Description, mail
            If attempt = 1 Then
                DoEvents
                Err.Clear
            Else
                MsgBox "Outlook reports this message was changed while AES was updating the subject." & vbCrLf & vbCrLf & _
                       "Close the draft, reopen it, and send again. Details are in the AES log.", _
                       vbExclamation, "AES"
                ApplySubjectWithConflictRetry = False
                Exit Function
            End If
        Else
            MSCANCore.LogError MODULE_NAME & ".ApplySubjectWithConflictRetry", Err.Number, Err.Description
            ApplySubjectWithConflictRetry = False
            Exit Function
        End If
        On Error GoTo EH
    Next attempt

    ApplySubjectWithConflictRetry = False
    Exit Function

EH:
    MSCANCore.LogError MODULE_NAME & ".ApplySubjectWithConflictRetry", Err.Number, Err.Description
    ApplySubjectWithConflictRetry = False
End Function

Private Sub ApplyReadNotifyWithConflictRetry( _
    ByVal mail As Outlook.MailItem, _
    ByVal enabled As Boolean, _
    ByVal suffix As String, _
    ByVal useTagDefault As Boolean)

    On Error GoTo EH

    Dim attempt As Long
    Dim ok As Boolean
    Dim tag As String
    Dim idx As Long

    For attempt = 1 To 2
        On Error Resume Next
        Err.Clear
        ok = False

        If useTagDefault Then
            tag = MSCANCore.DetectClassification(mail.Subject)
            idx = MSCANCore.GetClassificationLevel(tag)
            If idx < 1 Or Not MSCANCore.RequiresTracking(idx) Then
                Exit Sub
            End If
            ok = MSCANReadNotify.ApplyReadNotifyIfRequested(mail, True, "")
        Else
            ok = MSCANReadNotify.ApplyReadNotifyIfRequested(mail, enabled, suffix)
        End If

        If Err.Number <> 0 Then
            If IsMessageChangedError(Err.Number, Err.Description) Then
                LogConflict "ReadNotify(attempt " & attempt & ")", Err.Number, Err.Description, mail
            Else
                MSCANCore.LogError MODULE_NAME & ".ApplyReadNotifyWithConflictRetry", Err.Number, Err.Description
                Exit Sub
            End If
            ok = False
        End If
        On Error GoTo EH

        If ok Then Exit Sub

        If attempt = 1 Then
            LogSendDiagnostics mail, "ReadNotify-retry"
            DoEvents
        Else
            MSCANCore.Log "ReadNotify: giving up after conflict/failure — send continues without re-stamp."
            MSCANModLogging.WriteLog "ReadNotify: giving up after conflict/failure — send continues without re-stamp."
        End If
    Next attempt
    Exit Sub

EH:
    MSCANCore.LogError MODULE_NAME & ".ApplyReadNotifyWithConflictRetry", Err.Number, Err.Description
End Sub

Private Sub SchedulePostSendFooterScan(ByVal mailItem As Outlook.mailItem)
    On Error Resume Next
    MSCANModQueueManager.SchedulePostSendFooterScan mailItem.Subject
End Sub

'===============================================================================
' FORM DISPLAY
'===============================================================================

Private Sub ShowClassificationForm(ByVal mailItem As Outlook.mailItem)
    '''Shows the classification dialog and stores results in module variables

    On Error GoTo ErrorHandler

    Dim originalSubject As String
    Dim isAliniant As Boolean

    originalSubject = ""
    isAliniant = False
    If Not mailItem Is Nothing Then
        originalSubject = mailItem.Subject
        isAliniant = MSCANReadNotify.IsAliniantSend(mailItem)
    End If

    m_FormCancelled = True
    m_FormSelectedTag = ""
    m_FormUpdatedSubject = ""
    m_FormReadNotify = False
    m_FormReadNotifySuffix = ""

    m_FormCancelled = MSCANClassificationDialog.ShowDialog( _
        originalSubject, m_FormSelectedTag, m_FormUpdatedSubject, _
        m_FormReadNotify, isAliniant, m_FormReadNotifySuffix)

    Exit Sub

ErrorHandler:
    MSCANCore.LogError MODULE_NAME & ".ShowClassificationForm", Err.Number, Err.Description
    m_FormCancelled = True
End Sub

'===============================================================================
' STARTUP HANDLER
'===============================================================================

Public Sub HandleStartup()
    '''Handles application startup - initialization tasks
    
    On Error GoTo ErrorHandler
    
    MSCANCore.Log "========================================="
    MSCANCore.Log "MSCAN Started - Outlook Session Begin"
    MSCANCore.Log "========================================="
    MSCANCore.Log "User: " & Environ$("USERNAME")
    MSCANCore.Log "Computer: " & Environ$("COMPUTERNAME")
    MSCANCore.Log "Outlook Version: " & Application.Version
    
    ' Verify configuration
    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    MSCANCore.Log "Classification tags loaded: " & tagCount
    
    Exit Sub
    
ErrorHandler:
    ' Show error to user since this is startup
    Debug.Print "STARTUP ERROR: " & Err.Description
    MSCANCore.LogError MODULE_NAME & ".HandleStartup", Err.Number, Err.Description
End Sub


'===============================================================================
' QUIT HANDLER
'===============================================================================

Public Sub HandleQuit()
    '''Handles application quit - cleanup tasks
    
    On Error Resume Next
    
    MSCANCore.Log "========================================="
    MSCANCore.Log "MSCAN Stopped - Outlook Session End"
    MSCANCore.Log "========================================="
End Sub

'===============================================================================
' MANUAL CLASSIFICATION
'===============================================================================

Public Sub ClassifySelectedEmail()
    '''Allows user to classify or reclassify the currently selected email
    '''Can be called from a toolbar button or macro
    
    On Error GoTo ErrorHandler
    
    Dim explorer As Outlook.explorer
    Set explorer = Application.ActiveExplorer
    
    If explorer Is Nothing Then
        MsgBox "No active explorer window.", vbExclamation, "MSCAN"
        Exit Sub
    End If
    
    If explorer.selection.count = 0 Then
        MsgBox "Please select an email to classify.", vbInformation, "MSCAN"
        Exit Sub
    End If
    
    Dim selectedItem As Object
    Set selectedItem = explorer.selection.Item(1)
    
    If Not TypeOf selectedItem Is Outlook.mailItem Then
        MsgBox "Please select an email message.", vbInformation, "MSCAN"
        Exit Sub
    End If
    
    Dim mailItem As Outlook.mailItem
    Set mailItem = selectedItem
    
    ' Show classification form
    ShowClassificationForm mailItem
    
    If m_FormCancelled Then
        Exit Sub
    End If
    
    ' Apply classification
    mailItem.Subject = m_FormUpdatedSubject
    mailItem.Save
    MSCANReadNotify.ApplyReadNotifyIfRequested mailItem, m_FormReadNotify, m_FormReadNotifySuffix
    
    MSCANCore.Log "Manual classification applied: " & m_FormSelectedTag
    MsgBox "Classification applied successfully.", vbInformation, "MSCAN"
    
    Exit Sub
    
ErrorHandler:
    MSCANCore.LogError MODULE_NAME & ".ClassifySelectedEmail", Err.Number, Err.Description
    MsgBox "Error classifying email: " & Err.Description, vbExclamation, "MSCAN"
End Sub

'===============================================================================
' INSPECTOR ITEM CLASSIFICATION
'===============================================================================

Public Sub ClassifyOpenEmail()
    '''Classifies the email currently open in an inspector window
    '''Can be called from a toolbar button
    
    On Error GoTo ErrorHandler
    
    Dim inspector As Outlook.inspector
    Set inspector = Application.ActiveInspector
    
    If inspector Is Nothing Then
        MsgBox "No email is currently open.", vbExclamation, "MSCAN"
        Exit Sub
    End If
    
    Dim currentItem As Object
    Set currentItem = inspector.currentItem
    
    If Not TypeOf currentItem Is Outlook.mailItem Then
        MsgBox "The current item is not an email.", vbInformation, "MSCAN"
        Exit Sub
    End If
    
    Dim mailItem As Outlook.mailItem
    Set mailItem = currentItem
    
    ' Show classification form
    ShowClassificationForm mailItem
    
    If m_FormCancelled Then
        Exit Sub
    End If
    
    ' Apply classification
    mailItem.Subject = m_FormUpdatedSubject
    MSCANReadNotify.ApplyReadNotifyIfRequested mailItem, m_FormReadNotify, m_FormReadNotifySuffix
    
    MSCANCore.Log "Inspector classification applied: " & m_FormSelectedTag
    
    Exit Sub
    
ErrorHandler:
    MSCANCore.LogError MODULE_NAME & ".ClassifyOpenEmail", Err.Number, Err.Description
    MsgBox "Error classifying email: " & Err.Description, vbExclamation, "MSCAN"
End Sub

