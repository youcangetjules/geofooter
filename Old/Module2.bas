Attribute VB_Name = "Module2"
' Module2
' UI and String Utility Functions

Option Explicit

' === Show classification dialog and return the updated subject ===
Public Function ShowSecurityRatingDialog(ByVal emailSubject As String) As String
    On Error GoTo Fail
    Dim frm As SecurityRatingForm
    Set frm = New SecurityRatingForm

    ' Pass current subject so the form can detect any existing classification
    frm.subjectText = emailSubject

    ' Tidy up caption for clarity
    Dim truncatedSubject As String
    truncatedSubject = emailSubject
    If Len(truncatedSubject) > 50 Then
        truncatedSubject = left$(truncatedSubject, 47) & "..."
    End If
    frm.caption = "Classify Email: " & truncatedSubject

    ' Show the form modally so execution pauses
    frm.Show vbModal

    If frm.IsCancelled Then
        ' User cancelled – return empty string so caller can stop send
        ShowSecurityRatingDialog = ""
    Else
        ' Return the fully updated subject line (no duplication)
        ShowSecurityRatingDialog = frm.UpdatedSubject
    End If

Cleanup:
    Unload frm
    Set frm = Nothing
    Exit Function

Fail:
    Module1.WriteLog "ShowSecurityRatingDialog error: " & Err.Number & " - " & Err.Description
    ShowSecurityRatingDialog = emailSubject
    Resume Cleanup
End Function

' Removes any known classification tag from anywhere in the subject
Public Function StripSecurityTag(ByVal Subject As String) As String
    On Error GoTo Fail
    Dim tags As Variant, tag As Variant
    tags = Array( _
        "[NR/E]", _
        "[SEC1:EXTERNAL/UNRATED]", _
        "[SEC1:(C)PUBLIC]", _
        "[SEC2:(C)CiC/UNENCRYPTED]", _
        "[SEC3:(C)UNENCRYPTED/CIVILIAN SENSITIVE]", _
        "[SEC4:(M)RESTRICTED/ENCRYPTED]", _
        "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]", _
        "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]" _
    )
    
    StripSecurityTag = Subject
    For Each tag In tags
        StripSecurityTag = Replace(StripSecurityTag, tag, "", , , vbTextCompare)
    Next tag
    
    ' Clean up any extra spaces left behind
    StripSecurityTag = Trim$(StripSecurityTag)
    
    Exit Function

Fail:
    Module1.WriteLog "StripSecurityTag error: " & Err.Number & " - " & Err.Description
    StripSecurityTag = Subject
End Function


' === Append a footer to the body (unchanged) ===
Public Sub AddFooter(mail As Outlook.MailItem)
    On Error Resume Next

    Dim footerText As String
    footerText = vbCrLf & "--" & vbCrLf & "Confidential: This email is subject to company footer policy."

    mail.Body = mail.Body & footerText
    mail.Save
End Sub

' Returns the first known classification tag found anywhere in the subject, or "" if none
Public Function DetectSecurityTag(ByVal Subject As String) As String
    On Error GoTo Fail
    Dim tags As Variant, tag As Variant
    
    ' List of valid classification tags
    tags = Array( _
        "[NR/E]", _
        "[SEC1:EXTERNAL/UNRATED]", _
        "[SEC1:(C)PUBLIC]", _
        "[SEC2:(C)CiC/UNENCRYPTED]", _
        "[SEC3:(C)UNENCRYPTED/CIVILIAN SENSITIVE]", _
        "[SEC4:(M)RESTRICTED/ENCRYPTED]", _
        "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]", _
        "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]" _
    )
    
    DetectSecurityTag = ""
    
    For Each tag In tags
        If InStr(1, Subject, tag, vbTextCompare) > 0 Then
            DetectSecurityTag = tag
            Exit Function
        End If
    Next tag
    
    Exit Function

Fail:
    Module1.WriteLog "DetectSecurityTag error: " & Err.Number & " - " & Err.Description
    DetectSecurityTag = ""
End Function


