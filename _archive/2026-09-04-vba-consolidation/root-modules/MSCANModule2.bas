Attribute VB_Name = "MSCANModule2"
' Module2
' UI and String Utility Functions

Option Explicit

' === Show classification dialog and return the updated subject ===
Public Function ShowSecurityRatingDialog(ByVal emailSubject As String) As String
    On Error GoTo Fail
    Dim frm As MSCANSecurityRatingForm
    Set frm = New MSCANSecurityRatingForm

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
        ' User cancelled ? return empty string so caller can stop send
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
    MSCANModule1.WriteLog "ShowSecurityRatingDialog error: " & Err.Number & " - " & Err.Description
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
        "[SEC3:(C)CIVILIAN SENSITIVE/UNENCRYPTED]", _
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
    MSCANModule1.WriteLog "StripSecurityTag error: " & Err.Number & " - " & Err.Description
    StripSecurityTag = Subject
End Function


' === Append a footer to the body (incoming) ===
Public Sub addFooter(mail As Outlook.MailItem)
    On Error Resume Next
    
    Dim footerHTML As String
    footerHTML = "<br><hr><small><i><!-- MSCAN Footer Start -->Confidential: This email is subject to company footer policy.<!-- MSCAN Footer End --></i></small>"
    
    Select Case mail.BodyFormat
        Case olFormatHTML
            Dim pos As Long
            pos = InStrRev(mail.HTMLBody, "</body>")
            If pos > 0 Then
                mail.HTMLBody = left(mail.HTMLBody, pos - 1) & footerHTML & Mid(mail.HTMLBody, pos)
            Else
                mail.HTMLBody = mail.HTMLBody & footerHTML
            End If
        
        Case olFormatRichText
            ' Append plain text for RTF
            If InStr(mail.Body, "Confidential: This email is subject to company footer policy.") = 0 Then
                mail.Body = mail.Body & vbCrLf & "--" & vbCrLf & "Confidential: This email is subject to company footer policy."
            End If
        
        Case Else
            ' Plain text
            If InStr(mail.Body, "Confidential: This email is subject to company footer policy.") = 0 Then
                mail.Body = mail.Body & vbCrLf & "--" & vbCrLf & "Confidential: This email is subject to company footer policy."
            End If
    End Select
    
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
        "[SEC3:(C)COMMERCIAL-IN-CONFIDENCE-SENSITIVE/UNENCRYPTED]", _
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
    MSCANModule1.WriteLog "DetectSecurityTag error: " & Err.Number & " - " & Err.Description
    DetectSecurityTag = ""
End Function




