VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SecurityRatingForm 
   Caption         =   "Select Security Classification"
   ClientHeight    =   2310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   OleObjectBlob   =   "SecurityRatingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SecurityRatingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SecurityRatingForm (UserForm Code-Behind)

Option Explicit

' Public properties to be read by the calling module after the form is hidden.
Public IsCancelled As Boolean
Public SelectedRating As String
Public subjectText As String  ' Pass this before showing the form
Public UpdatedSubject As String ' This will be returned to caller

' --- List of classifications ---
Private Function GetSecurityClassifications() As Variant
    GetSecurityClassifications = Array( _
        "[NR/E]", _
        "[SEC1:EXTERNAL/UNRATED]", _
        "[SEC1:(C)PUBLIC]", _
        "[SEC2:(C)CiC/UNENCRYPTED]", _
        "[SEC3:(C)UNENCRYPTED/CIVILIAN SENSITIVE]", _
        "[SEC4:(M)RESTRICTED/ENCRYPTED]", _
        "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]", _
        "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]" _
    )
End Function

' --- Detect existing classification in subject ---
Private Function FindClassificationInSubject(subjectText As String) As String
    Dim cls As Variant
    For Each cls In GetSecurityClassifications()
        If InStr(1, subjectText, cls, vbTextCompare) > 0 Then
            FindClassificationInSubject = cls
            Exit Function
        End If
    Next cls
    FindClassificationInSubject = ""  ' Not found
End Function

Private Sub UserForm_Initialize()
    ' Populate the dropdown list of classifications
    With Me.cmbRating
        .Clear
        .AddItem "Select a classification..." ' Index 0
        .AddItem "[NR/E]"
        .AddItem "[SEC1:EXTERNAL/UNRATED]"
        .AddItem "[SEC1:(C)PUBLIC]"
        .AddItem "[SEC2:(C)CiC/UNENCRYPTED]"
        .AddItem "[SEC3:(C)UNENCRYPTED/CIVILIAN SENSITIVE]"
        .AddItem "[SEC4:(M)RESTRICTED/ENCRYPTED]"
        .AddItem "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]"
        .AddItem "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]"

        ' Default to the first item
        .ListIndex = 0
    End With

    ' Pre-select the classification if already detected
    If Len(Me.subjectText) > 0 Then
        Dim existingTag As String
        existingTag = Module2.DetectSecurityTag(Me.subjectText)
        If Len(existingTag) > 0 Then
            Dim i As Long
            For i = 1 To Me.cmbRating.ListCount - 1
                If Me.cmbRating.List(i) = existingTag Then
                    Me.cmbRating.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End If

    ' Default state
    Me.IsCancelled = True
    Me.SelectedRating = ""
End Sub


' --- OK button ---
Private Sub cmdOK_Click()
    Dim oldCls As String
    Dim newCls As String
    Dim subj As String

    If Me.cmbRating.ListIndex <= 0 Then
        MsgBox "You must select a valid classification before proceeding.", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    newCls = CStr(Me.cmbRating.Value)
    oldCls = FindClassificationInSubject(Me.subjectText)
    subj = Trim(Me.subjectText)

    ' --- Update logic ---
    If oldCls = "" Then
        ' No previous classification: prepend new one
        subj = newCls & " " & subj
    ElseIf StrComp(oldCls, newCls, vbTextCompare) <> 0 Then
        ' Different classification: replace old with new
        subj = Replace(subj, oldCls, newCls, , , vbTextCompare)
    Else
        ' Same classification: do nothing (avoid duplication)
    End If

    Me.SelectedRating = newCls
    Me.UpdatedSubject = subj
    Me.IsCancelled = False
    Me.Hide
End Sub

' --- Cancel button ---
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

' --- Handle window close (X) ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.IsCancelled = True
    End If
End Sub


