VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MSCANSecurityRatingForm
   Caption         =   "Select Security Classification"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   StartUpPosition =   1  'CenterOwner
   Begin {978C9E23-D4B0-11CE-BF41-00AA00574A59} lblPrompt
      Caption         =   "Select the security classification for this email:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5400
   End
   Begin {8BD21D03-EC42-11CE-9E0D-00AA00574A59} cmbRating
      Height          =   315
      Left            =   120
      Style           =   2  'fmStyleDropDownList
      TabIndex        =   1
      Top             =   480
      Width           =   5400
   End
   Begin {978C9E22-D4B0-11CE-BF41-00AA00574A59} cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2580
      Width           =   1215
   End
   Begin {978C9E22-D4B0-11CE-BF41-00AA00574A59} cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4260
      TabIndex        =   3
      Top             =   2580
      Width           =   1215
   End
End
Attribute VB_Name = "MSCANSecurityRatingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancelled As Boolean
Public SelectedRating As String
Public subjectText As String
Public UpdatedSubject As String

Private Function GetSecurityClassifications() As Variant
    GetSecurityClassifications = Array( _
        "[U)NR/E]", _
        "[SEC1:(U)EXTERNAL/UNRATED]", _
        "[SEC1:(U)PUBLIC]", _
        "[SEC2:(U)COMMERCIAL-IN-CONFIDENCE/UNENCRYPTED]", _
        "[SEC3:(U)COMMERCIAL-IN-CONFIDENCE-SENSITIVE/ENCRYPTED-PARTIAL]", _
        "[SEC4:(M)RESTRICTED/ENCRYPTED]", _
        "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]", _
        "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]" _
    )
End Function

Private Function FindClassificationInSubject(subjectText As String) As String
    Dim cls As Variant
    For Each cls In GetSecurityClassifications()
        If InStr(1, subjectText, cls, vbTextCompare) > 0 Then
            FindClassificationInSubject = cls
            Exit Function
        End If
    Next cls
    FindClassificationInSubject = ""
End Function

Private Sub UserForm_Initialize()
    With Me.cmbRating
        .Clear
        .AddItem "Select a classification..."
        .AddItem "[U)NR/E]"
        .AddItem "[SEC1:(U)EXTERNAL/UNRATED]"
        .AddItem "[SEC1:(U)PUBLIC]"
        .AddItem "[SEC2:(U)COMMERCIAL-IN-CONFIDENCE/UNENCRYPTED]"
        .AddItem "[SEC3:(U)COMMERCIAL-IN-CONFIDENCE-SENSITIVE/ENCRYPTED-PARTIAL]"
        .AddItem "[SEC4:(M)RESTRICTED/ENCRYPTED]"
        .AddItem "[SEC5:(M)MARKED/CLASSIFIED/ENCRYPTED]"
        .AddItem "[SEC6:(M)MARKED/TRACKED/CLASSIFIED/ENCRYPTED]"
        .ListIndex = 0
    End With

    If Len(Me.subjectText) > 0 Then
        Dim existingTag As String
        existingTag = MSCANModule2.DetectSecurityTag(Me.subjectText)
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

    Me.IsCancelled = True
    Me.SelectedRating = ""
End Sub

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

    If oldCls = "" Then
        subj = newCls & " " & subj
    ElseIf StrComp(oldCls, newCls, vbTextCompare) <> 0 Then
        subj = Replace(subj, oldCls, newCls, , , vbTextCompare)
    End If

    Me.SelectedRating = newCls
    Me.UpdatedSubject = subj
    Me.IsCancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.IsCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.IsCancelled = True
    End If
End Sub
