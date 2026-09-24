VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClassification
   Caption         =   "Select Security Classification"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClassification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmClassification - Security Classification Selection Form
' Controls are created at runtime so the .frm imports without a .frx file.
'===============================================================================
Option Explicit

Private lblPrompt As MSForms.Label
Private lblSubjectCaption As MSForms.Label
Private txtSubject As MSForms.TextBox
Private lblDescription As MSForms.Label

Private WithEvents lstClassifications As MSForms.ListBox
Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton

Private m_SelectedIndex As Long
Private m_SelectedTag As String
Private m_UpdatedSubject As String
Private m_OriginalSubject As String
Private m_IsCancelled As Boolean
Private m_ControlsBuilt As Boolean

Public Property Let SubjectText(ByVal value As String)
    m_OriginalSubject = value
End Property

Public Property Get SubjectText() As String
    SubjectText = m_OriginalSubject
End Property

Public Property Get selectedIndex() As Long
    selectedIndex = m_SelectedIndex
End Property

Public Property Get SelectedTag() As String
    SelectedTag = m_SelectedTag
End Property

Public Property Get SelectedRating() As String
    SelectedRating = m_SelectedTag
End Property

Public Property Get UpdatedSubject() As String
    UpdatedSubject = m_UpdatedSubject
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = m_IsCancelled
End Property

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    BuildControls

    m_IsCancelled = True
    m_SelectedIndex = 0
    m_SelectedTag = ""
    m_UpdatedSubject = ""
    lblDescription.Caption = ""
    PopulateClassificationList
    Me.StartUpPosition = 1
    Exit Sub

ErrorHandler:
    MsgBox "Error initializing form: " & Err.Description, vbExclamation, "MSCAN"
End Sub

Private Sub BuildControls()
    If m_ControlsBuilt Then Exit Sub

    Set lblPrompt = Me.Controls.Add("Forms.Label.1", "lblPrompt", True)
    With lblPrompt
        .Left = 8
        .Top = 8
        .Width = 432
        .Height = 17
        .Caption = "Select the security classification for this email:"
    End With

    Set lblSubjectCaption = Me.Controls.Add("Forms.Label.1", "lblSubjectCaption", True)
    With lblSubjectCaption
        .Left = 8
        .Top = 32
        .Width = 60
        .Height = 17
        .Caption = "Subject:"
    End With

    Set txtSubject = Me.Controls.Add("Forms.TextBox.1", "txtSubject", True)
    With txtSubject
        .Left = 72
        .Top = 30
        .Width = 368
        .Height = 21
        .Enabled = False
        .Locked = True
    End With

    Set lstClassifications = Me.Controls.Add("Forms.ListBox.1", "lstClassifications", True)
    With lstClassifications
        .Left = 8
        .Top = 60
        .Width = 432
        .Height = 208
    End With

    Set lblDescription = Me.Controls.Add("Forms.Label.1", "lblDescription", True)
    With lblDescription
        .Left = 8
        .Top = 276
        .Width = 432
        .Height = 41
        .Caption = "(shows description of selected classification)"
        .WordWrap = True
    End With

    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK", True)
    With btnOK
        .Left = 264
        .Top = 348
        .Width = 81
        .Height = 25
        .Caption = "OK"
        .Default = True
    End With

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel", True)
    With btnCancel
        .Left = 356
        .Top = 348
        .Width = 81
        .Height = 25
        .Caption = "Cancel"
        .Cancel = True
    End With

    m_ControlsBuilt = True
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next

    txtSubject.Text = m_OriginalSubject

    Dim existingTag As String
    Dim i As Long

    existingTag = MSCANCore.GetExistingClassification(m_OriginalSubject)

    If Len(existingTag) > 0 Then
        For i = 0 To lstClassifications.ListCount - 1
            If StrComp(lstClassifications.List(i), existingTag, vbTextCompare) = 0 Then
                lstClassifications.ListIndex = i
                Exit For
            End If
        Next i
    End If

    lstClassifications.SetFocus
End Sub

Private Sub lstClassifications_Click()
    On Error Resume Next

    Dim idx As Long

    If lstClassifications.ListIndex < 0 Then
        lblDescription.Caption = ""
        Exit Sub
    End If

    idx = lstClassifications.ListIndex + 1
    lblDescription.Caption = MSCANCore.GetTagDescription(idx)
End Sub

Private Sub lstClassifications_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOK_Click
End Sub

Private Sub btnOK_Click()
    On Error GoTo ErrorHandler

    If lstClassifications.ListIndex < 0 Then
        MsgBox "Please select a security classification.", vbExclamation, "MSCAN"
        lstClassifications.SetFocus
        Exit Sub
    End If

    m_SelectedIndex = lstClassifications.ListIndex + 1
    m_SelectedTag = MSCANCore.GetTag(m_SelectedIndex)
    m_UpdatedSubject = MSCANCore.ApplyClassification(m_OriginalSubject, m_SelectedIndex)
    m_IsCancelled = False

    MSCANCore.Log "Classification selected: " & m_SelectedTag
    Me.Hide
    Exit Sub

ErrorHandler:
    MsgBox "Error applying classification: " & Err.Description, vbExclamation, "MSCAN"
End Sub

Private Sub btnCancel_Click()
    m_IsCancelled = True
    m_SelectedIndex = 0
    m_SelectedTag = ""
    m_UpdatedSubject = ""
    MSCANCore.Log "Classification cancelled by user"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        m_IsCancelled = True
        m_SelectedIndex = 0
        m_SelectedTag = ""
        m_UpdatedSubject = ""
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub PopulateClassificationList()
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim tagCount As Long

    lstClassifications.Clear
    tagCount = MSCANCore.GetTagCount()

    For i = 1 To tagCount
        lstClassifications.AddItem MSCANCore.GetTag(i)
    Next i
    Exit Sub

ErrorHandler:
    MsgBox "Error loading classifications: " & Err.Description, vbExclamation, "MSCAN"
End Sub
