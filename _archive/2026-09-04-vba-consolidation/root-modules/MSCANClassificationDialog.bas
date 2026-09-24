Attribute VB_Name = "MSCANClassificationDialog"
'===============================================================================
' MSCANClassificationDialog - Classification picker (Python UI + InputBox fallback)
'===============================================================================
Option Explicit

Private Const RESULT_FILE_NAME As String = "classify_result.json"
Private Const DIALOG_SCRIPT_DEFAULT As String = "C:\GeoFooter\VBA\aes_classify_dialog.py"

Public Function ShowDialog( _
    ByVal originalSubject As String, _
    ByRef selectedTag As String, _
    ByRef updatedSubject As String, _
    Optional ByRef readNotify As Boolean = False, _
    Optional ByVal isAliniantSend As Boolean = False, _
    Optional ByRef readNotifySuffix As String = "") As Boolean

    On Error GoTo ErrorHandler

    ShowDialog = True
    selectedTag = ""
    updatedSubject = ""
    readNotify = False
    readNotifySuffix = ""

    Dim tagCount As Long
    tagCount = MSCANCore.GetTagCount()
    If tagCount = 0 Then
        MsgBox "No classification tags are configured.", vbCritical, "AES"
        Exit Function
    End If

    Dim defaultChoice As Long
    defaultChoice = GetDefaultChoiceIndex(originalSubject)

    If TryShowPythonDialog(originalSubject, defaultChoice, isAliniantSend, selectedTag, updatedSubject, readNotify, readNotifySuffix) Then
        If Len(selectedTag) > 0 Then
            MSCANCore.Log "Classification selected (Python): " & selectedTag & _
                " | ReadNotify=" & IIf(readNotify, "Yes", "No") & _
                IIf(Len(readNotifySuffix) > 0, " (" & readNotifySuffix & ")", "")
            ShowDialog = False
        End If
        Exit Function
    End If

    ' Fallback: InputBox (legacy) — default ReadNotify from TRACKED levels when Aliniant
    ShowDialog = ShowInputBoxDialog(originalSubject, defaultChoice, selectedTag, updatedSubject)
    If Not ShowDialog Then
        readNotify = isAliniantSend And MSCANCore.RequiresTracking(MSCANCore.GetTagIndex(selectedTag))
        If readNotify Then readNotifySuffix = ".read-notify.com"
    End If
    Exit Function

ErrorHandler:
    MSCANCore.LogError "MSCANClassificationDialog.ShowDialog", Err.Number, Err.Description
    MsgBox "Could not show classification dialog: " & Err.Description, vbCritical, "AES"
    ShowDialog = True
End Function

Private Function GetDefaultChoiceIndex(ByVal originalSubject As String) As Long
    Dim existingTag As String
    Dim i As Long

    GetDefaultChoiceIndex = 1
    existingTag = MSCANCore.GetExistingClassification(originalSubject)
    If Len(existingTag) = 0 Then Exit Function

    For i = 1 To MSCANCore.GetTagCount()
        If StrComp(MSCANCore.GetTag(i), existingTag, vbTextCompare) = 0 Then
            GetDefaultChoiceIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function TryShowPythonDialog( _
    ByVal originalSubject As String, _
    ByVal defaultChoice As Long, _
    ByVal isAliniantSend As Boolean, _
    ByRef selectedTag As String, _
    ByRef updatedSubject As String, _
    ByRef readNotify As Boolean, _
    ByRef readNotifySuffix As String) As Boolean

    On Error GoTo EH

    TryShowPythonDialog = False
    selectedTag = ""
    updatedSubject = ""
    readNotify = False
    readNotifySuffix = ""

    Dim pythonExe As String
    Dim scriptPath As String
    Dim outPath As String

    pythonExe = ResolvePythonExe()
    scriptPath = ResolveDialogScript()
    If Len(pythonExe) = 0 Or Len(scriptPath) = 0 Then
        MSCANCore.Log "Classification Python dialog unavailable; using InputBox fallback."
        Exit Function
    End If

    outPath = Environ$("LOCALAPPDATA") & "\GeoFooter\" & RESULT_FILE_NAME
    EnsureParentFolder outPath

    On Error Resume Next
    Kill outPath
    Err.Clear
    On Error GoTo EH

    Dim aliniantFlag As String
    If isAliniantSend Then
        aliniantFlag = "true"
    Else
        aliniantFlag = "false"
    End If

    Dim cmd As String
    cmd = """" & pythonExe & """ """ & scriptPath & """" & _
          " --subject " & ShellQuote(originalSubject) & _
          " --default " & CStr(defaultChoice) & _
          " --aliniant " & aliniantFlag & _
          " --out " & ShellQuote(outPath)

    MSCANCore.Log "Classification dialog launch: " & pythonExe

    Dim shell As Object
    Dim exitCode As Long
    Set shell = CreateObject("WScript.Shell")
    ' pythonw.exe = no console. WindowStyle 1 = show the Qt dialog (0 hides it and blocks Send).
    exitCode = shell.Run(cmd, 1, True)

    If exitCode <> 0 Then
        MSCANCore.Log "Classification Python dialog exit code " & exitCode & "; fallback."
        Exit Function
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(outPath) Then
        MSCANCore.Log "Classification result file missing; fallback."
        Exit Function
    End If

    Dim raw As String
    raw = ReadAllText(outPath)
    If Len(raw) = 0 Then
        MSCANCore.Log "Classification result file empty; fallback."
        Exit Function
    End If

    Dim cancelled As Boolean
    Dim idx As Long
    Dim tag As String
    Dim subj As String
    Dim rn As Boolean
    Dim rnSuffix As String

    If Not ParseClassifyResult(raw, cancelled, idx, tag, subj, rn, rnSuffix) Then
        MSCANCore.Log "Classification result JSON parse failed; fallback."
        Exit Function
    End If

    TryShowPythonDialog = True
    If cancelled Then
        selectedTag = ""
        updatedSubject = ""
        readNotify = False
        readNotifySuffix = ""
        Exit Function
    End If

    If idx < 1 Or idx > MSCANCore.GetTagCount() Then
        TryShowPythonDialog = False
        Exit Function
    End If

    selectedTag = MSCANCore.GetTag(idx)
    If Len(tag) > 0 Then selectedTag = tag
    If Len(subj) > 0 Then
        updatedSubject = subj
    Else
        updatedSubject = MSCANCore.ApplyClassification(originalSubject, idx)
    End If
    readNotify = rn
    readNotifySuffix = rnSuffix
    If readNotify And Len(readNotifySuffix) = 0 Then readNotifySuffix = ".read-notify.com"
    Exit Function

EH:
    MSCANCore.LogError "MSCANClassificationDialog.TryShowPythonDialog", Err.Number, Err.Description
    TryShowPythonDialog = False
End Function

Private Function ShowInputBoxDialog( _
    ByVal originalSubject As String, _
    ByVal defaultChoice As Long, _
    ByRef selectedTag As String, _
    ByRef updatedSubject As String) As Boolean

    On Error GoTo ErrorHandler

    Dim tagCount As Long
    Dim listText As String
    Dim i As Long

    ShowInputBoxDialog = True
    selectedTag = ""
    updatedSubject = ""

    tagCount = MSCANCore.GetTagCount()
    listText = "Subject:" & vbCrLf & originalSubject & vbCrLf & vbCrLf
    listText = listText & "Enter the number for the security classification:" & vbCrLf & vbCrLf

    For i = 1 To tagCount
        listText = listText & CStr(i) & ". " & MSCANCore.GetTag(i) & vbCrLf
        listText = listText & "   " & MSCANCore.GetTagDescription(i) & vbCrLf & vbCrLf
    Next i

    listText = listText & "Leave blank and click Cancel to abort sending."

    Do
        Dim choice As String
        Dim idx As Long
        Dim confirm As VbMsgBoxResult

        choice = InputBox(listText, "Select Security Classification", CStr(defaultChoice))
        If Len(choice) = 0 Then Exit Function

        idx = CLng(Val(choice))
        If idx < 1 Or idx > tagCount Then
            MsgBox "Please enter a number between 1 and " & tagCount & ".", vbExclamation, "AES"
            GoTo ContinueLoop
        End If

        selectedTag = MSCANCore.GetTag(idx)
        updatedSubject = MSCANCore.ApplyClassification(originalSubject, idx)

        confirm = MsgBox( _
            "Apply this classification?" & vbCrLf & vbCrLf & _
            selectedTag & vbCrLf & vbCrLf & _
            "New subject:" & vbCrLf & updatedSubject, _
            vbQuestion + vbYesNoCancel, "Confirm Classification")

        If confirm = vbYes Then
            MSCANCore.Log "Classification selected (InputBox): " & selectedTag
            ShowInputBoxDialog = False
            Exit Function
        ElseIf confirm = vbCancel Then
            Exit Function
        End If

ContinueLoop:
    Loop

    Exit Function

ErrorHandler:
    MSCANCore.LogError "MSCANClassificationDialog.ShowInputBoxDialog", Err.Number, Err.Description
    ShowInputBoxDialog = True
End Function

Private Function ResolvePythonExe() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Dialog MUST use pythonw.exe — python.exe always opens a console.
    Dim paths As Variant
    paths = Array( _
        "C:\Python313\pythonw.exe", _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\pythonw.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\pythonw.exe", _
        "C:\GeoFooter\.venv\Scripts\pythonw.exe", _
        Environ$("ProgramFiles") & "\Python\pythonw.exe")

    Dim p As Variant
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolvePythonExe = CStr(p)
            Exit Function
        End If
    Next p

    ' Last resort: derive pythonw.exe next to a found python.exe
    Dim pyPaths As Variant
    pyPaths = Array( _
        "C:\Python313\python.exe", _
        Environ$("USERPROFILE") & "\AppData\Local\Programs\Python\Python313\python.exe", _
        Environ$("LOCALAPPDATA") & "\Programs\Python\Python313\python.exe", _
        "C:\GeoFooter\.venv\Scripts\python.exe")

    Dim py As Variant
    Dim sibling As String
    For Each py In pyPaths
        If fso.FileExists(CStr(py)) Then
            sibling = fso.GetParentFolderName(CStr(py)) & "\pythonw.exe"
            If fso.FileExists(sibling) Then
                ResolvePythonExe = sibling
                Exit Function
            End If
        End If
    Next py

    ResolvePythonExe = ""
End Function

' Launch with no console. Uses a temporary WScript launcher so neither
' PowerShell nor cmd.exe windows appear.
Private Function RunHiddenProcess(ByVal commandLine As String) As Long
    On Error GoTo EH

    Dim fso As Object
    Dim folder As String
    Dim vbsPath As String
    Dim ts As Object
    Dim shell As Object
    Dim exitCode As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = Environ$("LOCALAPPDATA") & "\GeoFooter"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    vbsPath = folder & "\aes_launch_hidden.vbs"

    Set ts = fso.CreateTextFile(vbsPath, True, False)
    ts.Write "Dim sh, code" & vbCrLf
    ts.Write "Set sh = CreateObject(""WScript.Shell"")" & vbCrLf
    ts.Write "code = sh.Run(" & VbsQuote(commandLine) & ", 0, True)" & vbCrLf
    ts.Write "WScript.Quit code" & vbCrLf
    ts.Close

    Set shell = CreateObject("WScript.Shell")
    ' wscript.exe has no console; //B suppresses script UI.
    exitCode = shell.Run("wscript.exe //B //Nologo """ & vbsPath & """", 0, True)
    RunHiddenProcess = exitCode
    Exit Function

EH:
    MSCANCore.LogError "MSCANClassificationDialog.RunHiddenProcess", Err.Number, Err.Description
    RunHiddenProcess = -1
End Function

Private Function VbsQuote(ByVal value As String) As String
    ' Produce a VBScript string literal: "...." with quotes doubled.
    VbsQuote = """" & Replace(value, """", """""") & """"
End Function

Private Function ResolveDialogScript() As String
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim paths As Variant
    paths = Array( _
        DIALOG_SCRIPT_DEFAULT, _
        "C:\GeoFooter\aes_classify_dialog.py", _
        Environ$("LOCALAPPDATA") & "\GeoFooter\aes_classify_dialog.py")

    Dim p As Variant
    For Each p In paths
        If fso.FileExists(CStr(p)) Then
            ResolveDialogScript = CStr(p)
            Exit Function
        End If
    Next p
    ResolveDialogScript = ""
End Function

Private Sub EnsureParentFolder(ByVal filePath As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parent As String
    parent = fso.GetParentFolderName(filePath)
    If Len(parent) = 0 Then Exit Sub
    If Not fso.FolderExists(parent) Then fso.CreateFolder parent
End Sub

Private Function ShellQuote(ByVal value As String) As String
    ShellQuote = """" & Replace(value, """", """""") & """"
End Function

Private Function ReadAllText(ByVal filePath As String) As String
    On Error GoTo EH
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2)
    If Not ts.AtEndOfStream Then
        ReadAllText = ts.ReadAll
    Else
        ReadAllText = ""
    End If
    ts.Close
    Exit Function
EH:
    ReadAllText = ""
End Function

Private Function ParseClassifyResult( _
    ByVal raw As String, _
    ByRef cancelled As Boolean, _
    ByRef idx As Long, _
    ByRef tag As String, _
    ByRef subj As String, _
    ByRef readNotify As Boolean, _
    ByRef readNotifySuffix As String) As Boolean

    On Error GoTo EH

    cancelled = True
    idx = 0
    tag = ""
    subj = ""
    readNotify = False
    readNotifySuffix = ""

    Dim cancelledText As String
    cancelledText = JsonGetBoolText(raw, "cancelled")
    If Len(cancelledText) = 0 Then
        ParseClassifyResult = False
        Exit Function
    End If
    cancelled = (StrComp(cancelledText, "true", vbTextCompare) = 0)

    idx = CLng(Val(JsonGetNumberText(raw, "index")))
    tag = JsonGetString(raw, "tag")
    subj = JsonGetString(raw, "subject")

    Dim rnText As String
    rnText = JsonGetBoolText(raw, "read_notify")
    readNotify = (StrComp(rnText, "true", vbTextCompare) = 0)
    readNotifySuffix = JsonGetString(raw, "read_notify_suffix")

    ParseClassifyResult = True
    Exit Function
EH:
    ParseClassifyResult = False
End Function

Private Function JsonGetString(ByVal raw As String, ByVal key As String) As String
    On Error Resume Next
    Dim pattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim chunk As String

    pattern = """" & key & """"
    startPos = InStr(1, raw, pattern, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, raw, ":", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, raw, """", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + 1
    endPos = InStr(startPos, raw, """", vbBinaryCompare)
    If endPos = 0 Then Exit Function
    chunk = Mid$(raw, startPos, endPos - startPos)
    JsonGetString = Replace(chunk, "\/", "/")
    JsonGetString = Replace(JsonGetString, "\""", """")
    JsonGetString = Replace(JsonGetString, "\n", vbLf)
    JsonGetString = Replace(JsonGetString, "\r", "")
    JsonGetString = Replace(JsonGetString, "\\", "\")
End Function

Private Function JsonGetNumberText(ByVal raw As String, ByVal key As String) As String
    On Error Resume Next
    Dim pattern As String
    Dim startPos As Long
    Dim i As Long
    Dim ch As String
    Dim buf As String

    pattern = """" & key & """"
    startPos = InStr(1, raw, pattern, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, raw, ":", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + 1
    Do While startPos <= Len(raw) And (Mid$(raw, startPos, 1) = " " Or Mid$(raw, startPos, 1) = vbTab)
        startPos = startPos + 1
    Loop
    buf = ""
    For i = startPos To Len(raw)
        ch = Mid$(raw, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Then
            buf = buf & ch
        Else
            Exit For
        End If
    Next i
    JsonGetNumberText = buf
End Function

Private Function JsonGetBoolText(ByVal raw As String, ByVal key As String) As String
    On Error Resume Next
    Dim pattern As String
    Dim startPos As Long
    Dim chunk As String

    pattern = """" & key & """"
    startPos = InStr(1, raw, pattern, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = InStr(startPos, raw, ":", vbBinaryCompare)
    If startPos = 0 Then Exit Function
    chunk = LCase$(Mid$(raw, startPos + 1, 10))
    If InStr(1, chunk, "true", vbTextCompare) > 0 Then
        JsonGetBoolText = "true"
    ElseIf InStr(1, chunk, "false", vbTextCompare) > 0 Then
        JsonGetBoolText = "false"
    End If
End Function

Public Sub QuickTest()
    Dim selectedTag As String
    Dim updatedSubject As String
    Dim readNotify As Boolean
    Dim readNotifySuffix As String
    Dim cancelled As Boolean

    cancelled = ShowDialog("Test email subject line", selectedTag, updatedSubject, readNotify, True, readNotifySuffix)

    If cancelled Then
        MsgBox "Cancelled.", vbInformation, "AES"
    Else
        MsgBox "Selected: " & selectedTag & vbCrLf & _
               "ReadNotify: " & IIf(readNotify, "Yes", "No") & vbCrLf & _
               "Suffix: " & readNotifySuffix & vbCrLf & vbCrLf & _
               "Updated subject:" & vbCrLf & updatedSubject, vbInformation, "AES"
    End If
End Sub
