Attribute VB_Name = "Module1"
' Module1
' Robust geolocation footer processing for Outlook emails

Option Explicit

' --- CONFIGURATION ---
Private Const PYTHON_EXE As String = "C:\Python313\python.exe"
Private Const PYTHON_SCRIPT As String = "C:\GeoFooter\geolocate_headers.py"
Private Const BASE_DIR As String = "C:\GeoFooter"
' --- END CONFIGURATION ---

' =========================
' Logging
' =========================
Public Sub WriteLog(msg As String)
    On Error Resume Next
    Dim logFolder As String: logFolder = BASE_DIR & "\Logs"
    EnsureFolderExists logFolder
    Dim logFile As String: logFile = logFolder & "\VBA_Log.txt"
    
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logFile, 8, True)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & msg
    ts.Close
    
    Debug.Print "LOG: " & msg
    Set ts = Nothing
    Set fso = Nothing
End Sub

Private Function EnsureFolderExists(ByVal folderPath As String) As Boolean
    On Error GoTo EH
    If Len(folderPath) = 0 Then EnsureFolderExists = False: Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    EnsureFolderExists = True
    Exit Function
EH:
    Debug.Print "EnsureFolderExists error: " & Err.Number & " - " & Err.Description & " (" & folderPath & ")"
    EnsureFolderExists = False
End Function

' =========================
' Main Processing
' =========================
Public Sub ProcessEmailSecurity(mail As Outlook.MailItem)
    On Error GoTo EH
    
    ' Step 1: Apply default footer (Module2.AddFooter)
    Module2.AddFooter mail
    WriteLog "ProcessEmailSecurity: Default footer applied."

    ' Step 2: Extract headers
    Dim headers As String: headers = GetInternetHeaders(mail)
    If Len(headers) = 0 Then
        WriteLog "ProcessEmailSecurity: No headers extracted, skipping Python geolocation."
        Exit Sub
    End If

    ' Step 3: Write headers to file
    Dim headerFile As String, footerFile As String
    headerFile = BASE_DIR & "\headers\headers_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    WriteHeadersToFile headers, headerFile
    
    ' Step 4: Prepare footer file path
    footerFile = Replace(headerFile, "\headers\", "\output\")
    footerFile = Replace(footerFile, "headers_", "footer_")
    footerFile = Replace(footerFile, ".txt", ".html")
    
    ' Step 5: Run Python script
    Dim actualFooterPath As String
    actualFooterPath = RunGeolocationPythonScript(headerFile, footerFile)
    
    ' Step 6: Wait for footer to exist
    Dim attempts As Integer: attempts = 0
    Do While Dir(actualFooterPath) = "" And attempts < 10
        Application.Wait (Now + TimeValue("0:00:01"))
        attempts = attempts + 1
    Loop
    
    ' Step 7: Insert generated footer if it exists
    If Dir(actualFooterPath) <> "" Then InsertFooterIntoMail mail, actualFooterPath
    
    Exit Sub
EH:
    WriteLog "ProcessEmailSecurity error: " & Err.Number & " - " & Err.Description
End Sub

' =========================
' Internet Headers
' =========================
Public Function GetInternetHeaders(mail As Outlook.MailItem) As String
    On Error GoTo EH
    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    
    Dim propAccessor As Outlook.PropertyAccessor
    Set propAccessor = mail.PropertyAccessor
    
    Dim headers As String
    headers = ""
    On Error Resume Next
    headers = CStr(propAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS))
    If Err.Number <> 0 Then
        WriteLog "GetInternetHeaders: PropertyAccessor failed: " & Err.Number & " - " & Err.Description
        headers = ""
    End If
    On Error GoTo EH
    
    WriteLog "GetInternetHeaders: Retrieved headers length = " & Len(headers)
    GetInternetHeaders = headers
    Exit Function
EH:
    WriteLog "GetInternetHeaders error: " & Err.Number & " - " & Err.Description
    GetInternetHeaders = ""
End Function

Private Sub WriteHeadersToFile(headers As String, filePath As String)
    On Error GoTo EH
    EnsureFolderExists CreateObject("Scripting.FileSystemObject").GetParentFolderName(filePath)
    WriteTextUtf8 filePath, headers
    Exit Sub
EH:
    WriteLog "WriteHeadersToFile error: " & Err.Number & " - " & Err.Description
End Sub

' =========================
' Python Geolocation
' =========================
Private Function RunGeolocationPythonScript(headerFilePath As String, footerFilePath As String) As String
    On Error GoTo EH
    EnsureFolderExists BASE_DIR & "\output"
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(PYTHON_EXE) Then
        WriteLog "RunGeolocationPythonScript: Python not found at " & PYTHON_EXE
        Exit Function
    End If
    If Not fso.FileExists(PYTHON_SCRIPT) Then
        WriteLog "RunGeolocationPythonScript: Python script not found at " & PYTHON_SCRIPT
        Exit Function
    End If
    
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim pythonErrLog As String: pythonErrLog = BASE_DIR & "\output\python_error.log"
    Dim cmd As String
    cmd = "cmd /c """ & PYTHON_EXE & " " & PYTHON_SCRIPT & " """ & _
          headerFilePath & """ """ & footerFilePath & """ 2> " & pythonErrLog & """"
    WriteLog "Running Python command: " & cmd
    sh.Run cmd, 0, True
    WriteLog "Python command executed. Footer should be at: " & footerFilePath
    
    RunGeolocationPythonScript = footerFilePath
    Exit Function
EH:
    WriteLog "RunGeolocationPythonScript error: " & Err.Number & " - " & Err.Description
    RunGeolocationPythonScript = ""
End Function

' =========================
' Insert Footer into Mail
' =========================
Private Sub InsertFooterIntoMail(mail As Outlook.MailItem, footerPath As String)
    On Error GoTo EH
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(footerPath) Then
        WriteLog "InsertFooterIntoMail: Footer not found: " & footerPath
        Exit Sub
    End If
    
    Dim footerHtml As String: footerHtml = ReadTextUtf8(footerPath)
    If Len(footerHtml) = 0 Then
        WriteLog "InsertFooterIntoMail: Footer empty or unreadable: " & footerPath
        Exit Sub
    End If
    
    ' Ensure mail is HTML
    If mail.BodyFormat <> olFormatHTML Then
        mail.BodyFormat = olFormatHTML
        mail.HTMLBody = "<pre>" & mail.Body & "</pre>"
    End If
    
    ' Insert footer before </body> if exists
    Dim bodyHtml As String: bodyHtml = mail.HTMLBody
    Dim pos As Long: pos = InStrRev(LCase$(bodyHtml), "</body>")
    If pos > 0 Then
        mail.HTMLBody = left$(bodyHtml, pos - 1) & footerHtml & Mid$(bodyHtml, pos)
    Else
        mail.HTMLBody = bodyHtml & "<hr>" & footerHtml
    End If
    
    mail.Save
    WriteLog "InsertFooterIntoMail: Footer inserted into email with subject: " & SafeSubject(mail)
    Exit Sub
EH:
    WriteLog "InsertFooterIntoMail error: " & Err.Number & " - " & Err.Description
End Sub

' =========================
' UTF-8 Read/Write Helpers
' =========================
Private Sub WriteTextUtf8(ByVal filePath As String, ByVal content As String)
    On Error GoTo EH
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2
    stm.Close
    Exit Sub
EH:
    WriteLog "WriteTextUtf8 error: " & Err.Number & " - " & Err.Description
End Sub

Private Function ReadTextUtf8(ByVal filePath As String) As String
    On Error GoTo EH
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextUtf8 = stm.ReadText(-1)
    stm.Close
    Exit Function
EH:
    WriteLog "ReadTextUtf8 error: " & Err.Number & " - " & Err.Description
    ReadTextUtf8 = ""
End Function

' =========================
' Safe Subject
' =========================
Private Function SafeSubject(mail As Outlook.MailItem) As String
    On Error Resume Next
    SafeSubject = mail.Subject
    If Err.Number <> 0 Then SafeSubject = "[No Subject]"
End Function


