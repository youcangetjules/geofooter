Attribute VB_Name = "MSCANCore"
'===============================================================================
' MSCANCore - Core Classification Engine
' Author: Julian Garrett (Rewritten)
' Purpose: Central module for email classification detection and manipulation
'===============================================================================
Option Explicit

'===============================================================================
' CONFIGURATION CONSTANTS
'===============================================================================

' Number of classification tags
Private Const TAG_COUNT As Long = 8

'===============================================================================
' MODULE-LEVEL VARIABLES
'===============================================================================

Private m_ClassificationTags(1 To 8) As String
Private m_TagDescriptions(1 To 8) As String
Private m_TagsInitialized As Boolean

'===============================================================================
' INITIALIZATION
'===============================================================================

Private Sub InitializeTags()
    '''Initializes the classification tags array
    '''Modify these tags to match your organization's requirements
    
    If m_TagsInitialized Then Exit Sub
    
    ' Classification Tags
    m_ClassificationTags(1) = "[NR/E]"
    m_ClassificationTags(2) = "[SEC1: (C) NOT RATED /EXTERNAL]"
    m_ClassificationTags(3) = "[SEC2: (C) COMMERCIAL-IN-CONFIDENCE /UNENCRYPTED]"
    m_ClassificationTags(4) = "[SEC3: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/EDIT-ENCRYPTED]"
    m_ClassificationTags(5) = "[SEC4: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]"
    m_ClassificationTags(6) = "[SEC5: (D) OFFICIAL-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]"
    m_ClassificationTags(7) = "[SEC6: (D) CLASSIFIED/SIGNED/ENCRYPTED/TRACKED]"
    m_ClassificationTags(8) = "[SEC7: (D) SOA/SIGNED/ENCRYPTED/TRACKED]"
    
    ' Tag Descriptions (friendly names for UI)
    m_TagDescriptions(1) = "Minimal Marking with no security ratings or markings or Aliniant Internal"
    m_TagDescriptions(2) = "Not Security Rated, Non-commercial and sent externally"
    m_TagDescriptions(3) = "Commercial-in-Confidence / Unencrypted"
    m_TagDescriptions(4) = "Commercial-in-Confidence-Sensitive / Signed / Edit-Encrypted"
    m_TagDescriptions(5) = "Commercial-in-Confidence-Sensitive / Signed / Encrypted / Tracked"
    m_TagDescriptions(6) = "Official-Sensitive / Signed / Encrypted / Tracked"
    m_TagDescriptions(7) = "Classified / Signed / Encrypted / Tracked"
    m_TagDescriptions(8) = "Secret or Above / Signed / Encrypted / Tracked"
    
    m_TagsInitialized = True
End Sub

'===============================================================================
' PUBLIC FUNCTIONS - TAG RETRIEVAL
'===============================================================================

Public Function GetTagCount() As Long
    '''Returns the number of configured classification tags
    GetTagCount = TAG_COUNT
End Function

Public Function GetTag(ByVal index As Long) As String
    '''Returns the classification tag at the specified index (1-based)
    '''
    '''Args:
    '''    index: 1-based index of the tag (1 to 8)
    '''
    '''Returns:
    '''    The tag string, or empty string if index is invalid
    
    InitializeTags
    
    If index < 1 Or index > TAG_COUNT Then
        GetTag = ""
        Exit Function
    End If
    
    GetTag = m_ClassificationTags(index)
End Function

Public Function GetTagDescription(ByVal index As Long) As String
    '''Returns the description for the tag at the specified index (1-based)
    '''
    '''Args:
    '''    index: 1-based index of the tag (1 to 8)
    '''
    '''Returns:
    '''    The description string, or empty string if index is invalid
    
    InitializeTags
    
    If index < 1 Or index > TAG_COUNT Then
        GetTagDescription = ""
        Exit Function
    End If
    
    GetTagDescription = m_TagDescriptions(index)
End Function

Public Function GetDisplayText(ByVal index As Long) As String
    '''Returns combined display text for listbox/combobox
    '''
    '''Args:
    '''    index: 1-based index of the tag (1 to 8)
    '''
    '''Returns:
    '''    Formatted display string combining tag and description
    
    InitializeTags
    
    If index < 1 Or index > TAG_COUNT Then
        GetDisplayText = ""
        Exit Function
    End If
    
    GetDisplayText = m_TagDescriptions(index) & " - " & m_ClassificationTags(index)
End Function

Public Function GetClassificationTags() As Variant
    '''Returns array of all valid classification tags
    '''
    '''Returns:
    '''    Variant array containing all classification tag strings
    
    InitializeTags
    GetClassificationTags = m_ClassificationTags
End Function

'===============================================================================
' PUBLIC FUNCTIONS - CLASSIFICATION DETECTION
'===============================================================================

Public Function DetectClassification(ByVal subjectLine As String) As String
    '''Detects and extracts classification tag from subject line
    '''
    '''Args:
    '''    subjectLine: The email subject to analyze
    '''
    '''Returns:
    '''    The classification tag if found, empty string otherwise
    
    InitializeTags
    
    If Len(Trim$(subjectLine)) = 0 Then
        DetectClassification = ""
        Exit Function
    End If
    
    ' Only known tags count as classified (unknown [SEC…] must not skip the send dialog).
    Dim i As Long
    For i = 1 To TAG_COUNT
        If InStr(1, subjectLine, m_ClassificationTags(i), vbTextCompare) > 0 Then
            DetectClassification = m_ClassificationTags(i)
            Exit Function
        End If
    Next i

    DetectClassification = ""
End Function

Public Function HasClassification(ByVal subjectLine As String) As Boolean
    '''Checks if subject line contains any classification tag
    '''
    '''Args:
    '''    subjectLine: The email subject to check
    '''
    '''Returns:
    '''    True if a classification tag is present, False otherwise
    
    HasClassification = (Len(DetectClassification(subjectLine)) > 0)
End Function

Public Function GetExistingClassification(ByVal subjectLine As String) As String
    '''Alias for DetectClassification for compatibility
    '''
    '''Args:
    '''    subjectLine: The email subject to check
    '''
    '''Returns:
    '''    The tag found, or empty string if none
    
    GetExistingClassification = DetectClassification(subjectLine)
End Function

Private Function ExtractTagByPattern(ByVal subjectLine As String) As String
    '''Used by RemoveClassification to strip leftover [SEC…] / [NR/E] bracket text.
    '''Not used for HasClassification / DetectClassification (known tags only).
    
    Dim startPos As Long
    Dim endPos As Long
    Dim searchText As String
    
    searchText = UCase$(subjectLine)
    
    ' First check for [NR/E]
    If InStr(1, searchText, "[NR/E]", vbTextCompare) > 0 Then
        ExtractTagByPattern = "[NR/E]"
        Exit Function
    End If
    
    ' Look for [SEC pattern
    startPos = InStr(1, searchText, "[SEC", vbTextCompare)
    
    If startPos = 0 Then
        ExtractTagByPattern = ""
        Exit Function
    End If
    
    ' Find the closing bracket
    endPos = InStr(startPos, searchText, "]", vbTextCompare)
    
    If endPos = 0 Then
        ExtractTagByPattern = ""
        Exit Function
    End If
    
    ' Extract the tag
    ExtractTagByPattern = Mid$(subjectLine, startPos, endPos - startPos + 1)
End Function

'===============================================================================
' PUBLIC FUNCTIONS - CLASSIFICATION APPLICATION
'===============================================================================

Public Function ApplyClassification(ByVal subjectLine As String, ByVal tagIndex As Long) As String
    '''Applies a classification tag to a subject line by index
    '''Removes any existing classification first
    '''
    '''Args:
    '''    subjectLine: Original subject line
    '''    tagIndex: 1-based index of the tag to apply (1 to 8)
    '''
    '''Returns:
    '''    New subject line with classification applied
    
    InitializeTags
    
    Dim cleanSubject As String
    Dim tag As String
    
    ' Validate index
    If tagIndex < 1 Or tagIndex > TAG_COUNT Then
        ApplyClassification = subjectLine
        Exit Function
    End If
    
    ' Remove any existing classification
    cleanSubject = RemoveClassification(subjectLine)
    
    ' Trim and get new tag
    cleanSubject = Trim$(cleanSubject)
    tag = m_ClassificationTags(tagIndex)
    
    ' Apply new classification at the start
    If Len(cleanSubject) > 0 Then
        ApplyClassification = tag & " " & cleanSubject
    Else
        ApplyClassification = tag
    End If
End Function

Public Function ApplyClassificationByTag(ByVal subjectLine As String, ByVal classificationTag As String) As String
    '''Applies a classification tag string to a subject line
    '''Removes any existing classification first
    '''
    '''Args:
    '''    subjectLine: Original subject line
    '''    classificationTag: Tag string to apply (e.g., "[SEC1: (C) NOT RATED /EXTERNAL]")
    '''
    '''Returns:
    '''    New subject line with classification applied
    
    Dim cleanSubject As String
    Dim idx As Long

    idx = GetTagIndex(classificationTag)
    If idx < 1 Then
        ' Reject unknown tags — do not stamp arbitrary bracket text onto subjects.
        ApplyClassificationByTag = subjectLine
        Exit Function
    End If

    ' Remove any existing classification
    cleanSubject = RemoveClassification(subjectLine)
    
    ' Trim and apply canonical known tag at the start
    cleanSubject = Trim$(cleanSubject)
    classificationTag = m_ClassificationTags(idx)

    If Len(cleanSubject) > 0 Then
        ApplyClassificationByTag = classificationTag & " " & cleanSubject
    Else
        ApplyClassificationByTag = classificationTag
    End If
End Function

Public Function RemoveClassification(ByVal subjectLine As String) As String
    '''Removes classification tag from subject line
    '''
    '''Args:
    '''    subjectLine: Subject line potentially containing a tag
    '''
    '''Returns:
    '''    Subject line with classification removed
    
    InitializeTags
    
    Dim result As String
    Dim i As Long
    
    result = subjectLine
    
    ' Remove known tags
    For i = 1 To TAG_COUNT
        result = Replace(result, m_ClassificationTags(i), "", 1, -1, vbTextCompare)
    Next i
    
    ' Also try to remove any pattern-matched tag
    Dim existingTag As String
    existingTag = ExtractTagByPattern(result)
    
    If Len(existingTag) > 0 Then
        result = Replace(result, existingTag, "", 1, -1, vbTextCompare)
    End If
    
    ' Clean up extra spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    
    RemoveClassification = Trim$(result)
End Function

'===============================================================================
' PUBLIC FUNCTIONS - VALIDATION
'===============================================================================

Public Function IsValidClassification(ByVal classificationTag As String) As Boolean
    '''Validates if a tag is a recognized classification
    '''
    '''Args:
    '''    classificationTag: Tag to validate
    '''
    '''Returns:
    '''    True if valid, False otherwise
    
    InitializeTags
    
    Dim i As Long
    For i = 1 To TAG_COUNT
        If StrComp(classificationTag, m_ClassificationTags(i), vbTextCompare) = 0 Then
            IsValidClassification = True
            Exit Function
        End If
    Next i
    
    IsValidClassification = False
End Function

Public Function GetClassificationLevel(ByVal classificationTag As String) As Long
    '''Returns the security level (1-8) for a classification tag
    '''Higher number = higher security
    '''
    '''Args:
    '''    classificationTag: Tag to evaluate
    '''
    '''Returns:
    '''    Level number (1-8) or 0 if not recognized
    
    InitializeTags
    
    Dim i As Long
    For i = 1 To TAG_COUNT
        If StrComp(classificationTag, m_ClassificationTags(i), vbTextCompare) = 0 Then
            GetClassificationLevel = i
            Exit Function
        End If
    Next i
    
    GetClassificationLevel = 0
End Function

Public Function GetTagIndex(ByVal classificationTag As String) As Long
    '''Returns the index (1-8) for a classification tag
    '''
    '''Args:
    '''    classificationTag: Tag to find
    '''
    '''Returns:
    '''    Index number (1-8) or 0 if not found
    
    GetTagIndex = GetClassificationLevel(classificationTag)
End Function

'===============================================================================
' LOGGING FUNCTIONS
'===============================================================================

Public Sub Log(ByVal message As String)
    '''Writes a message to the shared AES VBA log (MSCANModLogging / VBA_Log.txt).
    On Error Resume Next
    MSCANModLogging.WriteLog message
End Sub

Public Sub LogError(ByVal source As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    '''Logs an error with context to the shared AES VBA log.
    On Error Resume Next
    MSCANModLogging.WriteLogLevel "error", _
        "ERROR in " & source & " | #" & errorNumber & " | " & errorDescription
End Sub

'===============================================================================
' DIAGNOSTIC HELPER FUNCTIONS
'===============================================================================

Public Function IsLoggingEnabled() As Boolean
    '''Returns whether AES VBA logging is enabled (aes_logging.json / defaults).
    On Error Resume Next
    IsLoggingEnabled = MSCANModLogging.IsLoggingEnabled()
End Function

Public Function GetLogFilePath() As String
    '''Returns the shared AES VBA log path (%LOCALAPPDATA%\GeoFooter\Logs\VBA_Log.txt).
    On Error Resume Next
    GetLogFilePath = MSCANModLogging.logPath()
End Function

Public Function GetVersion() As String
    '''Returns the MSCAN version string
    GetVersion = "2.0.0"
End Function

Public Function GetDetectionPattern() As String
    '''Returns a description of the classification detection pattern
    '''Used for diagnostics — detection matches known tags only.
    GetDetectionPattern = "Known tags only: [NR/E] or [SEC1]…[SEC7] (exact)"
End Function

'===============================================================================
' UTILITY FUNCTIONS
'===============================================================================

Public Function GetTagDisplayName(ByVal classificationTag As String) As String
    '''Extracts a simplified display name from a classification tag
    '''
    '''Args:
    '''    classificationTag: Full tag string
    '''
    '''Returns:
    '''    Simplified display name
    
    Dim index As Long
    index = GetTagIndex(classificationTag)
    
    If index > 0 Then
        GetTagDisplayName = m_TagDescriptions(index)
    Else
        GetTagDisplayName = classificationTag
    End If
End Function

Public Function GetTagShortCode(ByVal classificationTag As String) As String
    '''Extracts the short code from a classification tag
    '''E.g., "[SEC1: (C) NOT RATED /EXTERNAL]" returns "C"
    '''
    '''Args:
    '''    classificationTag: Full tag string
    '''
    '''Returns:
    '''    Short code letter(s) or empty string
    
    Dim startPos As Long
    Dim endPos As Long
    
    ' Handle [NR/E] specially
    If InStr(1, classificationTag, "[NR/E]", vbTextCompare) > 0 Then
        GetTagShortCode = "NR"
        Exit Function
    End If
    
    ' Find the ( which precedes the code
    startPos = InStr(classificationTag, "(")
    If startPos = 0 Then
        GetTagShortCode = ""
        Exit Function
    End If
    
    ' Find the closing )
    endPos = InStr(startPos, classificationTag, ")")
    If endPos = 0 Then
        GetTagShortCode = ""
        Exit Function
    End If
    
    GetTagShortCode = Mid$(classificationTag, startPos + 1, endPos - startPos - 1)
End Function

Public Function GetSecurityCategory(ByVal tagIndex As Long) As String
    '''Returns the security category for a tag
    '''
    '''Args:
    '''    tagIndex: 1-based index of the tag
    '''
    '''Returns:
    '''    "C" for Commercial, "D" for Defence, "NR" for Not Rated
    
    Select Case tagIndex
        Case 1
            GetSecurityCategory = "NR"
        Case 2 To 5
            GetSecurityCategory = "C"
        Case 6 To 8
            GetSecurityCategory = "D"
        Case Else
            GetSecurityCategory = ""
    End Select
End Function

Public Function RequiresEncryption(ByVal tagIndex As Long) As Boolean
    '''Determines if a classification level requires encryption
    '''
    '''Args:
    '''    tagIndex: 1-based index of the tag
    '''
    '''Returns:
    '''    True if encryption is required
    
    Select Case tagIndex
        Case 4 To 8
            RequiresEncryption = True
        Case Else
            RequiresEncryption = False
    End Select
End Function

Public Function RequiresTracking(ByVal tagIndex As Long) As Boolean
    '''Determines if a classification level requires tracking
    '''
    '''Args:
    '''    tagIndex: 1-based index of the tag
    '''
    '''Returns:
    '''    True if tracking is required
    
    Select Case tagIndex
        Case 5 To 8
            RequiresTracking = True
        Case Else
            RequiresTracking = False
    End Select
End Function

Public Function RequiresSigning(ByVal tagIndex As Long) As Boolean
    '''Determines if a classification level requires digital signing
    '''
    '''Args:
    '''    tagIndex: 1-based index of the tag
    '''
    '''Returns:
    '''    True if signing is required
    
    Select Case tagIndex
        Case 4 To 8
            RequiresSigning = True
        Case Else
            RequiresSigning = False
    End Select
End Function
