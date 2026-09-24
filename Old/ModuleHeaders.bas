Attribute VB_Name = "ModuleHeaders"
Option Explicit

' Deprecated: use Module1.GetInternetHeaders instead.
Public Function GetInternetHeaders_Legacy(mail As Outlook.MailItem) As String
    On Error GoTo ErrorHandler
    Dim tag As String
    tag = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    GetInternetHeaders_Legacy = CStr(mail.PropertyAccessor.GetProperty(tag))
    Exit Function

ErrorHandler:
    GetInternetHeaders_Legacy = ""
End Function
