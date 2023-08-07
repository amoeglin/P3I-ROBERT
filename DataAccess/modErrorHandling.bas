Attribute VB_Name = "modErrorHandling"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F85491100AD"
'
Option Explicit

' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
'##ModelId=3F85491100BB
Public Const MyObjectError1 = 1000
'##ModelId=3F85491100CB
Public Const MyObjectError2 = 1010
'##ModelId=3F85491100CC
Public Const MyObjectErrorN = 1234
'##ModelId=3F85491100CD
Public Const MyUnhandledError = 9999

' This function will retrieve an error description from a resource
' file (.RES).  The ErrorNum is the index of the string
' in the resource file.  Called by RaiseError
'##ModelId=3F85491100CE
Private Function GetErrorTextFromResource(ErrorNum As Long) As String
      On Error GoTo GetErrorTextFromResourceError
      Dim strMsg As String
     
      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)

      Exit Function
GetErrorTextFromResourceError:
      If Err.Number <> 0 Then
            GetErrorTextFromResource = "Une erreur inconnue est survenue!"
      End If
End Function

'There are a number of methods for retrieving the error
'message.  The following method uses a resource file to
'retrieve strings indexed by the error number you are
'raising.
'##ModelId=3F85491100DB
Public Sub RaiseError(ErrorNumber As Long, Source As String)
      Dim strErrorText As String

      strErrorText = GetErrorTextFromResource(ErrorNumber)

      'raise an error back to the client
      If strErrorText = "Une erreur inconnue est survenue!" Then
        Err.Raise vbObjectError + ErrorNumber, Source, strErrorText
      Else
        Err.Raise vbObjectError + ErrorNumber, Source, Err.Description
      End If
End Sub


