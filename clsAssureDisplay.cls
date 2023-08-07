VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssureDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A760C0007"
'This is a Display object and it contains a collection of fields (DB columns)
'##ModelId=5C8A760C0130
Private mAssureFields As New AssureFields

'##ModelId=5C8A760C018E
Private mID As Integer
'##ModelId=5C8A760C01AD
Private mName As String
'##ModelId=5C8A760C01CC
Private mDescription As String
'##ModelId=5C8A760C01EB
Private mUserName As String
'##ModelId=5C8A760C0217
Private mIsDefault As Boolean

'collection of Assure Fields
'##ModelId=5C8A760C0256
Public Property Get AssureFields() As AssureFields
   Set AssureFields = mAssureFields
End Property

'##ModelId=5C8A760C0237
Public Property Set AssureFields(ByVal vData As AssureFields)
   Set mAssureFields = vData
End Property

'Properties
'##ModelId=5C8A760C0275
Public Property Let ID(ByVal vData As Integer)
    mID = vData
End Property

'##ModelId=5C8A760C02A4
Public Property Get ID() As Integer
    ID = mID
End Property

'##ModelId=5C8A760C02B4
Public Property Let Name(ByVal vData As String)
    mName = vData
End Property

'##ModelId=5C8A760C02E2
Public Property Get Name() As String
    Name = mName
End Property

'##ModelId=5C8A760C02F2
Public Property Let Description(ByVal vData As String)
    mDescription = vData
End Property

'##ModelId=5C8A760C0321
Public Property Get Description() As String
    Description = mDescription
End Property

'##ModelId=5C8A760C0340
Public Property Let UserName(ByVal vData As String)
    mUserName = vData
End Property

'##ModelId=5C8A760C0360
Public Property Get UserName() As String
    UserName = mUserName
End Property

'##ModelId=5C8A760C037F
Public Property Let IsDefault(ByVal vData As Boolean)
    mIsDefault = vData
End Property

'##ModelId=5C8A760C039E
Public Property Get IsDefault() As Boolean
    IsDefault = mIsDefault
End Property

'##ModelId=5C8A760C03BD
Public Function ContainsFieldID(fieldID As Integer) As Boolean
  
  Dim field As AssureField
  
  ContainsFieldID = False
  
  For Each field In mAssureFields
    If field.ID = fieldID Then
      ContainsFieldID = True
      Exit For
    End If
  Next

End Function

'##ModelId=5C8A760D0004
Public Function ContainsFieldName(FieldName As String) As Boolean
  
  Dim field As AssureField
  
  ContainsFieldName = False
  
  For Each field In mAssureFields
    If LCase(field.DBQuery) = LCase(FieldName) Then
      ContainsFieldName = True
      Exit For
    End If
  Next

End Function




'### transfer the following methods to a transfer layer

'##ModelId=5C8A760D0033
Public Sub DeleteDisplay()
End Sub

'##ModelId=5C8A760D0043
Public Sub SaveDisplay()
End Sub

'##ModelId=5C8A760D0062
Public Function GetDefaultDisplay(arg1 As String) As AssureDisplay
End Function

'##ModelId=5C8A760D0091
Public Sub GetDisplay(arg1 As String)
End Sub



'### other methods:

'Private Sub cmdEmployeeDelete_Click()
'   If lstEmployees.ListIndex > -1 Then
'      sbMain.Employees.Delete _
'      Left(lstEmployees.text, 6)
'   End If
'   cmdListEmployees.Value = True
'End Sub
'
'Private Sub cmdListEmployees_Click()
'   Dim emp As Employee
'   lstEmployees.Clear
'   For Each emp In sbMain.Employees
'      lstEmployees.AddItem emp.ID & ", " & emp.Name _
'      & ", " & emp.Salary
'   Next
'End Sub

