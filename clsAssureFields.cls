VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssureFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsAssureField"
Attribute VB_Ext_KEY = "Member0" ,"clsAssureField"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A760B00D3"

'##ModelId=5C8A760B01DC
Private mCol As Collection


'##ModelId=5C8A760B01FC
Public Function Add(ID As Integer, DispalyField As String, DBQuery As String, SpreadColor As String, DataFillRequired As Boolean, IsDateColumn As Boolean) As AssureField

    Dim objAssureField As AssureField
    Set objAssureField = New AssureField
    
    objAssureField.ID = ID
    objAssureField.DispalyField = DispalyField
    objAssureField.DBQuery = DBQuery
    objAssureField.SpreadColor = SpreadColor
    objAssureField.DataFillRequired = DataFillRequired
    objAssureField.IsDateColumn = IsDateColumn
    
    mCol.Add objAssureField, ID

    Set Add = objAssureField
    Set objAssureField = Nothing

End Function

'##ModelId=5C8A760B028D
Public Sub AddFieldObject(field As AssureField)
    
    mCol.Add field, CStr(field.ID)

End Sub

'##ModelId=5C8A760B02AD
Public Property Get Item(vntIndexKey As Variant) As AssureField
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

'##ModelId=5C8A760B02D6
Public Property Get Count() As Long
    Count = mCol.Count
End Property

'##ModelId=5C8A760B02E5
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

'##ModelId=5C8A760B0314
Public Sub Clear()
    Set mCol = New Collection
End Sub

'##ModelId=5C8A760B0324
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Property

'##ModelId=5C8A760B0334
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

'##ModelId=5C8A760B0343
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

