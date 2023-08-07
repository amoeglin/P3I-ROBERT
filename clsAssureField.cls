VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssureField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A760A00A6"

'##ModelId=5C8A760A01B0
Private mvarID As Integer
'##ModelId=5C8A760A01CF
Private mvarDispalyField As String
'##ModelId=5C8A760A01EE
Private mvarDispalyFieldNoBrackets As String
'##ModelId=5C8A760A020D
Private mvarDBQuery As String
'##ModelId=5C8A760A0228
Private mvarSpreadColor As String
'##ModelId=5C8A760A0247
Private mvarDataFillRequired As Boolean
'##ModelId=5C8A760A025E
Private mvarIsDateColumn As Boolean

'##ModelId=5C8A760A027D
Public Property Let IsDateColumn(ByVal vData As Boolean)
    mvarIsDateColumn = vData
End Property

'##ModelId=5C8A760A02AC
Public Property Get IsDateColumn() As Boolean
    IsDateColumn = mvarIsDateColumn
End Property

'##ModelId=5C8A760A02CD
Public Property Let DataFillRequired(ByVal vData As Boolean)
    mvarDataFillRequired = vData
End Property

'##ModelId=5C8A760A02E6
Public Property Get DataFillRequired() As Boolean
    DataFillRequired = mvarDataFillRequired
End Property

'##ModelId=5C8A760A0305
Public Property Let SpreadColor(ByVal vData As String)
    mvarSpreadColor = vData
End Property

'##ModelId=5C8A760A0325
Public Property Get SpreadColor() As String
    SpreadColor = mvarSpreadColor
End Property

'##ModelId=5C8A760A0344
Public Property Let DBQuery(ByVal vData As String)
    mvarDBQuery = vData
End Property

'##ModelId=5C8A760A0363
Public Property Get DBQuery() As String
    DBQuery = mvarDBQuery
End Property

'##ModelId=5C8A760A0382
Public Property Let DispalyField(ByVal vData As String)
    mvarDispalyField = vData
End Property

'##ModelId=5C8A760A03B1
Public Property Get DispalyField() As String
    DispalyField = mvarDispalyField
End Property

'##ModelId=5C8A760A03D0
Public Property Let DispalyFieldNoBrackets(ByVal vData As String)
    mvarDispalyFieldNoBrackets = vData
End Property

'##ModelId=5C8A760B0027
Public Property Get DispalyFieldNoBrackets() As String
    DispalyFieldNoBrackets = mvarDispalyFieldNoBrackets
End Property

'##ModelId=5C8A760B0037
Public Property Let ID(ByVal vData As Integer)
    mvarID = vData
End Property

'##ModelId=5C8A760B0065
Public Property Get ID() As Integer
    ID = mvarID
End Property



