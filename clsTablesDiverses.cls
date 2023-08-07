VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTablesDiverses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75E303E3"
Option Explicit

'##ModelId=5C8A75E40105
Private NbTableDiverse As Integer
'##ModelId=5C8A75E40126
Private TableDiverse() As defTableDiverse

'##ModelId=5C8A75E40127
Public autoMode As Boolean
'

'##ModelId=5C8A75E40143
Private Sub Class_Initialize()
  ReDim TableDiverse(1)
  NbTableDiverse = 0
End Sub

'##ModelId=5C8A75E40162
Private Sub Class_Terminate()
  Clear
End Sub

'##ModelId=5C8A75E40172
Public Sub Clear()
  Erase TableDiverse
  NbTableDiverse = 0
End Sub

'##ModelId=5C8A75E40182
Public Sub FillCombo(cbo As ComboBox)
  Dim i As Integer

  cbo.Clear
  
  For i = 0 To NbTableDiverse - 1
    cbo.AddItem TableDiverse(i).nomTable
  Next
  
  cbo.ListIndex = 0
End Sub

'##ModelId=5C8A75E401B0
Public Sub AddTableDiverse(nomTable As String, orderBy As String, champsNonNullImport As String, champs As String)
  ReDim Preserve TableDiverse(NbTableDiverse + 1)
  
  Dim newDef As defTableDiverse
  
  Set newDef = New defTableDiverse
  
  With newDef
    .nomTable = nomTable
    .orderBy = orderBy
    .champsNonNullImport = champsNonNullImport
    .champs = champs
  End With
  
  Set TableDiverse(NbTableDiverse) = newDef
  
  NbTableDiverse = NbTableDiverse + 1
End Sub

'##ModelId=5C8A75E4021A
Public Function FindTableDiverse(nom As String) As defTableDiverse
  Dim i As Integer
  
  For i = 0 To NbTableDiverse - 1
    If TableDiverse(i).nomTable = nom Then
      Set FindTableDiverse = TableDiverse(i)
    End If
  Next
End Function

'##ModelId=5C8A75E40230
Public Function TableInfo(idx As Integer) As defTableDiverse
  Set TableInfo = TableDiverse(idx)
End Function
