VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDetailAffichagePeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A76010245"
Option Explicit

'##ModelId=5C8A76010354
Public DonneesBrutes As Boolean
'##ModelId=5C8A76010364
Public Avant As Boolean
'##ModelId=5C8A76010383
Public Apres As Boolean
'##ModelId=5C8A76010393
Public Ecart As Boolean
'##ModelId=5C8A760103B2
Public DejaAmorti As Boolean
'##ModelId=5C8A760103C2
Public ResteAAmortir As Boolean

'##ModelId=5C8A760103E1
Private Sub Class_Initialize()
  DonneesBrutes = False
  Avant = True
  Apres = True
  Ecart = False
  DejaAmorti = False
  ResteAAmortir = False
End Sub
