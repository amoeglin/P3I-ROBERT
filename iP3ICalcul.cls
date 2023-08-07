VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iP3ICalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75DB0013"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' calcul des Provisions concerant les assurés selectionnés
'
'##ModelId=5C8A75DB0119
Public Sub CalculProvisionsAssures(bRevalo As Boolean, numPeriode As Long, NumGroupe As Long, Logger As clsLogger, Optional autoMode As Boolean = False, Optional Avant As Boolean = True, Optional Apres As Boolean = False)

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' vérifie la présence des paramètres pour permettre les calculs sans erreur
'
'##ModelId=5C8A75DB01A6
Public Function CheckParametresAssures(numPeriode As Long, NumGroupe As Long, Logger As clsLogger) As Boolean

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' controle des resultats
'
'##ModelId=5C8A75DB0265
Public Function CheckResultatCalcul(numPeriode As Long, NumGroupe As Long, Logger As clsLogger) As Boolean

End Function

