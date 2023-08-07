VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecordsetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75DF0330"
Option Explicit

'##ModelId=5C8A75E000A0
Public autoMode As Boolean
'##ModelId=5C8A75E000D1
Public m_Logger As clsLogger

'********************************************************************
'
' Fonctions de calcul des provisions avec recordset déconnectés
'
' Appeler LoadRecordset(<NomTable>) avant de lancer les calculs
' pour optimiser les performances et ne pas casser les transactions
'
'********************************************************************

Private Type defRecordsetList
  tableName As String
  rs As ADODB.Recordset
  AgeMin As Integer
  AgeMax As Integer
End Type

'##ModelId=5C8A75E000DE
Private nbRecordset As Integer
'##ModelId=5C8A75E000F1
Private listRecordset() As defRecordsetList

'##ModelId=5C8A75E000FD
Public Sub InitRecordsetList()
  nbRecordset = 0
  Erase listRecordset
End Sub

'##ModelId=5C8A75E0011D
Public Function AgeMini(tableName As String) As Integer
  Dim i As Integer
  
  ' par securite, enleve 'P3IUser.' du nom de la table
  tableName = Replace(tableName, "P3IUser.", "")
  
  ' recherche si il existe
  For i = 0 To nbRecordset - 1
    If listRecordset(i).tableName = tableName Then
      AgeMini = listRecordset(i).AgeMin
      Exit Function
    End If
  Next
  
  AgeMini = -1
End Function

'##ModelId=5C8A75E0013C
Public Function AgeMaxi(tableName As String) As Integer
  Dim i As Integer
  
  ' par securite, enleve 'P3IUser.' du nom de la table
  tableName = Replace(tableName, "P3IUser.", "")
  
  ' recherche si il existe
  For i = 0 To nbRecordset - 1
    If listRecordset(i).tableName = tableName Then
      AgeMaxi = listRecordset(i).AgeMax
      Exit Function
    End If
  Next
  
  AgeMaxi = -1
End Function


'##ModelId=5C8A75E0015B
Public Function LoadRecordset(tableName As String, Optional idTable As Variant) As ADODB.Recordset
  On Error GoTo err_LoadRecordset
  
  Dim i As Integer, work_tableName As String, typeTable As Integer
  
  ' par securite, enleve 'P3IUser.' du nom de la table
  tableName = Replace(tableName, "P3IUser.", "")
  
  ' recherche si il existe
  For i = 0 To nbRecordset - 1
    If listRecordset(i).tableName = tableName Then
      Set LoadRecordset = listRecordset(i).rs
      Exit Function
    End If
  Next
  
  ' verifie que ce n'est pas une table stockée dans TableMortalite
  If tableName <> "TableMortalite" Then
    typeTable = m_dataHelper.GetParameter("SELECT TYPETABLE FROM ListeTableLoi WHERE NOMTABLE='" & tableName & "'")
    
    Select Case typeTable
      Case cdTypeTableMortalite, cdTypeTableGeneration
        work_tableName = "TableMortalite"
      Case Else
        work_tableName = tableName
    End Select
  Else
    typeTable = cdTypeTableMortalite
    work_tableName = tableName
  End If
  
  ' ouvre le recordset
  ReDim Preserve listRecordset(nbRecordset + 1)
  listRecordset(nbRecordset).tableName = tableName
  
  ' age min et max pour les lois et tables
  Select Case typeTable
    Case cdTypeTable_LoiMaintienIncapacite, cdTypeTable_LoiPassage, cdTypeTable_LoiMaintienInvalidite
      listRecordset(nbRecordset).AgeMin = m_dataHelper.GetParameter("SELECT MIN(Age) FROM " & tableName)
      listRecordset(nbRecordset).AgeMax = m_dataHelper.GetParameter("SELECT MAX(Age) FROM " & tableName)
    
    Case cdTypeTableCoeffBCACInval, cdTypeTableCoeffBCACIncap, cdTypeTable_MortaliteIncap, cdTypeTable_MortaliteInval
      If IsMissing(idTable) Then
        listRecordset(nbRecordset).AgeMin = m_dataHelper.GetParameter("SELECT MIN(Age) FROM " & tableName)
        listRecordset(nbRecordset).AgeMax = m_dataHelper.GetParameter("SELECT MAX(Age) FROM " & tableName)
      Else
        listRecordset(nbRecordset).AgeMin = m_dataHelper.GetParameter("SELECT MIN(Age) FROM " & tableName & " WHERE CleTable=" & idTable)
        listRecordset(nbRecordset).AgeMax = m_dataHelper.GetParameter("SELECT MAX(Age) FROM " & tableName & " WHERE CleTable=" & idTable)
      End If

    Case cdTypeTableMortalite, cdTypeTableGeneration
      listRecordset(nbRecordset).AgeMin = -1
      listRecordset(nbRecordset).AgeMax = -1
      
    Case Else
      listRecordset(nbRecordset).AgeMin = -1
      listRecordset(nbRecordset).AgeMax = -1
  End Select
  
  ' ouvre le recordset
  If tableName <> "TableMortalite" Then
    Set listRecordset(nbRecordset).rs = m_dataSource.OpenRecordset("SELECT * FROM " & work_tableName, Disconnected)
  Else
    Set listRecordset(nbRecordset).rs = m_dataSource.OpenRecordset("SELECT TM.*, LTL.TYPETABLE FROM " & work_tableName & " AS TM INNER JOIN ListeTableLoi AS LTL ON TM.NomTable = LTL.NOMTABLE ", Disconnected)
  End If
  
  Set LoadRecordset = listRecordset(nbRecordset).rs
  
  nbRecordset = nbRecordset + 1
  
  Exit Function
  
err_LoadRecordset:
  
  If Not autoMode Then
    MsgBox "Erreur " & Err & " dans LoadRecorset" & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " dans LoadRecorset" & vbLf & Err.Description
  End If
  
  Resume Next
End Function


'##ModelId=5C8A75E0019A
Public Sub CloseLoadedRecordset()
  Dim i As Integer
  
  ' recherche si il existe
  For i = 0 To nbRecordset - 1
    If Not listRecordset(i).rs Is Nothing Then
      listRecordset(i).rs.Close
      Set listRecordset(i).rs = Nothing
    End If
  Next
End Sub

'##ModelId=5C8A75E001A9
Private Sub Class_Initialize()
  InitRecordsetList
End Sub

'##ModelId=5C8A75E001C9
Private Sub Class_Terminate()
  CloseLoadedRecordset
End Sub
