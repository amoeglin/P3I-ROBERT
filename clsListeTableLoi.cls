VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListeTableLoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A76020057"
Option Explicit

'##ModelId=5C8A76020170
Public m_TABLECLE As Variant
'##ModelId=5C8A76020180
Public m_LIBTABLE As Variant
'##ModelId=5C8A7602019F
Public m_NOMTABLE As Variant
'##ModelId=5C8A760201AF
Public m_TYPETABLE As Variant
'##ModelId=5C8A760201CE
Public m_TableUtilisateur As Variant

'##ModelId=5C8A760201ED
Public autoMode As Boolean

'
' Constructors
'


'##ModelId=5C8A7602020C
Private Sub Class_Initialize()
    
  m_TABLECLE = Null
  m_LIBTABLE = Null
  m_NOMTABLE = Null
  m_TYPETABLE = Null
  m_TableUtilisateur = Null

End Sub


'##ModelId=5C8A76020217
Public Sub Copy(ByRef src As clsListeTableLoi)

  m_TABLECLE = src.m_TABLECLE
  m_LIBTABLE = src.m_LIBTABLE
  m_NOMTABLE = src.m_NOMTABLE
  m_TYPETABLE = src.m_TYPETABLE
  m_TableUtilisateur = src.m_TableUtilisateur

End Sub



'---------------------------------------------------
'##ModelId=5C8A76020236
Public Function Load(DB As DataAccess, idTable As Long) As Boolean

   Dim theData As ADODB.Recordset
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Load

   Set theData = DB.OpenRecordset("SELECT * FROM ListeTableLoi WHERE TABLECLE=" & idTable, Snapshot)
   If theData.EOF = False Then

      If IsNull(theData.fields("TABLECLE")) Then
         m_TABLECLE = 0
      Else
         m_TABLECLE = theData.fields("TABLECLE")
      End If

      If IsNull(theData.fields("LIBTABLE")) Then
         m_LIBTABLE = 0
      Else
         m_LIBTABLE = theData.fields("LIBTABLE")
      End If

      If IsNull(theData.fields("NOMTABLE")) Then
         m_NOMTABLE = 0
      Else
         m_NOMTABLE = theData.fields("NOMTABLE")
      End If

      If IsNull(theData.fields("TYPETABLE")) Then
         m_TYPETABLE = 0
      Else
         m_TYPETABLE = theData.fields("TYPETABLE")
      End If

      If IsNull(theData.fields("TableUtilisateur")) Then
         m_TableUtilisateur = 0
      Else
         m_TableUtilisateur = theData.fields("TableUtilisateur")
      End If

   Else
      ret = False
   End If

   Load = ret

   Exit Function

err_Load:
   MsgBox "Erreur dans clsListeTableLoi::Load()", vbCritical
   Load = False

End Function


'---------------------------------------------------
'##ModelId=5C8A7602026C
Public Function Save(DB As DataAccess) As Boolean
  Dim sql As String, ret As Boolean
  ret = True
  
  On Error GoTo err_Save
  
  Dim theCommand As New ADODB.Command, prm As ADODB.Parameter
  theCommand.ActiveConnection = DB.Connection
  theCommand.CommandType = adCmdText
  theCommand.CommandText = sql
  
  If (IsNull(m_TABLECLE) Or m_TABLECLE = 0) Then
    
    m_TABLECLE = m_dataHelper.GetParameterAsLong("SELECT MAX(IsNull(TABLECLE,0))+1 FROM ListeTableLoi")
  
    sql = "INSERT INTO ListeTableLoi(TABLECLE, LIBTABLE, NOMTABLE, TYPETABLE, TableUtilisateur) " _
          & "VALUES(?, ?, ?, ?, ?)"
  
    Set prm = theCommand.CreateParameter("@TABLECLE", adInteger, adParamInput, 0, m_TABLECLE)
    theCommand.Parameters.Append prm
    
  Else
    
    sql = "UPDATE ListeTableLoi SET LIBTABLE=?, NOMTABLE=?, TYPETABLE=?, TableUtilisateur=? WHERE TABLECLE=" & m_TABLECLE
    
  End If
  
  theCommand.CommandText = sql
  
  Set prm = theCommand.CreateParameter("@LIBTABLE", adVarChar, adParamInput, IIf(IsNull(m_LIBTABLE), 1, Len(m_LIBTABLE)), m_LIBTABLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@NOMTABLE", adVarChar, adParamInput, IIf(IsNull(m_NOMTABLE), 1, Len(m_NOMTABLE)), m_NOMTABLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@TYPETABLE", adInteger, adParamInput, 0, m_TYPETABLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@TableUtilisateur", adBoolean, adParamInput, 0, m_TableUtilisateur)
  theCommand.Parameters.Append prm
  
  theCommand.Execute
  
  Save = True
  
  Exit Function
  
err_Save:
  MsgBox "Erreur dans clsListeTableLoi::Save() : " & Err & vbLf & Err.Description, vbCritical
  Save = False
  Exit Function
  
  Resume Next
End Function


'##ModelId=5C8A7602028C
Private Function TestUtilisation(DB As DataAccess, ByVal sListFields As String, bIsCletable As Boolean, ByVal Value As String) As Boolean
  
  Dim astrSplitItems() As String
  Dim szSQL As String, i As Integer
  
  On Error GoTo err_TestUtilisation
  
  ' Liste des champs rejetés
  astrSplitItems = Split(UCase(sListFields), ",")
  For i = 0 To UBound(astrSplitItems)
    If bIsCletable Then
      szSQL = "SELECT count(*) FROM P3IUser.ParamCalcul WHERE " & astrSplitItems(i) & "=" & Value
    Else
      szSQL = "SELECT count(*) FROM P3IUser.ParamCalcul WHERE " & astrSplitItems(i) & "='" & Value & "'"
    End If
    
    If DB.CreateHelper.GetParameterAsLong(szSQL) <> 0 Then
      TestUtilisation = True
      Exit Function
    End If
  Next
  
  TestUtilisation = False
  
  Exit Function
  
err_TestUtilisation:
  TestUtilisation = False
  
  Exit Function
  
  Resume Next
End Function


'---------------------------------------------------
'##ModelId=5C8A76020303
Public Function Delete(DB As DataAccess, bDeleteDataFromTable As Boolean) As Boolean
  Dim ret As Boolean, bInTrans As Boolean
  ret = True
  
  bInTrans = False
  
  On Error GoTo err_Delete
  
  ' Vérification de l'utilisation de la table dans une période
  If TestUtilisation(DB, "PELMINCAP,PELMPASSAGE,PELMINVAL,PETRENTECONJOINT,PETRENTEEDUC,PELMINCAP_Retraite,PELMPASSAGE_Retraite,PELMINVAL_Retraite", False, m_NOMTABLE) _
     Or TestUtilisation(DB, "PECleTableMortBCACInval_MDC,PECleTableMortBCACIncap_MDC,PECleTableBCACInval_MDC," _
                          & "PECleTableBCACIncap_MDC,PECleTableMortBCACInval_MDC_Retraite," _
                          & "PECleTableMortBCACIncap_MDC_Retraite,PECleTableBCACInval_MDC_Retraite," _
                          & "PECleTableBCACIncap_MDC_Retraite", True, CStr(m_TABLECLE)) Then
    MsgBox "Impossible de supprimer la table " & m_NOMTABLE & " : elle est en cours d'utilisation !", vbCritical
    Delete = False
    Exit Function
  End If
  
  ' Suppression des données dans la table
  DB.BeginTrans
  
  bInTrans = True
  
  Dim szSQL As String
  
  Select Case m_TYPETABLE
    Case cdTypeTableMortalite, cdTypeTableGeneration
      szSQL = "DELETE FROM TableMortalite WHERE NomTable='" & m_NOMTABLE & "'"
    
    Case cdTypeTable_LoiMaintienIncapacite, cdTypeTable_LoiMaintienInvalidite, cdTypeTable_LoiPassage, cdTypeTable_BaremeAnneeStatutaire
      szSQL = "DROP TABLE " & m_NOMTABLE
  
    Case cdTypeTable_MortaliteIncap
      szSQL = "DELETE FROM MortIncap WHERE CleTable=" & m_TABLECLE
  
    Case cdTypeTable_MortaliteInval
      szSQL = "DELETE FROM MortInval WHERE CleTable=" & m_TABLECLE
      
    Case Else
      MsgBox "Impossible de supprimer la table " & m_NOMTABLE & " : type de table non supporté !", vbCritical
  End Select
  
  
  ' Suppression des données
  Dim theCommand As ADODB.Command
  
  Set theCommand = New ADODB.Command
  
  theCommand.ActiveConnection = DB.Connection
  theCommand.CommandType = adCmdText
  theCommand.CommandText = szSQL
  
  theCommand.Execute
  
  
  ' Suppression de la ligne dans ListeTableLoi
  theCommand.ActiveConnection = DB.Connection
  theCommand.CommandType = adCmdText
  theCommand.CommandText = "DELETE FROM ListeTableLoi WHERE TABLECLE= " & m_TABLECLE
  
  theCommand.Execute
  
  
  DB.CommitTrans
  bInTrans = False
  
  m_TABLECLE = 0
  
  Delete = ret
  
  Exit Function
  
err_Delete:
  If bInTrans Then DB.RollbackTrans
  bInTrans = False

  MsgBox "Erreur dans clsListeTableLoi::Delete()" & vbLf & Err.Description, vbCritical
  Delete = False
  
  Exit Function
  
  Resume Next
End Function




