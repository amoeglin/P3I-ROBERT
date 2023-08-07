VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'export des données d'une période
' vers une base MS Access
'
'FIXIT: Declare 'CommonDialog1' with an early-bound data type                              FixIT90210ae-R1672-R1B8ZE
Public Function ExportDBAccess(CommonDialog1 As Object, SrcDB As DataAccess, CleGroupe As Integer, NumPeriode As Long) As Boolean
  Dim DestDB As DAO.Database
  Dim rs As ADODB.Recordset
  Dim fWait As frmWait

  On Error GoTo GestionErreur
  
  ' création la base
  If CommonDialog1.FileName = "" Or CommonDialog1.FileName = "*.mdb" Then
    ExportDBAccess = False
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
  
  Set fWait = New frmWait
  
  fWait.Caption = "Export en cours..."
  
  fWait.ProgressBar1.Min = 0
  fWait.ProgressBar1.Value = 0
  fWait.ProgressBar1.Max = 37
  
  Screen.MousePointer = vbHourglass
 
  fWait.Show vbModeless
  fWait.Refresh
  
  ' création de la base
  If Dir(CommonDialog1.FileName) = "" Then
    FileCopy App.Path + "\P3I_Export_DB_Vide.mdb", CommonDialog1.FileName
  Else
    ' demande confirmation de la destruction des anciennes données
    If MsgBox("ATTENTION: le données vont écrasées celles existants dans " & CommonDialog1.FileName & " pour la période " & NumPeriode & "." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) = vbNo Then
      ExportDBAccess = False
      Exit Function
    End If
  End If

  '
  ' Export des tables
  '
  Set DestDB = DBEngine.OpenDatabase(CommonDialog1.FileName, , False)
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "AgeDepartRetraite", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "AgeSituationFamille", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Assure", " WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Assure_Retraite", " WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Capitaux_Moyens", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Categorie_GA_2006", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CATR9", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CATR9INVAL", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CDSITUAT", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CODECATINV", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CodePosition", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CodeProvision", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CodesCat", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Correspondance_CatOption", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "CorrespondanceGarantie", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Fractionnement", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Garantie", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "GarantieDC", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "IndemnisationIncapInval", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "LissageProvision", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "ListeTableLoi", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "ParamCalcul", " WHERE PEGPECLE=" & CleGroupe & " AND PENUMCLE=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "ParamRentes", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "PassageNCA", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Periode", " WHERE PEGPECLE=" & CleGroupe & " AND PENUMCLE=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "PlafondSS", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "PM_Retenue", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "ProvisionsOuverture", " WHERE GPECLE=" & CleGroupe & " AND NUMCLE=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "Reassurance", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "RisqueDeces", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "SituationFamille", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TBQREGA", " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeCalcul", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeDeTable", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeFractionnement", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeGarantie", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeLigne", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypePeriode", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeSexe", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  ExportDBAccess = ExportDBAccess And ExportTable(SrcDB, DestDB, "TypeTermeEchu", "", fWait)
  fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
  fWait.Refresh
  
  DestDB.Close
  
  fWait.Hide
  Unload fWait
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
    
  Exit Function
  
GestionErreur:
  fWait.Hide
  Unload fWait
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
  
  ExportDBAccess = False
  
  Exit Function
  
  Resume Next
  
End Function

'
' Export une table
'
Private Function ExportTable(SrcDB As DataAccess, DestDB As DAO.Database, sTable As String, sWhere As String, fWait As frmWait) As Boolean
  Dim rs As ADODB.Recordset
  Dim rs2 As DAO.Recordset, f As ADODB.Field
  
  On Error GoTo GestionErreur
  
  Screen.MousePointer = vbHourglass
  
  fWait.Label1(0).Caption = "Export de la table " & sTable & "..."
  fWait.Label1(1).Caption = "Veuillez patienter ..."
  fWait.Refresh
  Sleep 300
  
  ' supprime les données de la période actuelle
  DestDB.Execute "DELETE * FROM " & sTable & sWhere, dbFailOnError
  
  ' exporte les données (RECNO est laissé à la charge de la base MS Access
  Set rs = SrcDB.OpenRecordset("SELECT * FROM P3IUser." & sTable & sWhere, Snapshot)
  Set rs2 = DestDB.OpenRecordset(sTable, dbOpenTable)

  Do Until rs.EOF
    rs2.AddNew
    
    For Each f In rs.Fields
      If f.Name <> "RECNO" Then
        rs2.Fields(f.Name) = f
      End If
    Next
    
    rs2.Update
  
    If (rs.AbsolutePosition Mod 100) = 0 Then
      fWait.Label1(1).Caption = "Veuillez patienter (" & rs.AbsolutePosition & " / " & rs.RecordCount & ") ..."
      fWait.Refresh
      DoEvents
    End If
  
    rs.MoveNext
  Loop
  
  rs.Close
  rs2.Close
  
  Screen.MousePointer = vbDefault
  
  ExportTable = True
  
  Exit Function
  
GestionErreur:
  fWait.Visible = False
  Screen.MousePointer = vbDefault
  
  MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
  
  Screen.MousePointer = vbHourglass
  fWait.Visible = True
  
  Resume Next
  
End Function
