VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListeJeuxDonnees 
   Caption         =   "Lots de données disponibles"
   ClientHeight    =   7995
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   13170
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox lblFillTime 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "----"
      Top             =   0
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CommandButton btnImportSASP3I 
      Caption         =   "&Importer depuis SASP3I"
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   5085
      Width           =   1935
   End
   Begin VB.CommandButton btnUtiliser 
      Caption         =   "&Utiliser dans P3I"
      Height          =   375
      Left            =   7065
      TabIndex        =   3
      Top             =   5085
      Width           =   1935
   End
   Begin VB.CommandButton btnExporter 
      Caption         =   "&Exporter"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   5085
      Width           =   1935
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Editer"
      Height          =   375
      Left            =   4050
      TabIndex        =   2
      Top             =   5085
      Width           =   1935
   End
   Begin VB.CommandButton btnImportXLS 
      Caption         =   "Importer depuis E&xcel"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   5085
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   9045
      TabIndex        =   0
      Top             =   5085
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11250
      Top             =   3870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   5400
      Top             =   5130
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM P3IUser.P3IUser.Assure WHERE RECNO=-1"
      Caption         =   "dtaPeriode"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin FPSpreadADO.fpSpread sprListe 
      Height          =   4965
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11040
      _Version        =   524288
      _ExtentX        =   19473
      _ExtentY        =   8758
      _StockProps     =   64
      BackColorStyle  =   1
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   1000000
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmListeJeuDonnees.frx":0000
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   10
      VisibleRows     =   100
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "frmListeJeuxDonnees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A68110341"

Option Explicit

'##ModelId=5C8A68120073
Public ImportFromExcel As Boolean
'##ModelId=5C8A68120092
Public ExcelFilePath As String
'##ModelId=5C8A681200A1
Public LotNumber As Integer

'##ModelId=5C8A681200C1
Private modeAuto As Boolean

'##ModelId=5C8A681200E0
Private frmNumPeriode As Long
'##ModelId=5C8A681200FF
Private frmNumeroLot As Long

'##ModelId=5C8A6812011F
Property Let SetModeAuto(val As Boolean)
    modeAuto = val
End Property

'##ModelId=5C8A6812013E
Private Sub SetNumeroLot()
  frmNumeroLot = 0
  
  If sprListe.ActiveRow < 0 Then Exit Sub
  
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 1
  
  frmNumeroLot = CLng(sprListe.text)
End Sub

'##ModelId=5C8A6812014D
Private Sub btnClose_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A6812016D
Private Sub btnEdit_Click()
  
  Dim frm As frmJeuxDonnees
  
  SetNumeroLot
  
  If frmNumeroLot = 0 Then Exit Sub
  
  Set frm = New frmJeuxDonnees
  
  frm.NumeroLot = frmNumeroLot
  frm.numPeriode = frmNumPeriode
  
  frm.Show vbModal
  
  Dim r As Long, tr As Long
  
  Screen.MousePointer = vbHourglass
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
  
  RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  sprListe.TopRow = tr
  sprListe.SetActiveCell 2, r
  
  sprListe.Row = r
  sprListe.SelModeSelected = True
  
  sprListe.SetFocus
  
  Screen.MousePointer = vbDefault
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' export du jeux de données sélectioné
'
'##ModelId=5C8A6812017C
Private Sub btnExporter_Click()
  
  SetNumeroLot
  
  If frmNumeroLot = 0 Then Exit Sub
  
  
  '
  ' Selection du type de provision
  '
  Dim frm As New frmTypeExport, sTypeProsision As String
  Dim delExist As Boolean
  
  ret_code = 0
  frm.Show vbModal
  If ret_code = 0 Then
    
    ' nom standardise Generali
    Select Case frm.frmTypeProvision
      Case 1
        sTypeProsision = "BILAN"
      Case 2
        sTypeProsision = "CLIENT"
      Case 3
        sTypeProsision = "SIMUL"
    End Select
    
    Select Case frm.frmDelExistant
      Case True
        delExist = True
      Case False
        delExist = False
    End Select
  
  End If
  
  'Unload frm
  Set frm = Nothing
  
  If delExist Then
    If MsgBox("Les données vont être exportées et vont écrasées celles qui existent déjà dans les tables TTLOGTRAIT et TTPROVCOLL..." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) <> vbYes Then
      Exit Sub
    End If
  End If
  
  On Error GoTo err_btnUtiliser
  
  Screen.MousePointer = vbHourglass
  
  ' cree la boite d'attente
  Dim m_fWait As New frmWait, fErreur As Boolean
    
  Load m_fWait
    
  m_fWait.Caption = "Export des données..."
  
  m_fWait.ProgressBar1.Min = 0
  m_fWait.ProgressBar1.Value = 0
  m_fWait.ProgressBar1.Max = 5
  
  If delExist = True Then
    m_fWait.Label1(0).Caption = "Suppression des données précédentes..."
    
    m_dataSource.Execute "DELETE FROM TTPROVCOLL" ' WHERE NUTRAITP3I=" & frmNumeroLot
  '  m_dataSource.Execute "TRUNCATE TABLE TTPROVCOLL"
    m_fWait.ProgressBar1.Value = 1
    m_fWait.Refresh
    DoEvents
    
    m_dataSource.Execute "DELETE FROM TTLOGTRAIT" ' WHERE NUTRAITP3I=" & frmNumeroLot
  '  m_dataSource.Execute "TRUNCATE TABLE TTLOGTRAIT"
    m_fWait.ProgressBar1.Value = 2
    m_fWait.Refresh
    DoEvents
  
  End If
  
  m_fWait.Label1(0).Caption = "Export en cours..."
  
  '
  ' on recopie les données d'origine (test pour livrer une boite noire au 1/9/2008)
  '
  CopyLot "P3ILOGTRAIT", "TTLOGTRAIT", "", ""
  m_fWait.ProgressBar1.Value = 3
  m_fWait.Refresh
  DoEvents
  
  CopyLot "P3IPROVCOLL", "TTPROVCOLL", " AND DataVersion=0", sTypeProsision
  m_fWait.ProgressBar1.Value = 4
  m_fWait.Refresh
  DoEvents
  
  m_dataSource.Execute "UPDATE TTLOGTRAIT SET NBLIGTRAIT=(SELECT count(*) FROM TTPROVCOLL), MTTRAIT=0"
  m_fWait.ProgressBar1.Value = 5
  m_fWait.Refresh
  DoEvents
  
  ' Creation du fichier top de signalisation
  Dim sFichierTop As String
  
  sFichierTop = GetSettingIni(SectionName, "Dir", "FichierTop", "#")
  
  If sFichierTop = "#" Then
    
    ' erreur de parametrage
    MsgBox "Vous devez spécifier le chemin complet du fichier 'top' dans le fichier de parametre " & vbLf & sFichierIni & vbLf & "Section [DB], Entrée FichierTop ", vbCritical
  
  Else
        
    ' signal la presence des données
    
    Dim FileNumber As Integer
    
    FileNumber = FreeFile   ' Get unused file
    
    fErreur = False
    Open sFichierTop For Output As #FileNumber   ' Create file name.
    Print #FileNumber, "OK " & Now   ' Output text.
    Close #FileNumber         ' Close file.
    
    If fErreur = True Then
      MsgBox "Impossible de créer le fichier de signalisation !", vbCritical
    End If
    
  End If
  
  ' ferme la boite d'attente
  Unload m_fWait
  
  Set m_fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
err_btnUtiliser:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  fErreur = True
  Resume Next
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie une date sous forme de date en long
' format YYYYMMDD
'
'##ModelId=5C8A6812018C
Private Function ConvertDate(dDate As Date) As Long
  
  If Not IsNull(dDate) Then
    ConvertDate = Year(dDate) * 10000 + Month(dDate) * 100 + Day(dDate)
  Else
    ConvertDate = 0
  End If
 
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie un time sous forme de date en long
' format HHMMSS
'
'##ModelId=5C8A681201BB
Private Function ConvertTime(tTime As Date) As Long
  
  If Not IsNull(tTime) Then
    ConvertTime = Hour(tTime) * 10000 + Minute(tTime) * 100 + Second(tTime)
  Else
    ConvertTime = 0
  End If
 
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie une ligne de sTable vers sOut
'
'##ModelId=5C8A681201EA
Private Sub CopyLot(sTable As String, sOut As String, sWhereIn As String, sTypeProsision As String)
  
  If frmNumeroLot = 0 Then Exit Sub
  
  On Error GoTo err_CopyLot
  
  Dim rsIn As ADODB.Recordset, rsOut As ADODB.Recordset, i As Integer, f As ADODB.field
  
  Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot & sWhereIn, Snapshot)
'  Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM " & sOut & " WHERE NUTRAITP3I=" & frmNumeroLot, Dynamic)
  Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM " & sOut, Dynamic)
    
  Do Until rsIn.EOF
    rsOut.AddNew
    
    For i = 0 To rsIn.fields.Count - 1
      Set f = rsIn.fields(i)
      
      ' conversion des champs de type Date ou Heure
'      If Left(f.Name, 2) = "DT" Then
'        If Not IsNull(f.Value) Then
'          rsOut.fields(f.Name).Value = ConvertDate(f.Value)
'        Else
'          rsOut.fields(f.Name).Value = 0
'        End If
'
'      ElseIf Left(f.Name, 2) = "HH" Then
'        If Not IsNull(f.Value) Then
'          rsOut.fields(f.Name).Value = ConvertTime(f.Value)
'        Else
'          rsOut.fields(f.Name).Value = 0
'        End If
'
'      ElseIf (UCase(f.Name) <> "COMMENTAIRE") And (UCase(f.Name) <> "DATAVERSION") Then
      If (UCase(f.Name) <> "COMMENTAIRE") And (UCase(f.Name) <> "DATAVERSION") And FieldExistsInRS(rsOut, f.Name) Then
        
        If Not IsNull(f.Value) Then
          rsOut.fields(f.Name).Value = f.Value
        Else
          If rsOut.fields(f.Name).Type = adChar Or rsOut.fields(f.Name).Type = adVarChar Then
            rsOut.fields(f.Name).Value = " "
          Else
            rsOut.fields(f.Name).Value = 0
          End If
        End If
      
      End If
    
    Next
    
    
    If sOut = "TTPROVCOLL" Then
      rsOut.fields("TYPEPROVISION").Value = sTypeProsision
    End If
  
    For i = 0 To rsOut.fields.Count - 1
      Set f = rsOut.fields(i)
          
      If IsNull(f.Value) Or IsEmpty(f.Value) Then
        If f.Type = adChar Or f.Type = adVarChar Then
          f.Value = " "
        Else
          f.Value = 0
        End If
      End If
          
    Next
    
    rsOut.Update
    
    rsIn.MoveNext
  Loop
  
  rsIn.Close
  rsOut.Close
  
  Exit Sub
  
err_CopyLot:
  'MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A68120257
Private Sub btnImportSASP3I_Click()
  On Error Resume Next
  
  Dim frm As New frmListeLotSASP3I
  
  frm.Show vbModal
  
  RefreshListe
End Sub

'##ModelId=5C8A68120267
Private Sub btnImportXLS_Click()
  Unload Me
  
  Dim objImport As iP3IGeneraliImport
  Dim CleGroupe As Long
  Dim codeRetour As Boolean
  
  CleGroupe = GroupeCle ' en dur
    
  ' efface les assurés de la periode en cours
  Dim rq As String
  Dim rs As ADODB.Recordset
  
  ' base d'import
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "InputPath", App.Path)
    
  ' charge l'object d'import
  Dim txtObjetImport As String, sectionObjectImport As String, type_periode As Integer
  
  type_periode = m_dataHelper.GetParameterAsLong("SELECT PETYPEPERIODE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode)
  
  If type_periode = eRevalo Or type_periode = eProvisionRetraiteRevalo Then
    sectionObjectImport = "ObjetImportRevalo"
  Else
    sectionObjectImport = "ObjetImport"
  End If
  
  txtObjetImport = GetSettingIni(CompanyName, SectionName, sectionObjectImport, "#")
  
  If txtObjetImport = "#" Then
    MsgBox "La section " & sectionObjectImport & " n'est pas présente," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
    Exit Sub
  End If
  
  On Error GoTo errImport
  Set objImport = CreateObject(txtObjetImport)
  
'  rq = "SELECT PEDATEDEB, PEDATEFIN, PENBJOURMAX, PENBJOURDC, PEAGERETRAITE, PEDATEEXT " _
'       & " FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode

  rq = "SELECT PEDATEDEB, PEDATEFIN, PENBJOURMAX, PENBJOURDC, 65 as PEAGERETRAITE, PEDATEEXT " _
      & " FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode

  
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF() Then
  
    ' lance l'import
    frmMain.Enabled = False
    
    DoEvents
  ' recherche du mode P3I individuel
  Dim m_bP3I_Individuel As Boolean
  'm_bP3I_Individuel = IIf(GetSettingIni(CompanyName, SectionName, "P3I_Individuel", "0") <> "0", True, False)
 
   m_bP3I_Individuel = CBool(m_dataHelper.GetParameterAsDouble("SELECT PEP3I_INDIVIDUEL FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle))

    
    codeRetour = objImport.DoImport(CommonDialog1, m_dataSource, CleGroupe, frmNumPeriode, _
                 Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                 rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, m_bP3I_Individuel)
    
    frmMain.Enabled = True
  End If
  rs.Close
  
  Set objImport = Nothing
  
  ' rempli la liste avec les articles importés
  RefreshListe
  
  Screen.MousePointer = vbHourglass
 
  If codeRetour = True Then
    ' lancement des calculs
    'Call CalculProvisionsAssures   ' appel de la fonction de calcul des provisions pour les assurés
    
    ' mets à jour la date d'extraction
    ' NE PAS FAIRE CA CETTE DATE SERT A STOCKER LA DATE DE CLOTURE
    'rq = "UPDATE Periode SET PEDATEEXT = #" & Format(Now(), "mm/dd/yyyy") & "# WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode
    'theDB.Execute rq
  Else
    MsgBox "L'opération d'import a été INTERROMPUE !" & vbLf & "Aucun article n'a été ajouté à la période n°." & frmNumPeriode, vbExclamation
  End If
    
  ' rafraichit la liste
  RefreshListe
  
  Screen.MousePointer = vbDefault
  
  Unload Me
  
  Exit Sub
  
errImport:
  frmMain.Enabled = True
  MsgBox "Erreur : " & Err.Description & vbLf & "Objet = " & txtObjetImport, vbCritical
  Resume Next
End Sub

'##ModelId=5C8A68120276
Private Sub btnUtiliser_Click()
  SetNumeroLot
  
  If frmNumeroLot = 0 Then Exit Sub
  
  Unload Me
  
  Dim objImport As iP3IGeneraliImport
  Dim CleGroupe As Long
  Dim codeRetour As Boolean
  
  CleGroupe = GroupeCle ' en dur
    
  ' efface les assurés de la periode en cours
  Dim rq As String
  Dim rs As ADODB.Recordset
  
  ' base d'import
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "InputPath", App.Path)
    
  ' charge l'object d'import
  Dim txtObjetImport As String
  
  txtObjetImport = GetSettingIni(CompanyName, SectionName, "ObjetImportSASP3I", "#")
  
  If txtObjetImport = "#" Then
    MsgBox "La section ObjetImport n'est pas présente," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
    Exit Sub
  End If
  
  On Error GoTo errImport
  Set objImport = CreateObject(txtObjetImport)
  
  rq = "SELECT PEDATEDEB, PEDATEFIN, PENBJOURMAX, PENBJOURDC, 65 as PEAGERETRAITE, PEDATEEXT, PETYPEPERIODE " _
       & " FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode
  
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF() Then
  
    ' lance l'import
    frmMain.Enabled = False
    
    DoEvents
    
    If IsNull(objImport) Then
      rq = ""
    End If
    
    'NEW Import STATUTAIRE
    
    Dim PeriodeType As Integer
    PeriodeType = rs.fields("PETYPEPERIODE")
    
    Dim frm As New frmStatImport
    If PeriodeType = 6 Then
      frm.PeriodeType = cPeriodeStat
    Else
      frm.PeriodeType = cPeriodeStandard
    End If
    
    frm.Show vbModal
    
    If frm.Success = True Then
      If frm.ImportType = cImportStandard Then
      
        'Import of type Standard
        
        codeRetour = objImport.DoImportSASP3I(frmNumeroLot, m_logPath, m_dataSource, CleGroupe, frmNumPeriode, _
                 Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                 rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, False)
      
      Else
      
        'Import of type STATUTAIRE
        'PathSexFileExcel & SexAllMale , NumPeriodeNonStat && NumPeriodeStat -- CategoryCodeSTAT
        
        SetCategoryCodeStatVariable
        
'        'get the STAT category code
'        Dim rsCat As New ADODB.Recordset
'        Dim cnt As Integer
'
'        cnt = 1
'        CategoryCodeSTAT = ""
'
'        'Set rsCat = m_dataSource.OpenRecordset("Select Categorie From Statutaire_Categorie Where Description = 'CodeStatutaire'", Snapshot)
'        Set rsCat = m_dataSource.OpenRecordset("Select Categorie From Statutaire_Categorie", Snapshot)
'        If Not rsCat.EOF() Then
'          Do Until rsCat.EOF
'            If cnt = 1 Then
'              CategoryCodeSTAT = rsCat.fields("Categorie")
'            Else
'              CategoryCodeSTAT = CategoryCodeSTAT & "," & rsCat.fields("Categorie")
'            End If
'            rsCat.MoveNext
'          Loop
'
'          rsCat.Close
'        End If
        
        If CategoryCodeSTAT <> "" Then
        
        'Set all variables that are required for STAT treatment
        objImport.SetStatutaireVariables NumPeriodeStat, NumPeriodeNonStat, PathSexFileExcel, CategoryCodeSTAT, SexAllMale, TwoLotImport
        
        codeRetour = objImport.DoImportSASP3I(frmNumeroLot, m_logPath, m_dataSource, CleGroupe, frmNumPeriode, _
                 Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                 rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, False)
        Else
          '### msgbox : no catCode
          MsgBox "Le code catégorie pour les assurées du type Statutaire n'est pas renseigné", vbOKOnly, "Code Category Manquant"
          codeRetour = False
        End If
      
      End If
      
    Else
      'we did not launch the import
      codeRetour = False
      
    End If
        
    frmMain.Enabled = True
  End If
  
  rs.Close
  
  Set objImport = Nothing
  
  ' rempli la liste avec les articles importés
  RefreshListe
  
  Screen.MousePointer = vbHourglass
 
  If codeRetour = True Then
    ' lancement des calculs
    'Call CalculProvisionsAssures   ' appel de la fonction de calcul des provisions pour les assurés
    
    ' mets à jour la date d'extraction
    ' NE PAS FAIRE CA CETTE DATE SERT A STOCKER LA DATE DE CLOTURE
    'rq = "UPDATE Periode SET PEDATEEXT = #" & Format(Now(), "mm/dd/yyyy") & "# WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode
    'theDB.Execute rq
  Else
    MsgBox "L'opération d'import a été INTERROMPUE !" & vbLf & "Aucun article n'a été ajouté à la période n°." & frmNumPeriode, vbExclamation
  End If
    
  ' rafraichit la liste
  'RefreshListe
  
  Screen.MousePointer = vbDefault
  
  Unload Me
  
  Exit Sub
  
errImport:
  frmMain.Enabled = True
  MsgBox "Erreur : " & Err.Description & vbLf & "Objet = " & txtObjetImport, vbCritical
  Resume Next
End Sub

'##ModelId=5C8A68120295
Private Sub Form_Load()

  frmNumPeriode = numPeriode
  
  If modeAuto = True Then
    btnImportSASP3I.Enabled = False
    btnEdit.Enabled = False
  
  End If
  
  m_dataSource.SetDatabase dtaPeriode
  
  ' chargement du masque du spread
  sprListe.LoadFromFile App.Path & "\ListeJeuxDonnees.ss7"
  
  RefreshListe
End Sub

'##ModelId=5C8A681202A5
Private Sub RefreshListe()
  
  Dim rq As String, rs As ADODB.Recordset
  Dim filter As String
  Dim i As Integer
  
  Dim debut As Date, fin As Date
  
  debut = Now
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre
  Me.Caption = "Jeux de données disponibles"
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = False
  'sprListe.VirtualMode = True
  'sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  'sprListe.VScrollSpecial = True
  'sprListe.VScrollSpecialType = 0
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT NUTRAITP3I, DTTRAIT as [Date], CONVERT(CHAR(8), HHTRAIT, 8) as [Heure], NUTRAITP3I as [Identifiant du lot], RTRIM(NUTRAIT) as [N° traitement AGI], " _
       & " DTDEBPER as [Debut], DTFINPER as [Fin], NBLIGTRAIT as [nb lignes], MTTRAIT as [Montant], NUTRAIT, IDTABLESAS, Commentaire" _
       & " FROM P3ILOGTRAIT ORDER BY NUTRAITP3I DESC"
          
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
    
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If dtaPeriode.Recordset.EOF = False Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
  
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
    'sprListe.VirtualMaxRows = dtaPeriode.Recordset.RecordCount
  
    dtaPeriode.Recordset.MoveFirst
  Else
    sprListe.MaxRows = 0
    'sprListe.VirtualMaxRows = 0
    sprListe.ColWidth(1) = 0
    sprListe.Visible = True
    sprListe.ReDraw = True

    Screen.MousePointer = vbDefault
    
    GoTo pas_de_donnee
  End If
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
     
  For i = 2 To sprListe.MaxCols
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 2
  Next i
 
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = sprListe.MaxCols - 1
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = sprListe.MaxCols
  sprListe.Col2 = sprListe.MaxCols
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
pas_de_donnee:

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True
  
  Screen.MousePointer = vbDefault

  fin = Now
  
  lblFillTime.text = "Remplissage : " & DateDiff("s", debut, fin) & " s"

  Exit Sub

err_RefreshListe:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A681202B5
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 130
 
  topbtn = Me.ScaleHeight - btnHeight
  
  sprListe.Height = Maximum(topbtn - 100, 0)
  
  PlacePremierBoutton btnImportXLS, topbtn
  
  PlaceBoutton btnImportSASP3I, btnImportXLS, topbtn
  
  PlaceBoutton btnEdit, btnImportSASP3I, topbtn
  
  PlaceBoutton btnUtiliser, btnEdit, topbtn
  
  PlaceBoutton btnExporter, btnUtiliser, topbtn
  
  PlaceBoutton btnClose, btnExporter, topbtn

End Sub

'##ModelId=5C8A681202D4
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A68120312
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "Commentaire" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  End If
End Sub

