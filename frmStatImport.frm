VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStatImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Standard ou Statutaire"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15315
   Icon            =   "frmStatImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   15315
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   120
      Top             =   0
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
      RecordSource    =   ""
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGetSex 
      Caption         =   "Sélectionner"
      Height          =   375
      Left            =   12360
      TabIndex        =   8
      Top             =   2000
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   13080
      TabIndex        =   5
      Top             =   8880
      Width           =   1935
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Importer"
      Height          =   375
      Left            =   10920
      TabIndex        =   4
      Top             =   8880
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paramètres Import Statutaire"
      Height          =   7215
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   14775
      Begin VB.TextBox txtSexe 
         Height          =   285
         Left            =   3240
         TabIndex        =   7
         Top             =   600
         Width           =   8535
      End
      Begin FPSpreadADO.fpSpread sprListe 
         Height          =   4845
         Left            =   480
         TabIndex        =   9
         Top             =   2040
         Width           =   13845
         _Version        =   524288
         _ExtentX        =   24421
         _ExtentY        =   8546
         _StockProps     =   64
         BackColorStyle  =   1
         DAutoSizeCols   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OperationMode   =   2
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmStatImport.frx":1BB2
         ScrollBarTrack  =   3
         AppearanceStyle =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Le sexe sera forcé à homme si aucun fichier Excel n'est sélectionné"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   7575
      End
      Begin VB.Label lblSelPer 
         Caption         =   "Label2"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1560
         Width           =   11175
      End
      Begin VB.Label Label1 
         Caption         =   "Sexe de l'assurée du type Statutaire :"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         ToolTipText     =   "Le fichier Excel qui contienne l'information concernant du sexe de l'assurée"
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type d'Import"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14775
      Begin VB.OptionButton optImport 
         Caption         =   "Import avec éclatement en 2 lots (Statutaire et Non-Statutaire)"
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   7695
      End
      Begin VB.OptionButton optImport 
         Caption         =   "Import Standard en 1 lot"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmStatImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6835016A"
Option Explicit
Option Base 0

'##ModelId=5C8A68350264
Public PeriodeType As String
'##ModelId=5C8A68350283
Public ImportType As String
'##ModelId=5C8A683502A2
Public Success As Boolean
'Public PathSexeStat As String


'##ModelId=5C8A683502B2
Private Sub Form_Load()

  NumPeriodeStat = 0
  NumPeriodeNonStat = 0
  Success = False
  PathSexFileExcel = ""
  
  ImportType = cImportStandard
  
  SetFormToStandard
  
  If PeriodeType = cPeriodeStat Then
    NumPeriodeStat = numPeriode
    lblSelPer = "Sélectionnez la période sur laquelle toutes les assurées du type ""NON STATUTAIRE"" seront importées :"
  Else
    NumPeriodeNonStat = numPeriode
    lblSelPer = "Sélectionnez la période sur laquelle toutes les assurées du type ""STATUTAIRE"" seront importées :"
  End If
    
  FillGrid
  
End Sub

'##ModelId=5C8A683502D1
Private Sub cmdClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A683502E1
Private Sub cmdImport_Click()

'validation for STAT Import
If optImport(1).Value = True Then
  If NumPeriodeNonStat = 0 Then
    MsgBox "Aucune période pour l'import des assurées du type Non-Statutaire a été sélectionnée. S'il vous plait sélectionnez la période sur laquelle vous voulez importer les assurées du type Statutaire !", vbOKOnly, "Période Non-Statutaire manquant"
    Exit Sub
  End If
  
  If NumPeriodeStat = 0 Then
    MsgBox "Aucune période pour l'import des assurées du type Statutaire a été sélectionnée. S'il vous plait sélectionnez la période sur laquelle vous voulez importer les assurées du type Statutaire !", vbOKOnly, "Période Statutaire manquant"
    Exit Sub
  End If
  
'  If InStr(txtSexe.text, ".xls") <= 0 Then
'    MsgBox "Le fichier qui fournit l'information concernant le sexe de l'assurée n'est pas valable pour cette opération. S'il vous plait sélectionnez un fichier du type Excel !", vbOKOnly, "Mauvaise type du fichier !"
'    Exit Sub
'  End If
  
  '### Test the Excel File if it is correctly formated -- UNC PATH ONLY ???
  PathSexFileExcel = txtSexe.text
  
  'Load the Sexe Recordset for the function
  If PathSexFileExcel <> "" Then
    Dim SrcDB As DAO.Database
    Dim rsSexe As DAO.Recordset
    
    On Error Resume Next
    
    Set SrcDB = OpenDatabase(PathSexFileExcel, dbDriverNoPrompt, True, cdExcelExtendedPropertiesDAO)
    
    If SrcDB Is Nothing Then
      MsgBox "S'il vous plait, sélectionnez un fichier du type Excel (.xls)", vbOKOnly, "Mauvaise type du fichier"
      Exit Sub
    End If
    
    If Err.Number <> 0 Then
      MsgBox "Erreur : " & Err.Description & " Numero : " & Err.Number, vbOKOnly, "Erreur"
      Exit Sub
    End If
    
    SrcDB.QueryTimeout = 120
    
    'Various other test - not absolutely required right now
    
'    Set rsSexe = SrcDB.OpenRecordset("Select * From DONNEES_LOT", dbOpenSnapshot) 'Select * From DONNEES_LOT Where AssID = 1
    
'    If Err.Number = 3011 Then
'      MsgBox "Il y a un problème avec le fichier Excel que vous avez sélectionné. Vérifiez si la zone ""DONNEES_LOT"" a été correctement définie.", vbOKOnly, "Zone non défini"
'      Exit Sub
'    End If
'
'    rsSexe.MoveLast
'    rsSexe.MoveFirst
'
'    If rsSexe.RecordCount <= 1 Then
'      MsgBox "Le fichier sélectionné est vide !", vbOKOnly, "Fichier Vide"
'      Exit Sub
'    End If
    
    rsSexe.Close
    
    On Error GoTo 0
      
'    Dim Sexe As String
'
'    If Not rsSexe.EOF() Then
'      Sexe = rsSexe.fields("Sexe")
'      rsSexe.Close
'    End If

'    Do Until rsSexe.EOF
'      MsgBox rsSexe.fields("Sexe")
'      rsSexe.MoveNext
'    Loop

  
  End If 'If PathSexFileExcel <> "" Then
  
  If PathSexFileExcel = "" Then
    SexAllMale = True
  Else
    SexAllMale = False
  End If
  
  
  'Launch Import
  Success = True
  Unload Me

Else
  'Standard Import
  Success = True
  Unload Me

End If

End Sub

'##ModelId=5C8A683502F0
Private Sub SetFormToStandard()

  Me.Height = 2400
  Frame2.Visible = False
  cmdClose.top = 1360
  cmdImport.top = 1360

End Sub

'##ModelId=5C8A6835030F
Private Sub SetFormToStat()

  Me.Height = 9930
  Frame2.Visible = True
  cmdClose.top = 8880
  cmdImport.top = 8880
  
End Sub

'##ModelId=5C8A6835032F
Private Sub cmdGetSex_Click()

  Dim fName As String

  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Sélectionner un fichier Excel qui contienne l'information concernant du sexe de l'assurée"
  'CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.filter = "Fichiers Excel|*.xls"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
  
  fName = CommonDialog1.filename
  
  If InStr(fName, ".xls") <= 0 Then
    MsgBox "Le fichier sélectionné n'est pas valable pour cette opération. S'il vous plait sélectionnez un fichier du type Excel !", vbOKOnly, "Mauvaise type du fichier !"
    txtSexe.text = ""
    Exit Sub
  Else
    txtSexe.text = fName
  End If
  
End Sub

'##ModelId=5C8A6835033E
Private Sub optImport_Click(Index As Integer)

  Select Case Index
  Case 0 'Standard
    SetFormToStandard
    ImportType = cImportStandard
    TwoLotImport = False
  Case 1  'Statutaire
    SetFormToStat
    ImportType = cImportStat
    TwoLotImport = True
  
  End Select

End Sub

'##ModelId=5C8A6835035E
Private Sub FillGrid()

  Dim r As Long, tr As Long
  
  Screen.MousePointer = vbHourglass
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
  
  RefreshListe
  
  sprListe.TopRow = tr
  sprListe.SetActiveCell 2, r
  
  sprListe.Row = r
  sprListe.SelModeSelected = True
  
  Screen.MousePointer = vbDefault

End Sub

'##ModelId=5C8A6835036D
Private Sub RefreshListe()
  
  Dim rq As String
  
  m_dataSource.SetDatabase dtaPeriode
  sprListe.DataSource = dtaPeriode
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = True
  sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  
  DoEvents
  
  '& "P.StatutArchRest as [Statut] " _

  rq = "SELECT P.RECNO, P.PENUMCLE as [Numéro Période], " _
        & "CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + TP.Libelle as [Type], " _
        & "CAST(P.IdTypeCalcul as VARCHAR) + ' - ' + TC.Libelle as [Type Calcul], " _
        & "P.PEDATEDEB as [Début], " _
        & "P.PEDATEFIN as [Fin], " _
        & "P.PECOMMENTAIRE as Commentaire " _
        & "FROM P3IUser.Periode P LEFT JOIN P3IUser.TypePeriode TP ON TP.IdTypePeriode=P.PETYPEPERIODE " _
        & "LEFT JOIN P3IUser.TypeCalcul TC ON TC.IdTypeCalcul=P.IdTypeCalcul " _
        & "WHERE (P.StatutArchRest <> 'Archivée' or P.StatutArchRest is NULL) And P.PEGPECLE = " & GroupeCle    'P.StatutArchRest <> 'Archivée' And
        
  If PeriodeType = cPeriodeStat Then
    rq = rq & " And P.PETYPEPERIODE <> 6"
  Else
    rq = rq & " And P.PETYPEPERIODE = 6"
  End If
  
  rq = rq & " ORDER BY P.PENUMCLE DESC "
  
  sprListe.Visible = False

  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
  
  Set sprListe.DataSource = dtaPeriode
  
  ' mets à jours les n° de ligne dans le spread
  If Not dtaPeriode.Recordset.EOF Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
    
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
  Else
    sprListe.MaxRows = 0
  End If

  sprListe.Refresh
  
  SetColonneDataFill 3, True
  SetColonneDataFill 8, True
  
  sprListe.ColWidth(2) = 8
  sprListe.ColWidth(3) = 20  'Type
  sprListe.ColWidth(4) = 10
  sprListe.ColWidth(5) = 10
  sprListe.ColWidth(6) = 10
  sprListe.ColWidth(7) = 50
  'sprListe.ColWidth(8) = 9
  
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = 7
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = 3
  sprListe.Col2 = 3
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.Col = 7
  sprListe.Col2 = 7
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
  sprListe.ActiveCellHighlightStyle = ActiveCellHighlightStyleOff 'switch off rectangle around highlighted cell
  
  On Error Resume Next
     
  sprListe.OperationMode = OperationModeNormal ' OperationModeSingle ' OperationModeNormal
  sprListe.EditMode = True
  sprListe.Enabled = True
    
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True

  Me.SetFocus
  sprListe.SetFocus
  
End Sub

'##ModelId=5C8A6835038C
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  sprListe.sheet = sprListe.ActiveSheet
  sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
End Sub

'##ModelId=5C8A683503BB
Private Sub sprListe_Click(ByVal Col As Long, ByVal Row As Long)
  
  SetNumPeriode
  
  Dim r As Long, tr As Long
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
    
  Screen.MousePointer = vbHourglass
    
  sprListe.TopRow = tr
  sprListe.SetActiveCell 2, r
  
  sprListe.Row = r
  sprListe.SelModeSelected = True
  
  sprListe.SetFocus
  
  Screen.MousePointer = vbDefault
  
End Sub

'##ModelId=5C8A68360012
Private Sub SetNumPeriode()
  
  If sprListe.ActiveRow < 0 Then Exit Sub
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 2
  
  If PeriodeType = cPeriodeStat Then
    NumPeriodeNonStat = CLng(sprListe.text)
  Else
    NumPeriodeStat = CLng(sprListe.text)
  End If
  
End Sub

'##ModelId=5C8A68360031
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)

  Dim comment As Variant, i As Integer
  Dim archive As Variant
  
  If dtaPeriode.Recordset.fields(Col - 1).Name = "Type" Then

    sprListe.BlockMode = True
    sprListe.Col = -1
    sprListe.Row = Row
    sprListe.Col2 = -1
    sprListe.Row2 = Row
      
    sprListe.GetDataFillData comment, vbString
    
    If Len(comment) > 0 Then
        Select Case CInt(Left(comment, 1))
          Case eProvision  ' Provision
            sprListe.BackColor = jaune_clair
          
          Case eCapitalConstitutifRente  ' Rente
            sprListe.BackColor = vert_clair
          
          Case eRevalo  ' Revalo
            sprListe.BackColor = bleu_clair
            
          Case Else
            sprListe.BackColor = orange_clair
        End Select
        
        sprListe.ForeColor = noir
        
    End If
        
    sprListe.BlockMode = False
    
  Else
    
    sprListe.GetDataFillData comment, vbString
    If comment = "" Then
      sprListe.Col = Col
      sprListe.Row = Row
      sprListe.Value = ""

      Cancel = True
    End If
  
  End If
  
  'set background color for archived items
'  If dtaPeriode.Recordset.fields(Col - 1).Name = "Statut" Then
'
'    sprListe.GetDataFillData archive, vbString
'
'    If Len(archive) > 0 Then
'      If Left$(LCase(archive), 4) = "arch" Then
'        sprListe.BlockMode = True
'        sprListe.Col = -1
'        sprListe.Row = Row
'        sprListe.Col2 = -1
'        sprListe.Row2 = Row
'        sprListe.BackColor = LTRED
'
'        sprListe.ForeColor = noir
'
'        sprListe.BlockMode = False
'      End If
'    End If
'
'  End If
  
End Sub

'##ModelId=5C8A6836009E
Private Sub sprListe_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

  If NewRow = -1 Then
        Exit Sub
    End If
    
    'change color back to original color
    Dim typePeriode As String
    sprListe.Col = 3
    sprListe.Row = Row
    typePeriode = sprListe.text
    
    sprListe.Col = -1
    sprListe.ForeColor = noir
    
    If Len(typePeriode) > 0 Then
        Select Case CInt(Left(typePeriode, 1))
          Case eProvision  ' Provision
            sprListe.BackColor = jaune_clair
          
          Case eCapitalConstitutifRente  ' Rente
            sprListe.BackColor = vert_clair
          
          Case eRevalo  ' Revalo
            sprListe.BackColor = bleu_clair
            
          Case Else
            sprListe.BackColor = orange_clair
        End Select
    End If
    
    'change background color for archived items
    Dim statut As String
    sprListe.Col = 10
    sprListe.Row = Row
    statut = sprListe.text

    sprListe.Col = -1
    sprListe.ForeColor = noir

    If Len(statut) > 0 Then
      If Left$(LCase(statut), 4) = "arch" Then
        sprListe.BackColor = LTRED
      End If
    End If
    
    'change background color to black for the row that receives the focus
    sprListe.Row = Row
    sprListe.ForeColor = noir
    
    sprListe.Row = NewRow
    sprListe.BackColor = noir
    sprListe.ForeColor = blanc
    
End Sub
