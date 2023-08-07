VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAutoImportTables 
   Caption         =   "Import des tables de paramétrages"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21930
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   21930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   345
      Left            =   15960
      TabIndex        =   7
      Top             =   10560
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10170
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20010
      _ExtentX        =   35295
      _ExtentY        =   17939
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Périodes"
      TabPicture(0)   =   "frmAutoImportTables.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtaPeriode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "sprListe"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboCATR9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdImport"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Tables à importer"
      TabPicture(1)   =   "frmAutoImportTables.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dtaCATR9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "sprCATR9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboCATR92"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdImport2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdImport 
         Caption         =   "Importer dans toutes les périodes sélectionnées"
         Height          =   330
         Left            =   5760
         TabIndex        =   10
         Top             =   9240
         Width           =   4425
      End
      Begin VB.ComboBox cboCATR9 
         Height          =   315
         ItemData        =   "frmAutoImportTables.frx":0038
         Left            =   1920
         List            =   "frmAutoImportTables.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   9240
         Width           =   3480
      End
      Begin VB.CommandButton cmdImport2 
         Caption         =   "&Importer"
         Height          =   330
         Left            =   -71040
         TabIndex        =   2
         Top             =   675
         Width           =   1185
      End
      Begin VB.ComboBox cboCATR92 
         Height          =   315
         ItemData        =   "frmAutoImportTables.frx":003C
         Left            =   -74820
         List            =   "frmAutoImportTables.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   675
         Width           =   3480
      End
      Begin FPSpreadADO.fpSpread sprCATR9 
         Bindings        =   "frmAutoImportTables.frx":0040
         Height          =   6270
         Left            =   -74865
         TabIndex        =   3
         Top             =   1125
         Width           =   9015
         _Version        =   524288
         _ExtentX        =   15901
         _ExtentY        =   11060
         _StockProps     =   64
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
         OperationMode   =   3
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmAutoImportTables.frx":0057
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaCATR9 
         Height          =   330
         Left            =   -70050
         Top             =   360
         Visible         =   0   'False
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
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
         Caption         =   "dtaCATR9"
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
         Height          =   7605
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   18165
         _Version        =   524288
         _ExtentX        =   32041
         _ExtentY        =   13414
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
         SpreadDesigner  =   "frmAutoImportTables.frx":0493
         ScrollBarTrack  =   3
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaPeriode 
         Height          =   330
         Left            =   4800
         Top             =   600
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
      Begin VB.Label Label2 
         Caption         =   "Tables à afficher :"
         Height          =   240
         Left            =   480
         TabIndex        =   9
         Top             =   9240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Sélectionner les périodes destinataires"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label26 
         Caption         =   "Tables à afficher :"
         Height          =   240
         Left            =   -74820
         TabIndex        =   4
         Top             =   405
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAutoImportTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fmAction As Integer
Private fmTableDiverse As clsTablesDiverses

Private Sub Form_Load()

  Screen.MousePointer = vbHourglass

  ' activate first tab
  SSTab1.Tab = 0
  
  ' liste des tables diverses
  InitTableDiverse
  fmTableDiverse.FillCombo cboCATR9
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub InitTableDiverse()
  
  If fmTableDiverse Is Nothing Then
    Set fmTableDiverse = New clsTablesDiverses
  Else
    fmTableDiverse.Clear
  End If
  
  fmTableDiverse.AddTableDiverse "AgeDepartRetraite", "DateDebut", "DateDebut", "DateDebut, DateFin, AgeMinimum, AgeMinimumTauxPlein"
  
  fmTableDiverse.AddTableDiverse "Capitaux_Moyens", "Cat_Opt_Synth", "Cat_Opt_Synth", "Cat_Opt_Synth,Capital_Moyen"
  
  fmTableDiverse.AddTableDiverse "CATR9", "Categorie", "Categorie", "Categorie, Incap, Passage, PassageSuivantNCA"
  
  fmTableDiverse.AddTableDiverse "CATR9INVAL", "Categorie", "Categorie", "Categorie"
      
  fmTableDiverse.AddTableDiverse "CDSITUAT", "Code_CIE,Code_APP,CDSITUAT", "Code_CIE,Code_APP,CDSITUAT", "Code_CIE,Code_APP,CDSITUAT,LBSITUAT,LBSITUATCOURT,TAUXPM"
  
  fmTableDiverse.AddTableDiverse "CodesCat", "Code_CIE,Code_APP,Code_Cat_Contrat", "Code_CIE,Code_APP,Code_Cat_Contrat,NumParamCalcul", "Code_CIE,Code_APP,Code_Cat_Contrat,NumParamCalcul,Periode,RgrtPdtSup,RgrtPdt,Lib_Cat_Contrat,Nbj_Selection"
  
  fmTableDiverse.AddTableDiverse "CodeCatInv", "CDCHOIXPREST", "CDCHOIXPREST,LBCHOIXPREST,CDCATINV,LBCATINV,CategorieInval", "CDCHOIXPREST,LBCHOIXPREST,CDCATINV,LBCATINV,CategorieInval"
  
  fmTableDiverse.AddTableDiverse "CoeffAmortissement", "IdTypeCalcul,AnneeMoisCalcul,DateDebut", "IdTypeCalcul,AnneeMoisCalcul,DateDebut", "IdTypeCalcul,AnneeMoisCalcul, DateDebut, DateFin, CoeffAmortissement"
  
  fmTableDiverse.AddTableDiverse "Correspondance_CatOption", "CodeOption", "CodeOption", "CodeOption,Categorie,Cat_Opt_Synth"
        
  fmTableDiverse.AddTableDiverse "PassageNCA", "NCA", "NCA", "NCA, Passage"
  
  fmTableDiverse.AddTableDiverse "ParamRentes", "DateDebut", "TauxTechnique, TauxRevalo, FraisGestion, TableHomme, TableFemme", "DateDebut, DateFin, TauxTechnique, TauxRevalo, FraisGestion, TableHomme, TableFemme"
  
  fmTableDiverse.AddTableDiverse "PM_Retenue", "AnneeSurvenance", "AnneeSurvenance", "AnneeSurvenance,PMAvecCorrectif,PMReassAvecCorrectif"
  
  'fmTableDiverse.AddTableDiverse "Produits", "Annee,Mois,Code_Produit", "Annee,Mois,Code_Produit", "Annee,Mois,Code_Produit,Lib_Produit,Code_Cat_Contrat,Lib_Cat_Contrat,Code_Famille_Comptable,Lib_Famille_Comptable"
  
  fmTableDiverse.AddTableDiverse "Reassurance", "Regime, Categorie, NCA", "Regime, Categorie", "Regime, Categorie, NCA, NomNCA, NomReass, NomTraite, Effet, Resiliation, AnneeSinistre, Taux, Captive, ReprisePassif"
  
  'fmTableDiverse.AddTableDiverse "REGA01", "Code_GE", "Code_GE", "Code_GE,Lib_Long_GE,GAR_GENERALE,Code_GS,TYPE_REGLEMENT,Code_PROV"
  
  fmTableDiverse.AddTableDiverse "TBQREGA", "Code_CIE,Code_APP,Code_GE", "Code_CIE,Code_APP,Code_GE,Code_Prov", "Code_CIE,Code_APP,Code_GE,Lib_Court_GE,Lib_Long_GE,Code_GS,Lib_Court_GS,Lib_Long_GS,Code_GI,Lib_Court_GI,Lib_Long_GI,GAR_GENERALE,Code_Prov"
  
  fmTableDiverse.AddTableDiverse "DonneesSociales", "DateDebut", "SmicHoraireBrut,  CoefficientSmic,  TauxIJNettes, PlafondAnnuelSS,  TauxRembSS, TauxGarantieAssureur", "DateDebut,  DateFin,  SmicHoraireBrut,  CoefficientSmic,  TauxIJNettes, PlafondAnnuelSS,  TauxRembSS, TauxGarantieAssureur"
  
  'fmTableDiverse.AddTableDiverse "PSAP_Baremes", "Garantie, DebutPeriode", "Garantie, DebutPeriode", "Garantie, DebutPeriode, FinPeriode, Taux"
  
  
End Sub

Private Sub cmdImport_Click()
  Dim nomTable As String
  
  If DroitAdmin = False Then Exit Sub
  If cboCATR9.ListIndex = -1 Then Exit Sub
  
  nomTable = cboCATR9.List(cboCATR9.ListIndex)
  
  ImportGenerique CommonDialog1, ProgressBar1, nomTable, NumPeriode
  ' rafraichi la page
  Dim i As Integer
  
  i = cboCATR9.ListIndex
  
  fmTableDiverse.FillCombo cboCATR9
  
  cboCATR9.ListIndex = i
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set fmTableDiverse = Nothing
  Screen.MousePointer = vbDefault
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  If SSTab1.Tab = 1 Then
    fmTableDiverse.FillCombo cboCATR9
  End If
End Sub

Private Sub cboCATR9_Click()
  Dim rq As String
    
  If cboCATR9.ListIndex <> -1 Then
    sprCATR9.ReDraw = False
    
    m_dataSource.SetDatabase dtaCATR9
        
    On Error Resume Next
        
    Dim defTable As defTableDiverse
    
    Set defTable = fmTableDiverse.TableInfo(cboCATR9.ListIndex)
    If defTable.champs = "" Then
      defTable.champs = "*"
    End If
        
    rq = "SELECT " & defTable.champs & " FROM " & defTable.nomTable & " WHERE NumPeriode=" & NumPeriode & " And GroupeCle=" & GroupeCle & " ORDER BY " & defTable.orderBy
        
    dtaCATR9.RecordSource = m_dataHelper.ValidateSQL(rq)
    dtaCATR9.Refresh
    
    Set sprCATR9.DataSource = dtaCATR9
    
    If Not dtaCATR9.Recordset.EOF Then
      dtaCATR9.Recordset.MoveLast
      dtaCATR9.Recordset.MoveFirst
      
      sprCATR9.Refresh
      
      sprCATR9.MaxRows = dtaCATR9.Recordset.RecordCount
      
      Dim i As Integer

      For i = 2 To sprCATR9.MaxCols
        sprCATR9.Col = i
        sprCATR9.DataColCnt = True
      Next
      
      dtaCATR9.Refresh
    Else
      sprCATR9.MaxRows = 0
    End If
    
    sprCATR9.Refresh
    
    ' largeur des colonnes
    LargeurMaxColonneSpread sprCATR9
    
    sprCATR9.ReDraw = True
    
    On Error GoTo 0
  End If
End Sub

Private Sub sprCATR9_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  
  If DataField = "CoeffAmortissement" Then
    sprCATR9.BlockMode = True
    sprCATR9.Col = Col
    sprCATR9.Row = -1
    sprCATR9.Col2 = Col
    sprCATR9.Row2 = -1
    
    sprCATR9.TypeNumberDecPlaces = 4
    sprCATR9.BlockMode = False
  End If

End Sub


