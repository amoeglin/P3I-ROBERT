VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDetailPeriode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Détails de la Periode"
   ClientHeight    =   8520
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmDetailPeriode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   1080
      Top             =   0
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   582
      ConnectMode     =   3
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
      Caption         =   "datPrimaryRS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7530
      Left            =   90
      TabIndex        =   11
      Top             =   495
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   13282
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Général"
      TabPicture(0)   =   "frmDetailPeriode.frx":1BB2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Tables Diverses"
      TabPicture(1)   =   "frmDetailPeriode.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "btnExportCATR9"
      Tab(1).Control(1)=   "btnPrintCATR9"
      Tab(1).Control(2)=   "btnImportCATR9"
      Tab(1).Control(3)=   "cboCATR9"
      Tab(1).Control(4)=   "sprCATR9"
      Tab(1).Control(5)=   "dtaCATR9"
      Tab(1).Control(6)=   "Label26"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Correctif BCAC Affiné"
      TabPicture(2)   =   "frmDetailPeriode.frx":1BEA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtMonth"
      Tab(2).Control(1)=   "radBCACNew"
      Tab(2).Control(2)=   "radBCACOld"
      Tab(2).Control(3)=   "Label3"
      Tab(2).Control(4)=   "Label2"
      Tab(2).ControlCount=   5
      Begin VB.TextBox txtMonth 
         Height          =   285
         Left            =   -72120
         TabIndex        =   43
         Text            =   "24"
         Top             =   3360
         Width           =   495
      End
      Begin VB.OptionButton radBCACNew 
         Caption         =   "Nouvelle méthode : application en fonction du nombre mois glissants entre la date du sinistre et la date d'arrêté"
         Height          =   615
         Left            =   -74400
         TabIndex        =   45
         Top             =   2760
         Value           =   -1  'True
         Width           =   6495
      End
      Begin VB.OptionButton radBCACOld 
         Caption         =   $"frmDetailPeriode.frx":1C06
         Height          =   735
         Left            =   -74400
         TabIndex        =   44
         Top             =   1560
         Width           =   6255
      End
      Begin VB.Frame Frame11 
         Caption         =   "P3I Individuel - Sécurité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   38
         Top             =   2520
         Width           =   4110
         Begin VB.CheckBox chkLocked 
            Caption         =   "Verrouiller la période (interdire les modifications)"
            Height          =   780
            Index           =   0
            Left            =   1800
            TabIndex        =   41
            Top             =   240
            Width           =   2175
         End
         Begin VB.CheckBox P3I_Individuel 
            Caption         =   "P3I_INDIVIDUEL"
            Height          =   780
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Type de période"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4455
         TabIndex        =   35
         Top             =   2520
         Width           =   4695
         Begin VB.ComboBox cboTypeProvision 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   270
            Width           =   4380
         End
         Begin VB.ComboBox cboTypeCalcul 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   630
            Width           =   4380
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Commentaires"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   180
         TabIndex        =   32
         Top             =   450
         Width           =   8970
         Begin VB.TextBox txtFields 
            Height          =   1545
            Index           =   6
            Left            =   135
            MaxLength       =   1024
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   270
            Width           =   8670
         End
      End
      Begin VB.CommandButton btnExportCATR9 
         Caption         =   "&Exporter"
         Height          =   330
         Left            =   -68295
         TabIndex        =   29
         Top             =   675
         Width           =   1185
      End
      Begin VB.CommandButton btnPrintCATR9 
         Caption         =   "&Imprimer"
         Height          =   330
         Left            =   -71265
         TabIndex        =   28
         Top             =   675
         Width           =   1455
      End
      Begin VB.CommandButton btnImportCATR9 
         Caption         =   "&Importer"
         Height          =   330
         Left            =   -69555
         TabIndex        =   27
         Top             =   675
         Width           =   1185
      End
      Begin VB.ComboBox cboCATR9 
         Height          =   315
         ItemData        =   "frmDetailPeriode.frx":1C9B
         Left            =   -74820
         List            =   "frmDetailPeriode.frx":1C9D
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   675
         Width           =   3480
      End
      Begin VB.Frame Frame2 
         Caption         =   "Jeux de Paramètres de calcul"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Left            =   180
         TabIndex        =   21
         Top             =   5310
         Width           =   8970
         Begin VB.CommandButton btnDefaultParam 
            Caption         =   "&Par Défaut"
            Height          =   330
            Left            =   7425
            TabIndex        =   34
            Top             =   1530
            Width           =   1230
         End
         Begin VB.CommandButton btnDelParam 
            Caption         =   "&Supprimer"
            Height          =   330
            Left            =   7425
            TabIndex        =   25
            Top             =   1125
            Width           =   1230
         End
         Begin VB.CommandButton btnEditParam 
            Caption         =   "&Modifier"
            Height          =   330
            Left            =   7425
            TabIndex        =   24
            Top             =   720
            Width           =   1230
         End
         Begin VB.CommandButton btnAddPAram 
            Caption         =   "&Ajouter"
            Height          =   330
            Left            =   7425
            TabIndex        =   23
            Top             =   315
            Width           =   1230
         End
         Begin MSComctlLib.ListView lvParamCalcul 
            Height          =   1545
            Left            =   135
            TabIndex        =   22
            Top             =   315
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Relancez l'IMPORT si vous modifiez ces valeurs :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1545
         Left            =   180
         TabIndex        =   12
         Top             =   3690
         Width           =   8970
         Begin VB.TextBox txtFields 
            DataField       =   "PENBJOURMAX"
            DataSource      =   "datPrimaryRS"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   1
            Left            =   7965
            TabIndex        =   5
            Text            =   "180"
            Top             =   675
            Width           =   405
         End
         Begin VB.TextBox txtFields 
            Height          =   330
            Index           =   5
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1125
            Width           =   1260
         End
         Begin VB.TextBox txtFields 
            DataField       =   "PENBJOURDC"
            DataSource      =   "datPrimaryRS"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   2
            Left            =   7965
            TabIndex        =   4
            Text            =   "180"
            Top             =   315
            Width           =   405
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Bindings        =   "frmDetailPeriode.frx":1C9F
            Height          =   330
            Left            =   1530
            TabIndex        =   1
            Top             =   315
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   185270275
            CurrentDate     =   36114
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   1530
            TabIndex        =   2
            Top             =   720
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   185270275
            CurrentDate     =   36114
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fin de période"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   19
            Top             =   765
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Date d'arreté"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   5
            Left            =   180
            TabIndex        =   18
            Top             =   1170
            Width           =   960
         End
         Begin VB.Label lblLabels 
            Caption         =   "Début de période"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   17
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "jours"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   8415
            TabIndex        =   16
            Top             =   720
            Width           =   345
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Non prise en compte des prestations payées au delà de "
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   1
            Left            =   3825
            TabIndex        =   15
            Top             =   720
            Width           =   4065
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Nb jours écoulés depuis le décès pour calcul de la PSAP DECES"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   18
            Left            =   3240
            TabIndex        =   14
            Top             =   360
            Width           =   4605
         End
         Begin VB.Label Label4 
            Caption         =   "jours"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   8415
            TabIndex        =   13
            Top             =   360
            Width           =   345
         End
      End
      Begin FPSpreadADO.fpSpread sprCATR9 
         Bindings        =   "frmDetailPeriode.frx":1CBA
         Height          =   6270
         Left            =   -74865
         TabIndex        =   30
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
         SpreadDesigner  =   "frmDetailPeriode.frx":1CD1
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaCATR9 
         Height          =   330
         Left            =   -70080
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
      Begin VB.Label Label3 
         Caption         =   "Nombre de mois glissants :"
         Height          =   255
         Left            =   -74160
         TabIndex        =   46
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Choix de l'application du correctif BCAC Affiné : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   42
         Top             =   840
         Width           =   6255
      End
      Begin VB.Label Label26 
         Caption         =   "Tables à afficher :"
         Height          =   240
         Left            =   -74820
         TabIndex        =   31
         Top             =   405
         Width           =   1455
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9540
      TabIndex        =   9
      Top             =   8085
      Width           =   9540
      Begin VB.CommandButton btnFicheParam 
         Caption         =   "&Fiche Paramètre"
         Height          =   345
         Left            =   2160
         TabIndex        =   39
         Top             =   45
         Width           =   1380
      End
      Begin VB.CommandButton btnProvision 
         Caption         =   "&Provisions à l'ouverture"
         Height          =   345
         Left            =   90
         TabIndex        =   6
         Top             =   45
         Width           =   2010
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fermer"
         Height          =   345
         Left            =   4905
         TabIndex        =   8
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Enregistrer"
         Height          =   345
         Left            =   3825
         TabIndex        =   7
         Top             =   45
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8955
         Top             =   -45
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Base de données source"
         FileName        =   "*.mdb"
         Filter          =   "*.mdb"
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   6120
         TabIndex        =   20
         Top             =   45
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "RECNO"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   4815
      TabIndex        =   0
      Top             =   -45
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblGroupe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Groupe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   10
      Top             =   90
      Width           =   9330
   End
End
Attribute VB_Name = "frmDetailPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67DC02C1"
Option Explicit

'##ModelId=5C8A67DC03AB
Private fmAction As Integer

'##ModelId=5C8A67DC03CD
Private fmTableDiverse As clsTablesDiverses

'##ModelId=5C8A67DC03CE
Private Sub InitTableDiverse()
  If fmTableDiverse Is Nothing Then
    Set fmTableDiverse = New clsTablesDiverses
  Else
    fmTableDiverse.Clear
  End If
  
  
  fmTableDiverse.AddTableDiverse "AgeDepartRetraite", "DateDebut", "DateDebut", "DateDebut, DateFin, AgeMinimum, AgeMinimumTauxPlein"
  
  fmTableDiverse.AddTableDiverse "AgeDepartRetraiteInval1", "DateDebut", "DateDebut", "DateDebut, DateFin, AgeMinimum, AgeMinimumTauxPlein"
  
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

'##ModelId=5C8A67DD0002
Private Sub SetNumParamCalcul()
  NumParamCalcul = 0
  
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  NumParamCalcul = CLng(lvParamCalcul.SelectedItem.SubItems(1))
End Sub

'##ModelId=5C8A67DD0021
Private Sub btnAddPAram_Click()
  NumParamCalcul = -1
  frmParamCalcul.Show vbModal, Me

  ' Liste des paramètres de calcul
  RefreshListParamCalcul
End Sub

'##ModelId=5C8A67DD0032
Private Sub btnDefaultParam_Click()
  
  DefaultParam True

End Sub


'##ModelId=5C8A67DD0040
Private Sub DefaultParam(bMsg As Boolean)
 
  If bMsg = False Or MsgBox("Voulez-vous remplacer les paramètres existants par ceux par défaut ?", vbQuestion + vbYesNo, "Paramètres de calcul par défaut") <> vbYes Then Exit Sub
  
  Dim frm As New frmTypeExport, typeParam As String
  
  frm.frmSignalisation.Visible = False
  frm.frmModification.Visible = False
  frm.Caption = "Type de paramètres à charger"
  
  ret_code = 0
  frm.Show vbModal
  If ret_code = 0 Then
    
    ' nom standardise Generali
    Select Case frm.frmTypeProvision
      Case 1
        typeParam = "BILAN"
        
      Case 2
        typeParam = "CLIENT"
        
      Case 3
        typeParam = "SIMULATION"
    End Select
  
  Else
  
    Set frm = Nothing
    Exit Sub
    
  End If
  
  Set frm = Nothing
  
  If bMsg = False Or MsgBox("ATTENTION: Les paramètres existants vont définitivement être écrasés par ceux par défaut." & vbLf & vbLf & "Voulez-vous continuer ?", vbCritical + vbYesNo, "Paramètres de calcul par défaut") <> vbYes Then Exit Sub
      
  ' efface les paramètres existant
  m_dataSource.Execute "DELETE FROM ParamCalcul WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode
  
  ' charge les paramètres par défaut
  Dim aRet() As String
  
  If EnumSections(aRet, sFichierIni) Then
    Dim i As Integer, num As Long
    Dim sRet As String, sName As String
    Dim theParam As New clsParamCalcul
    
    sName = DEFAULT_PARAM_SECTION & typeParam & "_"
    
    sRet = vbNullString
    For i = LBound(aRet) To UBound(aRet)
      If Len(aRet(i)) > Len(sName) Then
        If UCase(Left(aRet(i), Len(sName))) = UCase(sName) Then
          num = CLng(mID(aRet(i), Len(sName) + 1))
          
          theParam.LoadFromIni num, typeParam
          theParam.SaveToDB GroupeCle, numPeriode
        End If
      End If
    Next
    
    Set theParam = Nothing
  End If
  
  Erase aRet
 
  
  ' Liste des paramètres de calcul
  RefreshListParamCalcul
End Sub

'##ModelId=5C8A67DD0060
Private Sub btnDelParam_Click()
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  On Error GoTo err_del
  
  SetNumParamCalcul

  If MsgBox("Voulez vous vraiement supprimer les paramètres de calcul : " & NumParamCalcul & " ?", vbQuestion + vbYesNo) = vbYes Then
    m_dataSource.Execute "DELETE FROM ParamCalcul WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode & " AND PENUMPARAMCALCUL=" & NumParamCalcul
    
    ' Liste des paramètres de calcul
    RefreshListParamCalcul
  End If
  
  Exit Sub
  
err_del:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=5C8A67DD007F
Private Sub btnEditParam_Click()
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  SetNumParamCalcul
  frmParamCalcul.Show vbModal, Me
  
  ' Liste des paramètres de calcul
  RefreshListParamCalcul
End Sub

'##ModelId=5C8A67DD008E
Private Sub btnExportCATR9_Click()
  If cboCATR9.ListIndex = -1 Then Exit Sub
  
  ExportTableToExcelFile cboCATR9.List(cboCATR9.ListIndex) & "_Periode_" & numPeriode & ".xls", _
                         cboCATR9.List(cboCATR9.ListIndex), _
                         cboCATR9.List(cboCATR9.ListIndex), sprCATR9, CommonDialog1, "", False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Export des paramètres sous forme de chaine
'
'##ModelId=5C8A67DD009E
Private Sub btnFicheParam_Click()
  Dim sParam As String
  Dim theParam As clsParamCalcul
  Dim Item As ListItem
  
  ' Liste des détails de la période
  sParam = "Paramètres de calcul de la période n°" & numPeriode & vbLf & vbLf
  
  sParam = sParam & "Commentaires = " & txtFields(6).text & vbLf & vbLf
  
  sParam = sParam & "Calculs pour P3I " & IIf(P3I_Individuel(0).Value = vbChecked, "INDIVIDUEL", "Collective") & vbLf & vbLf
  
  sParam = sParam & "Sécurité = " & IIf(chkLocked(0).Value = vbChecked, "Période vérouillée", "Période non vérouillée") & vbLf & vbLf
  
  sParam = sParam & "Type de provision = " & cboTypeProvision.text & vbLf
  sParam = sParam & "Type de calcul = " & cboTypeCalcul.text & vbLf & vbLf
  
  sParam = sParam & "Date de début de période = " & Format(DTPicker1.Value, "dd/MM/yyyy") & vbLf
  sParam = sParam & "Date de fin de période = " & Format(DTPicker2.Value, "dd/MM/yyyy") & vbLf
  sParam = sParam & "Date d'arreté = " & txtFields(5).text & vbLf & vbLf
  
  sParam = sParam & "Nb jours écoulés depuis le décès pour calcul de la PSAP DECES = " & txtFields(2).text & " jours" & vbLf
  sParam = sParam & "Non prise en compte des prestations payées au delà de = " & txtFields(1).text & " jours" & vbLf & vbLf
  
  ' Provision à l'ouverture
  Dim df As Integer, ns As String
    
  df = Year(DTPicker2.Value)

  ' charge les valeurs
  Dim rs As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM ProvisionsOuverture WHERE GPECLE = " & GroupeCle & " AND NUMCLE = " & numPeriode, Snapshot)
  If Not rs.EOF Then
    sParam = sParam & "======================================================================" & vbLf
    sParam = sParam & "Provisions à l'ouverture" & vbLf
    
    sParam = sParam & "Année " & df & " = " & rs.fields("PROV_ANn").Value & " €" & vbLf
    sParam = sParam & "Année " & df - 1 & " = " & rs.fields("PROV_ANn1").Value & " €" & vbLf
    sParam = sParam & "Année " & df - 2 & " = " & rs.fields("PROV_ANn2").Value & " €" & vbLf
    sParam = sParam & "Année " & df - 3 & " = " & rs.fields("PROV_ANn3").Value & " €" & vbLf
    sParam = sParam & "Année " & df - 4 & " = " & rs.fields("PROV_ANn4").Value & " €" & vbLf
    sParam = sParam & "Année " & df - 5 & " = " & rs.fields("PROV_ANn5").Value & " €" & vbLf
  End If
  rs.Close

  
  ' Liste des paramètres de calcul
  For Each Item In lvParamCalcul.ListItems
    Set theParam = New clsParamCalcul
    
    NumParamCalcul = CLng(Item.SubItems(1))
    theParam.LoadFromDB GroupeCle, numPeriode, NumParamCalcul
    
    sParam = sParam & theParam.DumpParam & vbLf
  Next
  
  ' Affichage
  Dim frm As frmDisplayLog
  
  Set frm = New frmDisplayLog
  
  frm.bFichierLogIsText = True
  frm.FichierLog_FileName = "FicheParametres_Periode_" & numPeriode & ".txt"
  frm.FichierLog = sParam
  frm.m_sFichierIni = sFichierIni
  
  frm.Show vbModal
End Sub


'##ModelId=5C8A67DD00BD
Private Sub btnImportCATR9_Click()
  Dim nomTable As String
  
  If DroitAdmin = False Then Exit Sub
  If cboCATR9.ListIndex = -1 Then Exit Sub
  
  nomTable = cboCATR9.List(cboCATR9.ListIndex)
  
  ImportGenerique CommonDialog1, ProgressBar1, nomTable, numPeriode
  ' rafraichi la page
  Dim i As Integer
  
  i = cboCATR9.ListIndex
  
  fmTableDiverse.FillCombo cboCATR9
  
  cboCATR9.ListIndex = i
End Sub

'##ModelId=5C8A67DD00CD
Private Sub btnPrintCATR9_Click()
  Dim bUsePrintDlg As Integer
  
  If cboCATR9.ListIndex = -1 Then Exit Sub

  Dim svgFontBold As Boolean
  
  On Error GoTo err_print
  
  bUsePrintDlg = GetSettingIni(CompanyName, "Parametre", "Print_UsePrintDlg", 1)
  If bUsePrintDlg = 1 Then
    CommonDialog1.CancelError = True
    CommonDialog1.PrinterDefault = False
    CommonDialog1.Flags = cdlPDReturnDC + cdlPDNoSelection + cdlPDNoPageNums
    CommonDialog1.Orientation = cdlLandscape
    CommonDialog1.ShowPrinter
  End If
  
  With sprCATR9
    Printer.Orientation = vbPRORLandscape
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    
    If bUsePrintDlg = 1 Then
      .hDCPrinter = CommonDialog1.hDC
    Else
      .hDCPrinter = Printer.hDC
    End If
    .PrintAbortMsg = "Impression en cours - Annuler pour interrompre"
    .PrintJobName = cboCATR9
    .PrintHeader = "/c - Table " & cboCATR9 & " - /n"
    .PrintFooter = "/l Provisions Incapacité Invalidtité /c page /p /r Imprimé le " & Format(Now(), "dd/mm/yyyy") & " /n   "
    
    .PrintBorder = True
    .PrintColor = True
    .PrintGrid = True
    .PrintShadows = False
    .PrintUseDataMax = False
    
    svgFontBold = .FontBold
    .FontBold = False
    
    .PrintMarginTop = 0
    .PrintMarginBottom = 0
    .PrintMarginLeft = 0
    .PrintMarginRight = 0
    .PrintOrientation = SS_PRINTORIENT_PORTRAIT
    
    .PrintType = SS_PRINT_ALL
    .Action = SS_ACTION_SMARTPRINT
  
    .FontBold = svgFontBold
  End With

err_print:
End Sub

'##ModelId=5C8A67DD00DD
Private Sub btnProvision_Click()
  frmListProvOuverture.Show vbModal
End Sub

'##ModelId=5C8A67DD00FC
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
        
    rq = "SELECT " & defTable.champs & " FROM " & defTable.nomTable & " WHERE NumPeriode=" & numPeriode & " And GroupeCle=" & GroupeCle & " ORDER BY " & defTable.orderBy
        
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


'##ModelId=5C8A67DD011B
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




'##ModelId=5C8A67DD015A
Private Sub cmdUpdate_Click()
  'On Error GoTo errcmdUpdate
   
  Dim curBookmark As Variant
  
  Screen.MousePointer = vbHourglass
  
  curBookmark = datPrimaryRS.Recordset.bookmark
   
  If fmAction = 1 Then
    'datPrimaryRS.Recordset.Edit
  End If
  
  ' Commentaires
  txtFields(6) = Trim(txtFields(6))
  If txtFields(6) = "" Then
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez saisir un commentaire", vbCritical
    SSTab1.Tab = 1
    txtFields(6).SetFocus
    datPrimaryRS.Recordset.CancelUpdate
    Exit Sub
  End If
  datPrimaryRS.Recordset.fields("PECOMMENTAIRE") = txtFields(6)
      
  datPrimaryRS.Recordset.fields("PEDATEDEB").Value = DTPicker1.Value
  datPrimaryRS.Recordset.fields("PEDATEFIN").Value = DTPicker2.Value
   
  datPrimaryRS.Recordset.fields("PENBJOURMAX").Value = txtFields(1)
  datPrimaryRS.Recordset.fields("PENBJOURDC").Value = txtFields(2)
  
  ' type de période
  If cboTypeProvision.ListIndex <> -1 Then
    datPrimaryRS.Recordset.fields("PETYPEPERIODE").Value = cboTypeProvision.ItemData(cboTypeProvision.ListIndex)
  Else
    datPrimaryRS.Recordset.fields("PETYPEPERIODE").Value = 1
  End If
  
  ' type de période
  If cboTypeCalcul.ListIndex <> -1 Then
    datPrimaryRS.Recordset.fields("IdTypeCalcul").Value = cboTypeCalcul.ItemData(cboTypeCalcul.ListIndex)
  Else
    datPrimaryRS.Recordset.fields("IdTypeCalcul").Value = 1
  End If
  
  ' Activation P3I_INDIVIDUEL et Sécurité
  If DroitAdmin Then
    datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL") = IIf(P3I_Individuel(0) = vbChecked, True, False)

    datPrimaryRS.Recordset.fields("PELOCKED") = IIf(chkLocked(0) = vbChecked, True, False)
  End If
  
  'New Fields For BCAC Affiné
  datPrimaryRS.Recordset.fields("IsBCACNew").Value = IIf(radBCACOld.Value = True, False, True)
  datPrimaryRS.Recordset.fields("NumberOfMonth").Value = txtMonth.text
  
  On Error GoTo errcmdUpdate
  datPrimaryRS.Recordset.Update
  datPrimaryRS.Recordset.bookmark = curBookmark
  
  On Error GoTo 0
  
  Screen.MousePointer = vbDefault
  
  If datPrimaryRS.Recordset.fields("PENUMCLE") > 0 Then
    ' on doit sauvegarder avant d'editer les params de calcul
    btnAddPAram.Enabled = True
    btnEditParam.Enabled = True
    btnDelParam.Enabled = True
    btnDefaultParam.Enabled = True
    
    If fmAction = 0 Then
      ' chargement des paramètres de calcul par défaut
      DefaultParam False
    End If
  End If
  
  fmAction = 1
  
  ' Unload Me
  
  Exit Sub
  
errcmdUpdate:
  MsgBox "erreur :" & Err.Description
  On Error GoTo 0
End Sub

'##ModelId=5C8A67DD017A
Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub




'##ModelId=5C8A67DD0188
Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Errreur : " & ErrorNumber & vbLf & Description, vbExclamation
  'Response = 0  'Throw away the error
End Sub

'##ModelId=5C8A67DD0234
Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1)
End Sub

'##ModelId=5C8A67DD0254
Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
      Screen.MousePointer = vbDefault
  End Select
  Screen.MousePointer = vbHourglass
End Sub

'##ModelId=5C8A67DD0282
Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Dim nomTable As String
  Dim sel As Long
  
  Screen.MousePointer = vbHourglass
  
  ProgressBar1.Visible = False
  
  If archiveMode Then
    btnAddPAram.Enabled = False
    btnEditParam.Enabled = False
    btnDelParam.Enabled = False
    cmdUpdate.Enabled = False
    btnImportCATR9.Enabled = False
    btnDefaultParam.Enabled = False
  End If
 
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  m_dataSource.SetDatabase datPrimaryRS
  
  ' activate first tab
  SSTab1.Tab = 0
  
  ' ajoute un enregistrement si besoin
  If numPeriode = -1 Then
    
    btnProvision.Visible = False
    
    datPrimaryRS.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Periode")
    datPrimaryRS.Refresh
    
    datPrimaryRS.Recordset.AddNew
    
    datPrimaryRS.Recordset.fields("IsBCACNew").Value = True
    datPrimaryRS.Recordset.fields("NumberOfMonth").Value = 24
    
    radBCACNew.Value = True
    txtMonth.text = 24
    
    
    datPrimaryRS.Recordset.fields("PEGPECLE").Value = GroupeCle
    txtFields(1) = 180
    datPrimaryRS.Recordset.fields("PENBJOURMAX").Value = 180
    txtFields(2) = 30
    datPrimaryRS.Recordset.fields("PENBJOURDC").Value = 30
    
    txtFields(6) = "Nouvelle période..."

    ' numero de periode
    Set rs = m_dataSource.OpenRecordset("SELECT MAX(PENUMCLE) FROM Periode WHERE PEGPECLE = " & GroupeCle, Snapshot)
    If Not rs.EOF Then
      datPrimaryRS.Recordset.fields("PENUMCLE").Value = IIf(IsNull(rs.fields(0)), 1, rs.fields(0) + 1)
    Else
      datPrimaryRS.Recordset.fields("PENUMCLE").Value = 1
    End If
    rs.Close
     
    fmAction = 0 ' add new
    numPeriode = datPrimaryRS.Recordset.fields("PENUMCLE").Value
     
    ' date d'extraction
    datPrimaryRS.Recordset.fields("PEDATEEXT").Value = Format(Now(), "dd/mm/yyyy")
    txtFields(5) = Format(datPrimaryRS.Recordset.fields("PEDATEEXT").Value, "dd/mm/yyyy")
    
    datPrimaryRS.Recordset.fields("PEDATEDEB").Value = Format(DateSerial(Year(Now()) - 1, 1, 1), "dd/mm/yyyy")
    DTPicker1.Value = datPrimaryRS.Recordset.fields("PEDATEDEB").Value
    datPrimaryRS.Recordset.fields("PEDATEFIN").Value = Format(DateSerial(Year(Now()) - 1, 12, 31), "dd/mm/yyyy")
    DTPicker2.Value = datPrimaryRS.Recordset.fields("PEDATEFIN").Value
    
    'valeurs par defaut
    datPrimaryRS.Recordset.fields("PENBJOURMAX").Value = GetSettingIni(CompanyName, SectionName, "DelaiPriseEnCompte", "180")
    
    ' P3I_Individuel
    'P3I_Individuel = vbUnchecked
    datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL").Value = IIf(GetSettingIni(CompanyName, SectionName, "P3I_Individuel", "0") <> "0", True, False)
    P3I_Individuel(0).Value = IIf(CBool(datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL").Value) = True, vbChecked, vbUnchecked)

    ' Sécurité
    chkLocked(0) = vbUnchecked
    
    ' Type
    m_dataHelper.FillCombo cboTypeProvision, "SELECT IdTypePeriode, Libelle FROM TypePeriode", 1
    m_dataHelper.FillCombo cboTypeCalcul, "SELECT IdTypeCalcul, Libelle FROM TypeCalcul", 1
    
    ' on doit suavegarder avant d'editer les params de calcul
    btnAddPAram.Enabled = False
    btnEditParam.Enabled = False
    btnDelParam.Enabled = False
    btnDefaultParam.Enabled = False
    
  Else
    
    fmAction = 1 ' edit
    
    datPrimaryRS.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Periode WHERE PENUMCLE = " & numPeriode & " And PEGPECLE = " & GroupeCle)
    datPrimaryRS.Refresh
  
    If Not datPrimaryRS.Recordset.EOF Then
      txtFields(5) = Format(datPrimaryRS.Recordset.fields("PEDATEEXT").Value, "dd/mm/yyyy")
      DTPicker1.Value = datPrimaryRS.Recordset.fields("PEDATEDEB").Value
      DTPicker2.Value = datPrimaryRS.Recordset.fields("PEDATEFIN").Value
      
      txtFields(1) = datPrimaryRS.Recordset.fields("PENBJOURMAX").Value
      txtFields(2) = datPrimaryRS.Recordset.fields("PENBJOURDC").Value
    
      If Not IsNull(datPrimaryRS.Recordset.fields("PECOMMENTAIRE")) Then
        txtFields(6) = datPrimaryRS.Recordset.fields("PECOMMENTAIRE")
      Else
        txtFields(6) = ""
      End If

      ' P3I Individuel
      'P3I_Individuel(0) = vbUnchecked
      If IsNull(datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL").Value) Then
      datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL").Value = False
      End If
      P3I_Individuel(0).Value = IIf(CBool(datPrimaryRS.Recordset.fields("PEP3I_INDIVIDUEL").Value) = True, vbChecked, vbUnchecked)
      
      ' Sécurité
      chkLocked(0) = IIf(CBool(datPrimaryRS.Recordset.fields("PELOCKED").Value) = True, vbChecked, vbUnchecked)
    
      ' Type
      If IsNull(datPrimaryRS.Recordset.fields("PETYPEPERIODE").Value) Then
        m_dataHelper.FillCombo cboTypeProvision, "SELECT IdTypePeriode, Libelle FROM TypePeriode", 1
      Else
        m_dataHelper.FillCombo cboTypeProvision, "SELECT IdTypePeriode, Libelle FROM TypePeriode", datPrimaryRS.Recordset.fields("PETYPEPERIODE").Value
      End If
    
      If IsNull(datPrimaryRS.Recordset.fields("IdTypeCalcul").Value) Then
        m_dataHelper.FillCombo cboTypeCalcul, "SELECT IdTypeCalcul, Libelle FROM TypeCalcul", 1
      Else
        m_dataHelper.FillCombo cboTypeCalcul, "SELECT IdTypeCalcul, Libelle FROM TypeCalcul", datPrimaryRS.Recordset.fields("IdTypeCalcul").Value
      End If
      
    End If
    
  End If
  
  ' rempli le nom du groupe
  Set rs = m_dataSource.OpenRecordset("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle, Snapshot)
  If Not rs.EOF Then
    lblGroupe = "Période n° " & datPrimaryRS.Recordset.fields("PENUMCLE").Value & " du Groupe '" & rs.fields("Nom") & "'"
  End If
  rs.Close
  
  
  ' liste des tables diverses
  InitTableDiverse
  fmTableDiverse.FillCombo cboCATR9
  
  
  'New Fields For BCAC Affiné
  radBCACOld.Value = True
  txtMonth.text = 24
  
  If Not IsNull(datPrimaryRS.Recordset.fields("IsBCACNew")) Then
    If datPrimaryRS.Recordset.fields("IsBCACNew") = True Then
      radBCACNew.Value = True
    End If
  End If
  
  If Not IsNull(datPrimaryRS.Recordset.fields("NumberOfMonth")) Then
    txtMonth.text = datPrimaryRS.Recordset.fields("NumberOfMonth")
  End If
 
  
  ' Liste des paramètres de calcul
  RefreshListParamCalcul
  
  ' Sécurité
  chkLocked(0).Enabled = DroitAdmin
  
  Screen.MousePointer = vbDefault
  
End Sub

'##ModelId=5C8A67DD0292
Private Sub RefreshListParamCalcul()
  Dim curSel As Long
  
  ' sauvegarde la selection
  If lvParamCalcul.SelectedItem Is Nothing Then
    curSel = -1
  Else
    curSel = lvParamCalcul.SelectedItem.Index
  End If
  
  ' Liste des paramètres de calcul
  m_dataHelper.Affiche_Liste_Choix lvParamCalcul, "SELECT PENUMPARAMCALCUL, PENUMPARAMCALCUL, PENUMPARAMCALCUL as Code, PENOMPARAM as Nom" _
                                                  & " FROM ParamCalcul " _
                                                  & " WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode _
                                                  & " ORDER BY Code"
  
  
  LargeurAutomatique Me, lvParamCalcul
  lvParamCalcul.ColumnHeaders(1).Width = 0 ' caché (cette colonne ne paut pas être centrée)
  
  lvParamCalcul.ColumnHeaders(2).Alignment = lvwColumnCenter
  lvParamCalcul.ColumnHeaders(2).Width = 600
  
  lvParamCalcul.ColumnHeaders(3).Width = lvParamCalcul.ColumnHeaders(3).Width + 100 'lvParamCalcul.Width - 100 - lvParamCalcul.ColumnHeaders(2).Width
  
  ' restaure la selection
  If curSel <> -1 And curSel <= lvParamCalcul.ListItems.Count Then
    lvParamCalcul.ListItems(curSel).Selected = True
    lvParamCalcul.ListItems(curSel).EnsureVisible
  End If
  
End Sub

'##ModelId=5C8A67DD02B2
Private Sub Form_Unload(Cancel As Integer)
  Set fmTableDiverse = Nothing
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67DD02D1
Private Sub lvParamCalcul_DblClick()
  btnEditParam_Click
End Sub

'##ModelId=5C8A67DD02E0
Private Sub SSTab1_Click(PreviousTab As Integer)
  
  ' Tables Diverses
  If SSTab1.Tab = 1 Then
    fmTableDiverse.FillCombo cboCATR9
  End If

End Sub

