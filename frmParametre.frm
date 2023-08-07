VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmParametre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramêtres Généraux"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "frmParametre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   90
      TabIndex        =   27
      Top             =   8010
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Top             =   8010
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   13917
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "Incap / Inval"
      TabPicture(0)   =   "frmParametre.frx":1BB2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label41"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNbDecimalePM"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnSave"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtNbDecimaleCalcul"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Rentes"
      TabPicture(1)   =   "frmParametre.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "btnSave2"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Maintien Décès"
      TabPicture(2)   =   "frmParametre.frx":1BEA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame17"
      Tab(2).Control(2)=   "chkPMGDForcerInval"
      Tab(2).Control(3)=   "Frame16"
      Tab(2).Control(4)=   "Frame15"
      Tab(2).Control(5)=   "Frame14"
      Tab(2).Control(6)=   "Frame13"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Lois de maintien"
      TabPicture(3)   =   "frmParametre.frx":1C06
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "btnDelete"
      Tab(3).Control(1)=   "btnImportLoiMantien"
      Tab(3).Control(2)=   "btnExportLoiMantien"
      Tab(3).Control(3)=   "vaSpread1"
      Tab(3).Control(4)=   "dtaListeTable"
      Tab(3).Control(5)=   "btnPrint"
      Tab(3).Control(6)=   "cboListeTable"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Tables diverses"
      TabPicture(4)   =   "frmParametre.frx":1C22
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "btnExportCATR9"
      Tab(4).Control(1)=   "btnPrintCATR9"
      Tab(4).Control(2)=   "btnImportCATR9"
      Tab(4).Control(3)=   "cboCATR9"
      Tab(4).Control(4)=   "Frame11"
      Tab(4).Control(5)=   "sprCATR9"
      Tab(4).Control(6)=   "dtaCATR9"
      Tab(4).Control(7)=   "Label26"
      Tab(4).ControlCount=   8
      TabCaption(5)   =   "Taux de provisions"
      TabPicture(5)   =   "frmParametre.frx":1C3E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Risque Statutaire"
      TabPicture(6)   =   "frmParametre.frx":1C5A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label5(11)"
      Tab(6).Control(1)=   "Label1"
      Tab(6).Control(2)=   "Label4"
      Tab(6).Control(3)=   "Label13"
      Tab(6).Control(4)=   "Label14"
      Tab(6).Control(5)=   "Label43"
      Tab(6).Control(6)=   "Label44"
      Tab(6).Control(7)=   "Label45"
      Tab(6).Control(8)=   "Label46"
      Tab(6).Control(9)=   "cmdSaveStatParams"
      Tab(6).Control(10)=   "txtAgeMax"
      Tab(6).Control(11)=   "cmbAnneeBareme"
      Tab(6).Control(12)=   "txtMATMaxSemaine"
      Tab(6).Control(13)=   "txtCLMMaxSemaine"
      Tab(6).Control(14)=   "txtCLDMaxSemaine"
      Tab(6).Control(15)=   "txtAgeRet"
      Tab(6).Control(16)=   "txtATMaxSemaine"
      Tab(6).Control(17)=   "txtAgeMin"
      Tab(6).Control(18)=   "txtMOMaxSemaine"
      Tab(6).ControlCount=   19
      Begin VB.TextBox txtMOMaxSemaine 
         Height          =   285
         Left            =   -72840
         TabIndex        =   183
         Text            =   "4"
         Top             =   2040
         Width           =   465
      End
      Begin VB.TextBox txtAgeMin 
         Height          =   285
         Left            =   -72840
         TabIndex        =   182
         Text            =   "4"
         Top             =   1680
         Width           =   465
      End
      Begin VB.TextBox txtATMaxSemaine 
         Height          =   285
         Left            =   -72840
         TabIndex        =   181
         Text            =   "4"
         Top             =   3480
         Width           =   465
      End
      Begin VB.TextBox txtAgeRet 
         Height          =   285
         Left            =   -72840
         TabIndex        =   180
         Text            =   "4"
         Top             =   960
         Width           =   465
      End
      Begin VB.TextBox txtCLDMaxSemaine 
         Height          =   285
         Left            =   -72840
         TabIndex        =   179
         Text            =   "4"
         Top             =   3120
         Width           =   465
      End
      Begin VB.TextBox txtCLMMaxSemaine 
         Height          =   285
         Left            =   -72840
         TabIndex        =   178
         Text            =   "4"
         Top             =   2760
         Width           =   465
      End
      Begin VB.TextBox txtMATMaxSemaine 
         Height          =   285
         Left            =   -72840
         TabIndex        =   177
         Text            =   "4"
         Top             =   2400
         Width           =   465
      End
      Begin VB.ComboBox cmbAnneeBareme 
         Height          =   315
         Left            =   -72840
         Style           =   2  'Dropdown List
         TabIndex        =   176
         Top             =   600
         Width           =   1125
      End
      Begin VB.TextBox txtAgeMax 
         Height          =   285
         Left            =   -72840
         TabIndex        =   175
         Text            =   "4"
         Top             =   1320
         Width           =   465
      End
      Begin VB.CommandButton cmdSaveStatParams 
         Caption         =   "Enregistrer les paramètres des risques statutaires"
         Height          =   375
         Left            =   -68880
         TabIndex        =   173
         Top             =   7080
         Width           =   3975
      End
      Begin VB.CommandButton btnDelete 
         Caption         =   "&Supprimer"
         Height          =   330
         Left            =   -65775
         TabIndex        =   172
         Top             =   450
         Width           =   1185
      End
      Begin VB.CommandButton btnImportLoiMantien 
         Caption         =   "&Importer"
         Height          =   330
         Left            =   -68295
         TabIndex        =   171
         Top             =   450
         Width           =   1185
      End
      Begin VB.Frame Frame3 
         Caption         =   "Taux de Provisions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7350
         Left            =   -74865
         TabIndex        =   161
         Top             =   405
         Width           =   10275
         Begin VB.CommandButton btnSupprimerTauxProvision 
            Caption         =   "&Supprimer"
            Height          =   330
            Left            =   8865
            TabIndex        =   170
            Top             =   315
            Width           =   1185
         End
         Begin VB.CommandButton btnExportTauxProvision 
            Caption         =   "&Exporter"
            Height          =   330
            Left            =   7560
            TabIndex        =   169
            Top             =   315
            Width           =   1185
         End
         Begin VB.CommandButton btnPrintTaux 
            Caption         =   "&Imprimer"
            Height          =   330
            Left            =   6255
            TabIndex        =   168
            Top             =   315
            Width           =   1185
         End
         Begin VB.TextBox txtCommentProvision 
            Height          =   540
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   164
            Top             =   720
            Width           =   10005
         End
         Begin VB.ComboBox cboTableTauxProvision 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   315
            Width           =   6000
         End
         Begin FPSpreadADO.fpSpread vaSpread2 
            Bindings        =   "frmParametre.frx":1C76
            Height          =   5865
            Left            =   135
            TabIndex        =   162
            Top             =   1350
            Width           =   10005
            _Version        =   524288
            _ExtentX        =   17648
            _ExtentY        =   10345
            _StockProps     =   64
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
            OperationMode   =   3
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmParametre.frx":1C91
            AppearanceStyle =   0
         End
         Begin MSAdodcLib.Adodc dtaProvision 
            Height          =   330
            Left            =   7065
            Top             =   720
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
            Caption         =   "dtaProvision"
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
      End
      Begin VB.Frame Frame12 
         Caption         =   "Coefficients de provisions précalculés"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3795
         Left            =   -74865
         TabIndex        =   153
         Top             =   3915
         Width           =   10275
         Begin VB.CommandButton btnExportBCAC 
            Caption         =   "&Exporter"
            Height          =   330
            Left            =   8955
            TabIndex        =   158
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton btnImportBCAC 
            Caption         =   "&Importer"
            Height          =   330
            Left            =   7695
            TabIndex        =   157
            Top             =   270
            Width           =   1185
         End
         Begin VB.ComboBox cboBCAC 
            Height          =   315
            Left            =   765
            Style           =   2  'Dropdown List
            TabIndex        =   156
            Top             =   270
            Width           =   4335
         End
         Begin VB.CommandButton btnPrintBCAC 
            Caption         =   "&Imprimer"
            Height          =   330
            Left            =   5175
            TabIndex        =   155
            Top             =   270
            Width           =   1185
         End
         Begin VB.CommandButton btnDelBCAC 
            Caption         =   "&Supprimer"
            Height          =   330
            Left            =   6435
            TabIndex        =   154
            Top             =   270
            Width           =   1185
         End
         Begin FPSpreadADO.fpSpread sprBCAC 
            Height          =   2985
            Left            =   90
            TabIndex        =   159
            Top             =   675
            Width           =   10050
            _Version        =   524288
            _ExtentX        =   17727
            _ExtentY        =   5265
            _StockProps     =   64
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
            OperationMode   =   3
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmParametre.frx":20CD
            AppearanceStyle =   0
         End
         Begin VB.Label Label27 
            Caption         =   "Tables :"
            Height          =   240
            Left            =   135
            TabIndex        =   160
            Top             =   315
            Width           =   600
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Jeux de Paramètres de calcul pour les nouvelles études"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   405
         TabIndex        =   148
         Top             =   4725
         Width           =   9690
         Begin VB.OptionButton rdoBilan 
            Caption         =   "Bilan"
            Height          =   285
            Left            =   225
            TabIndex        =   167
            Top             =   360
            Width           =   1140
         End
         Begin VB.OptionButton rdoSimul 
            Caption         =   "Simulation"
            Height          =   285
            Left            =   2925
            TabIndex        =   166
            Top             =   360
            Width           =   1140
         End
         Begin VB.OptionButton rdoClient 
            Caption         =   "Client"
            Height          =   285
            Left            =   1575
            TabIndex        =   165
            Top             =   360
            Width           =   1005
         End
         Begin VB.CommandButton btnAddPAram 
            Caption         =   "&Ajouter"
            Height          =   330
            Left            =   7920
            TabIndex        =   151
            Top             =   945
            Width           =   1140
         End
         Begin VB.CommandButton btnEditParam 
            Caption         =   "&Modifier"
            Height          =   330
            Left            =   7920
            TabIndex        =   150
            Top             =   1395
            Width           =   1140
         End
         Begin VB.CommandButton btnDelParam 
            Caption         =   "&Supprimer"
            Height          =   330
            Left            =   7920
            TabIndex        =   149
            Top             =   1845
            Width           =   1140
         End
         Begin MSComctlLib.ListView lvParamCalcul 
            Height          =   1545
            Left            =   180
            TabIndex        =   152
            Top             =   810
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
      Begin VB.Frame Frame5 
         Caption         =   "Reprise de revalorisation sur les rentes d'invalidités"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -69465
         TabIndex        =   135
         Top             =   630
         Width           =   4695
         Begin VB.TextBox txtDureeIndex 
            Height          =   285
            Left            =   1710
            TabIndex        =   138
            Top             =   1260
            Width           =   675
         End
         Begin VB.TextBox txtTxIndex 
            Height          =   285
            Left            =   1710
            TabIndex        =   137
            Top             =   360
            Width           =   675
         End
         Begin VB.TextBox txtTMO 
            Height          =   285
            Left            =   1710
            TabIndex        =   136
            Top             =   810
            Width           =   675
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   2
            Left            =   2385
            TabIndex        =   144
            Top             =   1260
            Width           =   285
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   1
            Left            =   2430
            TabIndex        =   143
            Top             =   810
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   7
            Left            =   2430
            TabIndex        =   142
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblLabels 
            Caption         =   "Durée d'indexation"
            Height          =   255
            Index           =   2
            Left            =   225
            TabIndex        =   141
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux d'indexation"
            Height          =   255
            Index           =   16
            Left            =   225
            TabIndex        =   140
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux financier TME"
            Height          =   255
            Index           =   17
            Left            =   225
            TabIndex        =   139
            Top             =   810
            Width           =   1500
         End
      End
      Begin VB.CommandButton btnExportLoiMantien 
         Caption         =   "&Exporter"
         Height          =   330
         Left            =   -67035
         TabIndex        =   19
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtNbDecimaleCalcul 
         Height          =   285
         Left            =   4275
         TabIndex        =   1
         Text            =   "=(NbDecimaleCalcul)"
         Top             =   540
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Frame Frame17 
         Caption         =   "Méthode de calcul des provisions maintien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -74865
         TabIndex        =   128
         Top             =   3150
         Width           =   5100
         Begin VB.OptionButton rdoCotisationsExonerees 
            Caption         =   "Prime Exonérée"
            Height          =   240
            Left            =   3150
            TabIndex        =   130
            Top             =   270
            Width           =   1860
         End
         Begin VB.OptionButton rdoCapitauxConstitif 
            Caption         =   "Capitaux constitutif sous risque"
            Height          =   240
            Left            =   315
            TabIndex        =   129
            Top             =   270
            Width           =   2715
         End
      End
      Begin VB.CheckBox chkPMGDForcerInval 
         Caption         =   "Provisions Maintien en Garantie Décès : forcer en Invalidité les assurés en Incapacité depuis plus de 36 mois sans passage"
         Height          =   420
         Left            =   -69510
         TabIndex        =   127
         Top             =   3465
         Visible         =   0   'False
         Width           =   4740
      End
      Begin VB.Frame Frame16 
         Caption         =   "Coefficients de provisions maintien"
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3480
         Left            =   -69690
         TabIndex        =   98
         Top             =   360
         Width           =   5100
         Begin VB.TextBox txtAgeLimiteCalulDC 
            Height          =   330
            Left            =   4230
            TabIndex        =   132
            Text            =   "0"
            Top             =   630
            Width           =   465
         End
         Begin VB.CommandButton btnCalclInvalBCAC 
            Caption         =   "Calculer Coeff Inval"
            Height          =   375
            Left            =   2655
            TabIndex        =   103
            Top             =   1845
            Width           =   1965
         End
         Begin VB.CommandButton btnCalclIncapBCAC 
            Caption         =   "Calculer Coeff Incap"
            Height          =   375
            Left            =   585
            TabIndex        =   102
            Top             =   1845
            Width           =   1965
         End
         Begin VB.CommandButton btnSaveDC 
            Caption         =   "&Enregistrer les paramètres"
            Height          =   375
            Left            =   3015
            TabIndex        =   117
            Top             =   180
            Width           =   1965
         End
         Begin VB.OptionButton rdoCalculCoeffBCAC 
            Caption         =   "Calcul du coefficient"
            Height          =   240
            Left            =   315
            TabIndex        =   108
            Top             =   360
            Width           =   1860
         End
         Begin VB.OptionButton rdoLireCoeffBCAC 
            Caption         =   "Utiliser les tables de coefficients précalculés du BCAC"
            Height          =   240
            Left            =   315
            TabIndex        =   107
            Top             =   2340
            Width           =   4605
         End
         Begin VB.ComboBox cboTableIncapCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   106
            Top             =   1035
            Width           =   3525
         End
         Begin VB.TextBox txtFraisGestionCalculDC 
            Height          =   330
            Left            =   2970
            TabIndex        =   105
            Text            =   "0"
            Top             =   630
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechnicCalculDC 
            Height          =   330
            Left            =   1080
            TabIndex        =   104
            Text            =   "4.5"
            Top             =   630
            Width           =   465
         End
         Begin VB.ComboBox cboTableInvalCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   1395
            Width           =   3525
         End
         Begin VB.ComboBox cboTableIncapPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   2655
            Width           =   3525
         End
         Begin VB.ComboBox cboTableInvalPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   3015
            Width           =   3525
         End
         Begin VB.Label Label5 
            Caption         =   "Age limite"
            Height          =   375
            Index           =   10
            Left            =   3825
            TabIndex        =   134
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label42 
            Caption         =   "ans"
            Height          =   240
            Left            =   4725
            TabIndex        =   133
            Top             =   675
            Width           =   330
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   420
            Index           =   9
            Left            =   315
            TabIndex        =   116
            Top             =   585
            Width           =   780
         End
         Begin VB.Label Label36 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   1800
            TabIndex        =   115
            Top             =   675
            Width           =   1140
         End
         Begin VB.Label Label35 
            Caption         =   "%"
            Height          =   240
            Left            =   1575
            TabIndex        =   114
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label34 
            Caption         =   "%"
            Height          =   240
            Left            =   3465
            TabIndex        =   113
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Incap"
            Height          =   240
            Index           =   8
            Left            =   315
            TabIndex        =   112
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Inval"
            Height          =   240
            Index           =   5
            Left            =   315
            TabIndex        =   111
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Table Incap"
            Height          =   375
            Index           =   6
            Left            =   495
            TabIndex        =   110
            Top             =   2700
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Table Inval"
            Height          =   375
            Index           =   7
            Left            =   495
            TabIndex        =   109
            Top             =   3060
            Width           =   960
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Pourcentage de provisions maintien à constituer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74865
         TabIndex        =   95
         Top             =   2475
         Width           =   5100
         Begin VB.OptionButton rdoAvecLissage 
            Caption         =   "Utiliser la table 'LissageProvision'"
            Height          =   240
            Left            =   1575
            TabIndex        =   97
            Top             =   270
            Width           =   3120
         End
         Begin VB.OptionButton rdoSansLissage 
            Caption         =   "100%"
            Height          =   240
            Left            =   315
            TabIndex        =   96
            Top             =   270
            Width           =   1275
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Rente de Conjoint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   -74865
         TabIndex        =   91
         Top             =   1125
         Width           =   5100
         Begin VB.TextBox txtAgeConjointRenteConjointDC 
            Height          =   330
            Left            =   4095
            TabIndex        =   124
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointViagereDC 
            Height          =   330
            Left            =   3960
            TabIndex        =   121
            Text            =   "4.5"
            Top             =   810
            Width           =   915
         End
         Begin VB.CheckBox chkForcerCapitalMoyenRteConjoitDC 
            Caption         =   "Forcer"
            Height          =   240
            Left            =   315
            TabIndex        =   119
            Top             =   855
            Width           =   780
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointTempoDC 
            Height          =   330
            Left            =   2205
            TabIndex        =   118
            Text            =   "4.5"
            Top             =   810
            Width           =   915
         End
         Begin VB.TextBox txtFraisGestionRenteConjointDC 
            Height          =   330
            Left            =   1665
            TabIndex        =   92
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label37 
            Caption         =   "Age du conjoint : +/-"
            Height          =   375
            Left            =   2565
            TabIndex        =   126
            Top             =   315
            Width           =   1500
         End
         Begin VB.Label Label38 
            Caption         =   "ans"
            Height          =   240
            Left            =   4590
            TabIndex        =   125
            Top             =   315
            Width           =   330
         End
         Begin VB.Label Label40 
            Caption         =   "Viagère"
            Height          =   195
            Left            =   3240
            TabIndex        =   123
            Top             =   855
            Width           =   600
         End
         Begin VB.Label Label39 
            Caption         =   "Temporaire"
            Height          =   195
            Left            =   1305
            TabIndex        =   122
            Top             =   855
            Width           =   825
         End
         Begin VB.Label Label5 
            Caption         =   "Capital constitutif moyen"
            Height          =   240
            Index           =   4
            Left            =   135
            TabIndex        =   120
            Top             =   630
            Width           =   1770
         End
         Begin VB.Label Label33 
            Caption         =   "Frais de gestion"
            Height          =   195
            Left            =   135
            TabIndex        =   94
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label32 
            Caption         =   "%"
            Height          =   240
            Left            =   2205
            TabIndex        =   93
            Top             =   315
            Width           =   240
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frais de gestion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -74865
         TabIndex        =   84
         Top             =   360
         Width           =   5100
         Begin VB.TextBox txtFraisGestionRenteEducationDC 
            Height          =   330
            Left            =   4050
            TabIndex        =   86
            Text            =   "0"
            Top             =   225
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionCapitauxDecesDC 
            Height          =   330
            Left            =   1485
            TabIndex        =   85
            Text            =   "0"
            Top             =   225
            Width           =   465
         End
         Begin VB.Label Label31 
            Caption         =   "Rente Education"
            Height          =   375
            Left            =   2700
            TabIndex        =   90
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label30 
            Caption         =   "%"
            Height          =   240
            Left            =   4590
            TabIndex        =   89
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label29 
            Caption         =   "Capitaux Décès"
            Height          =   375
            Left            =   135
            TabIndex        =   88
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label28 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   87
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.CommandButton btnExportCATR9 
         Caption         =   "&Exporter"
         Height          =   330
         Left            =   -68295
         TabIndex        =   82
         Top             =   810
         Width           =   1185
      End
      Begin VB.CommandButton btnPrintCATR9 
         Caption         =   "&Imprimer"
         Height          =   330
         Left            =   -71265
         TabIndex        =   81
         Top             =   810
         Width           =   1455
      End
      Begin VB.CommandButton btnImportCATR9 
         Caption         =   "&Importer"
         Height          =   330
         Left            =   -69555
         TabIndex        =   80
         Top             =   810
         Width           =   1185
      End
      Begin VB.ComboBox cboCATR9 
         Height          =   315
         Left            =   -74820
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   810
         Width           =   3480
      End
      Begin VB.Frame Frame11 
         Caption         =   "Contrôle de l'unicité d'une catégorie entre CATR9 et CATR9INVAL"
         Height          =   1950
         Left            =   -74910
         TabIndex        =   74
         Top             =   5850
         Visible         =   0   'False
         Width           =   5235
         Begin VB.TextBox txtCATR9NonUnique 
            Height          =   1230
            Left            =   135
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   540
            Width           =   2850
         End
         Begin VB.CommandButton btnCATR9Unique 
            Caption         =   "Contrôle de l'unicité"
            Height          =   330
            Left            =   3240
            TabIndex        =   75
            Top             =   945
            Width           =   1770
         End
         Begin VB.Label Label22 
            Caption         =   "Catégories présentes dans les 2 tables :"
            Height          =   240
            Left            =   180
            TabIndex        =   76
            Top             =   270
            Width           =   3075
         End
      End
      Begin FPSpreadADO.fpSpread vaSpread1 
         Bindings        =   "frmParametre.frx":2537
         Height          =   6855
         Left            =   -74865
         TabIndex        =   73
         Top             =   900
         Width           =   10230
         _Version        =   524288
         _ExtentX        =   18045
         _ExtentY        =   12091
         _StockProps     =   64
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
         OperationMode   =   3
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmParametre.frx":2553
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaListeTable 
         Height          =   330
         Left            =   -66945
         Top             =   495
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
         Caption         =   "dtaListeTable"
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
      Begin VB.CommandButton btnSave2 
         Caption         =   "&Enregistrer les paramètres pour les Rentes"
         Height          =   375
         Left            =   -68160
         TabIndex        =   72
         Top             =   7110
         Width           =   3135
      End
      Begin VB.Frame Frame6 
         Caption         =   "Paramètres par défaut si Catégorie/Survenance absents"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -74685
         TabIndex        =   51
         Top             =   630
         Width           =   5055
         Begin VB.Frame Frame10 
            Caption         =   "Fractionnement"
            ClipControls    =   0   'False
            Height          =   1095
            Left            =   2340
            TabIndex        =   67
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton rdoSemestrielEducation 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   71
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelEducation 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   70
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelEducation 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   69
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoTrimestrielEducation 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   68
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Paiement"
            ClipControls    =   0   'False
            Height          =   1095
            Left            =   135
            TabIndex        =   64
            Top             =   1350
            Width           =   1725
            Begin VB.OptionButton rdoPaiementAvanceEducation 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   66
               Top             =   360
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementEchuEducation 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   65
               Top             =   720
               Width           =   1275
            End
         End
         Begin VB.TextBox txtTauxTechniqueRenteEducation 
            Height          =   330
            Left            =   1395
            TabIndex        =   24
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionRenteEducation 
            Height          =   330
            Left            =   4230
            TabIndex        =   25
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.ComboBox cboTableRenteEducation 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   900
            Width           =   4785
         End
         Begin VB.Label Label19 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   55
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label20 
            Caption         =   "%"
            Height          =   240
            Left            =   1935
            TabIndex        =   54
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label21 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   53
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   3
            Left            =   135
            TabIndex        =   52
            Top             =   405
            Width           =   1185
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Rente de Conjoint (N'EST PLUS UTILISER)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2625
         Left            =   -74685
         TabIndex        =   46
         Top             =   3465
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame Frame8 
            Caption         =   "Fractionnement"
            ClipControls    =   0   'False
            Height          =   1140
            Left            =   2340
            TabIndex        =   59
            Top             =   1305
            Width           =   2580
            Begin VB.OptionButton rdoTrimestrielConjoint 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   63
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelConjoint 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   62
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelConjoint 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   61
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoSemestrielConjoint 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   60
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Paiement"
            ClipControls    =   0   'False
            Height          =   1140
            Left            =   135
            TabIndex        =   56
            Top             =   1305
            Width           =   1725
            Begin VB.OptionButton rdoPaiementEchuConjoint 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   58
               Top             =   720
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementAvanceConjoint 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   57
               Top             =   360
               Width           =   1275
            End
         End
         Begin VB.TextBox txtTauxTechniqueRenteConjoint 
            Height          =   330
            Left            =   1395
            TabIndex        =   21
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionRenteConjoint 
            Height          =   330
            Left            =   4230
            TabIndex        =   22
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.ComboBox cboTableRenteConjoint 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   900
            Width           =   4785
         End
         Begin VB.Label Label25 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   50
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label24 
            Caption         =   "%"
            Height          =   240
            Left            =   1935
            TabIndex        =   49
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   48
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   2
            Left            =   135
            TabIndex        =   47
            Top             =   405
            Width           =   1185
         End
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "&Enregistrer les paramètres de calculs"
         Height          =   375
         Left            =   6930
         TabIndex        =   44
         Top             =   7320
         Width           =   3135
      End
      Begin VB.TextBox txtNbDecimalePM 
         Height          =   285
         Left            =   9315
         TabIndex        =   2
         Text            =   "=(NbDecimalePM)"
         Top             =   540
         Width           =   510
      End
      Begin VB.Frame Frame1 
         Caption         =   "Loi de maintien en Incapacité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   360
         TabIndex        =   34
         Top             =   945
         Width           =   9690
         Begin VB.ComboBox cboTablePassage 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1395
            Width           =   3525
         End
         Begin VB.ComboBox cboTableIncap 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   990
            Width           =   3525
         End
         Begin VB.TextBox Txtduree 
            Height          =   285
            Left            =   4095
            TabIndex        =   3
            Text            =   "36"
            Top             =   225
            Width           =   465
         End
         Begin VB.TextBox TxtFranchise 
            Height          =   285
            Left            =   4095
            TabIndex        =   4
            Text            =   "0"
            Top             =   585
            Width           =   465
         End
         Begin VB.CommandButton btnCalcIncapInval 
            Caption         =   "&Calculer les provisions de Passage"
            Height          =   510
            Left            =   6075
            TabIndex        =   10
            Top             =   1080
            Width           =   3345
         End
         Begin VB.TextBox txtFraisGestionIncap 
            Height          =   285
            Left            =   1530
            TabIndex        =   6
            Text            =   "0"
            Top             =   630
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueIncap 
            Height          =   285
            Left            =   1530
            TabIndex        =   5
            Text            =   "4.5"
            Top             =   270
            Width           =   465
         End
         Begin VB.CommandButton btnCalcIncap 
            Caption         =   "&Calculer les provisions Incapacite"
            Height          =   510
            Left            =   6075
            TabIndex        =   9
            Top             =   360
            Width           =   3345
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table de passage"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   146
            Top             =   1440
            Width           =   1275
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table incapacité"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   145
            Top             =   1035
            Width           =   1230
         End
         Begin VB.Label Label3 
            Caption         =   "Durée"
            Height          =   240
            Left            =   3330
            TabIndex        =   42
            Top             =   270
            Width           =   510
         End
         Begin VB.Label Label9 
            Caption         =   "mois"
            Height          =   285
            Left            =   4635
            TabIndex        =   41
            Top             =   270
            Width           =   330
         End
         Begin VB.Label Label10 
            Caption         =   "Franchise"
            Height          =   195
            Left            =   3330
            TabIndex        =   40
            Top             =   630
            Width           =   690
         End
         Begin VB.Label Label11 
            Caption         =   "mois"
            Height          =   240
            Left            =   4635
            TabIndex        =   39
            Top             =   630
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   38
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label6 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   225
            TabIndex        =   37
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label Label7 
            Caption         =   "%"
            Height          =   240
            Left            =   2070
            TabIndex        =   36
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label8 
            Caption         =   "%"
            Height          =   240
            Left            =   2070
            TabIndex        =   35
            Top             =   675
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Loi de maintien en Invalidité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   405
         TabIndex        =   28
         Top             =   3060
         Width           =   9690
         Begin VB.ComboBox cboTableInval 
            Height          =   315
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   990
            Width           =   3525
         End
         Begin VB.CommandButton btnCalcInval 
            Caption         =   "&Calculer les provisions Invalidité"
            Height          =   510
            Left            =   6030
            TabIndex        =   15
            Top             =   495
            Width           =   3345
         End
         Begin VB.TextBox txtFraisGestionInval 
            Height          =   285
            Left            =   1530
            TabIndex        =   13
            Text            =   "0"
            Top             =   630
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueInval 
            Height          =   285
            Left            =   1530
            TabIndex        =   12
            Text            =   "4.5"
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtRetraite 
            Height          =   285
            Left            =   4635
            TabIndex        =   11
            Text            =   "65"
            Top             =   225
            Width           =   375
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table invalidité"
            Height          =   255
            Index           =   11
            Left            =   180
            TabIndex        =   147
            Top             =   1035
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   3
            Left            =   5040
            TabIndex        =   45
            Top             =   270
            Width           =   285
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   33
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label12 
            Caption         =   "Frais de gestion"
            Height          =   195
            Left            =   180
            TabIndex        =   32
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "%"
            Height          =   240
            Left            =   2070
            TabIndex        =   31
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label17 
            Caption         =   "%"
            Height          =   240
            Left            =   2070
            TabIndex        =   30
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label16 
            Caption         =   "Age de départ à la retraite"
            Height          =   240
            Left            =   2565
            TabIndex        =   29
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "Im&primer"
         Height          =   330
         Left            =   -69825
         TabIndex        =   18
         Top             =   450
         Width           =   1455
      End
      Begin VB.ComboBox cboListeTable 
         Height          =   315
         Left            =   -74865
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   4875
      End
      Begin FPSpreadADO.fpSpread sprCATR9 
         Bindings        =   "frmParametre.frx":298F
         Height          =   6495
         Left            =   -74865
         TabIndex        =   77
         Top             =   1260
         Width           =   10230
         _Version        =   524288
         _ExtentX        =   18045
         _ExtentY        =   11456
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
         SpreadDesigner  =   "frmParametre.frx":29A6
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaCATR9 
         Height          =   330
         Left            =   -67080
         Top             =   810
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
      Begin VB.Label Label46 
         Caption         =   "AT Max Semaine"
         Height          =   255
         Left            =   -74640
         TabIndex        =   191
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label45 
         Caption         =   "Age assuré maximum"
         Height          =   255
         Left            =   -74640
         TabIndex        =   190
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "Age retraite"
         Height          =   255
         Left            =   -74640
         TabIndex        =   189
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label43 
         Caption         =   "CLD Max Semaine"
         Height          =   255
         Left            =   -74640
         TabIndex        =   188
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "CLM Max Semaine"
         Height          =   255
         Left            =   -74640
         TabIndex        =   187
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "MAT Max Semaine"
         Height          =   255
         Left            =   -74640
         TabIndex        =   186
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "MO Max Semaine"
         Height          =   255
         Left            =   -74640
         TabIndex        =   185
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Age assuré minimum"
         Height          =   255
         Left            =   -74640
         TabIndex        =   184
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Années des barèmes"
         Height          =   195
         Index           =   11
         Left            =   -74640
         TabIndex        =   174
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label41 
         Caption         =   "Nombre de décimales pour les calculs intermédiares"
         Height          =   285
         Left            =   540
         TabIndex        =   131
         Top             =   585
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Label Label26 
         Caption         =   "Tables à afficher :"
         Height          =   240
         Left            =   -74820
         TabIndex        =   79
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Nombre de décimales pour les Provisions calculées"
         Height          =   285
         Left            =   5580
         TabIndex        =   43
         Top             =   585
         Width           =   3705
      End
      Begin VB.Label Label2 
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   1155
         Width           =   2085
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10215
      Top             =   7965
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
End
Attribute VB_Name = "frmParametre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67B102AF"
Option Explicit

'##ModelId=5C8A67B1037A
Private fmTableDiverse As clsTablesDiverses
'##ModelId=5C8A67B1037B
Private archiveOk As Boolean

'##ModelId=5C8A67B1039A
Private Sub InitTableDiverse()
  If fmTableDiverse Is Nothing Then
    Set fmTableDiverse = New clsTablesDiverses
  Else
    fmTableDiverse.Clear
  End If
  
  fmTableDiverse.AddTableDiverse "AgeSituationFamille", "Age", "Not IsNull(Age)", "Age, IsCadre, CVD0, M0, CVD1, M1, NombreEnfant, AgeMoyenEnfant"
          
  'fmTableDiverse.AddTableDiverse "CATR9", "Categorie", "Categorie"
        
  'fmTableDiverse.AddTableDiverse "CATR9INVAL", "Categorie", "Categorie", "Categorie"
  
  fmTableDiverse.AddTableDiverse "CodePosition", "Position", "Position", "Position, Libelle"
  
  fmTableDiverse.AddTableDiverse "CodeProvision", "CodeProv", "CodeProv", "CodeProv, Libelle"
      
  fmTableDiverse.AddTableDiverse "CorrespondanceGarantie", "RegimeIncap, CategorieIncap", "RegimeIncap, CategorieIncap", "RegimeIncap,CategorieIncap,RegimeDeces, CategorieDeces, RegimeRenteConjointTemporaire,CategorieRenteConjointTemporaire,RegimeRenteConjointViagere,CategorieRenteConjointViagere,RegimeRenteEduc,CategorieRenteEduc,LissageProvision"
  
  fmTableDiverse.AddTableDiverse "GarantieDC", "Regime, Categorie", "Regime, Categorie", "Regime, Categorie,  CVD0, M0, CVD1, M1, MajorationEnfant, MajorationAccident"
      
  fmTableDiverse.AddTableDiverse "IndemnisationIncapInval", "Regime, Categorie", "Regime, Categorie", "Regime, Categorie, PourcentIndemnisation"
      
  fmTableDiverse.AddTableDiverse "LissageProvision", "Annee", "Annee", "Annee, Pourcentage"
  
  'fmTableDiverse.AddTableDiverse "PassageNCA", "NCA", "NCA"
      
  fmTableDiverse.AddTableDiverse "PlafondSS", "Annee", "Annee", "Annee, Montant"

  fmTableDiverse.AddTableDiverse "PSAP_Baremes", "Garantie, DebutPeriode", "Garantie, DebutPeriode", "Garantie, DebutPeriode, FinPeriode, Taux"

  'fmTableDiverse.AddTableDiverse "Reassurance", "Regime, Categorie, NCA", "Regime, Categorie"
      
  fmTableDiverse.AddTableDiverse "RisqueDeces", "Libelle", "Code", "Code,Libelle,AnnualisationZero"
      
  fmTableDiverse.AddTableDiverse "SituationFamille", "Libelle", "CleSituationFamille", "CleSituationFamille, Libelle"
      
  fmTableDiverse.AddTableDiverse "TauxRenteConjoint", "Regime, Categorie", "Regime, Categorie", "Regime,Categorie,AgeTerme,Taux_65x_x25,Taux_k,CapitalMoyen"
  
  fmTableDiverse.AddTableDiverse "TauxRenteEducation", "Categorie", "Categorie", "Categorie, B1, B2, B3, B4, B1_PCT, B2_PCT, B3_PCT, B4_PCT"
  
  fmTableDiverse.AddTableDiverse "AgeSituationFamille", "Age", "Not IsNull(Age)", "Age, IsCadre, CVD0, M0, CVD1, M1, NombreEnfant, AgeMoyenEnfant"
  
  'fmTableDiverse.AddTableDiverse "Statutaire", "StatId", "StatId", "StatId, TableAnnee, Collect, TypeSisnistre, Sexe, AgeMalade, Semaine, PM_AT, PM_CLD, PM_CLM, PM_CLM_CLD, PM_MO, PM_MO_CLM, PM_MO_CLD, PM_Total"
  
  fmTableDiverse.AddTableDiverse "Statutaire_Garantie", "ID", "ID", "ID, Garantie"
  
  fmTableDiverse.AddTableDiverse "Statutaire_Garantie_Code", "IDGarantie", "ID", "ID, IDGarantie, GarantieCode"
  
  fmTableDiverse.AddTableDiverse "Statutaire_Categorie", "Categorie", "", "Categorie, Description, TypeCollective"
  
  
End Sub

'##ModelId=5C8A67B103A9
Private Sub btnCalcIncap_Click()
  
  Dim x As Integer, Anciennete As Integer, Duree As Integer, franchise As Integer, xdepartRetraite As Integer
  Dim iIncap As Double, fraisGestionIncap As Double
  Dim iInval As Double, fraisGestionInval As Double
  
  Dim garantie As String, feuilleLueIncap As String, feuilleLue As String, feuilleLueInval As String, rq As String
  
  Dim rs As ADODB.Recordset
  
  Screen.MousePointer = vbHourglass
  
  '****************************************************
  '* paramètres de la fonction de calcul INCAPACITE   *
  '****************************************************
  
  garantie = "Incapacité"
  x = 0                     ' âge           n'est pas passé en paramètres
  Anciennete = 0            ' ancienneté    n'est pas passé en paramètres
  Duree = Txtduree          ' durée en mois dans l'état d'incapacité
  franchise = TxtFranchise  ' franchise en mois
  
  ' recupere le nom de la table
  If cboTableIncap.ListIndex = -1 Then
    MsgBox "Vous devez choisir une table Loi de maintien Incapacité !"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncap.ItemData(cboTableIncap.ListIndex)
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF Then
    feuilleLue = rs.fields("NOMTABLE")
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Table 'loi de maintien incapacité' INVALIDE !"
    rs.Close
    Exit Sub
  End If
  
  rs.Close
  
  iIncap = m_dataHelper.GetDouble2(txtTauxTechniqueIncap) / 100           ' taux technique
  fraisGestionIncap = m_dataHelper.GetDouble2(txtFraisGestionIncap) / 100 ' frais de gestion
  
  iInval = 0
  fraisGestionInval = 0
  
  xdepartRetraite = txtRetraite              ' âge de départ en retraite

  
  feuilleLueIncap = " "                      ' utilisée pour calculer les provisions de passage Incap/Inval
  feuilleLueInval = " "                      ' utilisée pour calculer les provisions de passage Incap/Inval
   
  ' charge les modules de calcul
  Dim ModuleCalcul_Provisions As clsCalcul_Provisions
  Dim recordsetList As clsRecordsetList
  
  Set recordsetList = New clsRecordsetList
  Set ModuleCalcul_Provisions = New clsCalcul_Provisions

  Set ModuleCalcul_Provisions.recordsetList = recordsetList
  
  ' lance la sub de calcul qui remplie la table destination
  Call ModuleCalcul_Provisions.CalcTableauProv(x, Anciennete, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap, iInval, fraisGestionInval, xdepartRetraite, garantie, ProgressBar1, feuilleLueIncap, feuilleLueInval)
 
  Set ModuleCalcul_Provisions = Nothing
  Set recordsetList = Nothing
 
  Screen.MousePointer = vbDefault

End Sub

  '*******************************************************************
  '* paramètres de la fonction de calcul PASSAGE INCAPACITE / INVAL  *
  '*******************************************************************

'##ModelId=5C8A67B103B9
Private Sub btnCalcIncapInval_Click()

  Dim x As Integer, Anciennete As Integer, Duree As Integer, franchise As Integer, xdepartRetraite As Integer
  Dim iIncap As Double, fraisGestionIncap As Double
  Dim iInval As Double, fraisGestionInval As Double
  
  Dim garantie As String, feuilleLueIncap As String, feuilleLue As String, feuilleLueInval As String, rq As String
    
  Dim rs As ADODB.Recordset
  
  Screen.MousePointer = vbHourglass
  
  
  garantie = "Passage"
  x = 0                     ' âge           n'est pas passé en paramètres
  Anciennete = 0            ' ancienneté    n'est pas passé en paramètres
  Duree = Txtduree          ' durée en mois dans l'état d'incapacité
  franchise = TxtFranchise  ' franchise en mois
  

  
  ' recupere le nom de la table incapacité
  If cboTableIncap.ListIndex = -1 Then
    MsgBox "Vous devez choisir une table Loi de maintien Incapacité !"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncap.ItemData(cboTableIncap.ListIndex)
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF Then
    feuilleLueIncap = rs.fields("NOMTABLE")
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Table 'loi de maintien incapacité' INVALIDE !"
    rs.Close
    Exit Sub
  End If
  rs.Close
  
  
  ' recupere le nom de la table PASSAGE incapacité / invalidité
  If cboTablePassage.ListIndex = -1 Then
    MsgBox "Vous devez choisir une table Loi de maintien Passage Incapacité/Invalidité !"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTablePassage.ItemData(cboTablePassage.ListIndex)
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF Then
    feuilleLue = rs.fields("NOMTABLE")
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Table 'loi de Passage Incapacité/Invalidité' INVALIDE !"
    rs.Close
    Exit Sub
  End If
  rs.Close
  
  
  ' recupere le nom de la table invalidité
  If cboTableInval.ListIndex = -1 Then
    MsgBox "Vous devez choisir une Loi de maintien Invalidité !"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInval.ItemData(cboTableInval.ListIndex)
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF Then
    feuilleLueInval = rs.fields("NOMTABLE")
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Table 'loi de maintien Invalidité' INVALIDE !"
    rs.Close
    Exit Sub
  End If
  
  rs.Close
  
  iIncap = m_dataHelper.GetDouble2(txtTauxTechniqueIncap) / 100           ' taux technique
  fraisGestionIncap = m_dataHelper.GetDouble2(txtFraisGestionIncap) / 100 ' frais de gestion
  
  iInval = m_dataHelper.GetDouble2(txtTauxTechniqueInval) / 100           ' taux technique Inval
  fraisGestionInval = m_dataHelper.GetDouble2(txtFraisGestionInval) / 100 ' frais de gestion Inval
   
  xdepartRetraite = txtRetraite             ' âge de départ en retraite

  
  ' feuilleLueIncap = chargée précédemment  utilisée pour calculer les provisions de passage Incap/Inval
  ' feuilleLueInval = = chargée précédemment  utilisée pour calculer les provisions de passage Incap/Inval
  
  ' charge les modules de calcul
  Dim ModuleCalcul_Provisions As clsCalcul_Provisions
  Dim recordsetList As clsRecordsetList
  
  Set recordsetList = New clsRecordsetList
  Set ModuleCalcul_Provisions = New clsCalcul_Provisions

  Set ModuleCalcul_Provisions.recordsetList = recordsetList
   
  ' lance la sub de calcul qui remplie la table destination
  Call ModuleCalcul_Provisions.CalcTableauProv(x, Anciennete, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap, iInval, fraisGestionInval, xdepartRetraite, garantie, ProgressBar1, feuilleLueIncap, feuilleLueInval)
  
  Set ModuleCalcul_Provisions = Nothing
  Set recordsetList = Nothing
  
  Screen.MousePointer = vbDefault
End Sub

'****************************************************
'* paramètres de la fonction de calcul INVALIDITE   *
'****************************************************
'##ModelId=5C8A67B103C9
Private Sub btnCalcInval_Click()

  Dim x As Integer, Anciennete As Integer, Duree As Integer, franchise As Integer, xdepartRetraite As Integer
  Dim iIncap As Double, fraisGestionIncap As Double
  Dim iInval As Double, fraisGestionInval As Double
  
  Dim garantie As String, feuilleLueIncap As String, feuilleLue As String, feuilleLueInval As String, rq As String
    
  Dim rs As ADODB.Recordset
  
  Screen.MousePointer = vbHourglass
  
  
  garantie = "Invalidité"
  x = 0                     ' âge                 n'est pas passé en paramètres
  Anciennete = 0            ' ancienneté          n'est pas passé en paramètres
  Duree = 0                 ' durée en année      forcée à 0 calculée ultérieurement dans la fonction CalcTableauProv
  franchise = 0             ' franchise en années n'est pas passé en paramètres
  
  ' recupere le nom de la table invalidité
  If cboTableInval.ListIndex = -1 Then
    MsgBox "Vous devez choisir une Loi de maintien Invalidité !"
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInval.ItemData(cboTableInval.ListIndex)
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  If Not rs.EOF Then
    feuilleLue = rs.fields("NOMTABLE")
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Table 'loi de maintien Invalacité' INVALIDE !"
    rs.Close
    Exit Sub
  End If
  
  rs.Close
  
  
  iIncap = 0
  fraisGestionIncap = 0
  
  iInval = m_dataHelper.GetDouble2(txtTauxTechniqueInval) / 100           ' taux technique Inval
  fraisGestionInval = m_dataHelper.GetDouble2(txtFraisGestionInval) / 100 ' frais de gestion Inval
  
  xdepartRetraite = txtRetraite             ' âge de départ en retraite

  feuilleLueIncap = " "                     ' utilisée pour calculer les provisions de passage Incap/Inval
  feuilleLueInval = " "                     ' utilisée pour calculer les provisions de passage Incap/Inval
   
  ' charge les modules de calcul
  Dim ModuleCalcul_Provisions As clsCalcul_Provisions
  Dim recordsetList As clsRecordsetList
  
  Set recordsetList = New clsRecordsetList
  Set ModuleCalcul_Provisions = New clsCalcul_Provisions

  Set ModuleCalcul_Provisions.recordsetList = recordsetList
  
  ' lance la sub de calcul qui remplie la table destination
  Call ModuleCalcul_Provisions.CalcTableauProv(x, Anciennete, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap, iInval, fraisGestionInval, xdepartRetraite, garantie, ProgressBar1, feuilleLueIncap, feuilleLueInval)

  Set ModuleCalcul_Provisions = Nothing
  Set recordsetList = Nothing
  
  Screen.MousePointer = vbDefault

End Sub

'##ModelId=5C8A67B20000
Private Sub DoCalculCoeeficienntProvisionMaintienDC(bIncap As Boolean)
  Dim nomTable As String, libTable As String, typeTable As Integer
  Dim n As Integer, cleTableBCAC As Long
  Dim fReplaceExistingTable As Boolean
  Dim rsProvisionBCAC As ADODB.Recordset, rsTableloi As ADODB.Recordset
  Dim module_Calcul_PM_MaintienDeces As clsCalcul_PM_MaintienDeces
  Dim recordset_list As clsRecordsetList
  Dim cls As clsListeTableLoi
  
  Set cls = New clsListeTableLoi
  
  ' nom par défaut
  If bIncap Then
    nomTable = "P3I_PMIJ" & txtAgeLimiteCalulDC & "-i" & txtTauxTechnicCalculDC & "%-g" & txtFraisGestionCalculDC & "%"
  Else
    nomTable = "P3I_PMInv" & txtAgeLimiteCalulDC & "-i" & txtTauxTechnicCalculDC & "%-g" & txtFraisGestionCalculDC & "%"
  End If
  
  ' demande le nom de la table
  n = 0
  fReplaceExistingTable = False
  libTable = nomTable
  Do
    If bIncap Then
      libTable = InputBox("Entrez un nom pour la table de provision Maintien en garantie Décès :" & IIf(n <> 0, vbLf & "=> Le nom de la table doit être unique !", ""), "PM Maintien Incapacité", libTable)
    Else
      libTable = InputBox("Entrez un nom pour la table de provision Maintien en garantie Décès :" & IIf(n <> 0, vbLf & "=> Le nom de la table doit être unique !", ""), "PM Maintien Invalidité", libTable)
    End If
    
    If libTable = "" Then
      Exit Sub
    End If
    
    ' test l'unicité du nom
    cleTableBCAC = m_dataHelper.GetParameterAsLongWithParam("SELECT TABLECLE FROM ListeTableLoi WHERE LIBTABLE=?", libTable)
    If cleTableBCAC <> 0 Then
      n = 1
      If MsgBox("Une table existe avec ce nom. Voulez-vous la remplacer ?", vbQuestion + vbYesNo) = vbYes Then
        n = 0
        fReplaceExistingTable = True
      End If
    End If
  Loop While n = 1
  
  ' lance le calcul
  If bIncap Then
    typeTable = cdTypeTableCoeffBCACIncap
  Else
    typeTable = cdTypeTableCoeffBCACInval
  End If
    
  Screen.MousePointer = vbHourglass
    
  ' ouvre les tables de destinations
  Set rsProvisionBCAC = m_dataSource.OpenRecordset("ProvisionBCAC", table)
  Set rsTableloi = m_dataSource.OpenRecordset("SELECT * FROM ListeTableLoi", Dynamic)
  
  If fReplaceExistingTable Then
    cleTableBCAC = m_dataHelper.GetParameterAsLongWithParam("SELECT TABLECLE FROM ListeTableLoi WHERE LIBTABLE=?", libTable)
    
    cls.Load m_dataSource, cleTableBCAC
    
    ' met à jour le type de table
    cls.m_TYPETABLE = typeTable
    cls.Save m_dataSource
    
    ' vide la table existante
    m_dataSource.Execute "DELETE FROM ProvisionBCAC WHERE CleTable=" & cleTableBCAC
  Else
    ' recupere la clé de la table
    cls.m_LIBTABLE = libTable
    If bIncap Then
      cls.m_NOMTABLE = "P3I_PMIJMDC_" & m_dataHelper.GetParameterAsLong("SELECT MAX(IsNull(TABLECLE,0))+1 FROM ListeTableLoi")
    Else
      cls.m_NOMTABLE = "P3I_PMInvMDC_" & m_dataHelper.GetParameterAsLong("SELECT MAX(IsNull(TABLECLE,0))+1 FROM ListeTableLoi")
    End If
    cls.m_TYPETABLE = typeTable
    cls.m_TableUtilisateur = True
    
    cls.Save m_dataSource
    
    cleTableBCAC = cls.m_TABLECLE
  End If
  
  ' lance le calcul
  Set module_Calcul_PM_MaintienDeces = New clsCalcul_PM_MaintienDeces
  Set recordset_list = New clsRecordsetList
  
  Set module_Calcul_PM_MaintienDeces.recordsetList = recordset_list
  
  ' calcul de la table
  If bIncap Then
    module_Calcul_PM_MaintienDeces.CalculTableCoeffPMMaintienDecesIncap _
                          GetSettingIni(CompanyName, SectionName, "LoiIncapacite", "#"), _
                          GetSettingIni(CompanyName, SectionName, "LoiIncapaciteDC", "#"), _
                          GetSettingIni(CompanyName, SectionName, "LoiPassage", "#"), _
                          GetSettingIni(CompanyName, SectionName, "LoiInvalidite", "#"), _
                          GetSettingIni(CompanyName, SectionName, "LoiInvaliditeDC", "#"), _
                          m_dataHelper.GetDouble2(txtTauxTechnicCalculDC), _
                          m_dataHelper.GetDouble2(txtFraisGestionCalculDC), _
                          m_dataHelper.GetDouble2(txtTauxTechnicCalculDC), _
                          m_dataHelper.GetDouble2(txtFraisGestionCalculDC), _
                          rsProvisionBCAC, cleTableBCAC, ProgressBar1, m_dataHelper.GetDouble2(txtAgeLimiteCalulDC)
  Else
    module_Calcul_PM_MaintienDeces.CalculTableCoeffPMMaintienDecesInval _
                          GetSettingIni(CompanyName, SectionName, "LoiInvalidite", "#"), _
                          GetSettingIni(CompanyName, SectionName, "LoiInvaliditeDC", "#"), _
                          m_dataHelper.GetDouble2(txtTauxTechnicCalculDC), _
                          m_dataHelper.GetDouble2(txtFraisGestionCalculDC), _
                          rsProvisionBCAC, cleTableBCAC, ProgressBar1, m_dataHelper.GetDouble2(txtAgeLimiteCalulDC)
  End If
  
  ' clean up
  Set module_Calcul_PM_MaintienDeces = Nothing
  
  recordset_list.CloseLoadedRecordset
  Set recordset_list = Nothing

  rsProvisionBCAC.Close
  rsTableloi.Close
  
  ' rafraichie la liste des tables
  SSTab1_Click SSTab1.Tab
  
  ' selectionne la table que l'on vient de calculee
  For n = 0 To cboBCAC.ListCount
    If cboBCAC.ItemData(n) = cleTableBCAC Then
      cboBCAC.ListIndex = n
      Exit For
    End If
  Next
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67B2001F
Private Sub btnCalclIncapBCAC_Click()
  DoCalculCoeeficienntProvisionMaintienDC True
End Sub

'##ModelId=5C8A67B2002F
Private Sub btnCalclInvalBCAC_Click()
  DoCalculCoeeficienntProvisionMaintienDC False
End Sub

'##ModelId=5C8A67B2004E
Private Sub btnCATR9Unique_Click()
  Dim rs As ADODB.Recordset, txt As String
  
  txtCATR9NonUnique = ""
  
  Set rs = m_dataSource.OpenRecordset("SELECT CATR9.Categorie FROM CATR9 CATR9 INNER JOIN CATR9INVAL CATR9INVAL ON CATR9.Categorie = CATR9INVAL.Categorie", Snapshot)
  
  If rs.EOF Then
    txt = "Aucune catégorie en erreur"
  Else
    Do Until rs.EOF
      txt = txt & "Catégorie " & rs.fields("Categorie") & vbLf
      
      rs.MoveNext
    Loop
  End If
  
  rs.Close
  
  txtCATR9NonUnique = txt
End Sub

'##ModelId=5C8A67B2005E
Private Sub btnDelBCAC_Click()
  If DroitAdmin = False Then Exit Sub
  If cboBCAC.ListIndex = -1 Then Exit Sub
  
  On Error GoTo err_delbcac
  
  If m_dataHelper.GetParameterAsDouble("SELECT COUNT(*) FROM ParamCalcul WHERE PECleTableBCACInval_MDC=" & cboBCAC.ItemData(cboBCAC.ListIndex) & " OR PECleTableBCACIncap_MDC=" & cboBCAC.ItemData(cboBCAC.ListIndex)) Then
    MsgBox "La table est utilisée par une période. Vous ne pouvez la supprimer !", vbCritical
    Exit Sub
  End If
  
  If MsgBox("Voulez-vous vraiement supprimer la table '" & cboBCAC.List(cboBCAC.ListIndex) & "' ?", vbQuestion + vbYesNo) = vbYes Then
    m_dataSource.Execute "DELETE FROM ProvisionBCAC WHERE CleTable=" & cboBCAC.ItemData(cboBCAC.ListIndex)
    m_dataSource.Execute "DELETE FROM ListeTableLoi WHERE TABLECLE=" & cboBCAC.ItemData(cboBCAC.ListIndex)
  End If
    
  ' table de coefficients précalculés
  m_dataHelper.FillCombo cboBCAC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval & " OR TYPETABLE = " & cdTypeTableCoeffBCACIncap, -1, False, True
  If cboBCAC.ListCount <> 0 Then
    cboBCAC.ListIndex = 0
  End If
  
  SSTab1_Click SSTab1.Tab
  
  Exit Sub

err_delbcac:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=5C8A67B2006D
Private Function GetTypeParam() As String
    ' type de paramètre
  GetTypeParam = "BILAN"
  If rdoBilan.Value = True Then
    GetTypeParam = "BILAN"
  ElseIf rdoClient.Value = True Then
    GetTypeParam = "CLIENT"
  ElseIf rdoSimul.Value = True Then
    GetTypeParam = "SIMULATION"
  End If
End Function

'##ModelId=5C8A67B2007D
Private Sub btnAddPAram_Click()
  Dim frm As New frmParamCalculIni
  
  NumParamCalcul = -1
  
  ' type de paramètre
  frm.typeParam = GetTypeParam
  
  frm.Show vbModal, Me
  
  Set frm = Nothing
  
  ' Liste des paramètres de calcul
  RefreshListParamCalculIni
End Sub

'##ModelId=5C8A67B2009C
Private Sub btnDelete_Click()
  If cboListeTable.ListIndex = -1 Then Exit Sub
    
  Dim idTable As Long
  
  idTable = cboListeTable.ItemData(cboListeTable.ListIndex)
  
  If 0 <> m_dataHelper.GetParameterAsDouble("SELECT TableUtilisateur FROM ListeTableLoi WHERE TABLECLE=" & idTable) Then
    Dim cls As clsListeTableLoi
    
    Set cls = New clsListeTableLoi
    
    cls.Load m_dataSource, idTable
    
    cls.Delete m_dataSource, True
    If archiveOk Then cls.Delete m_dataSourceArchive, True
  
    ' Rafraichissement du combe liste des tables et sélection de la table qui vient d'être importée
    RefreshListeTableLoi -1
  End If

End Sub

'##ModelId=5C8A67B200AC
Private Sub btnDelParam_Click()
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  On Error GoTo err_del
  
  SetNumParamCalcul

  If MsgBox("Voulez vous vraiement supprimer les paramètres de calcul : " & NumParamCalcul & " de type " & GetTypeParam & " ?", vbQuestion + vbYesNo) = vbYes Then
    Call DeleteSection(CompanyName, DEFAULT_PARAM_SECTION & GetTypeParam & "_" & NumParamCalcul)
    
    ' Liste des paramètres de calcul
    RefreshListParamCalculIni
  End If
  
  Exit Sub
  
err_del:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=5C8A67B200BB
Private Sub btnEditParam_Click()
  Dim frm As New frmParamCalculIni
  
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  SetNumParamCalcul
  
  ' type de paramètre
  frm.typeParam = GetTypeParam
  
  frm.Show vbModal, Me
  
  Set frm = Nothing
  
  ' Liste des paramètres de calcul
  RefreshListParamCalculIni
End Sub

'##ModelId=5C8A67B200CB
Private Sub btnImportLoiMantien_Click()
  Dim idTable As Long, idTypeTable As Integer, work_table As String
  Dim cListe As clsListeTableLoi, bTableExists As Boolean
  
  Set cListe = New clsListeTableLoi
  
  ' Choix du type de table, du fichier, ...
  Dim FM As frmImportTable
  
  Set FM = New frmImportTable
  
  FM.Show vbModal
  If FM.ret_code = 0 Then
    
    ' Création/Modification de la ligne dans ListeTableLoi
    bTableExists = False
    idTable = m_dataHelper.GetParameterAsLong("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE='" & FM.txtNomTable & "'")
    If idTable <> 0 Then
      cListe.Load m_dataSource, idTable
      
      bTableExists = True
      
      If cListe.m_TableUtilisateur = False Then
        MsgBox "La table '" & FM.txtNomTable & "' est une table système et ne pas être modifiée !", vbExclamation, "Import table..."
        Unload FM
        Exit Sub
      End If
      
      If MsgBox("La table '" & FM.txtNomTable & "' existe déjà." & vbLf & "Voulez-vous la remplacer ?", vbQuestion + vbYesNo, "Import table...") = vbNo Then
        Unload FM
        Exit Sub
      End If
      
    End If
    
    m_dataSource.BeginTrans "Import table"
    If archiveOk Then m_dataSourceArchive.BeginTrans
    
    ' Création de la table / suppression des données existantes
    idTypeTable = FM.cboTypeTable.ItemData(FM.cboTypeTable.ListIndex)
    
    If idTable = 0 Then
      cListe.m_NOMTABLE = FM.txtNomTable
      cListe.m_TYPETABLE = idTypeTable
      cListe.m_TableUtilisateur = True
    End If
    cListe.m_LIBTABLE = FM.txtLibelleTable
    
    cListe.Save m_dataSource
    If archiveOk Then cListe.Save m_dataSourceArchive
    
    idTable = cListe.m_TABLECLE
    
    Select Case idTypeTable
      
      Case cdTypeTable_BaremeAnneeStatutaire
      
        work_table = FM.txtNomTable
        If bTableExists = True Then
          m_dataSource.Execute "DROP TABLE " & work_table
          If archiveOk Then m_dataSourceArchive.Execute "DROP TABLE " & work_table
        End If
        
         If CreateTable(idTypeTable, work_table, m_dataSource) = False Then
          MsgBox "Impossible de créer la table " & work_table & " !", vbCritical
          m_dataSource.RollbackTrans "Import table"
          Unload FM
          Exit Sub
        End If
        
        If archiveOk Then
          If CreateTable(idTypeTable, work_table, m_dataSourceArchive) = False Then
            MsgBox "Impossible de créer la table " & work_table & " dans la base Archive !", vbCritical
            m_dataSourceArchive.RollbackTrans
            Unload FM
            Exit Sub
          End If
        End If

        ' Import des données - Bulk Insert
        Dim csvFile As String
        csvFile = Right(FM.txtFilename, Len(FM.txtFilename) - InStrRev(FM.txtFilename, "\"))
        
        If BulkInsertStat(FM.txtNomTable, csvFile) Then
          m_dataSource.CommitTrans "Import table"
          If archiveOk Then m_dataSourceArchive.CommitTrans
        Else
          m_dataSource.RollbackTrans "Import table"
          If archiveOk Then m_dataSourceArchive.RollbackTrans
        End If
        
        If Not (FM Is Nothing) Then
          Unload FM
        End If
      
        MsgBox "L'import est fini !"
        
        ' Rafraichissement du combe liste des tables et sélection de la table qui vient d'être importée
        RefreshListeTableLoi idTable
  
        Exit Sub
        
      Case cdTypeTable_LoiMaintienIncapacite, cdTypeTable_LoiPassage, cdTypeTable_LoiMaintienInvalidite, cdTypeTable_LoiDependance
        work_table = FM.txtNomTable
        If bTableExists = True Then
          m_dataSource.Execute "DROP TABLE " & work_table
          If archiveOk Then m_dataSourceArchive.Execute "DROP TABLE " & work_table
        End If
        
        If CreateTable(idTypeTable, work_table, m_dataSource) = False Then
          MsgBox "Impossible de créer la table " & work_table & " !", vbCritical
          m_dataSource.RollbackTrans "Import table"
          Unload FM
          Exit Sub
        End If
        
        If archiveOk Then
          If CreateTable(idTypeTable, work_table, m_dataSourceArchive) = False Then
            MsgBox "Impossible de créer la table " & work_table & " dans la vase Archive !", vbCritical
            m_dataSourceArchive.RollbackTrans
            Unload FM
            Exit Sub
          End If
        End If
        
      Case cdTypeTableMortalite, cdTypeTableGeneration
        work_table = "TableMortalite"
        m_dataSource.Execute "DELETE FROM TableMortalite WHERE NOMTABLE='" & cListe.m_NOMTABLE & "'"
        If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM TableMortalite WHERE NOMTABLE='" & cListe.m_NOMTABLE & "'"
        
      Case cdTypeTable_MortaliteIncap
        work_table = "MortIncap"
        m_dataSource.Execute "DELETE FROM MortIncap WHERE CleTable='" & idTable & "'"
        If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM MortIncap WHERE CleTable='" & idTable & "'"
        
      Case cdTypeTable_MortaliteInval
        work_table = "MortInval"
        m_dataSource.Execute "DELETE FROM MortInval WHERE CleTable='" & idTable & "'"
        If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM MortInval WHERE CleTable='" & idTable & "'"
        
    End Select
    
    ' Import des données
    If ImportTableMortalite(FM.txtFilename, work_table, idTable, FM.txtNomTable) Then
      m_dataSource.CommitTrans "Import table"
      If archiveOk Then m_dataSourceArchive.CommitTrans
    Else
      m_dataSource.RollbackTrans "Import table"
      If archiveOk Then m_dataSourceArchive.RollbackTrans
    End If
    
  Else   'If FM.ret_code = 0 Then...
    Set FM = Nothing
  End If
    
  If Not (FM Is Nothing) Then
    Unload FM
  End If

  ' Rafraichissement du combe liste des tables et sélection de la table qui vient d'être importée
  RefreshListeTableLoi idTable
  
End Sub


'##ModelId=5C8A67B200DB
Private Function RefreshListeTableLoi(ByVal idTable As Long)
  
  ' Rafraichissement du combe liste des tables et sélection de la table qui vient d'être importée
  cboListeTable.Clear
  m_dataHelper.FillCombo cboListeTable, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE<> " & cdTypeTableCoeffBCACIncap & " AND TYPETABLE <> " & cdTypeTableCoeffBCACInval, idTable, False, True

End Function


'##ModelId=5C8A67B20109
Private Sub btnSupprimerTauxProvision_Click()
  If cboTableTauxProvision.ListIndex = -1 Then Exit Sub

  Dim CleTable As Long, nomTable As String
  
  CleTable = cboTableTauxProvision.ItemData(cboTableTauxProvision.ListIndex)
  nomTable = cboTableTauxProvision.List(cboTableTauxProvision.ListIndex)

  If MsgBox("Voulez-vous vraiement supprimer la table " & nomTable & " ?", vbQuestion + vbYesNo) = vbYes Then
    ' Suppression de la table
    m_dataSource.Execute "DROP TABLE " & nomTable
    
    ' Suppression de la reference dasn ListeTableLoi
    m_dataSource.Execute "DELETE FROM ListeTableLoi WHERE TABLECLE=" & CleTable
    
    ' rafraichi la liste des tables de coefficients précalculés
    SSTab1_Click SSTab1.Tab
  End If
End Sub

'##ModelId=5C8A67B20119
Private Sub cmdSaveStatParams_Click()
  btnSave_Click
End Sub

'##ModelId=5C8A67B20129
Private Sub lvParamCalcul_DblClick()
  btnEditParam_Click
End Sub

'##ModelId=5C8A67B20138
Private Sub btnExportBCAC_Click()
  If cboBCAC.ListIndex = -1 Then Exit Sub
  
  Dim CleTable As Long
  
  CleTable = cboBCAC.ItemData(cboBCAC.ListIndex)
    
  ' affiche les coefficients avec toute la précision
  AfficheProvisionTableBCAC CleTable, True
  
  ExportTableToExcelFile cboBCAC.List(cboBCAC.ListIndex) & ".xls", _
                         "TableBCAC", "TableBCAC", sprBCAC, CommonDialog1, "", True
  
  ' affiche les coefficients arrondis
  AfficheProvisionTableBCAC CleTable, False
End Sub

'##ModelId=5C8A67B20158
Private Sub btnExportCATR9_Click()
  If cboCATR9.ListIndex = -1 Then Exit Sub
  
  ExportTableToExcelFile cboCATR9.List(cboCATR9.ListIndex) & ".xls", _
                         cboCATR9.List(cboCATR9.ListIndex), _
                         cboCATR9.List(cboCATR9.ListIndex), sprCATR9, CommonDialog1, "", False
End Sub

'##ModelId=5C8A67B20167
Private Sub btnExportLoiMantien_Click()
  If cboListeTable.ListIndex = -1 Then Exit Sub
  
  Dim nom_table As String
  Dim typeTable As String
  
  nom_table = m_dataHelper.GetParameterAsStringCRW("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE=" & cboListeTable.ItemData(cboListeTable.ListIndex))
  
'  typeTable = m_dataHelper.GetParameterAsDouble("SELECT TYPETABLE FROM ListeTableLoi WHERE TABLECLE=" & cboListeTable.ItemData(cboListeTable.ListIndex))
'
'  If typeTable = cdTypeTable_BaremeAnneeStatutaire Then
'
'  Else
'
'  End If
  
  ExportTableToExcelFile cboListeTable.List(cboListeTable.ListIndex) & ".xls", _
                         cboListeTable.List(cboListeTable.ListIndex), _
                         IIf(nom_table = "", "LoiDeMaintien", nom_table), vaSpread1, CommonDialog1, "", False, True, ProgressBar1
                         
  ProgressBar1.Visible = False
  
End Sub

'##ModelId=5C8A67B20177
Private Sub btnImportBCAC_Click()
  If DroitAdmin = False Then Exit Sub
  
  ImportBCAC CommonDialog1, ProgressBar1
  
  ' rafraichi la page
  m_dataHelper.FillCombo cboBCAC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval & " OR TYPETABLE = " & cdTypeTableCoeffBCACIncap, -1, False, True
  If cboBCAC.ListCount <> 0 Then
    cboBCAC.ListIndex = 0
  End If
  
  SSTab1_Click SSTab1.Tab
End Sub

'##ModelId=5C8A67B20196
Private Sub btnImportCATR9_Click()
  Dim nomTable As String
  
  If DroitAdmin = False Then Exit Sub
  If cboCATR9.ListIndex = -1 Then Exit Sub
  
  nomTable = cboCATR9.List(cboCATR9.ListIndex)
  
  If nomTable = "CodePosition" Or nomTable = "CodeProvision" Then Exit Sub
  
  ImportGenerique CommonDialog1, ProgressBar1, nomTable, -1

  ' rafraichi la page
  Dim i As Integer
  
  i = cboCATR9.ListIndex
  
  SSTab1_Click SSTab1.Tab
  
  cboCATR9.ListIndex = i
End Sub

'##ModelId=5C8A67B201A6
Private Sub btnPrint_Click()
  Dim bUsePrintDlg As Integer
  
  If cboListeTable.ListIndex = -1 Then Exit Sub

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
  
  With vaSpread1
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
    .PrintJobName = cboListeTable
    .PrintHeader = "/c - Lois de maintien : " & cboListeTable & " - /n"
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
    .PrintOrientation = SS_PRINTORIENT_LANDSCAPE
    
    .PrintType = SS_PRINT_ALL
    .Action = SS_ACTION_SMARTPRINT
  
    .FontBold = svgFontBold
  End With

err_print:
End Sub

'##ModelId=5C8A67B201B5
Private Sub btnPrintBCAC_Click()
  Dim bUsePrintDlg As Integer
  
  If cboBCAC.ListIndex = -1 Then Exit Sub

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
  
  With sprBCAC
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
    .PrintJobName = cboBCAC
    .PrintHeader = "/c - Table des coefficients précalculés du BCAC : " & cboBCAC & " - /n"
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
    .PrintOrientation = SS_PRINTORIENT_LANDSCAPE
    
    .PrintType = SS_PRINT_ALL
    .Action = SS_ACTION_SMARTPRINT
  
    .FontBold = svgFontBold
  End With

err_print:
End Sub

'##ModelId=5C8A67B201D5
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

'##ModelId=5C8A67B201E4
Private Sub btnPrintTaux_Click()
  Dim bUsePrintDlg As Integer
  
  If cboTableTauxProvision.ListIndex = -1 Then Exit Sub

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
  
  With vaSpread2
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
    .PrintJobName = cboTableTauxProvision
    .PrintHeader = "/c - Taux de provisions : " & cboTableTauxProvision & " - /n"
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
    .PrintOrientation = SS_PRINTORIENT_LANDSCAPE
    
    .PrintType = SS_PRINT_ALL
    .Action = SS_ACTION_SMARTPRINT
  
    .FontBold = svgFontBold
  End With

err_print:
End Sub

'##ModelId=5C8A67B20203
Private Sub btnSave_Click()
  'PHM_SXO
  'valeurs par defaut
  Dim sel As Integer
  Dim rq As String, rs As ADODB.Recordset
  
  ' Incap / Inval
  If SSTab1.Tab = 0 Then
    ' rempli les combo de liste des tables
    ' incapacite
    If cboTableIncap.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncap.ItemData(cboTableIncap.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "LoiIncapacite", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' passage
    If cboTablePassage.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTablePassage.ItemData(cboTablePassage.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "LoiPassage", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' invalidite
    If cboTableInval.ListIndex <> -1 Then
      
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInval.ItemData(cboTableInval.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "LoiInvalidite", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' taux
    Call SaveSettingIni(CompanyName, SectionName, "TauxIncapacite", txtTauxTechniqueIncap)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionIncapacite", txtFraisGestionIncap)

    Call SaveSettingIni(CompanyName, SectionName, "TauxInvalidité", txtTauxTechniqueInval)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionInvalidite", txtFraisGestionInval)
    Call SaveSettingIni(CompanyName, SectionName, "AgeRetraite", txtRetraite)
    Call SaveSettingIni(CompanyName, SectionName, "DureeIncap", Txtduree)
    Call SaveSettingIni(CompanyName, SectionName, "FranchiseIncap", TxtFranchise)
    
    Call SaveSettingIni(CompanyName, SectionName, "NbDecimalPM", txtNbDecimalePM)
    Call SaveSettingIni(CompanyName, SectionName, "NbDecimalCalcul", txtNbDecimaleCalcul)
  End If
  
  'Statutaire
'  If SSTab1.Tab = 6 Then
'
'    Call SaveSettingIni(CompanyName, SectionName, "StatAgeRetraite", txtAgeRet)
'    Call SaveSettingIni(CompanyName, SectionName, "StatAgeMin", txtAgeMin)
'    Call SaveSettingIni(CompanyName, SectionName, "StatAgeMax", txtAgeMax)
'    Call SaveSettingIni(CompanyName, SectionName, "StatMOMaxSemaine", txtMOMaxSemaine)
'    Call SaveSettingIni(CompanyName, SectionName, "StatCLMMaxSemaine", txtCLMMaxSemaine)
'    Call SaveSettingIni(CompanyName, SectionName, "StatCLDMaxSemaine", txtCLDMaxSemaine)
'    Call SaveSettingIni(CompanyName, SectionName, "StatMATMaxSemaine", txtMATMaxSemaine)
'    Call SaveSettingIni(CompanyName, SectionName, "StatATMaxSemaine", txtATMaxSemaine)
'
'    Call SaveSettingIni(CompanyName, SectionName, "StatAnneeBareme", cmbAnneeBareme.text)
'
'  End If
  
  ' Rentes
  If SSTab1.Tab = 1 Then
    ' rente conjoint
    If cboTableRenteConjoint.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableRenteConjoint.ItemData(cboTableRenteConjoint.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "RenteConjoint", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' rente education
    If cboTableRenteEducation.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableRenteEducation.ItemData(cboTableRenteEducation.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "RenteEducation", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' taux
    Call SaveSettingIni(CompanyName, SectionName, "TauxIndexation", txtTxIndex)
    Call SaveSettingIni(CompanyName, SectionName, "DureeIndexation", txtDureeIndex)
    Call SaveSettingIni(CompanyName, SectionName, "TMO", txtTMO)
    
    Call SaveSettingIni(CompanyName, SectionName, "TauxTechniqueRenteConjoint", txtTauxTechniqueRenteConjoint)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionRenteConjoint", txtFraisGestionRenteConjoint)
    Call SaveSettingIni(CompanyName, SectionName, "TauxTechniqueRenteEducation", txtTauxTechniqueRenteEducation)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionRenteEducation", txtFraisGestionRenteEducation)
    
    ' options pour rentes
    Call SaveSettingIni(CompanyName, SectionName, "AnnuelConjoint", IIf(rdoAnnuelConjoint, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "SemestrielConjoint", IIf(rdoSemestrielConjoint, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "TrimestrielConjoint", IIf(rdoTrimestrielConjoint, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "MensuelConjoint", IIf(rdoMensuelConjoint, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "PaiementAvanceConjoint", IIf(rdoPaiementAvanceConjoint, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "PaiementEchuConjoint", IIf(rdoPaiementEchuConjoint, 1, 0))
  
    Call SaveSettingIni(CompanyName, SectionName, "AnnuelEducation", IIf(rdoAnnuelEducation, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "SemestrielEducation", IIf(rdoSemestrielEducation, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "TrimestrielEducation", IIf(rdoTrimestrielEducation, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "MensuelEducation", IIf(rdoMensuelEducation, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "PaiementAvanceEducation", IIf(rdoPaiementAvanceEducation, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "PaiementEchuEducation", IIf(rdoPaiementEchuEducation, 1, 0))
  End If
  
  ' Mantien DC
  If SSTab1.Tab = 2 Then
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionCapitauxDecesDC", txtFraisGestionCapitauxDecesDC)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionRenteEducationDC", txtFraisGestionRenteEducationDC)
    
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionRenteConjointDC", txtFraisGestionRenteConjointDC)
    Call SaveSettingIni(CompanyName, SectionName, "CapitalMoyenRenteConjointTempoDC", txtCapitalMoyenRenteConjointTempoDC)
    Call SaveSettingIni(CompanyName, SectionName, "CapitalMoyenRenteConjointViagereDC", txtCapitalMoyenRenteConjointViagereDC)
    Call SaveSettingIni(CompanyName, SectionName, "ForcerCapitalMoyenRteConjoitDC", IIf(chkForcerCapitalMoyenRteConjoitDC.Value = vbChecked, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "AgeConjointRenteConjointDC", txtAgeConjointRenteConjointDC)
    
    Call SaveSettingIni(CompanyName, SectionName, "LissageProvision", IIf(rdoAvecLissage, 1, 0))
    
    Call SaveSettingIni(CompanyName, SectionName, "RecalculBCAC", IIf(rdoCalculCoeffBCAC, 1, 0))
    Call SaveSettingIni(CompanyName, SectionName, "PMGDForcerInval", IIf(chkPMGDForcerInval = vbChecked, 1, 0))
    
    Call SaveSettingIni(CompanyName, SectionName, "MethodeCalcul", IIf(rdoCapitauxConstitif = True, 0, 1))
    
    Call SaveSettingIni(CompanyName, SectionName, "TauxTechniqueDC", txtTauxTechnicCalculDC)
    Call SaveSettingIni(CompanyName, SectionName, "FraisGestionDC", txtFraisGestionCalculDC)
    Call SaveSettingIni(CompanyName, SectionName, "AgeLimiteCalulDC", txtAgeLimiteCalulDC)
    
    ' recalcul incap
    If cboTableIncapCalculDC.ListIndex <> -1 Then
      Call SaveSettingIni(CompanyName, SectionName, "LoiIncapaciteDC", cboTableIncapCalculDC.ItemData(cboTableIncapCalculDC.ListIndex))
    End If
    
    ' recalcul inval
    If cboTableInvalCalculDC.ListIndex <> -1 Then
      Call SaveSettingIni(CompanyName, SectionName, "LoiInvaliditeDC", cboTableInvalCalculDC.ItemData(cboTableInvalCalculDC.ListIndex))
    End If
    
    ' precalcul incap
    If cboTableIncapPrecalculDC.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncapPrecalculDC.ItemData(cboTableIncapPrecalculDC.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "LoiIncapacitePrecalculDC", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
    
    ' precalcul inval
    If cboTableInvalPrecalculDC.ListIndex <> -1 Then
      rq = "SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInvalPrecalculDC.ItemData(cboTableInvalPrecalculDC.ListIndex)
      Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
      If Not rs.EOF() Then
        Call SaveSettingIni(CompanyName, SectionName, "LoiInvaliditePrecalculDC", rs.fields("NOMTABLE"))
      End If
      rs.Close
    End If
  End If
  
End Sub

'##ModelId=5C8A67B20213
Private Sub btnSaveDC_Click()
  btnSave_Click
End Sub

'##ModelId=5C8A67B20223
Private Sub btnExportTauxProvision_Click()
  If cboTableTauxProvision.ListIndex = -1 Then Exit Sub
  
  ExportTableToExcelFile cboTableTauxProvision.List(cboTableTauxProvision.ListIndex) & ".xls", _
                         cboTableTauxProvision.List(cboTableTauxProvision.ListIndex), _
                         "TauxDeProvision", vaSpread2, CommonDialog1, "", False
End Sub

'##ModelId=5C8A67B20232
Private Sub cboBCAC_Click()
  Dim CleTable As Long
  
  If cboBCAC.ListIndex <> -1 Then
    CleTable = cboBCAC.ItemData(cboBCAC.ListIndex)
    
    AfficheProvisionTableBCAC CleTable, False
  End If
End Sub

'##ModelId=5C8A67B20252
Private Sub AfficheProvisionTableBCAC(CleTable As Long, bForExport As Boolean)
  ' matrice des provisions (anciennete, age)
  Dim AgeMini As Integer, ancMini As Integer, AgeMax As Integer, AncMax As Integer
  Dim i As Integer
  
  Screen.MousePointer = vbHourglass
  
  ' recuprére les ages
  Dim rs As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT Min(Age) as AgeMini, Max(Age) as AgeMax, " _
                                      & " Min(Anciennete) as AncMini, Max(Anciennete) as AncMax " _
                                      & " FROM ProvisionBCAC WHERE CleTable=" & CleTable, Snapshot)
  
  If rs.EOF Then
    sprBCAC.MaxCols = 0
    sprBCAC.MaxRows = 0
    Screen.MousePointer = vbDefault
    rs.Close
    Exit Sub
  End If
  
  If IsNull(rs.fields("AgeMini")) Then
    sprBCAC.MaxCols = 0
    sprBCAC.MaxRows = 0
    Screen.MousePointer = vbDefault
    rs.Close
    Exit Sub
  End If
  
  AgeMini = rs.fields("AgeMini")
  AgeMax = rs.fields("AgeMax")
  
  ancMini = rs.fields("AncMini")
  AncMax = rs.fields("AncMax")
  
  rs.Close
  
  sprBCAC.ReDraw = False
  
  sprBCAC.MaxRows = 0
  sprBCAC.MaxCols = 0
  
  sprBCAC.MaxRows = AgeMax - AgeMini + 1
  sprBCAC.MaxCols = AncMax - ancMini + 1
  
  ' titre
  sprBCAC.Row = SpreadHeader
  sprBCAC.Col = SpreadHeader
  sprBCAC.text = "Age"
  
  ' labelise les entetes
  For i = AgeMini To AgeMax
    sprBCAC.Col = SpreadHeader
    sprBCAC.Row = i - AgeMini + 1
    sprBCAC.text = i
  Next
  
  For i = ancMini To AncMax
    sprBCAC.Row = SpreadHeader
    sprBCAC.Col = i - ancMini + 1
    sprBCAC.text = IIf(i = ancMini, "Anc=" & i, i)
  Next
  
  ' recupère les provisions
  Set rs = m_dataSource.OpenRecordset("SELECT Anciennete, Age, Provision FROM ProvisionBCAC " _
                                     & " WHERE CleTable=" & CleTable & " ORDER BY Anciennete, Age", Disconnected)
  
  Do Until rs.EOF
    ' anciennete
    sprBCAC.Col = rs.fields(0) - ancMini + 1
    sprBCAC.Row = rs.fields(1) - AgeMini + 1

    sprBCAC.CellType = CellTypeEdit
    'sprBCAC.TypeNumberDecPlaces = 4
    sprBCAC.TypeHAlign = TypeHAlignCenter

    If Not IsNull(rs.fields(2)) Then
      'sprBCAC.Value = Format(rs.Fields(2), "0.0000")
      If IsNull(rs.fields(2)) Then
        sprBCAC.Value = ""
      Else
        sprBCAC.Value = Format(rs.fields(2), IIf(bForExport, "0.00000000", "0.0000"))
      End If
    Else
      sprBCAC.Value = 0#
    End If
    
    rs.MoveNext
  Loop
  
  rs.Close
  
  LargeurMaxColonneSpread sprBCAC
  
  sprBCAC.ReDraw = True
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67B20280
Private Sub rdoBilan_Click()
  RefreshListParamCalculIni
End Sub

'##ModelId=5C8A67B202A0
Private Sub rdoClient_Click()
  RefreshListParamCalculIni
End Sub

'##ModelId=5C8A67B202AF
Private Sub rdoSimulation_Click()
  RefreshListParamCalculIni
End Sub

'##ModelId=5C8A67B202CF
Private Sub sprCATR9_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  Dim comment As Variant
  
  sprCATR9.GetDataFillData comment, vbString
  If comment = "" Then
    sprCATR9.Col = Col
    sprCATR9.Row = Row
    sprCATR9.Value = ""
    
    Cancel = True
  End If
End Sub

'##ModelId=5C8A67B2033C
Private Sub cboCATR9_Click()
  Dim rq As String
    
  If cboCATR9.ListIndex <> -1 Then
    sprCATR9.ReDraw = False
    
    m_dataSource.SetDatabase dtaCATR9
        
    On Error Resume Next
        
    rq = "SELECT * FROM " & fmTableDiverse.TableInfo(cboCATR9.ListIndex).nomTable & " ORDER BY " & fmTableDiverse.TableInfo(cboCATR9.ListIndex).orderBy
        
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
        sprCATR9.DataFillEvent = True
      Next
      
      dtaCATR9.Refresh
    Else
      sprCATR9.MaxRows = 0
    End If
    
    sprCATR9.Refresh
    
    ' largeur des colonnes
    LargeurMaxColonneSpread sprCATR9
    
    sprCATR9.ReDraw = True
  
    If fmTableDiverse.TableInfo(cboCATR9.ListIndex).nomTable = "CodePosition" Or _
      fmTableDiverse.TableInfo(cboCATR9.ListIndex).nomTable = "CodeProvision" Or _
      fmTableDiverse.TableInfo(cboCATR9.ListIndex).nomTable = "Statutaire_Garantie" Then
      btnImportCATR9.Enabled = False
    Else
      btnImportCATR9.Enabled = True
    End If
       
    On Error GoTo 0
  End If
End Sub


'##ModelId=5C8A67B2034C
Private Sub cboListeTable_Click()
  Dim nomTable As String
  
  If cboListeTable.ListIndex <> -1 Then
    vaSpread1.ReDraw = False
    
    m_dataSource.SetDatabase dtaListeTable
        
    Dim rq As String, typeDeTable As Integer
    
    ' fabrique la requete en fct de la table
    typeDeTable = m_dataHelper.GetParameterAsDouble("SELECT TYPETABLE FROM ListeTableLoi WHERE TABLECLE=" & cboListeTable.ItemData(cboListeTable.ListIndex))
    Select Case typeDeTable
      Case cdTypeTableMortalite, cdTypeTableGeneration ' TableMortalité
        ' table etant une partie de TableMortalite
        rq = "SELECT * FROM TableMortalite WHERE NomTable='" & m_dataHelper.GetParameter("SELECT NOMTABLE FROM listeTableLoi WHERE TABLECLE = " & cboListeTable.ItemData(cboListeTable.ListIndex)) & "' ORDER BY Naissance"
      
      Case cdTypeTable_MortaliteIncap ' Mortalité Incap
        rq = "SELECT * FROM MortIncap WHERE CleTable=" & cboListeTable.ItemData(cboListeTable.ListIndex) & " ORDER BY Age"
      
      Case cdTypeTable_MortaliteInval ' Mortalité Inval
        rq = "SELECT * FROM MortInval WHERE CleTable=" & cboListeTable.ItemData(cboListeTable.ListIndex) & " ORDER BY Age"
        
      Case cdTypeTable_BaremeAnneeStatutaire
        Dim myTable As String
        'myTable = m_dataHelper.GetParameter("SELECT NOMTABLE FROM listeTableLoi WHERE TABLECLE = " & cboListeTable.ItemData(cboListeTable.ListIndex)) & " ORDER BY StatId"
        myTable = m_dataHelper.GetParameter("SELECT NOMTABLE FROM listeTableLoi WHERE TABLECLE = " & cboListeTable.ItemData(cboListeTable.ListIndex))
        
        rq = "SELECT * FROM " & myTable & " ORDER BY StatId"
        
        If m_dataHelper.GetParameterAsLong("SELECT Count(*) FROM " & myTable) > 1000 Then
          MsgBox "Cette table peut contenir plusieurs milliers de lignes - seules les 500 premières lignes seront affichées !"
          
          rq = "SELECT top 500 * FROM " & myTable & " ORDER BY StatId"
        End If
      
      Case Else
        ' table a part entiere
        rq = "SELECT * FROM " & m_dataHelper.GetParameter("SELECT NOMTABLE FROM listeTableLoi WHERE TABLECLE = " & cboListeTable.ItemData(cboListeTable.ListIndex)) & " ORDER BY Age"
    End Select
    
    On Error Resume Next
    dtaListeTable.RecordSource = m_dataHelper.ValidateSQL(rq)
    dtaListeTable.Refresh
    
    Set vaSpread1.DataSource = dtaListeTable
    
    If Not dtaListeTable.Recordset.EOF Then
      dtaListeTable.Recordset.MoveLast
      dtaListeTable.Recordset.MoveFirst
      
      vaSpread1.MaxRows = dtaListeTable.Recordset.RecordCount
    Else
      vaSpread1.MaxRows = 0
    End If
    
    LargeurMaxColonneSpread vaSpread1
    
    If typeDeTable = cdTypeTableMortalite Or typeDeTable = cdTypeTableGeneration _
       Or typeDeTable = cdTypeTable_MortaliteIncap Or typeDeTable = cdTypeTable_MortaliteInval Then
      vaSpread1.ColWidth(1) = 0 ' on cache NomTable ou CleTable
    End If
    
    vaSpread1.ReDraw = True
    
    If 0 <> m_dataHelper.GetParameterAsDouble("SELECT TableUtilisateur FROM ListeTableLoi WHERE TABLECLE=" & cboListeTable.ItemData(cboListeTable.ListIndex)) Then
      btnDelete.Enabled = True
    Else
      btnDelete.Enabled = False
    End If
    
    On Error GoTo 0
  
  End If
End Sub


'##ModelId=5C8A67B2035B
Private Sub ShowDouble(f As ADODB.field, t As TextBox)
  If Not IsNull(f) Then
    Dim d As Double
    
    d = m_dataHelper.GetDouble2(f.Value)
    If d <= 1 Then
      t = d * 100
    Else
      t = d
    End If
  Else
    t = ""
  End If
End Sub

'##ModelId=5C8A67B2038A
Private Sub cboTableTauxProvision_Click()
  Screen.MousePointer = vbHourglass
    
  vaSpread2.ReDraw = False
    
  m_dataSource.SetDatabase dtaProvision
  dtaProvision.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM " & cboTableTauxProvision & " WHERE Age > 0")
  
  dtaProvision.Refresh
  
  Set vaSpread2.DataSource = dtaProvision

  If Not dtaProvision.Recordset.EOF Then
    dtaProvision.Recordset.MoveLast
    dtaProvision.Recordset.MoveFirst
    
    vaSpread2.MaxRows = dtaProvision.Recordset.RecordCount
  Else
    vaSpread2.MaxRows = 0
  End If
  
  ' rempli les cases taux
  Dim rs As ADODB.Recordset
  
'  Set rs = m_dataSource.OpenRecordset("SELECT Anc5, Anc6, Anc7, Anc8, Anc10 FROM " & cboTableTauxProvision & " WHERE Age = -1", Snapshot)
'
'  If Not rs.EOF Then
'    If rs.fields("Anc10") <> "Invalidité" Then
'      Call ShowDouble(rs.fields("Anc5"), Text2)
'      Call ShowDouble(rs.fields("Anc6"), Text1)
'    Else
'      Call ShowDouble(rs.fields("Anc7"), Text2)
'      Call ShowDouble(rs.fields("Anc8"), Text1)
'    End If
'  Else
'    Text2 = ""
'    Text1 = ""
'  End If

  txtCommentProvision.text = m_dataHelper.GetParameter("SELECT Comment FROM " & cboTableTauxProvision & " WHERE Age = -2") & vbCrLf _
                             & m_dataHelper.GetParameter("SELECT Comment FROM " & cboTableTauxProvision & " WHERE Age = -1")
  
  LargeurMaxColonneSpread vaSpread2
    
  vaSpread2.ReDraw = True
  
  Screen.MousePointer = vbDefault
End Sub


'##ModelId=5C8A67B2039A
Private Sub btnSave2_Click()
  btnSave_Click
End Sub

'##ModelId=5C8A67B203B9
Private Sub Command2_Click()
  Unload Me
End Sub


'##ModelId=5C8A67B203C9
Private Sub Form_Load()
  ' selectionne le premier tab
  
  'hide the Risque Stat tab
  SSTab1.TabVisible(6) = False
  
  SSTab1.Tab = 0
  Call SSTab1_Click(0) ' valeurs par defaut
  
  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  
  ProgressBar1.Visible = False
  
  ' init de la table des tables de la page "Tables diverses"
  InitTableDiverse
  
  ' liste des parametres de calcul
  rdoBilan.Value = True
  RefreshListParamCalculIni
  
  'create a connection to the Archives DB
  If DatabaseFileNameArchive <> "" Then
    If Not CreateArchiveConnection Then
      'problem creating connection to Archive DB
      MsgBox "Impossible d'ouvrir la base de données Archive!" & vbLf & "Source: frmParametre.Form_Load" _
      & vbLf & "Connection : " & DatabaseFileNameArchive, vbCritical
    Else
      If Not m_dataSourceArchive Is Nothing Then
        If m_dataSourceArchive.Connected Then
          archiveOk = True
        Else
          archiveOk = False
        End If
      End If
    End If
  End If

End Sub

'##ModelId=5C8A67B203D8
Private Sub Form_Resize()
  ' gna gna , gna
  
End Sub

'##ModelId=5C8A67B30000
Private Sub EnableCoeffBCAC()
  txtTauxTechnicCalculDC.Enabled = rdoCalculCoeffBCAC = True
  txtFraisGestionCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableIncapCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableInvalCalculDC.Enabled = rdoCalculCoeffBCAC = True
  
  cboTableIncapPrecalculDC.Enabled = rdoLireCoeffBCAC = True
  cboTableInvalPrecalculDC.Enabled = rdoLireCoeffBCAC = True
End Sub

'##ModelId=5C8A67B3001F
Private Sub Form_Unload(Cancel As Integer)
  Set fmTableDiverse = Nothing
  
  CloseArchiveConnection
End Sub

'##ModelId=5C8A67B3003F
Private Sub rdoCalculCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

'##ModelId=5C8A67B3004E
Private Sub rdoLireCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

'##ModelId=5C8A67B3005E
Private Sub SSTab1_Click(PreviousTab As Integer)
  Dim nomTable As String
  Dim sel As Long
  Dim rs As ADODB.Recordset
  
  ' provision
  Select Case SSTab1.Tab
    ' Incap / Inval
    Case 0
      txtNbDecimalePM = NbDecimalePM
      txtNbDecimaleCalcul = NbDecimaleCalcul
      
      'valeurs par defaut
      
      ' rempli les combo de liste des tables
      
      ' rempli le combo incap
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiIncapacite", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      m_dataHelper.FillCombo cboTableIncap, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 1", sel, False, True
      If sel = -1 Then
        cboTableIncap.ListIndex = 0
      End If
      
      ' passage
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiPassage", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTablePassage, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 2", sel, False, True
      If sel = -1 Then
        cboTablePassage.ListIndex = 0
      End If
      
      ' invalidite
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiInvalidite", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTableInval, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 3", sel, False, True
      If sel = -1 Then
        cboTableInval.ListIndex = 0
      End If
      
      txtTauxTechniqueIncap = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxIncapacite", "3.5"))
      txtFraisGestionIncap = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionIncapacite", "3.5"))
      txtTauxTechniqueInval = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxInvalidité", "3.5"))
      txtFraisGestionInval = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionInvalidite", "3.5"))
      txtRetraite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeRetraite", "65"))
      Txtduree = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "DureeIncap", "36"))
      TxtFranchise = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FranchiseIncap", "0"))
  
    ' Loi de maintien
    Case 3
      ' combo liste des tables
      m_dataHelper.FillCombo cboListeTable, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE<> " & cdTypeTableCoeffBCACIncap & " AND TYPETABLE <> " & cdTypeTableCoeffBCACInval & " ORDER BY LIBTABLE", -1, False, True
      
      If cboListeTable.ListCount <> 0 Then cboListeTable.ListIndex = 0
    
    ' provision
    Case 5
      Dim oCat As New ADOX.Catalog
      Dim oTable As ADOX.table
      
      txtCommentProvision.text = ""
      vaSpread2.MaxCols = 0
      vaSpread2.MaxRows = 0
      
      oCat.ActiveConnection = m_dataSource.Connection
      
      ' combo liste des table de provision precalculé
      cboTableTauxProvision.Clear
      
      For Each oTable In oCat.Tables
        If Left(oTable.Name, 5) = "PROV_" Then
          cboTableTauxProvision.AddItem oTable.Name
        End If
        If cboTableTauxProvision.ListCount <> 0 Then cboTableTauxProvision.ListIndex = 0
      Next
      
      Set oCat = Nothing
     
  
    ' Paramêtre Rentes
    Case 1
      ' rente conjoint
      nomTable = GetSettingIni(CompanyName, SectionName, "RenteConjoint", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTableRenteConjoint, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel, False, True
      If sel = -1 Then
        cboTableRenteConjoint.ListIndex = 0
      End If
      
      ' rente education
      nomTable = GetSettingIni(CompanyName, SectionName, "RenteEducation", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTableRenteEducation, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel, False, True
      If sel = -1 Then
        cboTableRenteEducation.ListIndex = 0
      End If
      
      ' taux pour rente
      txtTauxTechniqueRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxTechniqueRenteConjoint", "3.5"))
      txtFraisGestionRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionRenteConjoint", "3.5"))
      txtTauxTechniqueRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxTechniqueRenteEducation", "3.5"))
      txtFraisGestionRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionRenteEducation", "3.5"))
    
      txtTxIndex = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxIndexation", "3"))
      txtDureeIndex = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "DureeIndexation", "7"))
      txtTMO = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TMO", "1"))
      
      ' options pour rentes
      rdoAnnuelConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AnnuelConjoint", "0"))
      rdoSemestrielConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "SemestrielConjoint", "0"))
      rdoTrimestrielConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TrimestrielConjoint", "1"))
      rdoMensuelConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "MensuelConjoint", "0"))
      rdoPaiementAvanceConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "PaiementAvanceConjoint", "0"))
      rdoPaiementEchuConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "PaiementEchuConjoint", "1"))
    
      rdoAnnuelEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AnnuelEducation", "0"))
      rdoSemestrielEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "SemestrielEducation", "0"))
      rdoTrimestrielEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TrimestrielEducation", "1"))
      rdoMensuelEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "MensuelEducation", "0"))
      rdoPaiementAvanceEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "PaiementAvanceEducation", "0"))
      rdoPaiementEchuEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "PaiementEchuEducation", "1"))

    ' CATR9
    Case 4
      fmTableDiverse.FillCombo cboCATR9
      
    'Statutaire
'    Case 6
'
'      Dim annBar As String
'      Dim listAnnBar As String
'      Dim arrAnnBar() As String
'      Dim cnt As Integer
'      Dim ind As Integer
'
'      txtAgeRet = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatAgeRetraite", "62"))
'      txtAgeMin = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatAgeMin", "62"))
'      txtAgeMax = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatAgeMax", "62"))
'      txtMOMaxSemaine = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatMOMaxSemaine", "52"))
'      txtCLMMaxSemaine = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatCLMMaxSemaine", "156"))
'      txtCLDMaxSemaine = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatCLDMaxSemaine", "252"))
'      txtMATMaxSemaine = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatMATMaxSemaine", "366"))
'      txtATMaxSemaine = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatATMaxSemaine", "23"))
'
'      annBar = GetSettingIni(CompanyName, SectionName, "StatAnneeBareme", "2004")
'      listAnnBar = GetSettingIni(CompanyName, SectionName, "StatAnneeBaremeList", "2004,2015")
'
'      cmbAnneeBareme.Clear
'
'      arrAnnBar = Split(listAnnBar, ",")
'      For cnt = 0 To UBound(arrAnnBar)
'        cmbAnneeBareme.AddItem arrAnnBar(cnt)
'        'cmbAnneeBareme.ItemData(cmbAnneeBareme.NewIndex) = cnt
'        If arrAnnBar(cnt) = annBar Then
'          ind = cnt
'        End If
'      Next
'
'      cmbAnneeBareme.ListIndex = ind
   

    ' maintien deces
    Case 2
      txtFraisGestionCapitauxDecesDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionCapitauxDecesDC", "3.5"))
      txtFraisGestionRenteEducationDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionRenteEducationDC", "3.5"))
      
      txtFraisGestionRenteConjointDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionRenteConjointDC", "3.5"))
      txtCapitalMoyenRenteConjointTempoDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "CapitalMoyenRenteConjointTempoDC", ""))
      txtCapitalMoyenRenteConjointViagereDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "CapitalMoyenRenteConjointViagereDC", ""))
      chkForcerCapitalMoyenRteConjoitDC.Value = IIf(m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "ForcerCapitalMoyenRteConjoitDC", "0")) = 0, vbUnchecked, vbChecked)
      txtAgeConjointRenteConjointDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeConjointRenteConjointDC", "2"))
      
      rdoSansLissage = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "LissageProvision", "1")) = 0
      rdoAvecLissage = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "LissageProvision", "1")) = 1
      
      '
      ' recalcul
      '
      
      chkPMGDForcerInval = IIf(GetSettingIni(CompanyName, SectionName, "PMGDForcerInval", "0") = "0", vbUnchecked, vbChecked)
      rdoCalculCoeffBCAC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "RecalculBCAC", "0")) = 1
      
      rdoCapitauxConstitif = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "MethodeCalcul", "0")) = 0
      rdoCotisationsExonerees = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "MethodeCalcul", "0")) = 1
  
      ' rempli le combo incap
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiIncapaciteDC", "#")
      If nomTable <> "#" And IsNumeric(nomTable) Then
        sel = CLng(nomTable)
      Else
        sel = -1
      End If
      m_dataHelper.FillCombo cboTableIncapCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteIncap, sel, False, True
      If sel = -1 Then
        cboTableIncapCalculDC.ListIndex = 0
      End If
      
      ' rempli le combo invalidite
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiInvaliditeDC", "#")
      If nomTable <> "#" And IsNumeric(nomTable) Then
        sel = CLng(nomTable)
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTableInvalCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteInval, sel, False, True
      If sel = -1 Then
        cboTableInvalCalculDC.ListIndex = 0
      End If
      
      txtTauxTechnicCalculDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "TauxTechniqueDC", "3.5"))
      txtFraisGestionCalculDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "FraisGestionDC", "3.5"))
      
      txtAgeLimiteCalulDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeLimiteCalulDC", "60"))
      
      '
      ' lecture
      '
      rdoLireCoeffBCAC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "RecalculBCAC", "0")) = 0
      
      ' rempli le combo incap
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiIncapacitePrecalculDC", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      m_dataHelper.FillCombo cboTableIncapPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACIncap, sel, False, True
      If sel = -1 Then
        cboTableIncapPrecalculDC.ListIndex = 0
      End If
      
      ' rempli le combo invalidite
      nomTable = GetSettingIni(CompanyName, SectionName, "LoiInvaliditePrecalculDC", "#")
      If nomTable <> "#" Then
        Set rs = m_dataSource.OpenRecordset("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """", Snapshot)
        If Not rs.EOF Then
          sel = rs.fields(0)
        End If
        rs.Close
      Else
        sel = -1
      End If
      
      m_dataHelper.FillCombo cboTableInvalPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval, sel, False, True
      If sel = -1 Then
        cboTableInvalPrecalculDC.ListIndex = 0
      End If
      
      ' table de coefficients précalculés
      m_dataHelper.FillCombo cboBCAC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval & " OR TYPETABLE = " & cdTypeTableCoeffBCACIncap, -1, False, True
      If cboBCAC.ListCount <> 0 Then
        cboBCAC.ListIndex = 0
      End If
      
      btnDelBCAC.Enabled = DroitAdmin
      btnImportBCAC.Enabled = DroitAdmin
  End Select
End Sub

'##ModelId=5C8A67B3008C
Private Sub txtNbDecimaleCalcul_Change()
  setNbDecimalePM
End Sub

'##ModelId=5C8A67B300AC
Private Sub txtNbDecimalePM_Change()
  setNbDecimalePM
End Sub

'##ModelId=5C8A67B300BB
Private Sub setNbDecimalePM()
   NbDecimalePM = val(txtNbDecimalePM)
   'NbDecimaleCalcul = Val(txtNbDecimaleCalcul)
End Sub

'##ModelId=5C8A67B300CB
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' tagada
End Sub

'##ModelId=5C8A67B30109
Private Sub vaSpread2_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' tagada
End Sub

'##ModelId=5C8A67B30148
Private Sub SetNumParamCalcul()
  NumParamCalcul = 0
  
  If lvParamCalcul.SelectedItem Is Nothing Then Exit Sub
  
  NumParamCalcul = CLng(lvParamCalcul.SelectedItem.SubItems(1))
End Sub

'##ModelId=5C8A67B30158
Private Sub RefreshListParamCalculIni()
  Dim itmX As ListItem, typeParam As String
  Dim clmX As ColumnHeader
  
  ' sauvegarde la selection
  Dim curSel As Long
  
  If lvParamCalcul.SelectedItem Is Nothing Then
    curSel = -1
  Else
    curSel = lvParamCalcul.SelectedItem.Index
  End If
  
  ' remplissage de la liste avec les infos du .INI
  lvParamCalcul.View = lvwReport   'Determination de l'affichage sous forme de liste
  lvParamCalcul.ListItems.Clear     'Suppression des elements de la liste
  lvParamCalcul.ColumnHeaders.Clear 'Suppression des colonnes
  
  'Creation des colonnes
  Set clmX = lvParamCalcul.ColumnHeaders.Add(, , "Code", 100)
  Set clmX = lvParamCalcul.ColumnHeaders.Add(, , "Code", 100)
  Set clmX = lvParamCalcul.ColumnHeaders.Add(, , "Nom", 100)

  
  ' type de paramètre
  typeParam = GetTypeParam
  

  ' Liste des paramètres de calcul
'  Dim i As Integer
  
'  For i = 1 To 255
'    If GetSettingIni(CompanyName, DEFAULT_PARAM_SECTION & typeParam & "_" & i, "NumParamCalcul", "QQQQ") <> "QQQQ" Then
'      Set itmX = lvParamCalcul.ListItems.Add(, , CStr(i))
'      itmX.SubItems(1) = CStr(i)
'      itmX.SubItems(2) = GetSettingIni(CompanyName, DEFAULT_PARAM_SECTION & typeParam & "_" & i, "Nom", "")
'
'      If i = NumParamCalcul Then
'        itmX.Selected = True
'        itmX.EnsureVisible
'      End If
'    End If
'  Next
  
  Dim aRet() As String
  
  If EnumSections(aRet, sFichierIni) Then
    Dim i As Integer, num As Long
    Dim sRet As String, sName As String
    
    sName = DEFAULT_PARAM_SECTION & typeParam & "_"
    
    sRet = vbNullString
    For i = LBound(aRet) To UBound(aRet)
      If Len(aRet(i)) > Len(sName) Then
        If UCase(Left(aRet(i), Len(sName))) = UCase(sName) Then
          num = CLng(mID(aRet(i), Len(sName) + 1))
          
          Set itmX = lvParamCalcul.ListItems.Add(, , CStr(num))
          itmX.SubItems(1) = CStr(num)
          itmX.SubItems(2) = GetSettingIni(CompanyName, aRet(i), "Nom", "")
          
          If num = NumParamCalcul Then
            itmX.Selected = True
            itmX.EnsureVisible
          End If
        End If
      End If
    Next
  
  End If
  
  Erase aRet
  
  LargeurAutomatique Me, lvParamCalcul
  
  lvParamCalcul.ColumnHeaders(1).Width = 0
  lvParamCalcul.ColumnHeaders(2).Alignment = lvwColumnCenter
  lvParamCalcul.ColumnHeaders(2).Width = 600
  
  lvParamCalcul.ColumnHeaders(3).Width = lvParamCalcul.ColumnHeaders(3).Width + 100 'lvParamCalcul.Width - 100 - lvParamCalcul.ColumnHeaders(2).Width
  
  ' restaure la selection
  If curSel <> -1 And curSel <= lvParamCalcul.ListItems.Count Then
    lvParamCalcul.ListItems(curSel).Selected = True
    lvParamCalcul.ListItems(curSel).EnsureVisible
  End If

End Sub

