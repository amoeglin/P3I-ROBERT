VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmParamCalcul 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramètres de calcul"
   ClientHeight    =   10125
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmParamCalcul.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   11415
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
      Height          =   9090
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   16034
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
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
      TabPicture(0)   =   "frmParamCalcul.frx":1BB2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabels(6)"
      Tab(0).Control(1)=   "lblLabels(1)"
      Tab(0).Control(2)=   "lblLabels(3)"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(6)=   "txtFields(6)"
      Tab(0).Control(7)=   "Frame19"
      Tab(0).Control(8)=   "txtNomParamCalcul"
      Tab(0).Control(9)=   "txtCode"
      Tab(0).Control(10)=   "Frame11"
      Tab(0).Control(11)=   "maintiendep"
      Tab(0).Control(12)=   "maintienViag"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Rentes && Revalorisations"
      TabPicture(1)   =   "frmParamCalcul.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Maintien Garanties Décès"
      TabPicture(2)   =   "frmParamCalcul.frx":1BEA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame15"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame16"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame17"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Risque Statutaire"
      TabPicture(3)   =   "frmParamCalcul.frx":1C06
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame23"
      Tab(3).Control(1)=   "Frame22"
      Tab(3).Control(2)=   "Frame18"
      Tab(3).ControlCount=   3
      Begin VB.Frame maintienViag 
         Caption         =   "Loi de maintien Invalidité viagère "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69285
         TabIndex        =   218
         Top             =   8020
         Width           =   5460
         Begin VB.ComboBox cboTableViagere 
            Height          =   315
            Left            =   1920
            TabIndex        =   219
            Text            =   "cboTableViag"
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label Label34 
            Caption         =   "Table invalidité viagère"
            Height          =   255
            Left            =   120
            TabIndex        =   220
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Durée d'indemnisation en semaines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   -74640
         TabIndex        =   199
         Top             =   4200
         Width           =   10575
         Begin VB.TextBox txtCLM 
            Height          =   330
            Left            =   3360
            TabIndex        =   211
            Text            =   "156"
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txtCLD 
            Height          =   330
            Left            =   3360
            TabIndex        =   210
            Text            =   "252"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtMAT 
            Height          =   330
            Left            =   3360
            TabIndex        =   209
            Text            =   "23"
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox txtAT 
            Height          =   330
            Left            =   3360
            TabIndex        =   208
            Text            =   "366"
            Top             =   2280
            Width           =   2895
         End
         Begin VB.TextBox txtMO 
            Height          =   330
            Left            =   3360
            TabIndex        =   207
            Text            =   "52"
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label31 
            Caption         =   "AT - Accident du travail"
            Height          =   255
            Left            =   360
            TabIndex        =   217
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label30 
            Caption         =   "MAT - Maternité"
            Height          =   255
            Left            =   360
            TabIndex        =   216
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label Label29 
            Caption         =   "CLD - Congé de longue durée"
            Height          =   255
            Left            =   360
            TabIndex        =   215
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Label28 
            Caption         =   "CLM - Congé de longue maladie"
            Height          =   255
            Left            =   360
            TabIndex        =   214
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label27 
            Caption         =   "MO - Maladie ordinaire"
            Height          =   255
            Left            =   360
            TabIndex        =   213
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Barème"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74640
         TabIndex        =   198
         Top             =   1920
         Width           =   10575
         Begin VB.ComboBox cmbYear 
            Height          =   315
            Left            =   3360
            TabIndex        =   212
            Text            =   "Combo1"
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtAgeMax 
            Height          =   330
            Left            =   3360
            TabIndex        =   206
            Text            =   "64"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.TextBox txtAgeMin 
            Height          =   330
            Left            =   3360
            TabIndex        =   205
            Text            =   "20"
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label26 
            Caption         =   "Age maximum"
            Height          =   375
            Left            =   360
            TabIndex        =   204
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Age minimum"
            Height          =   255
            Left            =   360
            TabIndex        =   203
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label18 
            Caption         =   "Année de référence"
            Height          =   255
            Left            =   360
            TabIndex        =   202
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Age de départ en retraite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   197
         Top             =   720
         Width           =   10575
         Begin VB.TextBox txtAgeRet 
            Height          =   330
            Left            =   3360
            TabIndex        =   201
            Text            =   "62"
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label17 
            Caption         =   "Age retraite"
            Height          =   255
            Left            =   360
            TabIndex        =   200
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame maintiendep 
         Caption         =   "Loi de maintien Dépendance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -69285
         TabIndex        =   179
         Top             =   7080
         Width           =   5460
         Begin VB.ComboBox cboTableDependance 
            Height          =   315
            Left            =   1680
            TabIndex        =   181
            Text            =   "cboTableDependance"
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "Table dépendance"
            Height          =   255
            Left            =   120
            TabIndex        =   180
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Salariés"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -65730
         TabIndex        =   146
         Top             =   1845
         Width           =   1905
         Begin VB.CheckBox chkPortefeuilleSalaries 
            Caption         =   "Portefeuille ""Salariés"""
            Height          =   420
            Left            =   360
            TabIndex        =   147
            Top             =   360
            Width           =   1320
         End
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   -73410
         MaxLength       =   60
         TabIndex        =   3
         Top             =   540
         Width           =   1020
      End
      Begin VB.Frame Frame5 
         Caption         =   "Reprise de revalorisation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   -74595
         TabIndex        =   75
         Top             =   3690
         Width           =   5100
         Begin VB.TextBox txtDureeLissage 
            Height          =   285
            Left            =   4140
            TabIndex        =   86
            Top             =   1425
            Width           =   495
         End
         Begin VB.TextBox txtTMO 
            Height          =   285
            Left            =   1575
            TabIndex        =   83
            Top             =   1440
            Width           =   675
         End
         Begin VB.TextBox txtTxIndex 
            Height          =   285
            Left            =   1575
            TabIndex        =   77
            Top             =   630
            Width           =   675
         End
         Begin VB.TextBox txtDureeIndex 
            Height          =   285
            Left            =   4140
            TabIndex        =   80
            Top             =   630
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Indexation (Incapacité, Invalidité, Rentes (cjt, éduc, autres rentes)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   90
            TabIndex        =   150
            Top             =   315
            Width           =   4755
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cotisation revalo (Incapacité, Invalidité, Rentes)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   149
            Top             =   1125
            Width           =   3435
         End
         Begin VB.Label Label3 
            Caption         =   "Durée de lissage de la Prime Unique"
            Height          =   375
            Left            =   2760
            TabIndex        =   85
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   9
            Left            =   4680
            TabIndex        =   87
            Top             =   1485
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   8
            Left            =   4680
            TabIndex        =   81
            Top             =   675
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   7
            Left            =   2295
            TabIndex        =   78
            Top             =   675
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   6
            Left            =   2295
            TabIndex        =   84
            Top             =   1485
            Width           =   150
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux financier TME"
            Height          =   255
            Index           =   17
            Left            =   90
            TabIndex        =   82
            Top             =   1485
            Width           =   1410
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux d'indexation"
            Height          =   255
            Index           =   16
            Left            =   90
            TabIndex        =   76
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Durée d'indexation"
            Height          =   255
            Index           =   2
            Left            =   2745
            TabIndex        =   79
            Top             =   675
            Width           =   1365
         End
      End
      Begin VB.TextBox txtNomParamCalcul 
         Height          =   285
         Left            =   -70800
         MaxLength       =   60
         TabIndex        =   5
         Top             =   540
         Width           =   6870
      End
      Begin VB.Frame Frame19 
         Caption         =   "Calul de l'âge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -74865
         TabIndex        =   8
         Top             =   1845
         Width           =   5460
         Begin VB.CheckBox chkBridageAge 
            Caption         =   "Age bridé aux limites de table"
            Height          =   195
            Left            =   1350
            TabIndex        =   11
            Top             =   675
            Width           =   2400
         End
         Begin VB.OptionButton rdoAgeAnniversaire 
            Caption         =   "Anniversaire"
            Height          =   240
            Left            =   3195
            TabIndex        =   10
            Top             =   315
            Width           =   2175
         End
         Begin VB.OptionButton rdoAgeMillesime 
            Caption         =   "par différence de Millesime"
            Height          =   240
            Left            =   360
            TabIndex        =   9
            Top             =   315
            Width           =   2220
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Rente Education"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   315
         TabIndex        =   136
         Top             =   4410
         Width           =   5100
         Begin VB.TextBox txtFraisGestionRenteEducationDC 
            Height          =   330
            Left            =   1485
            TabIndex        =   137
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label10 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   135
            TabIndex        =   139
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   138
            Top             =   315
            Width           =   240
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Méthode de calcul des Provisions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   315
         TabIndex        =   135
         Top             =   5700
         Width           =   5100
         Begin VB.OptionButton rdoDCAucun 
            Caption         =   "Aucun"
            Height          =   375
            Left            =   135
            TabIndex        =   145
            Top             =   360
            Width           =   780
         End
         Begin VB.TextBox txtPctPMCalculeeDC 
            Height          =   330
            Left            =   4455
            TabIndex        =   105
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.OptionButton rdoCapitauxConstitif 
            Caption         =   "Capitaux constitutifs sous risque"
            Height          =   375
            Left            =   1035
            TabIndex        =   103
            Top             =   360
            Width           =   1770
         End
         Begin VB.OptionButton rdoPctPMCalculeeDC 
            Caption         =   "% de la Provision calculée"
            Height          =   375
            Left            =   2835
            TabIndex        =   104
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame13 
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
         Height          =   1500
         Left            =   315
         TabIndex        =   127
         Top             =   2400
         Width           =   5100
         Begin VB.TextBox txtFraisGestionRenteConjointDC 
            Height          =   330
            Left            =   1665
            TabIndex        =   96
            Text            =   "0"
            Top             =   315
            Width           =   465
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointTempoDC 
            Height          =   330
            Left            =   2205
            TabIndex        =   99
            Text            =   "4.5"
            Top             =   945
            Width           =   915
         End
         Begin VB.CheckBox chkForcerCapitalMoyenRteConjoitDC 
            Caption         =   "Forcer"
            Height          =   195
            Left            =   315
            TabIndex        =   98
            Top             =   990
            Width           =   780
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointViagereDC 
            Height          =   330
            Left            =   3960
            TabIndex        =   100
            Text            =   "4.5"
            Top             =   945
            Width           =   915
         End
         Begin VB.TextBox txtAgeConjointRenteConjointDC 
            Height          =   330
            Left            =   4095
            TabIndex        =   97
            Text            =   "0"
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label32 
            Caption         =   "%"
            Height          =   240
            Left            =   2205
            TabIndex        =   134
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label33 
            Caption         =   "Frais de gestion"
            Height          =   195
            Left            =   135
            TabIndex        =   133
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Capital constitutif moyen"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   132
            Top             =   720
            Width           =   1770
         End
         Begin VB.Label Label39 
            Caption         =   "Temporaire"
            Height          =   195
            Left            =   1305
            TabIndex        =   131
            Top             =   990
            Width           =   825
         End
         Begin VB.Label Label40 
            Caption         =   "Viagère"
            Height          =   195
            Left            =   3240
            TabIndex        =   130
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label38 
            Caption         =   "ans"
            Height          =   240
            Left            =   4590
            TabIndex        =   129
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label37 
            Caption         =   "Age du conjoint : +/-"
            Height          =   375
            Left            =   2565
            TabIndex        =   128
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Coefficients de provisions"
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
         Height          =   7575
         Left            =   5640
         TabIndex        =   118
         Top             =   600
         Width           =   5235
         Begin VB.TextBox txtAgeLimiteCalulDC_Retraite_Inval1 
            Height          =   330
            Left            =   3960
            TabIndex        =   227
            Text            =   "0"
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox txtAgeLimiteCalulDC_Inval1 
            Height          =   330
            Left            =   1800
            TabIndex        =   225
            Text            =   "0"
            Top             =   1800
            Width           =   465
         End
         Begin VB.TextBox txtAgeLimiteCalulDC_Retraite 
            Height          =   330
            Left            =   3960
            TabIndex        =   176
            Text            =   "0"
            Top             =   1350
            Width           =   465
         End
         Begin VB.ComboBox cboTableIncapCalculDC_Retraite 
            Height          =   315
            Left            =   1485
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   4035
            Width           =   3525
         End
         Begin VB.ComboBox cboTableInvalCalculDC_Retraite 
            Height          =   315
            Left            =   1485
            Style           =   2  'Dropdown List
            TabIndex        =   170
            Top             =   4440
            Width           =   3525
         End
         Begin VB.ComboBox cboTableIncapPrecalculDC_Retraite 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   6690
            Width           =   3525
         End
         Begin VB.ComboBox cboTableInvalPrecalculDC_Retraite 
            Height          =   315
            Left            =   1395
            Style           =   2  'Dropdown List
            TabIndex        =   164
            Top             =   7095
            Width           =   3525
         End
         Begin VB.TextBox txtAgeLimiteCalulDC 
            Height          =   330
            Left            =   1755
            TabIndex        =   110
            Text            =   "0"
            Top             =   1350
            Width           =   465
         End
         Begin VB.CheckBox chkPMGDForcerInval 
            Caption         =   "Forcer en Invalidité si Anc > 36 mois et pas de passage"
            Height          =   285
            Left            =   495
            TabIndex        =   106
            Top             =   270
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.ComboBox cboTableInvalPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   115
            Top             =   5925
            Width           =   3525
         End
         Begin VB.ComboBox cboTableIncapPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   5520
            Width           =   3525
         End
         Begin VB.ComboBox cboTableInvalCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   112
            Top             =   3315
            Width           =   3525
         End
         Begin VB.TextBox txtTauxTechnicCalculDC 
            Height          =   330
            Left            =   1755
            TabIndex        =   108
            Text            =   "0"
            Top             =   945
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionCalculDC 
            Height          =   330
            Left            =   3960
            TabIndex        =   109
            Text            =   "0"
            Top             =   945
            Width           =   465
         End
         Begin VB.ComboBox cboTableIncapCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   2910
            Width           =   3525
         End
         Begin VB.OptionButton rdoLireCoeffBCAC 
            Caption         =   "Utiliser les tables de coefficients précalculés du BCAC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   113
            Top             =   4890
            Width           =   4965
         End
         Begin VB.OptionButton rdoCalculCoeffBCAC 
            Caption         =   "Calcul du coefficient"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   135
            TabIndex        =   107
            Top             =   630
            Width           =   2580
         End
         Begin VB.Label Label41 
            Caption         =   "ans"
            Height          =   240
            Left            =   4560
            TabIndex        =   229
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label5 
            Caption         =   "Après réforme"
            Height          =   240
            Index           =   15
            Left            =   2760
            TabIndex        =   228
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "ans"
            Height          =   240
            Left            =   2280
            TabIndex        =   226
            Top             =   1800
            Width           =   330
         End
         Begin VB.Label Label5 
            Caption         =   "Age limite Inv cat 1"
            Height          =   240
            Index           =   14
            Left            =   240
            TabIndex        =   224
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Après réforme"
            Height          =   240
            Index           =   13
            Left            =   2700
            TabIndex        =   178
            Top             =   1395
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "ans"
            Height          =   240
            Left            =   4500
            TabIndex        =   177
            Top             =   1395
            Width           =   330
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables après réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   315
            TabIndex        =   175
            Top             =   3720
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables avant réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   315
            TabIndex        =   174
            Top             =   2595
            Width           =   3435
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Incap"
            Height          =   240
            Index           =   12
            Left            =   360
            TabIndex        =   173
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Inval"
            Height          =   375
            Index           =   11
            Left            =   360
            TabIndex        =   172
            Top             =   4485
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables après réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   315
            TabIndex        =   169
            Top             =   6375
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables avant réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   315
            TabIndex        =   168
            Top             =   5205
            Width           =   3435
         End
         Begin VB.Label Label5 
            Caption         =   "Table Incap"
            Height          =   375
            Index           =   10
            Left            =   315
            TabIndex        =   167
            Top             =   6735
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Table Inval"
            Height          =   240
            Index           =   9
            Left            =   315
            TabIndex        =   166
            Top             =   7140
            Width           =   960
         End
         Begin VB.Label Label6 
            Caption         =   "ans"
            Height          =   240
            Left            =   2295
            TabIndex        =   144
            Top             =   1395
            Width           =   330
         End
         Begin VB.Label Label5 
            Caption         =   "Age limite"
            Height          =   240
            Index           =   8
            Left            =   255
            TabIndex        =   143
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Table Inval"
            Height          =   240
            Index           =   7
            Left            =   315
            TabIndex        =   126
            Top             =   5970
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Table Incap"
            Height          =   375
            Index           =   6
            Left            =   315
            TabIndex        =   125
            Top             =   5565
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Inval"
            Height          =   375
            Index           =   5
            Left            =   315
            TabIndex        =   124
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Incap"
            Height          =   240
            Index           =   4
            Left            =   315
            TabIndex        =   123
            Top             =   2955
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   240
            Left            =   4500
            TabIndex        =   122
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            Height          =   240
            Left            =   2295
            TabIndex        =   121
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Frais de gestion"
            Height          =   240
            Left            =   2700
            TabIndex        =   120
            Top             =   990
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   240
            Index           =   1
            Left            =   255
            TabIndex        =   119
            Top             =   990
            Width           =   1185
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Pourcentage de provisions à constituer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   315
         TabIndex        =   117
         Top             =   7440
         Width           =   5100
         Begin VB.OptionButton rdoSansLissage 
            Caption         =   "100%"
            Height          =   240
            Left            =   315
            TabIndex        =   101
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton rdoAvecLissage 
            Caption         =   "Utiliser la table 'LissageProvision'"
            Height          =   240
            Left            =   1575
            TabIndex        =   102
            Top             =   315
            Width           =   3120
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Capitaux Décès"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   315
         TabIndex        =   93
         Top             =   585
         Width           =   5100
         Begin VB.TextBox txtPourcentageAccident 
            Height          =   330
            Left            =   2790
            TabIndex        =   140
            Text            =   "0"
            Top             =   720
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionCapitauxDecesDC 
            Height          =   330
            Left            =   2790
            TabIndex        =   94
            Text            =   "0"
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label15 
            Caption         =   "Pourcentage Décès par Accident"
            Height          =   240
            Left            =   135
            TabIndex        =   142
            Top             =   765
            Width           =   2490
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   240
            Left            =   3330
            TabIndex        =   141
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "%"
            Height          =   240
            Left            =   3330
            TabIndex        =   116
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label8 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   135
            TabIndex        =   95
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2670
         Left            =   -69150
         TabIndex        =   44
         Top             =   765
         Visible         =   0   'False
         Width           =   5100
         Begin VB.ComboBox cboTableRenteConjoint 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   900
            Width           =   4695
         End
         Begin VB.TextBox txtFraisGestionRenteConjoint 
            Height          =   285
            Left            =   4230
            TabIndex        =   49
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueRenteConjoint 
            Height          =   285
            Left            =   1485
            TabIndex        =   46
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.Frame Frame7 
            Caption         =   "Paiement"
            Height          =   1140
            Left            =   225
            TabIndex        =   52
            Top             =   1350
            Width           =   1725
            Begin VB.OptionButton rdoPaiementAvanceConjoint 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   53
               Top             =   360
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementEchuConjoint 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   54
               Top             =   720
               Width           =   1275
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Fractionnement"
            Height          =   1140
            Left            =   2340
            TabIndex        =   55
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton rdoSemestrielConjoint 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   57
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelConjoint 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   56
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelConjoint 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   59
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoTrimestrielConjoint 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   58
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   2
            Left            =   225
            TabIndex        =   45
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label23 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   48
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label24 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   47
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label25 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   50
            Top             =   405
            Width           =   240
         End
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
         Height          =   2670
         Left            =   -74595
         TabIndex        =   60
         Top             =   765
         Width           =   5100
         Begin VB.ComboBox cboTableRenteEducation 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   900
            Width           =   4695
         End
         Begin VB.TextBox txtFraisGestionRenteEducation 
            Height          =   285
            Left            =   4230
            TabIndex        =   65
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueRenteEducation 
            Height          =   285
            Left            =   1485
            TabIndex        =   62
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.Frame Frame9 
            Caption         =   "Paiement"
            Height          =   1140
            Left            =   225
            TabIndex        =   68
            Top             =   1350
            Width           =   1725
            Begin VB.OptionButton rdoPaiementEchuEducation 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   70
               Top             =   720
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementAvanceEducation 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   69
               Top             =   360
               Width           =   1275
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Fractionnement"
            Height          =   1140
            Left            =   2340
            TabIndex        =   71
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton rdoTrimestrielEducation 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   74
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelEducation 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   88
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelEducation 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   72
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoSemestrielEducation 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   73
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   3
            Left            =   225
            TabIndex        =   61
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label21 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   64
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label20 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   63
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label19 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   66
            Top             =   405
            Width           =   240
         End
      End
      Begin VB.TextBox txtFields 
         Height          =   915
         Index           =   6
         Left            =   -73425
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   900
         Width           =   9480
      End
      Begin VB.Frame Frame2 
         Caption         =   "Loi de maintien Incapacité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5865
         Left            =   -74880
         TabIndex        =   19
         Top             =   3000
         Width           =   5460
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   222
            Top             =   3000
            Width           =   735
         End
         Begin VB.Frame Chômage 
            Caption         =   "Chômage"
            Height          =   615
            Left            =   150
            TabIndex        =   190
            Top             =   5160
            Width           =   5190
            Begin VB.TextBox TxChomage 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   195
               Top             =   240
               Width           =   495
            End
            Begin VB.TextBox DureeChomage 
               Height          =   285
               Left            =   3120
               TabIndex        =   194
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "%"
               Height          =   255
               Index           =   34
               Left            =   1920
               TabIndex        =   196
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lblLabels 
               Caption         =   "taux technique"
               Height          =   255
               Index           =   33
               Left            =   120
               TabIndex        =   193
               Top             =   240
               Width           =   1140
            End
            Begin VB.Label lblLabels 
               Caption         =   "mois"
               Height          =   255
               Index           =   32
               Left            =   3720
               TabIndex        =   192
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lblLabels 
               Caption         =   "durée"
               Height          =   255
               Index           =   31
               Left            =   2520
               TabIndex        =   191
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame20 
            Caption         =   "Interpolation / Lissage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1470
            Left            =   135
            TabIndex        =   186
            Top             =   3645
            Width           =   5190
            Begin VB.CheckBox chkAnnulisationPassage 
               Caption         =   "Calcul de l'annualisation réduite pour la provision de passage"
               Height          =   330
               Left            =   135
               TabIndex        =   189
               Top             =   900
               Width           =   4695
            End
            Begin VB.CheckBox chkInterpolationIncap 
               Caption         =   "Interpoler les coeffs de provision Incapacité et Passage"
               Height          =   195
               Left            =   135
               TabIndex        =   188
               Top             =   315
               Width           =   4335
            End
            Begin VB.CheckBox chkLissagePassage 
               Caption         =   "Ne pas lisser les coeffs de Passage (pas de terme correcteur)"
               Height          =   195
               Left            =   135
               TabIndex        =   187
               Top             =   630
               Width           =   4695
            End
         End
         Begin VB.ComboBox cboTablePassage_Retraite 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   152
            Top             =   2520
            Width           =   3840
         End
         Begin VB.ComboBox cboTableIncap_Retraite 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   2115
            Width           =   3840
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   10
            Left            =   4500
            TabIndex        =   28
            Top             =   315
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   9
            Left            =   1440
            TabIndex        =   25
            Top             =   315
            Width           =   675
         End
         Begin VB.ComboBox cboTableIncap 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   990
            Width           =   3840
         End
         Begin VB.ComboBox cboTablePassage 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   1395
            Width           =   3840
         End
         Begin VB.Label Label35 
            Caption         =   "% des incap avec passage en inval cat 1 en utilisant la nouvelle table Age Départ retraite spécifique Inval1 à 64 ans"
            Height          =   495
            Left            =   960
            TabIndex        =   223
            Top             =   3000
            Width           =   4335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux technique"
            Height          =   255
            Index           =   35
            Left            =   0
            TabIndex        =   221
            Top             =   -840
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables après réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   135
            TabIndex        =   156
            Top             =   1800
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tables avant réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   135
            TabIndex        =   155
            Top             =   675
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table incapacité"
            Height          =   255
            Index           =   20
            Left            =   90
            TabIndex        =   154
            Top             =   2160
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table de passage"
            Height          =   210
            Index           =   19
            Left            =   90
            TabIndex        =   153
            Top             =   2580
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   10
            Left            =   3195
            TabIndex        =   27
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux technique"
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   24
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table de passage"
            Height          =   210
            Index           =   8
            Left            =   90
            TabIndex        =   22
            Top             =   1455
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table incapacité"
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   20
            Top             =   1035
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   26
            Top             =   360
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   2
            Left            =   5220
            TabIndex        =   29
            Top             =   360
            Width           =   150
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Loi de maintien Invalidité"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4065
         Left            =   -69285
         TabIndex        =   30
         Top             =   3000
         Width           =   5460
         Begin VB.TextBox Inval1Cat 
            Height          =   285
            Left            =   4320
            TabIndex        =   183
            Top             =   1440
            Width           =   525
         End
         Begin VB.TextBox txtAgeLimiteInvalCat1 
            Height          =   285
            Left            =   2430
            TabIndex        =   182
            Top             =   1440
            Width           =   450
         End
         Begin VB.TextBox txtAgeLimiteInvalCat1_Retraite 
            Height          =   285
            Left            =   3510
            TabIndex        =   161
            Top             =   2565
            Width           =   450
         End
         Begin VB.ComboBox cboTableInval_Retraite 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Top             =   2160
            Width           =   3840
         End
         Begin VB.Frame Frame21 
            Caption         =   "Interpolation"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   135
            TabIndex        =   39
            Top             =   2925
            Width           =   5145
            Begin VB.OptionButton rdoInterpolationInval_Age 
               Caption         =   "Interpolation sur l'âge"
               Height          =   285
               Left            =   180
               TabIndex        =   42
               Top             =   585
               Width           =   1905
            End
            Begin VB.OptionButton rdoInterpolationInval_CorrectionDuree 
               Caption         =   "Correction si durée restante <12 mois"
               Height          =   285
               Left            =   2070
               TabIndex        =   41
               Top             =   270
               Width           =   2940
            End
            Begin VB.OptionButton rdoInterpolationInval_AgeDuree 
               Caption         =   "Interpolation sur l'âge et la durée"
               Height          =   285
               Left            =   2070
               TabIndex        =   43
               Top             =   585
               Width           =   2850
            End
            Begin VB.OptionButton rdoInterpolationInval_NON 
               Caption         =   "Aucune interpolation"
               Height          =   285
               Left            =   180
               TabIndex        =   40
               Top             =   270
               Width           =   1770
            End
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   13
            Left            =   4455
            TabIndex        =   37
            Top             =   315
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   12
            Left            =   1440
            TabIndex        =   34
            Top             =   315
            Width           =   675
         End
         Begin VB.ComboBox cboTableInval 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1035
            Width           =   3840
         End
         Begin VB.Label Label16 
            Caption         =   "%"
            Height          =   240
            Left            =   5040
            TabIndex        =   185
            Top             =   1485
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "ans  Proportion->"
            Height          =   240
            Index           =   10
            Left            =   2970
            TabIndex        =   184
            Top             =   1485
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Age limite pour les invalides en 1ere catégorie"
            Height          =   255
            Index           =   26
            Left            =   135
            TabIndex        =   163
            Top             =   2610
            Width           =   3390
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   11
            Left            =   4050
            TabIndex        =   162
            Top             =   2610
            Width           =   330
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table après réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   135
            TabIndex        =   160
            Top             =   1845
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table avant réforme des retraites"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   135
            TabIndex        =   159
            Top             =   720
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table invalidité"
            Height          =   255
            Index           =   23
            Left            =   135
            TabIndex        =   158
            Top             =   2205
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            Caption         =   "invalides 1ère cat ->Age limite "
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   148
            Top             =   1485
            Width           =   2220
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   13
            Left            =   3195
            TabIndex        =   36
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux technique"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   33
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table invalidité"
            Height          =   255
            Index           =   11
            Left            =   135
            TabIndex        =   31
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   3
            Left            =   2160
            TabIndex        =   35
            Top             =   360
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   4
            Left            =   5175
            TabIndex        =   38
            Top             =   360
            Width           =   150
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Etat réglementaire C7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   -69285
         TabIndex        =   12
         Top             =   1845
         Width           =   3435
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   15
            Left            =   2295
            TabIndex        =   14
            Top             =   270
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   14
            Left            =   2295
            TabIndex        =   17
            Top             =   630
            Width           =   675
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux des revenus financiers"
            Height          =   255
            Index           =   15
            Left            =   180
            TabIndex        =   13
            Top             =   315
            Width           =   2040
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   14
            Left            =   180
            TabIndex        =   16
            Top             =   675
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   0
            Left            =   3015
            TabIndex        =   15
            Top             =   315
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   5
            Left            =   3015
            TabIndex        =   18
            Top             =   675
            Width           =   150
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Code"
         Height          =   255
         Index           =   3
         Left            =   -74730
         TabIndex        =   2
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom"
         Height          =   255
         Index           =   1
         Left            =   -71310
         TabIndex        =   4
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lblLabels 
         Caption         =   "Commentaires"
         Height          =   255
         Index           =   6
         Left            =   -74745
         TabIndex        =   6
         Top             =   945
         Width           =   1230
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
      ScaleWidth      =   11415
      TabIndex        =   91
      Top             =   9690
      Width           =   11415
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fermer"
         Height          =   345
         Left            =   5760
         TabIndex        =   90
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Enregistrer"
         Height          =   345
         Left            =   4680
         TabIndex        =   89
         Top             =   45
         Width           =   975
      End
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
      TabIndex        =   0
      Top             =   90
      Width           =   11265
   End
   Begin VB.Label lblLabels 
      Caption         =   "RECNO:"
      Height          =   255
      Index           =   0
      Left            =   4095
      TabIndex        =   92
      Top             =   45
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmParamCalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A68060266"
Option Explicit

'##ModelId=5C8A68060360
Private fmAction As Integer
'##ModelId=5C8A6806037F
Private fmRECNO As Long
'##ModelId=5C8A680603A1
Private theParam As clsParamCalcul
'

'##ModelId=5C8A680603A2
Private Sub cmdUpdate_Click()
  Dim sel As Long
  
  On Error GoTo errcmdUpdate
  
  
  'update all constants for Statutaire calculation
  m_dataSource.Execute "UPDATE TTLOGTRAIT SET NBLIGTRAIT=(SELECT count(*) FROM TTPROVCOLL), MTTRAIT=0"
  
   
  '  place les bonnes tables
 
  
  ' table incap
  If cboTableIncap.ListIndex <> -1 Then
    theParam.LoiIncapacite = theParam.GetNomTable(cboTableIncap.ItemData(cboTableIncap.ListIndex))
  Else
    MsgBox "Vous devez choisir une table d'incapacité", vbCritical
    Exit Sub
  End If
  
  If cboTableIncap_Retraite.ListIndex <> -1 Then
    theParam.LoiIncapacite_Retraite = theParam.GetNomTable(cboTableIncap_Retraite.ItemData(cboTableIncap_Retraite.ListIndex))
  Else
    MsgBox "Vous devez choisir une table d'incapacité pour la réforme des retraite", vbCritical
    Exit Sub
  End If
  
  ' table passage
  If cboTablePassage.ListIndex <> -1 Then
    theParam.LoiPassage = theParam.GetNomTable(cboTablePassage.ItemData(cboTablePassage.ListIndex))
  Else
    MsgBox "Vous devez choisir une table de passage", vbCritical
    Exit Sub
  End If
  
  If cboTablePassage_Retraite.ListIndex <> -1 Then
    theParam.LoiPassage_Retraite = theParam.GetNomTable(cboTablePassage_Retraite.ItemData(cboTablePassage_Retraite.ListIndex))
  Else
    MsgBox "Vous devez choisir une table de passage pour la réforme des retraite", vbCritical
    Exit Sub
  End If
  
  ' table inval
  If cboTableInval.ListIndex <> -1 Then
    theParam.LoiInvalidite = theParam.GetNomTable(cboTableInval.ItemData(cboTableInval.ListIndex))
  Else
    MsgBox "Vous devez choisir une table d'invalidité", vbCritical
    Exit Sub
  End If
  
  '###
  ' table inval viagere
  If cboTableViagere.ListIndex <> -1 Then
    theParam.LoiInvalidite_Viagere = theParam.GetNomTable(cboTableViagere.ItemData(cboTableViagere.ListIndex))
  Else
    MsgBox "Vous devez choisir une table de Invalidité viagère", vbCritical
    Exit Sub
  End If
   
  ' table dépendance
  If cboTableDependance.ListIndex <> -1 Then
    theParam.LoiDependance = theParam.GetNomTable(cboTableDependance.ItemData(cboTableDependance.ListIndex))
  Else
    MsgBox "Vous devez choisir une table de dépendance", vbCritical
    Exit Sub
  End If
  
  
  sel = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1)
  If sel < 50 Or sel > 70 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez entrer un chiffre entre 50 et 70 !", vbCritical
    txtAgeLimiteInvalCat1.SetFocus
    Exit Sub
  End If
    
  sel = m_dataHelper.GetDouble2(DureeChomage)
  If sel < 0 Or sel > 240 Then
    SSTab1.Tab = 0
    MsgBox "La durée d'indemnisation de la prestation chômage doit être comprise entre 0 et 240 mois(=20 ans)", vbCritical
    DureeChomage.SetFocus
    Exit Sub
  End If
    
  sel = m_dataHelper.GetDouble2(Inval1Cat)
  If sel < 0 Or sel > 100 Then
    SSTab1.Tab = 0
    MsgBox "La proportion des Invalides 1 ère Cat doit être comprise entre 0 et 100 % ", vbCritical
    Inval1Cat.SetFocus
    Exit Sub
  End If
    
    
  If cboTableInval_Retraite.ListIndex <> -1 Then
    theParam.LoiInvalidite_Retraite = theParam.GetNomTable(cboTableInval_Retraite.ItemData(cboTableInval_Retraite.ListIndex))
  Else
    MsgBox "Vous devez choisir une table d'invalidité pour la réforme des retraite", vbCritical
    Exit Sub
  End If
   
  sel = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1_Retraite)
  If sel < 50 Or sel > 70 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez entrer un chiffre entre 50 et 70 pour la réforme des retraite !", vbCritical
    txtAgeLimiteInvalCat1_Retraite.SetFocus
    Exit Sub
  End If
    
  If Not IsNumeric(txtCode) Then
    MsgBox "Vous devez entrer un nombre !", vbCritical
    SSTab1.Tab = 0
    txtCode.SetFocus
    Exit Sub
  End If
  
  If CLng(txtCode) <> NumParamCalcul _
     And m_dataHelper.GetParameterAsDouble("SELECT count(*) FROM ParamCalcul WHERE PENUMCLE = " & numPeriode & " And PEGPECLE = " & GroupeCle & " AND PENUMCLE=" & numPeriode & " AND PENUMPARAMCALCUL=" & txtCode & " AND RECNO<>" & fmRECNO) > 0 Then
    MsgBox "Ce code est déjà utilisé, veuillez le modifier !", vbCritical
    SSTab1.Tab = 0
    txtCode.SetFocus
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  theParam.CodeParamCalcul = CLng(txtCode)
  theParam.NomParamCalcul = txtNomParamCalcul
  theParam.Commentaire = txtFields(6)
  
  'statutaire
  theParam.statAgeRet = m_dataHelper.GetDouble2(txtAgeRet)
  theParam.statAgeMin = m_dataHelper.GetDouble2(txtAgeMin)
  theParam.statAgeMax = m_dataHelper.GetDouble2(txtAgeMax)
  theParam.statMO = m_dataHelper.GetDouble2(txtMO)
  theParam.statCLM = m_dataHelper.GetDouble2(txtCLM)
  theParam.statCLD = m_dataHelper.GetDouble2(txtCLD)
  theParam.statMAT = m_dataHelper.GetDouble2(txtMAT)
  theParam.statAT = m_dataHelper.GetDouble2(txtAT)
  
  'theParam.statYear = cmbYear.text
  If cmbYear.ListIndex <> -1 Then
    theParam.TableStatutaire = theParam.GetNomTable(cmbYear.ItemData(cmbYear.ListIndex))
  End If
  
   
  ' modifie les ',' en '.' ou l'inverse
  theParam.TauxIncapacite = m_dataHelper.GetDouble2(txtFields(9))
  theParam.FraisGestionIncapacite = m_dataHelper.GetDouble2(txtFields(10))
  theParam.TauxIncapPassageSpecifiqueInval1 = m_dataHelper.GetDouble2(txtFields(16))
  theParam.TauxInvalidite = m_dataHelper.GetDouble2(txtFields(12))
  theParam.FraisGestionInvalidite = m_dataHelper.GetDouble2(txtFields(13))
  theParam.AgeLimiteInvalCat1 = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1)
  theParam.AgeLimiteInvalCat1_Retraite = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1_Retraite)
  theParam.Inval1Cat = m_dataHelper.GetDouble2(Inval1Cat)
  theParam.DureeChomage = m_dataHelper.GetDouble2(DureeChomage)
  
  
  
  theParam.TauxRevenuC7 = m_dataHelper.GetDouble2(txtFields(15))
  theParam.FraisGestionC7 = m_dataHelper.GetDouble2(txtFields(14))
      
  theParam.TauxIndexation = m_dataHelper.GetDouble2(txtTxIndex)
  theParam.DureeIndexation = m_dataHelper.GetDouble2(txtDureeIndex)
  theParam.DureeLissage = m_dataHelper.GetDouble2(txtDureeLissage)
  theParam.TMO = m_dataHelper.GetDouble2(txtTMO)
  'Call m_dataHelper.GetDouble(datPrimaryRS.Recordset.fields("PEAGERETRAITE"), GetSettingIni(CompanyName, SectionName, "AgeRetraite", "65"))
  
  '*** reassurance
  ' rente education
  If cboTableRenteEducation.ListIndex <> -1 Then
    theParam.TableRenteEducation = theParam.GetNomTable(cboTableRenteEducation.ItemData(cboTableRenteEducation.ListIndex))
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez choisir une table de mortalite pour la Rente de Education", vbCritical
    SSTab1.Tab = 1
    cboTableRenteEducation.SetFocus
    Exit Sub
  End If
  
  theParam.TauxTechniqueRenteEducation = m_dataHelper.GetDouble2(txtTauxTechniqueRenteEducation)
  theParam.FraisGestionRenteEducation = m_dataHelper.GetDouble2(txtFraisGestionRenteEducation)
  
  ' paiement
  If rdoPaiementAvanceEducation Then
    theParam.PaiementRenteEducation = ePaiementAvance ' d'avance
  ElseIf rdoPaiementEchuEducation Then
    theParam.PaiementRenteEducation = ePaiementEchu ' echu
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez choisir un type de paiement pour la Rente de Education", vbCritical
    SSTab1.Tab = 1
    rdoPaiementAvanceEducation.SetFocus
    Exit Sub
  End If
  
  ' fractionnement
  If rdoAnnuelEducation Then
    theParam.FractionnementRenteEducation = eFractionnementAnnuel ' annuel
  ElseIf rdoSemestrielEducation Then
    theParam.FractionnementRenteEducation = eFractionnementSemestriel ' semestriel
  ElseIf rdoTrimestrielEducation Then
    theParam.FractionnementRenteEducation = eFractionnementTrimestriel ' trimestriel
  ElseIf rdoMensuelEducation Then
    theParam.FractionnementRenteEducation = eFractionnementMensuel ' mensuel
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez choisir un fractionnement pour la Rente de Education", vbCritical
    SSTab1.Tab = 1
    rdoAnnuelEducation.SetFocus
    Exit Sub
  End If
  
  ' rente conjoint
  If cboTableRenteConjoint.ListIndex <> -1 Then
    theParam.TableRenteConjoint = theParam.GetNomTable(cboTableRenteConjoint.ItemData(cboTableRenteConjoint.ListIndex))
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez choisir une table de mortalite pour la Rente de Conjoint", vbCritical
    SSTab1.Tab = 1
    cboTableRenteConjoint.SetFocus
    Exit Sub
  End If
  
  theParam.TauxTechniqueRenteConjoint = m_dataHelper.GetDouble2(txtTauxTechniqueRenteConjoint)
  theParam.FraisGestionRenteConjoint = m_dataHelper.GetDouble2(txtFraisGestionRenteConjoint)
  
  ' paiement
  rdoPaiementAvanceConjoint.Value = True
  If rdoPaiementAvanceConjoint Then
    theParam.PaiementRenteConjoint = ePaiementAvance ' d'avance
  ElseIf rdoPaiementEchuConjoint Then
    theParam.PaiementRenteConjoint = ePaiementEchu ' echu
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Vous devez choisir un type de paiement pour la Rente de Conjoint", vbCritical
    SSTab1.Tab = 1
    rdoPaiementAvanceConjoint.SetFocus
    Exit Sub
  End If
  
  ' fractionnement
  rdoTrimestrielConjoint.Value = True
  If rdoAnnuelConjoint Then
    theParam.FractionnementRenteConjoint = eFractionnementAnnuel ' annuel
  ElseIf rdoSemestrielConjoint Then
    theParam.FractionnementRenteConjoint = eFractionnementSemestriel ' semestriel
  ElseIf rdoTrimestrielConjoint Then
    theParam.FractionnementRenteConjoint = eFractionnementTrimestriel ' trimestriel
  ElseIf rdoMensuelConjoint Then
    theParam.FractionnementRenteConjoint = eFractionnementMensuel ' mensuel
  Else
    MsgBox "Vous devez choisir un fractionnement pour la Rente de Conjoint", vbCritical
    SSTab1.Tab = 1
    rdoAnnuelConjoint.SetFocus
    Exit Sub
  End If
  
  ' maintien deces
  theParam.FraisGestionCapitauxDecesDC = m_dataHelper.GetDouble2(txtFraisGestionCapitauxDecesDC)
  theParam.FraisGestionRenteEducationDC = m_dataHelper.GetDouble2(txtFraisGestionRenteEducationDC)
  
  theParam.PourcentageAccident = m_dataHelper.GetDouble2(txtPourcentageAccident)
  
  theParam.FraisGestionRenteConjointDC = m_dataHelper.GetDouble2(txtFraisGestionRenteConjointDC)
  theParam.CapitalMoyenRenteConjointTempoDC = m_dataHelper.GetDouble2(txtCapitalMoyenRenteConjointTempoDC)
  theParam.CapitalMoyenRenteConjointViagereDC = m_dataHelper.GetDouble2(txtCapitalMoyenRenteConjointViagereDC)
  theParam.ForcerCapitalMoyenRenteConjointDC = IIf(chkForcerCapitalMoyenRteConjoitDC.Value = vbChecked, 1, 0)
  theParam.AgeConjointRenteConjointDC = m_dataHelper.GetDouble2(txtAgeConjointRenteConjointDC)
  
  theParam.UtiliserTableLissageProvision = IIf(rdoSansLissage = True, False, True)
  
  '
  ' recalcul
  '
  
  'datPrimaryRS.Recordset.fields("PEPMGDForcerInval") = IIf(chkPMGDForcerInval = vbChecked, 1, 0)
  theParam.RecalculCoeffBCAC = rdoCalculCoeffBCAC = True
  
  theParam.AgeLimiteCalulDC = m_dataHelper.GetDouble2(txtAgeLimiteCalulDC)
  theParam.AgeLimiteCalulDC_Retraite = m_dataHelper.GetDouble2(txtAgeLimiteCalulDC_Retraite)
  theParam.AgeLimiteCalulDC_Inval1 = m_dataHelper.GetDouble2(txtAgeLimiteCalulDC_Inval1)
  theParam.AgeLimiteCalulDC_Retraite_Inval1 = m_dataHelper.GetDouble2(txtAgeLimiteCalulDC_Retraite_Inval1)
  
  If rdoCapitauxConstitif.Value = True Then
    theParam.MethodeCalculDC = eCapitauxConstitutifs
  ElseIf rdoPctPMCalculeeDC.Value = True Then
    theParam.MethodeCalculDC = ePctProvisionCalculee
    theParam.PctPMCalculeeDC = m_dataHelper.GetDouble2(txtPctPMCalculeeDC)
  ElseIf rdoDCAucun.Value = True Then
    theParam.MethodeCalculDC = ePasDeCalcul
  End If

  ' rempli le combo incap
  If cboTableIncapCalculDC.ListIndex <> -1 Then
    theParam.LoiIncapaciteDC = theParam.GetNomTable(cboTableIncapCalculDC.ItemData(cboTableIncapCalculDC.ListIndex))
  End If
  
  If cboTableIncapCalculDC_Retraite.ListIndex <> -1 Then
    theParam.LoiIncapaciteDC_Retraite = theParam.GetNomTable(cboTableIncapCalculDC_Retraite.ItemData(cboTableIncapCalculDC_Retraite.ListIndex))
  End If
  
  ' rempli le combo invalidite
  If cboTableInvalCalculDC.ListIndex <> -1 Then
    theParam.LoiInvaliditeDC = theParam.GetNomTable(cboTableInvalCalculDC.ItemData(cboTableInvalCalculDC.ListIndex))
  End If
  
  If cboTableInvalCalculDC_Retraite.ListIndex <> -1 Then
    theParam.LoiInvaliditeDC_Retraite = theParam.GetNomTable(cboTableInvalCalculDC_Retraite.ItemData(cboTableInvalCalculDC_Retraite.ListIndex))
  End If
  
  theParam.TauxTechnicCalculDC = m_dataHelper.GetDouble2(txtTauxTechnicCalculDC)
  theParam.FraisGestionCalculDC = m_dataHelper.GetDouble2(txtFraisGestionCalculDC)
  
  ' combo incap DC
  If cboTableIncapPrecalculDC.ListIndex <> -1 Then
    theParam.TableIncapacitePrecalculDC = theParam.GetNomTable(cboTableIncapPrecalculDC.ItemData(cboTableIncapPrecalculDC.ListIndex))
  End If
  
  If cboTableIncapPrecalculDC_Retraite.ListIndex <> -1 Then
    theParam.TableIncapacitePrecalculDC_Retraite = theParam.GetNomTable(cboTableIncapPrecalculDC_Retraite.ItemData(cboTableIncapPrecalculDC_Retraite.ListIndex))
  End If
  
  ' combo invalidite DC
  If cboTableInvalPrecalculDC.ListIndex <> -1 Then
    theParam.TableInvaliditePrecalculDC = theParam.GetNomTable(cboTableInvalPrecalculDC.ItemData(cboTableInvalPrecalculDC.ListIndex))
  End If
  
  If cboTableInvalPrecalculDC_Retraite.ListIndex <> -1 Then
    theParam.TableInvaliditePrecalculDC_Retraite = theParam.GetNomTable(cboTableInvalPrecalculDC_Retraite.ItemData(cboTableInvalPrecalculDC_Retraite.ListIndex))
  End If
  
  
  theParam.CalculAge_Anniversaire = IIf(rdoAgeMillesime.Value = True, False, True)


  
  ' méthode d'interpollation inval
  If rdoInterpolationInval_CorrectionDuree.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_CorrectionDuree
  ElseIf rdoInterpolationInval_Age.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_Age
  ElseIf rdoInterpolationInval_AgeDuree.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_AgeDuree
  Else
    theParam.InterpolationInvalidite = eInterpolationInval_NON
  End If
  
  
  ' la checkbox et le paramètre sont inversé : "[ ] Ne pas lisser..."
  theParam.LissageCoeffPassage = IIf(chkLissagePassage.Value = vbChecked, False, True)
  theParam.AnnualisationPassage = IIf(chkAnnulisationPassage.Value = vbChecked, True, False)
  theParam.BridageAgeLimiteTable = IIf(chkBridageAge.Value = vbChecked, True, False)
  
  
  theParam.PortefeuilleSalaries = IIf(chkPortefeuilleSalaries.Value = vbChecked, True, False)
  
  
  On Error GoTo errcmdUpdate
  
  If theParam.SaveToDB(GroupeCle, numPeriode) = False Then
    Screen.MousePointer = vbDefault
  
    Exit Sub
  End If
    
  Set theParam = Nothing
  
  Screen.MousePointer = vbDefault
  
  Unload Me
  
  Exit Sub
  
errcmdUpdate:
  MsgBox "Erreur " & Err & " :" & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=5C8A680603AE
Private Sub cmdClose_Click()
  Set theParam = Nothing
  
  Screen.MousePointer = vbDefault
  Unload Me
End Sub


'##ModelId=5C8A680603BE
Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Dim nomTable As String
  Dim sel As Long
  
  Screen.MousePointer = vbHourglass
  
  Set theParam = New clsParamCalcul
  
  ' activate first tab
  SSTab1.Tab = 0
  
  ' ajoute un enregistrement si besoin
  If NumParamCalcul = -1 Then
    fmRECNO = 0
    
    ' chargement des valeurs par défaut (section [P3I])
    theParam.LoadFromIni -1, "BILAN"
    
    ' numero de periode
    Set rs = m_dataSource.OpenRecordset("SELECT MAX(PENUMPARAMCALCUL) FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE=" & numPeriode, Snapshot)
    If Not rs.EOF Then
      NumParamCalcul = IIf(IsNull(rs.fields(0)), 1, rs.fields(0) + 1)
    Else
      NumParamCalcul = 1
    End If
    rs.Close
     
    fmAction = 0 ' add new
    
    theParam.CodeParamCalcul = NumParamCalcul
    theParam.NomParamCalcul = "Nouveaux paramètres de calcul..."
    theParam.Commentaire = ""
  
  Else
    
    fmAction = 1 ' edit
    
    theParam.LoadFromDB GroupeCle, numPeriode, NumParamCalcul
        
    fmRECNO = theParam.RECNO
    
  End If
  
  
  txtAgeRet.text = IIf(theParam.statAgeRet = 0, 62, theParam.statAgeRet)
  txtAgeMin.text = IIf(theParam.statAgeMin = 0, 20, theParam.statAgeMin)
  txtAgeMax.text = IIf(theParam.statAgeMax = 0, 64, theParam.statAgeMax)
  txtMO.text = IIf(theParam.statMO = 0, 52, theParam.statMO)
  txtCLM.text = IIf(theParam.statCLM = 0, 156, theParam.statCLM)
  txtCLD.text = IIf(theParam.statCLD = 0, 252, theParam.statCLD)
  txtMAT.text = IIf(theParam.statMAT = 0, 23, theParam.statMAT)
  txtAT.text = IIf(theParam.statAT = 0, 366, theParam.statAT)
  
  ' rempli le combo annee bareme statutaire
  cmbYear.Clear
  
  sel = theParam.GetCleTable(theParam.TableStatutaire)  '107
  m_dataHelper.FillCombo cmbYear, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_BaremeAnneeStatutaire, sel
  
  If sel = -1 Then
    cmbYear.ListIndex = 0
  End If
  
  
  '### old code - to be deleted
  
  'fill combo from INI
'  Dim annBar As String
'  Dim listAnnBar As String
'  Dim arrAnnBar() As String
'  Dim cnt As Integer
'  Dim ind As Integer
'  Dim dimens As Integer
'  Dim rsAnnBar As ADODB.Recordset
  
  'dimens = 0
  'annBar = GetSettingIni(CompanyName, SectionName, "StatAnneeBareme", "2004")
  'annBar = IIf(theParam.statYear = 0, 2015, theParam.statYear)
  
  'Set rsAnnBar = m_dataSource.OpenRecordset("SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_BaremeAnneeStatutaire, Disconnected)
  
  'listAnnBar = GetSettingIni(CompanyName, SectionName, "StatAnneeBaremeList", "2004,2015")
     
  'If rsAnnBar.EOF Then
'    Do While Not rsAnnBar.EOF
'      ReDim Preserve arrAnnBar(dimens)
'      arrAnnBar(dimens) = rsAnnBar.fields("LIBTABLE")
'      dimens = dimens + 1
'
'      rsAnnBar.MoveNext
'    Loop
  'Else
  
  'End If
   
  'cmbYear.Clear
      
'  arrAnnBar = Split(listAnnBar, ",")
  'For cnt = 0 To UBound(arrAnnBar)
  '  cmbYear.AddItem arrAnnBar(cnt)
    'cmbAnneeBareme.ItemData(cmbAnneeBareme.NewIndex) = cnt
  '  If arrAnnBar(cnt) = annBar Then
  '    ind = cnt
  '  End If
  'Next

  'cmbYear.ListIndex = ind
  
  '### end delete
  

  txtCode = theParam.CodeParamCalcul
  txtNomParamCalcul = theParam.NomParamCalcul

  txtFields(6) = theParam.Commentaire
  txtFields(9) = theParam.TauxIncapacite
  txtFields(10) = theParam.FraisGestionIncapacite
 
  'modif réforme retraire 2023
  txtFields(16) = theParam.TauxIncapPassageSpecifiqueInval1
  txtFields(12) = theParam.TauxInvalidite
  txtFields(13) = theParam.FraisGestionInvalidite
  
  txtAgeLimiteInvalCat1 = theParam.AgeLimiteInvalCat1
  txtAgeLimiteInvalCat1_Retraite = theParam.AgeLimiteInvalCat1_Retraite
  Inval1Cat = theParam.Inval1Cat
  DureeChomage = theParam.DureeChomage
  TxChomage = theParam.TauxIncapacite
  
  txtFields(15) = theParam.TauxRevenuC7
  txtFields(14) = theParam.FraisGestionC7
   
  txtTxIndex = theParam.TauxIndexation
  txtDureeIndex = theParam.DureeIndexation
  txtDureeLissage = theParam.DureeLissage
  txtTMO = theParam.TMO
    
  ' rente conjoint
  txtTauxTechniqueRenteConjoint = theParam.TauxTechniqueRenteConjoint
  txtFraisGestionRenteConjoint = theParam.FraisGestionRenteConjoint
    
  Select Case theParam.PaiementRenteConjoint
    Case ePaiementAvance
      rdoPaiementAvanceConjoint = True
    Case ePaiementEchu
      rdoPaiementEchuConjoint = True
  End Select
  
  Select Case theParam.FractionnementRenteConjoint
    Case eFractionnementAnnuel
      rdoAnnuelConjoint = True
    Case eFractionnementSemestriel
      rdoSemestrielConjoint = True
    Case eFractionnementTrimestriel
      rdoTrimestrielConjoint = True
    Case eFractionnementMensuel
      rdoMensuelConjoint = True
  End Select
    
    
  ' rente education
  txtTauxTechniqueRenteEducation = theParam.TauxTechniqueRenteEducation
  txtFraisGestionRenteEducation = theParam.FraisGestionRenteEducation
    
  Select Case theParam.PaiementRenteEducation
    Case ePaiementAvance
      rdoPaiementAvanceEducation = True
    Case ePaiementEchu
      rdoPaiementEchuEducation = True
  End Select
  
  Select Case theParam.FractionnementRenteEducation
    Case eFractionnementAnnuel
      rdoAnnuelEducation = True
    Case eFractionnementSemestriel
      rdoSemestrielEducation = True
    Case eFractionnementTrimestriel
      rdoTrimestrielEducation = True
    Case eFractionnementMensuel
      rdoMensuelEducation = True
  End Select
  
    
  ' maintien deces
  txtFraisGestionCapitauxDecesDC = theParam.FraisGestionCapitauxDecesDC
  txtFraisGestionRenteEducationDC = theParam.FraisGestionRenteEducationDC
  
  txtPourcentageAccident = theParam.PourcentageAccident
  
  txtFraisGestionRenteConjointDC = theParam.FraisGestionRenteConjointDC
  txtCapitalMoyenRenteConjointTempoDC = theParam.CapitalMoyenRenteConjointTempoDC
  txtCapitalMoyenRenteConjointViagereDC = theParam.CapitalMoyenRenteConjointViagereDC
  chkForcerCapitalMoyenRteConjoitDC.Value = IIf(theParam.ForcerCapitalMoyenRenteConjointDC = True, vbChecked, vbUnchecked)
  txtAgeConjointRenteConjointDC = theParam.AgeConjointRenteConjointDC
  
  rdoSansLissage = theParam.UtiliserTableLissageProvision = False
  rdoAvecLissage = theParam.UtiliserTableLissageProvision = True
  
  '
  ' recalcul coeff DC
  '
  
  'chkPMGDForcerInval = IIf(datPrimaryRS.Recordset.fields("PEPMGDForcerInval") = True, vbChecked, vbUnchecked)
  chkPMGDForcerInval = vbUnchecked
  
  rdoCalculCoeffBCAC = theParam.RecalculCoeffBCAC = True
  
  rdoCapitauxConstitif.Value = theParam.MethodeCalculDC = eCapitauxConstitutifs
  rdoPctPMCalculeeDC.Value = theParam.MethodeCalculDC = ePctProvisionCalculee
  txtPctPMCalculeeDC = theParam.PctPMCalculeeDC
  rdoDCAucun.Value = theParam.MethodeCalculDC = ePasDeCalcul

  If rdoCapitauxConstitif.Value = False And rdoPctPMCalculeeDC.Value = False And rdoDCAucun.Value = False Then
    rdoCapitauxConstitif.Value = True
  End If


  ' rempli le combo incap
  sel = theParam.GetCleTable(theParam.LoiIncapaciteDC)
  m_dataHelper.FillCombo cboTableIncapCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteIncap, sel
  If sel = -1 Then
    cboTableIncapCalculDC.ListIndex = 0
  End If
  
  sel = theParam.GetCleTable(theParam.LoiIncapaciteDC_Retraite)
  m_dataHelper.FillCombo cboTableIncapCalculDC_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteIncap, sel
  If sel = -1 Then
    cboTableIncapCalculDC_Retraite.ListIndex = 0
  End If
    
  ' rempli le combo invalidite
  sel = theParam.GetCleTable(theParam.LoiInvaliditeDC)
  m_dataHelper.FillCombo cboTableInvalCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteInval, sel
  If sel = -1 Then
    cboTableInvalCalculDC.ListIndex = 0
  End If
    
  sel = theParam.GetCleTable(theParam.LoiInvaliditeDC_Retraite)
  m_dataHelper.FillCombo cboTableInvalCalculDC_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteInval, sel
  If sel = -1 Then
    cboTableInvalCalculDC_Retraite.ListIndex = 0
  End If
    
  txtTauxTechnicCalculDC = theParam.TauxTechnicCalculDC
  txtFraisGestionCalculDC = theParam.FraisGestionCalculDC
  
  txtAgeLimiteCalulDC = theParam.AgeLimiteCalulDC
  txtAgeLimiteCalulDC_Retraite = theParam.AgeLimiteCalulDC_Retraite
  txtAgeLimiteCalulDC_Inval1 = theParam.AgeLimiteCalulDC_Inval1
  txtAgeLimiteCalulDC_Retraite_Inval1 = theParam.AgeLimiteCalulDC_Retraite_Inval1
    
  '
  ' lecture coeff DC
  '
  rdoLireCoeffBCAC = theParam.RecalculCoeffBCAC = False
    
  ' rempli le combo incap precalcul dc
  sel = theParam.GetCleTable(theParam.TableIncapacitePrecalculDC)
  m_dataHelper.FillCombo cboTableIncapPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACIncap, sel
  If sel = -1 Then
    cboTableIncapPrecalculDC.ListIndex = 0
  End If
          
  sel = theParam.GetCleTable(theParam.TableIncapacitePrecalculDC_Retraite)
  m_dataHelper.FillCombo cboTableIncapPrecalculDC_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACIncap, sel
  If sel = -1 Then
    cboTableIncapPrecalculDC_Retraite.ListIndex = 0
  End If
          
  ' rempli le combo invalidite
  sel = theParam.GetCleTable(theParam.TableInvaliditePrecalculDC)
  m_dataHelper.FillCombo cboTableInvalPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval, sel
  If sel = -1 Then
    cboTableInvalPrecalculDC.ListIndex = 0
  End If
      
  sel = theParam.GetCleTable(theParam.TableInvaliditePrecalculDC_Retraite)
  m_dataHelper.FillCombo cboTableInvalPrecalculDC_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval, sel
  If sel = -1 Then
    cboTableInvalPrecalculDC_Retraite.ListIndex = 0
  End If
      
      
  rdoAgeMillesime.Value = theParam.CalculAge_Anniversaire = False
  rdoAgeAnniversaire.Value = theParam.CalculAge_Anniversaire = True
  
  
  chkInterpolationIncap.Value = vbChecked
  ' la checkbox et le paramètre sont inversé : "[ ] Ne pas lisser..."
  chkLissagePassage.Value = IIf(theParam.LissageCoeffPassage = False, vbChecked, vbUnchecked)
  chkAnnulisationPassage.Value = IIf(theParam.AnnualisationPassage = True, vbChecked, vbUnchecked)
  chkBridageAge.Value = IIf(theParam.BridageAgeLimiteTable = True, vbChecked, vbUnchecked)
  
  ' interpolation de l'inval
  Select Case theParam.InterpolationInvalidite
    Case eInterpolationInval_CorrectionDuree
      rdoInterpolationInval_CorrectionDuree.Value = True
    Case eInterpolationInval_Age
      rdoInterpolationInval_Age.Value = True
    Case eInterpolationInval_AgeDuree
      rdoInterpolationInval_AgeDuree.Value = True
    Case Else
      rdoInterpolationInval_NON.Value = True
  End Select
    
  cmdUpdate.Enabled = m_dataHelper.GetParameterAsDouble("SELECT PELOCKED FROM Periode WHERE PENUMCLE = " & numPeriode & " And PEGPECLE = " & GroupeCle & " AND PENUMCLE=" & numPeriode) = 0
  
  ' Incap toujours interpolée
  chkInterpolationIncap.Enabled = False
  
  ' rempli le nom du groupe
  lblGroupe = "Période n° " & numPeriode & " du Groupe '" & NomGroupe & "'"
  
  ' rempli le combo incap
  sel = theParam.GetCleTable(theParam.LoiIncapacite)
  m_dataHelper.FillCombo cboTableIncap, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 1", sel
  If sel = -1 Then
    cboTableIncap.ListIndex = 0
  End If
  
  sel = theParam.GetCleTable(theParam.LoiIncapacite_Retraite)
  m_dataHelper.FillCombo cboTableIncap_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 1", sel
  If sel = -1 Then
    cboTableIncap_Retraite.ListIndex = 0
  End If
  
  ' passage
  sel = theParam.GetCleTable(theParam.LoiPassage)
  m_dataHelper.FillCombo cboTablePassage, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 2", sel
  If sel = -1 Then
    cboTablePassage.ListIndex = 0
  End If
  
  sel = theParam.GetCleTable(theParam.LoiPassage_Retraite)
  m_dataHelper.FillCombo cboTablePassage_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 2", sel
  If sel = -1 Then
    cboTablePassage_Retraite.ListIndex = 0
  End If
  
  ' invalidite
  sel = theParam.GetCleTable(theParam.LoiInvalidite)
  m_dataHelper.FillCombo cboTableInval, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 3", sel
  If sel = -1 Then
    cboTableInval.ListIndex = 0
  End If
  
  sel = theParam.GetCleTable(theParam.LoiInvalidite_Retraite)
  m_dataHelper.FillCombo cboTableInval_Retraite, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 3", sel
  If sel = -1 Then
    cboTableInval_Retraite.ListIndex = 0
  End If
  
  ' rempli le combo Dépendance
  sel = theParam.GetCleTable(theParam.LoiDependance)
  m_dataHelper.FillCombo cboTableDependance, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 30", sel
  If sel = -1 Then
    cboTableDependance.ListIndex = 0
  End If
  
  ' ### rempli le combo Invalidité viagère
  sel = theParam.GetCleTable(theParam.LoiInvalidite_Viagere)
  m_dataHelper.FillCombo cboTableViagere, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 3", sel
  If sel = -1 Then
    cboTableViagere.ListIndex = 0
  End If
  
  
  ' rente conjoint
  sel = theParam.GetCleTable(theParam.TableRenteConjoint)
  m_dataHelper.FillCombo cboTableRenteConjoint, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel
  If sel = -1 Then
    cboTableRenteConjoint.ListIndex = 0
  End If
  
  ' rente Education
  sel = theParam.GetCleTable(theParam.TableRenteEducation)
  m_dataHelper.FillCombo cboTableRenteEducation, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel
  If sel = -1 Then
    cboTableRenteConjoint.ListIndex = 0
  End If
   
  
  chkPortefeuilleSalaries.Value = IIf(theParam.PortefeuilleSalaries = True, vbChecked, vbUnchecked)
  
  
  Screen.MousePointer = vbDefault

End Sub

'##ModelId=5C8A680603DD
Private Sub Form_Unload(Cancel As Integer)
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A68070014
Private Sub EnableCoeffBCAC()
  txtTauxTechnicCalculDC.Enabled = rdoCalculCoeffBCAC = True
  txtFraisGestionCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableIncapCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableInvalCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableIncapCalculDC_Retraite.Enabled = rdoCalculCoeffBCAC = True
  cboTableInvalCalculDC_Retraite.Enabled = rdoCalculCoeffBCAC = True
  
  cboTableIncapPrecalculDC.Enabled = rdoLireCoeffBCAC = True
  cboTableInvalPrecalculDC.Enabled = rdoLireCoeffBCAC = True
  cboTableIncapPrecalculDC_Retraite.Enabled = rdoLireCoeffBCAC = True
  cboTableInvalPrecalculDC_Retraite.Enabled = rdoLireCoeffBCAC = True
End Sub


'##ModelId=5C8A68070034
Private Sub rdoCalculCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

'##ModelId=5C8A68070043
Private Sub rdoCapitauxConstitif_Click()
  If rdoCapitauxConstitif.Value = True Then
    txtPctPMCalculeeDC.Enabled = False
  Else
    txtPctPMCalculeeDC.Enabled = True
  End If
End Sub

'##ModelId=5C8A68070053
Private Sub rdoDCAucun_Click()
  If rdoCapitauxConstitif.Value = True Then
    txtPctPMCalculeeDC.Enabled = False
  Else
    txtPctPMCalculeeDC.Enabled = True
  End If
End Sub

'##ModelId=5C8A68070072
Private Sub rdoPctPMCalculeeDC_Click()
  If rdoCapitauxConstitif.Value = True Then
    txtPctPMCalculeeDC.Enabled = False
  Else
    txtPctPMCalculeeDC.Enabled = True
  End If
'  If rdoCotisationsExonerees = True Then
'    MsgBox "ATTENTION: cette méthode de calcul n'est pas encore implémentée !", vbCritical
'    rdoCotisationsExonerees = False
'    rdoCapitauxConstitif = True
'  End If
End Sub

'##ModelId=5C8A68070082
Private Sub rdoLireCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

