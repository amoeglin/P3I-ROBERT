VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmParamCalculIni 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paramètres de calcul par défaut"
   ClientHeight    =   7125
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11415
   Icon            =   "frmParamCalculIni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6270
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   11060
      _Version        =   393216
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
      TabPicture(0)   =   "frmParamCalculIni.frx":1BB2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFields(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame19"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtNomParamCalcul"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCode"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Rentes"
      TabPicture(1)   =   "frmParamCalculIni.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Maintien Garanties Décès"
      TabPicture(2)   =   "frmParamCalculIni.frx":1BEA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame12"
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(2)=   "Frame15"
      Tab(2).Control(3)=   "Frame13"
      Tab(2).Control(4)=   "Frame17"
      Tab(2).Control(5)=   "Frame16"
      Tab(2).ControlCount=   6
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
         TabIndex        =   138
         Top             =   3735
         Width           =   5100
         Begin VB.TextBox txtDureeIndex 
            Height          =   285
            Left            =   4140
            TabIndex        =   142
            Top             =   630
            Width           =   495
         End
         Begin VB.TextBox txtTxIndex 
            Height          =   285
            Left            =   1575
            TabIndex        =   141
            Top             =   630
            Width           =   675
         End
         Begin VB.TextBox txtTMO 
            Height          =   285
            Left            =   1575
            TabIndex        =   140
            Top             =   1440
            Width           =   675
         End
         Begin VB.TextBox txtDureeLissage 
            Height          =   285
            Left            =   4140
            TabIndex        =   139
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Durée d'indexation"
            Height          =   255
            Index           =   2
            Left            =   2745
            TabIndex        =   152
            Top             =   675
            Width           =   1365
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux d'indexation"
            Height          =   255
            Index           =   16
            Left            =   90
            TabIndex        =   151
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux financier TME"
            Height          =   255
            Index           =   17
            Left            =   90
            TabIndex        =   150
            Top             =   1485
            Width           =   1410
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   6
            Left            =   2295
            TabIndex        =   149
            Top             =   1485
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   7
            Left            =   2295
            TabIndex        =   148
            Top             =   675
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   8
            Left            =   4680
            TabIndex        =   147
            Top             =   675
            Width           =   330
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   9
            Left            =   4680
            TabIndex        =   146
            Top             =   1485
            Width           =   330
         End
         Begin VB.Label Label3 
            Caption         =   "Durée de lissage de la Prime Unique"
            Height          =   375
            Left            =   2760
            TabIndex        =   145
            Top             =   1380
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cotisation revalo (Incapacité, Invalidité, Rente)"
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
            TabIndex        =   144
            Top             =   1125
            Width           =   3435
         End
         Begin VB.Label lblLabels 
            Caption         =   "Indexation (Incapacité, Invalidité)"
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
            TabIndex        =   143
            Top             =   315
            Width           =   3435
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
         Left            =   9270
         TabIndex        =   133
         Top             =   2025
         Width           =   1905
         Begin VB.CheckBox chkPortefeuilleSalaries 
            Caption         =   "Portefeuille ""Salariés"""
            Height          =   420
            Left            =   360
            TabIndex        =   134
            Top             =   360
            Width           =   1320
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
         Height          =   780
         Left            =   -74685
         TabIndex        =   129
         Top             =   4680
         Width           =   5100
         Begin VB.OptionButton rdoPctPMCalculeeDC 
            Caption         =   "% de la Provision calculée"
            Height          =   330
            Left            =   2340
            TabIndex        =   132
            Top             =   270
            Width           =   1680
         End
         Begin VB.OptionButton rdoCapitauxConstitif 
            Caption         =   "Capitaux constitutifs sous risque"
            Height          =   330
            Left            =   270
            TabIndex        =   131
            Top             =   270
            Width           =   1995
         End
         Begin VB.TextBox txtPctPMCalculeeDC 
            Height          =   330
            Left            =   4140
            TabIndex        =   130
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   1590
         MaxLength       =   60
         TabIndex        =   2
         Top             =   540
         Width           =   1020
      End
      Begin VB.TextBox txtNomParamCalcul 
         Height          =   285
         Left            =   4200
         MaxLength       =   60
         TabIndex        =   4
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
         Left            =   135
         TabIndex        =   7
         Top             =   2025
         Width           =   5460
         Begin VB.CheckBox chkBridageAge 
            Caption         =   "Age bridé aux limites de table"
            Height          =   195
            Left            =   1350
            TabIndex        =   10
            Top             =   675
            Width           =   2400
         End
         Begin VB.OptionButton rdoAgeAnniversaire 
            Caption         =   "Anniversaire"
            Height          =   240
            Left            =   3195
            TabIndex        =   9
            Top             =   315
            Width           =   2175
         End
         Begin VB.OptionButton rdoAgeMillesime 
            Caption         =   "par différence de Millesime"
            Height          =   240
            Left            =   360
            TabIndex        =   8
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
         Height          =   690
         Left            =   -74685
         TabIndex        =   120
         Top             =   3825
         Width           =   5100
         Begin VB.TextBox txtFraisGestionRenteEducationDC 
            Height          =   330
            Left            =   1485
            TabIndex        =   121
            Text            =   "0"
            Top             =   225
            Width           =   465
         End
         Begin VB.Label Label10 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   135
            TabIndex        =   123
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label12 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   122
            Top             =   270
            Width           =   240
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
         Height          =   1365
         Left            =   -74685
         TabIndex        =   112
         Top             =   2160
         Width           =   5100
         Begin VB.TextBox txtFraisGestionRenteConjointDC 
            Height          =   330
            Left            =   1665
            TabIndex        =   84
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointTempoDC 
            Height          =   330
            Left            =   2205
            TabIndex        =   87
            Text            =   "4.5"
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkForcerCapitalMoyenRteConjoitDC 
            Caption         =   "Forcer"
            Height          =   240
            Left            =   315
            TabIndex        =   86
            Top             =   945
            Width           =   780
         End
         Begin VB.TextBox txtCapitalMoyenRenteConjointViagereDC 
            Height          =   330
            Left            =   3960
            TabIndex        =   88
            Text            =   "4.5"
            Top             =   900
            Width           =   915
         End
         Begin VB.TextBox txtAgeConjointRenteConjointDC 
            Height          =   330
            Left            =   4095
            TabIndex        =   85
            Text            =   "0"
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label32 
            Caption         =   "%"
            Height          =   240
            Left            =   2205
            TabIndex        =   119
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label33 
            Caption         =   "Frais de gestion"
            Height          =   195
            Left            =   135
            TabIndex        =   118
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Capital constitutif moyen"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   117
            Top             =   675
            Width           =   1770
         End
         Begin VB.Label Label39 
            Caption         =   "Temporaire"
            Height          =   195
            Left            =   1305
            TabIndex        =   116
            Top             =   945
            Width           =   825
         End
         Begin VB.Label Label40 
            Caption         =   "Viagère"
            Height          =   195
            Left            =   3240
            TabIndex        =   115
            Top             =   945
            Width           =   600
         End
         Begin VB.Label Label38 
            Caption         =   "ans"
            Height          =   240
            Left            =   4590
            TabIndex        =   114
            Top             =   315
            Width           =   330
         End
         Begin VB.Label Label37 
            Caption         =   "Age du conjoint : +/-"
            Height          =   375
            Left            =   2565
            TabIndex        =   113
            Top             =   315
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
         Height          =   3930
         Left            =   -69285
         TabIndex        =   103
         Top             =   585
         Width           =   5100
         Begin VB.TextBox txtAgeLimiteCalulDC 
            Height          =   330
            Left            =   1755
            TabIndex        =   95
            Text            =   "0"
            Top             =   1440
            Width           =   465
         End
         Begin VB.CheckBox chkPMGDForcerInval 
            Caption         =   "Forcer en Invalidité si Anc > 36 mois et pas de passage"
            Height          =   285
            Left            =   315
            TabIndex        =   91
            Top             =   315
            Visible         =   0   'False
            Width           =   4200
         End
         Begin VB.ComboBox cboTableInvalPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   3465
            Width           =   3525
         End
         Begin VB.ComboBox cboTableIncapPrecalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   3060
            Width           =   3525
         End
         Begin VB.ComboBox cboTableInvalCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   2250
            Width           =   3525
         End
         Begin VB.TextBox txtTauxTechnicCalculDC 
            Height          =   330
            Left            =   1755
            TabIndex        =   93
            Text            =   "0"
            Top             =   1035
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionCalculDC 
            Height          =   330
            Left            =   3960
            TabIndex        =   94
            Text            =   "0"
            Top             =   1035
            Width           =   465
         End
         Begin VB.ComboBox cboTableIncapCalculDC 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1845
            Width           =   3525
         End
         Begin VB.OptionButton rdoLireCoeffBCAC 
            Caption         =   "Utiliser les tables de coefficients précalculés du BCAC"
            Height          =   240
            Left            =   315
            TabIndex        =   98
            Top             =   2700
            Width           =   4605
         End
         Begin VB.OptionButton rdoCalculCoeffBCAC 
            Caption         =   "Calcul du coefficient"
            Height          =   240
            Left            =   315
            TabIndex        =   92
            Top             =   720
            Width           =   1860
         End
         Begin VB.Label Label6 
            Caption         =   "ans"
            Height          =   240
            Left            =   2295
            TabIndex        =   128
            Top             =   1485
            Width           =   330
         End
         Begin VB.Label Label5 
            Caption         =   "Age limite"
            Height          =   375
            Index           =   8
            Left            =   495
            TabIndex        =   127
            Top             =   1485
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Table Inval"
            Height          =   240
            Index           =   7
            Left            =   495
            TabIndex        =   111
            Top             =   3510
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Table Incap"
            Height          =   375
            Index           =   6
            Left            =   495
            TabIndex        =   110
            Top             =   3105
            Width           =   960
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Inval"
            Height          =   375
            Index           =   5
            Left            =   315
            TabIndex        =   109
            Top             =   2295
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Mortalité Incap"
            Height          =   240
            Index           =   4
            Left            =   315
            TabIndex        =   108
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   240
            Left            =   4500
            TabIndex        =   107
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label9 
            Caption         =   "%"
            Height          =   240
            Left            =   2295
            TabIndex        =   106
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Frais de gestion"
            Height          =   240
            Left            =   2700
            TabIndex        =   105
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   240
            Index           =   1
            Left            =   495
            TabIndex        =   104
            Top             =   1080
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
         Left            =   -69285
         TabIndex        =   102
         Top             =   4680
         Width           =   5100
         Begin VB.OptionButton rdoSansLissage 
            Caption         =   "100%"
            Height          =   240
            Left            =   315
            TabIndex        =   89
            Top             =   315
            Width           =   1275
         End
         Begin VB.OptionButton rdoAvecLissage 
            Caption         =   "Utiliser la table 'LissageProvision'"
            Height          =   240
            Left            =   1575
            TabIndex        =   90
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
         Left            =   -74685
         TabIndex        =   81
         Top             =   585
         Width           =   5100
         Begin VB.TextBox txtPourcentageAccident 
            Height          =   330
            Left            =   2790
            TabIndex        =   124
            Text            =   "0"
            Top             =   720
            Width           =   465
         End
         Begin VB.TextBox txtFraisGestionCapitauxDecesDC 
            Height          =   330
            Left            =   2790
            TabIndex        =   82
            Text            =   "0"
            Top             =   315
            Width           =   465
         End
         Begin VB.Label Label15 
            Caption         =   "Pourcentage Décès par Accident"
            Height          =   240
            Left            =   135
            TabIndex        =   126
            Top             =   765
            Width           =   2490
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   240
            Left            =   3330
            TabIndex        =   125
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "%"
            Height          =   240
            Left            =   3330
            TabIndex        =   101
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label8 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   135
            TabIndex        =   83
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
         TabIndex        =   46
         Top             =   765
         Visible         =   0   'False
         Width           =   5100
         Begin VB.ComboBox cboTableRenteConjoint 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   900
            Width           =   4695
         End
         Begin VB.TextBox txtFraisGestionRenteConjoint 
            Height          =   285
            Left            =   4230
            TabIndex        =   51
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueRenteConjoint 
            Height          =   285
            Left            =   1485
            TabIndex        =   48
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.Frame Frame7 
            Caption         =   "Paiement"
            Height          =   1140
            Left            =   225
            TabIndex        =   54
            Top             =   1350
            Width           =   1725
            Begin VB.OptionButton rdoPaiementAvanceConjoint 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   55
               Top             =   360
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementEchuConjoint 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   56
               Top             =   720
               Width           =   1275
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Fractionnement"
            Height          =   1140
            Left            =   2340
            TabIndex        =   57
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton rdoSemestrielConjoint 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   59
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelConjoint 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   58
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelConjoint 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   61
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoTrimestrielConjoint 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   60
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   2
            Left            =   225
            TabIndex        =   47
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label23 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   50
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label24 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   49
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label25 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   52
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
         TabIndex        =   62
         Top             =   765
         Width           =   5100
         Begin VB.ComboBox cboTableRenteEducation 
            Height          =   315
            Left            =   225
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   900
            Width           =   4695
         End
         Begin VB.TextBox txtFraisGestionRenteEducation 
            Height          =   285
            Left            =   4230
            TabIndex        =   67
            Text            =   "0"
            Top             =   360
            Width           =   465
         End
         Begin VB.TextBox txtTauxTechniqueRenteEducation 
            Height          =   285
            Left            =   1485
            TabIndex        =   64
            Text            =   "4.5"
            Top             =   360
            Width           =   465
         End
         Begin VB.Frame Frame9 
            Caption         =   "Paiement"
            Height          =   1140
            Left            =   225
            TabIndex        =   70
            Top             =   1350
            Width           =   1725
            Begin VB.OptionButton rdoPaiementEchuEducation 
               Caption         =   "A terme échu"
               Height          =   240
               Left            =   180
               TabIndex        =   72
               Top             =   720
               Width           =   1275
            End
            Begin VB.OptionButton rdoPaiementAvanceEducation 
               Caption         =   "D'avance"
               Height          =   240
               Left            =   180
               TabIndex        =   71
               Top             =   360
               Width           =   1275
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Fractionnement"
            Height          =   1140
            Left            =   2340
            TabIndex        =   73
            Top             =   1350
            Width           =   2580
            Begin VB.OptionButton rdoTrimestrielEducation 
               Caption         =   "Trimestriel"
               Height          =   240
               Left            =   135
               TabIndex        =   76
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoMensuelEducation 
               Caption         =   "Mensuel"
               Height          =   240
               Left            =   1395
               TabIndex        =   77
               Top             =   720
               Width           =   1095
            End
            Begin VB.OptionButton rdoAnnuelEducation 
               Caption         =   "Annuel"
               Height          =   240
               Left            =   135
               TabIndex        =   74
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton rdoSemestrielEducation 
               Caption         =   "Semestriel"
               Height          =   240
               Left            =   1395
               TabIndex        =   75
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Label Label5 
            Caption         =   "Taux technique"
            Height          =   375
            Index           =   3
            Left            =   225
            TabIndex        =   63
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label21 
            Caption         =   "Frais de gestion"
            Height          =   375
            Left            =   2970
            TabIndex        =   66
            Top             =   405
            Width           =   1185
         End
         Begin VB.Label Label20 
            Caption         =   "%"
            Height          =   240
            Left            =   2025
            TabIndex        =   65
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label19 
            Caption         =   "%"
            Height          =   240
            Left            =   4770
            TabIndex        =   68
            Top             =   405
            Width           =   240
         End
      End
      Begin VB.TextBox txtFields 
         Height          =   1050
         Index           =   6
         Left            =   1575
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
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
         Height          =   2985
         Left            =   135
         TabIndex        =   18
         Top             =   3195
         Width           =   5460
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
            Height          =   1410
            Left            =   135
            TabIndex        =   29
            Top             =   1440
            Width           =   5190
            Begin VB.CheckBox chkAnnualisationPassage 
               Caption         =   "Calcul de l'annualisation réduite de la provision de passage"
               Height          =   285
               Left            =   135
               TabIndex        =   153
               Top             =   945
               Width           =   4650
            End
            Begin VB.CheckBox chkInterpolationIncap 
               Caption         =   "Interpoler les coeffs de provision Incapacité et Passage"
               Height          =   195
               Left            =   135
               TabIndex        =   30
               Top             =   315
               Width           =   4335
            End
            Begin VB.CheckBox chkLissagePassage_NON 
               Caption         =   "Ne pas lisser les coeffs de Passage (pas de terme correcteur)"
               Height          =   195
               Left            =   135
               TabIndex        =   31
               Top             =   630
               Width           =   4695
            End
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   10
            Left            =   4500
            TabIndex        =   27
            Top             =   1080
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   9
            Left            =   1440
            TabIndex        =   24
            Top             =   1080
            Width           =   675
         End
         Begin VB.ComboBox cboTableIncap 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   270
            Width           =   3840
         End
         Begin VB.ComboBox cboTablePassage 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   675
            Width           =   3840
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   10
            Left            =   3195
            TabIndex        =   26
            Top             =   1125
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux technique"
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   23
            Top             =   1125
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table de passage"
            Height          =   210
            Index           =   8
            Left            =   90
            TabIndex        =   21
            Top             =   735
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table incapacité"
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   19
            Top             =   315
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   25
            Top             =   1125
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   2
            Left            =   5220
            TabIndex        =   28
            Top             =   1125
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
         Height          =   2985
         Left            =   5715
         TabIndex        =   32
         Top             =   3195
         Width           =   5460
         Begin VB.TextBox txtAgeLimiteInvalCat1 
            Height          =   285
            Left            =   3555
            TabIndex        =   136
            Top             =   675
            Width           =   450
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
            TabIndex        =   41
            Top             =   1440
            Width           =   5145
            Begin VB.OptionButton rdoInterpolationInval_Age 
               Caption         =   "Interpolation sur l'âge"
               Height          =   285
               Left            =   180
               TabIndex        =   44
               Top             =   585
               Width           =   1905
            End
            Begin VB.OptionButton rdoInterpolationInval_CorrectionDuree 
               Caption         =   "Correction si durée restante <12 mois"
               Height          =   285
               Left            =   2070
               TabIndex        =   43
               Top             =   270
               Width           =   2940
            End
            Begin VB.OptionButton rdoInterpolationInval_AgeDuree 
               Caption         =   "Interpolation sur l'âge et la durée"
               Height          =   285
               Left            =   2070
               TabIndex        =   45
               Top             =   585
               Width           =   2850
            End
            Begin VB.OptionButton rdoInterpolationInval_NON 
               Caption         =   "Aucune interpolation"
               Height          =   285
               Left            =   180
               TabIndex        =   42
               Top             =   270
               Width           =   1770
            End
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   13
            Left            =   4455
            TabIndex        =   39
            Top             =   1080
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   12
            Left            =   1440
            TabIndex        =   36
            Top             =   1080
            Width           =   675
         End
         Begin VB.ComboBox cboTableInval 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   270
            Width           =   3840
         End
         Begin VB.Label Label2 
            Caption         =   "ans"
            Height          =   240
            Index           =   10
            Left            =   4095
            TabIndex        =   137
            Top             =   720
            Width           =   330
         End
         Begin VB.Label lblLabels 
            Caption         =   "Age limite pour les invalides en 1ere catégorie"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   135
            Top             =   720
            Width           =   3390
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   13
            Left            =   3195
            TabIndex        =   38
            Top             =   1125
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux technique"
            Height          =   255
            Index           =   12
            Left            =   135
            TabIndex        =   35
            Top             =   1125
            Width           =   1185
         End
         Begin VB.Label lblLabels 
            Caption         =   "Table invalidité"
            Height          =   255
            Index           =   11
            Left            =   135
            TabIndex        =   33
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   3
            Left            =   2160
            TabIndex        =   37
            Top             =   1125
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   4
            Left            =   5175
            TabIndex        =   40
            Top             =   1125
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
         Left            =   5715
         TabIndex        =   11
         Top             =   2025
         Width           =   3435
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   15
            Left            =   2295
            TabIndex        =   13
            Top             =   270
            Width           =   675
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   14
            Left            =   2295
            TabIndex        =   16
            Top             =   630
            Width           =   675
         End
         Begin VB.Label lblLabels 
            Caption         =   "Taux des revenus financiers"
            Height          =   255
            Index           =   15
            Left            =   180
            TabIndex        =   12
            Top             =   315
            Width           =   2040
         End
         Begin VB.Label lblLabels 
            Caption         =   "Frais de gestion"
            Height          =   255
            Index           =   14
            Left            =   180
            TabIndex        =   15
            Top             =   675
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   0
            Left            =   3015
            TabIndex        =   14
            Top             =   315
            Width           =   150
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   240
            Index           =   5
            Left            =   3015
            TabIndex        =   17
            Top             =   675
            Width           =   150
         End
      End
      Begin VB.Label lblLabels 
         Caption         =   "Code"
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   1
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom"
         Height          =   255
         Index           =   1
         Left            =   3690
         TabIndex        =   3
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lblLabels 
         Caption         =   "Commentaires"
         Height          =   255
         Index           =   6
         Left            =   255
         TabIndex        =   5
         Top             =   945
         Width           =   1230
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   11415
      TabIndex        =   80
      Top             =   6285
      Width           =   11415
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fermer"
         Height          =   345
         Left            =   5760
         TabIndex        =   79
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Enregistrer"
         Height          =   345
         Left            =   4680
         TabIndex        =   78
         Top             =   45
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmParamCalculIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A680A017C"
Option Explicit

' type de parametre (BILAN, SIMULATION, CLIENT)
'##ModelId=5C8A680A0286
Public typeParam As String
'


'##ModelId=5C8A680A02B4
Private Sub cmdUpdate_Click()
  Dim nomTable As String, theParamSectionName As String
  Dim sel As Long
  
  On Error GoTo errcmdUpdate
   
  ' table incap
  If cboTableIncap.ListIndex = -1 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez choisir une table d'incapacité", vbCritical
    cboTableIncap.SetFocus
    Exit Sub
  End If
  
  ' table passage
  If cboTablePassage.ListIndex = -1 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez choisir une table de passage", vbCritical
    cboTablePassage.SetFocus
    Exit Sub
  End If
  
  ' table inval
  If cboTableInval.ListIndex = -1 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez choisir une table d'invalidité", vbCritical
    cboTableInval.SetFocus
    Exit Sub
  End If
  
  sel = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1)
  If sel < 50 Or sel > 70 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez entrer un chiffre entre 50 et 70 !", vbCritical
    txtAgeLimiteInvalCat1.SetFocus
    Exit Sub
  End If
    
   
  txtCode = Trim(txtCode)
  If Not IsNumeric(txtCode) Then
    SSTab1.Tab = 0
    MsgBox "Vous devez entrer un chiffre !", vbCritical
    txtCode.SetFocus
    Exit Sub
  End If
  
  If CLng(txtCode) > 255 Or CLng(txtCode) < 1 Then
    SSTab1.Tab = 0
    MsgBox "Vous devez entrer un chiffre entre 1 et 255 !", vbCritical
    txtCode.SetFocus
    Exit Sub
  End If
    
  If CLng(txtCode) <> NumParamCalcul And GetSettingIni(CompanyName, DEFAULT_PARAM_SECTION & txtCode, "NumParamCalcul", "QQQQ") <> "QQQQ" Then
    SSTab1.Tab = 0
    MsgBox "Ce code est déjà utilisé, veuillez le modifier !", vbCritical
    txtCode.SetFocus
    Exit Sub
  End If
  
  '*** reassurance
  ' rente education
  If cboTableRenteEducation.ListIndex = -1 Then
    Screen.MousePointer = vbDefault
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir une table de mortalite pour la Rente de Education", vbCritical
    cboTableRenteEducation.SetFocus
    Exit Sub
  End If
  
  ' paiement
  If rdoPaiementAvanceEducation = False And rdoPaiementEchuEducation = False Then
    Screen.MousePointer = vbDefault
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir un type de paiement pour la Rente de Education", vbCritical
    rdoPaiementAvanceEducation.SetFocus
    Exit Sub
  End If
  
  ' fractionnement
  If rdoAnnuelEducation = False And rdoSemestrielEducation = False _
    And rdoTrimestrielEducation = False And rdoMensuelEducation = False Then
    Screen.MousePointer = vbDefault
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir un fractionnement pour la Rente de Education", vbCritical
    rdoAnnuelEducation.SetFocus
    Exit Sub
  End If
  
  ' rente conjoint
  If cboTableRenteConjoint.ListIndex = -1 Then
    Screen.MousePointer = vbDefault
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir une table de mortalite pour la Rente de Conjoint", vbCritical
    cboTableRenteConjoint.SetFocus
    Exit Sub
  End If
  
  ' paiement
  If rdoPaiementAvanceConjoint = False And rdoPaiementEchuConjoint = False Then
    Screen.MousePointer = vbDefault
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir un type de paiement pour la Rente de Conjoint", vbCritical
    rdoPaiementAvanceConjoint.SetFocus
    Exit Sub
  End If
  
  ' fractionnement
  If rdoAnnuelConjoint = False And rdoSemestrielConjoint = False _
     And rdoTrimestrielConjoint = False And rdoMensuelConjoint = False Then
    SSTab1.Tab = 1
    MsgBox "Vous devez choisir un fractionnement pour la Rente de Conjoint", vbCritical
    rdoAnnuelConjoint.SetFocus
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass

  Dim theParam As clsParamCalcul

  Set theParam = New clsParamCalcul
  
  NumParamCalcul = CLng(txtCode)
  
  theParam.CodeParamCalcul = NumParamCalcul
  theParam.NomParamCalcul = txtNomParamCalcul
  theParam.Commentaire = txtFields(6)
   
  ' Incap
  theParam.TauxIncapacite = m_dataHelper.GetDouble2(txtFields(9))
  theParam.FraisGestionIncapacite = m_dataHelper.GetDouble2(txtFields(10))
  theParam.TauxIncapPassageSpecifiqueInval1 = m_dataHelper.GetDouble2(txtFields(16))
  
  ' inval
  theParam.TauxInvalidite = m_dataHelper.GetDouble2(txtFields(12))
  theParam.FraisGestionInvalidite = m_dataHelper.GetDouble2(txtFields(13))
  theParam.AgeLimiteInvalCat1 = m_dataHelper.GetDouble2(txtAgeLimiteInvalCat1)
  
  ' etat C7
  theParam.TauxRevenuC7 = m_dataHelper.GetDouble2(txtFields(15))
  theParam.FraisGestionC7 = m_dataHelper.GetDouble2(txtFields(14))
  
  theParam.TauxIndexation = m_dataHelper.GetDouble2(txtTxIndex)
  theParam.DureeIndexation = m_dataHelper.GetDouble2(txtDureeIndex)
  theParam.TMO = m_dataHelper.GetDouble2(txtTMO)
  theParam.DureeLissage = m_dataHelper.GetDouble2(txtDureeLissage)
    
  ' rente conjoint
  theParam.TauxTechniqueRenteConjoint = m_dataHelper.GetDouble2(txtTauxTechniqueRenteConjoint)
  theParam.FraisGestionRenteConjoint = m_dataHelper.GetDouble2(txtFraisGestionRenteConjoint)
  
  If rdoAnnuelConjoint.Value = True Then
    theParam.FractionnementRenteConjoint = eFractionnementAnnuel
  ElseIf rdoSemestrielConjoint.Value = True Then
    theParam.FractionnementRenteConjoint = eFractionnementSemestriel
  ElseIf rdoTrimestrielConjoint.Value = True Then
    theParam.FractionnementRenteConjoint = eFractionnementTrimestriel
  Else
    theParam.FractionnementRenteConjoint = eFractionnementMensuel
  End If
  
  If rdoPaiementAvanceConjoint.Value = True Then
    theParam.PaiementRenteConjoint = ePaiementAvance
  Else
    theParam.PaiementRenteConjoint = ePaiementEchu
  End If
  
  ' rente education
  theParam.TauxTechniqueRenteEducation = m_dataHelper.GetDouble2(txtTauxTechniqueRenteEducation)
  theParam.FraisGestionRenteEducation = m_dataHelper.GetDouble2(txtFraisGestionRenteEducation)
  
  If rdoAnnuelEducation.Value = True Then
    theParam.FractionnementRenteEducation = eFractionnementAnnuel
  ElseIf rdoSemestrielEducation.Value = True Then
    theParam.FractionnementRenteEducation = eFractionnementSemestriel
  ElseIf rdoTrimestrielEducation.Value = True Then
    theParam.FractionnementRenteEducation = eFractionnementTrimestriel
  Else
    theParam.FractionnementRenteEducation = eFractionnementMensuel
  End If
  
  If rdoPaiementAvanceEducation.Value = True Then
    theParam.PaiementRenteEducation = ePaiementAvance
  Else
    theParam.PaiementRenteEducation = ePaiementEchu
  End If

  ' maintien deces
  theParam.FraisGestionCapitauxDecesDC = m_dataHelper.GetDouble2(txtFraisGestionCapitauxDecesDC)
  theParam.FraisGestionRenteEducationDC = m_dataHelper.GetDouble2(txtFraisGestionRenteEducationDC)
  
  theParam.PourcentageAccident = m_dataHelper.GetDouble2(txtPourcentageAccident)
  
  theParam.FraisGestionRenteConjointDC = m_dataHelper.GetDouble2(txtFraisGestionRenteConjointDC)
  theParam.CapitalMoyenRenteConjointTempoDC = m_dataHelper.GetDouble2(txtCapitalMoyenRenteConjointTempoDC)
  theParam.CapitalMoyenRenteConjointViagereDC = m_dataHelper.GetDouble2(txtCapitalMoyenRenteConjointViagereDC)
  theParam.ForcerCapitalMoyenRenteConjointDC = IIf(chkForcerCapitalMoyenRteConjoitDC.Value = vbChecked, True, False)
  theParam.AgeConjointRenteConjointDC = m_dataHelper.GetDouble2(txtAgeConjointRenteConjointDC)
  
  theParam.UtiliserTableLissageProvision = rdoAvecLissage.Value = True
  
  '
  ' recalcul
  '
  
  'Call SaveSettingIni(CompanyName, theParamSectionName, "PMGDForcerInval", "0")
  theParam.RecalculCoeffBCAC = rdoCalculCoeffBCAC.Value = True
  
  theParam.MethodeCalculDC = IIf(rdoPctPMCalculeeDC.Value = True, ePctProvisionCalculee, eCapitauxConstitutifs)
  theParam.PctPMCalculeeDC = m_dataHelper.GetDouble2(txtPctPMCalculeeDC)
  
  ' combo incap
  If cboTableIncapCalculDC.ListIndex <> -1 Then
    theParam.LoiIncapaciteDC = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncapCalculDC.ItemData(cboTableIncapCalculDC.ListIndex))
  End If
  
  ' combo invalidite
  If cboTableInvalCalculDC.ListIndex <> -1 Then
    theParam.LoiInvaliditeDC = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInvalCalculDC.ItemData(cboTableInvalCalculDC.ListIndex))
  End If
  
  theParam.TauxTechnicCalculDC = m_dataHelper.GetDouble2(txtTauxTechnicCalculDC)
  theParam.FraisGestionCalculDC = m_dataHelper.GetDouble2(txtFraisGestionCalculDC)
  
  theParam.AgeLimiteCalulDC = m_dataHelper.GetDouble2(txtAgeLimiteCalulDC)
  
  '
  ' lecture
  '
  
  ' combo incap
  If cboTableIncapPrecalculDC.ListIndex <> -1 Then
    theParam.TableIncapacitePrecalculDC = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncapPrecalculDC.ItemData(cboTableIncapPrecalculDC.ListIndex))
  End If
  
  ' combo invalidite
  If cboTableInvalPrecalculDC.ListIndex <> -1 Then
    theParam.TableInvaliditePrecalculDC = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInvalPrecalculDC.ItemData(cboTableInvalPrecalculDC.ListIndex))
  End If
  
  
  theParam.CalculAge_Anniversaire = rdoAgeAnniversaire.Value = True


  theParam.LissageCoeffPassage = chkLissagePassage_NON.Value = vbUnchecked
  theParam.BridageAgeLimiteTable = chkBridageAge.Value = vbChecked
  
  
  theParam.PortefeuilleSalaries = IIf(chkPortefeuilleSalaries.Value = vbChecked, True, False)
  
  
  ' interpolation de l'inval
  If rdoInterpolationInval_CorrectionDuree.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_CorrectionDuree
  ElseIf rdoInterpolationInval_Age.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_Age
  ElseIf rdoInterpolationInval_AgeDuree.Value = True Then
    theParam.InterpolationInvalidite = eInterpolationInval_AgeDuree
  Else
    theParam.InterpolationInvalidite = eInterpolationInval_NON
  End If
    

  ' combo incap
  theParam.LoiIncapacite = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableIncap.ItemData(cboTableIncap.ListIndex))
  
  ' passage
  theParam.LoiPassage = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTablePassage.ItemData(cboTablePassage.ListIndex))
  
  ' invalidite
  theParam.LoiInvalidite = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableInval.ItemData(cboTableInval.ListIndex))
  
  ' rente conjoint
  theParam.TableRenteConjoint = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableRenteConjoint.ItemData(cboTableRenteConjoint.ListIndex))
  
  ' rente Education
  theParam.TableRenteEducation = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & cboTableRenteEducation.ItemData(cboTableRenteEducation.ListIndex))
  
    
  theParam.SaveToIni False, typeParam
  
  
  Screen.MousePointer = vbDefault
  
  Unload Me
  
  Exit Sub
  
errcmdUpdate:
  MsgBox "erreur " & Err & vbLf & Err.Description
  Exit Sub
  Resume Next
End Sub

'##ModelId=5C8A680A02D4
Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

'##ModelId=5C8A680A02E3
Private Sub Form_Load()
  Dim nomTable As String, theParamSectionName As String
  Dim sel As Long, rs As ADODB.Recordset
  Dim theParam As clsParamCalcul
  
  Screen.MousePointer = vbHourglass
  
  ' activate first tab
  SSTab1.Tab = 0
  
  ' chargement des parametres depuis le .INI
  Set theParam = New clsParamCalcul
  
  theParam.LoadFromIni NumParamCalcul, typeParam
  
  If NumParamCalcul = -1 Then
    ' Nouveau jeu de parametres
    theParamSectionName = SectionName
  
    For NumParamCalcul = 1 To 255
      If GetSettingIni(CompanyName, DEFAULT_PARAM_SECTION & NumParamCalcul, "NumParamCalcul", "QQQQ") = "QQQQ" Then
        Exit For
      End If
    Next
  
  Else
    
    ' jeu de parametres existants
    theParamSectionName = DEFAULT_PARAM_SECTION & NumParamCalcul
    
  End If
  
  txtCode = NumParamCalcul
  txtNomParamCalcul = theParam.NomParamCalcul
  txtFields(6) = theParam.Commentaire
   
  'valeur par defaut
  txtFields(9) = theParam.TauxIncapacite
  txtFields(10) = theParam.FraisGestionIncapacite
  txtFields(16) = theParam.TauxIncapPassageSpecifiqueInval1
  txtFields(12) = theParam.TauxInvalidite
  txtFields(13) = theParam.FraisGestionInvalidite
  
  txtAgeLimiteInvalCat1 = theParam.AgeLimiteInvalCat1
  
  txtFields(15) = theParam.TauxRevenuC7
  txtFields(14) = theParam.FraisGestionC7
  
  txtTxIndex = theParam.TauxIndexation
  txtDureeIndex = theParam.DureeIndexation
  txtTMO = theParam.TMO
  txtDureeLissage = theParam.DureeLissage
    
  ' rente conjoint
  txtTauxTechniqueRenteConjoint = theParam.TauxTechniqueRenteConjoint
  txtFraisGestionRenteConjoint = theParam.FraisGestionRenteConjoint
  
  Select Case theParam.FractionnementRenteConjoint
    Case eFractionnementAnnuel
      rdoAnnuelConjoint.Value = True
    Case eFractionnementSemestriel
      rdoSemestrielConjoint.Value = True
    Case eFractionnementTrimestriel
      rdoTrimestrielConjoint.Value = True
    Case Else
      rdoMensuelConjoint.Value = True
  End Select
  
  If theParam.PaiementRenteConjoint = ePaiementAvance Then
    rdoPaiementAvanceConjoint.Value = True
  Else
    rdoPaiementEchuConjoint.Value = True
  End If

  ' rente education
  txtTauxTechniqueRenteEducation = theParam.TauxTechniqueRenteEducation
  txtFraisGestionRenteEducation = theParam.FraisGestionRenteEducation
  
  Select Case theParam.FractionnementRenteEducation
    Case eFractionnementAnnuel
      rdoAnnuelEducation.Value = True
    Case eFractionnementSemestriel
      rdoSemestrielEducation.Value = True
    Case eFractionnementTrimestriel
      rdoTrimestrielEducation.Value = True
    Case Else
      rdoMensuelEducation.Value = True
  End Select
  
  If theParam.PaiementRenteEducation = ePaiementAvance Then
    rdoPaiementAvanceEducation.Value = True
  Else
    rdoPaiementEchuEducation.Value = True
  End If

  ' maintien deces
  txtFraisGestionCapitauxDecesDC = theParam.FraisGestionCapitauxDecesDC
  txtFraisGestionRenteEducationDC = theParam.FraisGestionRenteEducationDC
  
  txtPourcentageAccident = theParam.PourcentageAccident
  
  txtFraisGestionRenteConjointDC = theParam.FraisGestionRenteConjointDC
  txtCapitalMoyenRenteConjointTempoDC = theParam.CapitalMoyenRenteConjointTempoDC
  txtCapitalMoyenRenteConjointViagereDC = theParam.CapitalMoyenRenteConjointViagereDC
  chkForcerCapitalMoyenRteConjoitDC.Value = IIf(theParam.ForcerCapitalMoyenRenteConjointDC = True, vbChecked, vbUnchecked)
  txtAgeConjointRenteConjointDC = theParam.AgeConjointRenteConjointDC
  
  ' Lissage des provision
  rdoSansLissage.Value = theParam.UtiliserTableLissageProvision = False
  rdoAvecLissage.Value = theParam.UtiliserTableLissageProvision = True
  
  '
  ' recalcul
  '
  
  'chkPMGDForcerInval = IIf(GetSettingIni(CompanyName, theParamSectionName, "PMGDForcerInval", "0") = "0", vbUnchecked, vbChecked)
  chkPMGDForcerInval.Value = vbUnchecked
  rdoCalculCoeffBCAC.Value = theParam.RecalculCoeffBCAC = True
  
  rdoCapitauxConstitif.Value = theParam.MethodeCalculDC = eCapitauxConstitutifs
  rdoPctPMCalculeeDC.Value = theParam.MethodeCalculDC = ePctProvisionCalculee
  txtPctPMCalculeeDC = theParam.PctPMCalculeeDC

  If rdoCapitauxConstitif.Value = False And rdoPctPMCalculeeDC.Value = False Then
    rdoCapitauxConstitif.Value = True
  End If
  
  
  ' rempli le combo incap
  nomTable = theParam.LoiIncapaciteDC
  If nomTable <> "#" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
  Else
    sel = -1
  End If
  m_dataHelper.FillCombo cboTableIncapCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteIncap, sel
  If sel = -1 Then
    cboTableIncapCalculDC.ListIndex = 0
  End If
  
  ' rempli le combo invalidite
  nomTable = theParam.LoiInvaliditeDC
  If nomTable <> "#" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableInvalCalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTable_MortaliteInval, sel
  If sel = -1 Then
    cboTableInvalCalculDC.ListIndex = 0
  End If
  
  txtTauxTechnicCalculDC = theParam.TauxTechnicCalculDC
  txtFraisGestionCalculDC = theParam.FraisGestionCalculDC
  
  txtAgeLimiteCalulDC = theParam.AgeLimiteCalulDC
  
  '
  ' lecture
  '
  rdoLireCoeffBCAC = theParam.RecalculCoeffBCAC = False
  
  ' rempli le combo incap
  nomTable = theParam.TableIncapacitePrecalculDC
  If nomTable <> "#" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
  Else
    sel = -1
  End If
  m_dataHelper.FillCombo cboTableIncapPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACIncap, sel
  If sel = -1 Then
    cboTableIncapPrecalculDC.ListIndex = 0
  End If
  
  ' rempli le combo invalidite
  nomTable = theParam.TableInvaliditePrecalculDC
  If nomTable <> "#" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableInvalPrecalculDC, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = " & cdTypeTableCoeffBCACInval, sel
  If sel = -1 Then
    cboTableInvalPrecalculDC.ListIndex = 0
  End If

  ' calcul de l'age
  If theParam.CalculAge_Anniversaire = True Then
    rdoAgeMillesime.Value = False
    rdoAgeAnniversaire.Value = True
  Else
    rdoAgeMillesime.Value = True
    rdoAgeAnniversaire.Value = False
  End If


  chkInterpolationIncap.Value = vbChecked
  chkLissagePassage_NON.Value = IIf(theParam.LissageCoeffPassage, vbUnchecked, vbChecked)
  chkBridageAge.Value = IIf(theParam.BridageAgeLimiteTable, vbChecked, vbUnchecked)
  
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
    
  ' Incap toujours interpolée
  chkInterpolationIncap.Enabled = False
    
'*********
    
  ' rempli le combo incap
  nomTable = theParam.LoiIncapacite
  If nomTable <> "" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableIncap, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 1", sel
  If sel = -1 Then
    cboTableIncap.ListIndex = 0
  End If
  
  ' passage
  nomTable = theParam.LoiPassage
  If nomTable <> "" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTablePassage, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 2", sel
  If sel = -1 Then
    cboTablePassage.ListIndex = 0
  End If
  
  ' invalidite
  nomTable = theParam.LoiInvalidite
  If nomTable <> "" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableInval, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE = 3", sel
  If sel = -1 Then
    cboTableInval.ListIndex = 0
  End If
  
  ' rente conjoint
  nomTable = theParam.TableRenteConjoint
  If nomTable <> "" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableRenteConjoint, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel
  If sel = -1 Then
    cboTableRenteConjoint.ListIndex = 0
  End If
  
  ' rente Education
  nomTable = theParam.TableRenteEducation
  If nomTable <> "" Then
    sel = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = """ & nomTable & """")
  Else
    sel = -1
  End If
  
  m_dataHelper.FillCombo cboTableRenteEducation, "SELECT TABLECLE, LIBTABLE FROM ListeTableLoi WHERE TYPETABLE=" & cdTypeTableMortalite & " OR TYPETABLE=" & cdTypeTableGeneration, sel
  If sel = -1 Then
    cboTableRenteConjoint.ListIndex = 0
  End If
  
  
  chkPortefeuilleSalaries.Value = IIf(theParam.PortefeuilleSalaries = True, vbChecked, vbUnchecked)
  
  
  Screen.MousePointer = vbDefault

End Sub



'##ModelId=5C8A680A02F3
Private Sub Form_Unload(Cancel As Integer)
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A680A0312
Private Sub EnableCoeffBCAC()
  txtTauxTechnicCalculDC.Enabled = rdoCalculCoeffBCAC = True
  txtFraisGestionCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableIncapCalculDC.Enabled = rdoCalculCoeffBCAC = True
  cboTableInvalCalculDC.Enabled = rdoCalculCoeffBCAC = True
  
  cboTableIncapPrecalculDC.Enabled = rdoLireCoeffBCAC = True
  cboTableInvalPrecalculDC.Enabled = rdoLireCoeffBCAC = True
End Sub



'##ModelId=5C8A680A0322
Private Sub rdoCalculCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

'##ModelId=5C8A680A0341
Private Sub rdoCotisationsExonerees_Click()
  If rdoCapitauxConstitif.Value = True Then
    txtPctPMCalculeeDC.Enabled = False
  Else
    txtPctPMCalculeeDC.Enabled = True
  End If
End Sub

'##ModelId=5C8A680A0360
Private Sub rdoCapitauxConstitif_Click()
  If rdoCapitauxConstitif.Value = True Then
    txtPctPMCalculeeDC.Enabled = False
  Else
    txtPctPMCalculeeDC.Enabled = True
  End If
End Sub

'##ModelId=5C8A680A0380
Private Sub rdoLireCoeffBCAC_Click()
  EnableCoeffBCAC
End Sub

