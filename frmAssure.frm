VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAssure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saisie manuelle des Assurés"
   ClientHeight    =   6945
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmAssure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6615
      Width           =   5730
      _ExtentX        =   10107
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
      Caption         =   " Cliquez sur les flèches pour parcourir la table..."
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
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   61
      Top             =   6165
      Width           =   5730
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5325
      Left            =   45
      TabIndex        =   29
      Top             =   540
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   9393
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frmAssure.frx":1BB2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblLabels(15)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(13)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabels(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblLabels(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLabels(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLabels(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblLabels(45)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DataSociete"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "DTPicker5"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "DTPicker4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DTPicker3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "DTPicker2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DTPicker1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFields(15)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtFields(9)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtFields(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtFields(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtFields(6)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtFields(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtFields(35)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "DataGarantie"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "DBCombo1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "DBCombo2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frmAssure.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabels(27)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(26)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(25)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabels(24)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(23)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabels(22)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabels(21)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabels(20)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblLabels(19)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLabels(17)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblLabels(2)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblLabels(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblLabels(0)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblLabels(16)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblLabels(28)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblLabels(35)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblLabels(43)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtFields(27)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtFields(26)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtFields(25)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtFields(24)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtFields(23)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtFields(22)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtFields(21)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtFields(20)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtFields(19)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtFields(17)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtFields(2)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtFields(1)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtFields(0)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtFields(16)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtFields(3)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtFields(5)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtFields(28)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtFields(33)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Command1"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Page 3"
      TabPicture(2)   =   "frmAssure.frx":1BEA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(32)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblLabels(34)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblLabels(36)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblLabels(37)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblLabels(31)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblLabels(41)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblLabels(42)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblLabels(30)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblLabels(29)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblLabels(44)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "lblLabels(33)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "lblLabels(46)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lblLabels(47)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtFields(13)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtFields(18)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtFields(29)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtFields(30)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtFields(12)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtFields(31)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtFields(32)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtFields(11)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtFields(10)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtFields(34)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtFields(14)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtFields(36)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtFields(37)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "Page 4"
      TabPicture(3)   =   "frmAssure.frx":1C06
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblLabels(48)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblLabels(49)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblLabels(40)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblLabels(39)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblLabels(38)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lblLabels(18)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblLabels(66)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "DTPicker10"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "DTPicker6"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "DTPicker9"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "DTPicker8"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "DTPicker7"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cboSituationFamille"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "dtaSituationFamille"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Check1"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtFields(38)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "chkDateEstimee"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).ControlCount=   17
      TabCaption(4)   =   "Page 5"
      TabPicture(4)   =   "frmAssure.frx":1C22
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblLabels(58)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblLabels(59)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "lblLabels(60)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lblLabels(61)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lblLabels(62)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "lblLabels(63)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "lblLabels(64)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "lblLabels(65)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "lblLabels(50)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "lblLabels(51)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "lblLabels(52)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "lblLabels(53)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "lblLabels(56)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "lblLabels(57)"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "txtFields(45)"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "txtFields(46)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "txtFields(47)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "txtFields(48)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "txtFields(49)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "txtFields(50)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "txtFields(51)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "txtFields(52)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "txtFields(39)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "txtFields(40)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "txtFields(41)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "txtFields(42)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "txtFields(43)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "txtFields(44)"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).ControlCount=   28
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_RCJT_1R"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   37
         Left            =   -71550
         TabIndex        =   140
         Top             =   2655
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_REDUC_1R"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   36
         Left            =   -71550
         TabIndex        =   139
         Top             =   2970
         Width           =   1935
      End
      Begin VB.CheckBox chkDateEstimee 
         Alignment       =   1  'Right Justify
         Caption         =   "Date de paiement estimée"
         DataField       =   "PODATEPAIEMENTESTIMEE"
         DataSource      =   "datPrimaryRS"
         Height          =   240
         Left            =   225
         TabIndex        =   132
         Top             =   3645
         Width           =   3390
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCoeffBCAC"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   44
         Left            =   -71550
         TabIndex        =   129
         Top             =   4950
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPourcentLissage"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   43
         Left            =   -71550
         TabIndex        =   128
         Top             =   4620
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POMajoEnfant"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   42
         Left            =   -71550
         TabIndex        =   123
         Top             =   4305
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POTauxGarantieDC"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   41
         Left            =   -71550
         TabIndex        =   122
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PONbEnfant"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   40
         Left            =   -71550
         TabIndex        =   121
         Top             =   3675
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POAgeMoyenEnfant"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   39
         Left            =   -71550
         TabIndex        =   120
         Top             =   3990
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PORegimeDeces"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   52
         Left            =   -71550
         TabIndex        =   111
         Top             =   450
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCategorieDeces"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   51
         Left            =   -71550
         TabIndex        =   110
         Top             =   768
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PORegimeRenteConjointTempo"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   50
         Left            =   -71550
         TabIndex        =   109
         Top             =   1170
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCategorieRenteConjointTempo"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   49
         Left            =   -71550
         TabIndex        =   108
         Top             =   1500
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PORegimeRenteEduc"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   48
         Left            =   -71550
         TabIndex        =   107
         Top             =   2625
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PORegimeRenteConjointViagere"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   47
         Left            =   -71550
         TabIndex        =   106
         Top             =   1905
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCategorieRenteEduc"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   46
         Left            =   -71550
         TabIndex        =   105
         Top             =   2940
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCategorieRenteConjointViagere"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   45
         Left            =   -71550
         TabIndex        =   104
         Top             =   2220
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POSalaireAnnuel"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   38
         Left            =   3450
         TabIndex        =   98
         Top             =   900
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "L'assuré est Cadre"
         DataField       =   "POIsCadre"
         DataSource      =   "datPrimaryRS"
         Height          =   240
         Left            =   180
         TabIndex        =   97
         Top             =   630
         Width           =   3480
      End
      Begin MSDataListLib.DataCombo DBCombo2 
         Bindings        =   "frmAssure.frx":1C3E
         DataField       =   "POGARCLE"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   -72930
         TabIndex        =   96
         Top             =   1080
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "GALIB"
         BoundColumn     =   "GAGARCLE"
         Text            =   "DBComboGarantie"
      End
      Begin MSDataListLib.DataCombo DBCombo1 
         Bindings        =   "frmAssure.frx":1C59
         DataField       =   "POSTECLE"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   -72930
         TabIndex        =   95
         Top             =   720
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "SONOM"
         BoundColumn     =   "SOCLE"
         Text            =   "DBComboSociete"
      End
      Begin MSAdodcLib.Adodc DataGarantie 
         Height          =   330
         Left            =   -71355
         Top             =   3240
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "DataGarantie"
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
      Begin VB.TextBox txtFields 
         DataField       =   "POCATEGORIE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   35
         Left            =   -72945
         MaxLength       =   10
         TabIndex        =   23
         Top             =   4635
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPSAP"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   -71550
         TabIndex        =   92
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Valeurs possibles"
         Height          =   300
         Left            =   -71130
         TabIndex        =   91
         Top             =   2700
         Width           =   1425
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POTRAITE_RASSUR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   34
         Left            =   -71550
         TabIndex        =   89
         Top             =   4230
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_REDUC_1F"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   33
         Left            =   -71640
         TabIndex        =   86
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_RCJT_1F"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   28
         Left            =   -71640
         TabIndex        =   85
         Top             =   4005
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_RASSUR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   -71550
         TabIndex        =   82
         Top             =   3285
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPSAP_RASSUR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   -71550
         TabIndex        =   81
         Top             =   3915
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_INVAL_1R"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   32
         Left            =   -71550
         TabIndex        =   77
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_PASS_1R"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   31
         Left            =   -71550
         TabIndex        =   76
         Top             =   2025
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_INCAP_1R"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   -71550
         TabIndex        =   75
         Top             =   1710
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_REVALO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   30
         Left            =   -71550
         TabIndex        =   70
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_RI"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   29
         Left            =   -71550
         TabIndex        =   69
         Top             =   765
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_SORTIE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   18
         Left            =   -71550
         TabIndex        =   68
         Top             =   450
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCOT_REVALO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   -71550
         TabIndex        =   67
         Top             =   1395
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PONOM"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -73020
         MaxLength       =   60
         TabIndex        =   64
         Top             =   495
         Width           =   3330
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PONUMCLE"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -74325
         MaxLength       =   10
         TabIndex        =   62
         Top             =   495
         Width           =   1260
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POTYPEF"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   16
         Left            =   -71640
         TabIndex        =   59
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "RECNO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   0
         Left            =   -70050
         TabIndex        =   53
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POGPECLE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   -70050
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPERCLE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   2
         Left            =   -70050
         TabIndex        =   51
         Top             =   540
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PONUMCLE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   -72930
         MaxLength       =   10
         TabIndex        =   3
         Top             =   405
         Width           =   1350
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCONVENTION"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   6
         Left            =   -72945
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PONOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   -72945
         MaxLength       =   60
         TabIndex        =   7
         Top             =   1755
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POSEXE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   -72945
         MaxLength       =   1
         TabIndex        =   9
         Top             =   2085
         Width           =   675
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCSP"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   -72945
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POCAUSE"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   15
         Left            =   -72945
         MaxLength       =   10
         TabIndex        =   22
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "PODELAI"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   17
         Left            =   -71640
         TabIndex        =   31
         Top             =   1125
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPRESTATION"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   19
         Left            =   -71640
         TabIndex        =   33
         Top             =   1755
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPRESTATION_AN_PREC"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   20
         Left            =   -71640
         TabIndex        =   35
         Top             =   2085
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPRESTATION_AN"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   -71640
         TabIndex        =   37
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POSIT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   22
         Left            =   -71640
         TabIndex        =   39
         Top             =   2715
         Width           =   450
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_INCAP_1F"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   23
         Left            =   -71640
         TabIndex        =   41
         Top             =   3045
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_PASS_1F"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   24
         Left            =   -71640
         TabIndex        =   43
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_INVAL_1F"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   25
         Left            =   -71640
         TabIndex        =   45
         Top             =   3675
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   26
         Left            =   -71640
         TabIndex        =   47
         Top             =   4635
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         DataField       =   "POPM_VAR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   27
         Left            =   -71640
         TabIndex        =   49
         Top             =   4950
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "PONAIS"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   -72930
         TabIndex        =   13
         Top             =   2700
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "POEFFET"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   -72930
         TabIndex        =   15
         Top             =   3015
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         DataField       =   "POTERME"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   -72930
         TabIndex        =   17
         Top             =   3330
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         DataField       =   "POARRET"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   -72930
         TabIndex        =   19
         Top             =   3645
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker5 
         DataField       =   "POREPRISE"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   -72930
         TabIndex        =   21
         Top             =   3960
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSAdodcLib.Adodc DataSociete 
         Height          =   330
         Left            =   -71355
         Top             =   2835
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "DataSociete"
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
      Begin MSAdodcLib.Adodc dtaSituationFamille 
         Height          =   330
         Left            =   3735
         Top             =   585
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
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
         Caption         =   "dtaSituationFamille"
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
      Begin MSDataListLib.DataCombo cboSituationFamille 
         Bindings        =   "frmAssure.frx":1C73
         DataField       =   "POCleSituationFamille"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Left            =   3450
         TabIndex        =   100
         Top             =   1215
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "Libelle"
         BoundColumn     =   "CleSituationFamille"
         Text            =   "DBComboGarantie"
      End
      Begin MSComCtl2.DTPicker DTPicker7 
         DataField       =   "PODERNIERPAIEMENT"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   3450
         TabIndex        =   133
         Top             =   1890
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker8 
         DataField       =   "PODEBUT"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   3450
         TabIndex        =   134
         Top             =   2790
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker9 
         DataField       =   "POFIN"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   3450
         TabIndex        =   135
         Top             =   3150
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         DataField       =   "POPREMIER_PAIEMENT"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   3450
         TabIndex        =   143
         Top             =   2250
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin MSComCtl2.DTPicker DTPicker10 
         DataField       =   "PODATEENTREEINVAL"
         DataSource      =   "datPrimaryRS"
         Height          =   330
         Left            =   3450
         TabIndex        =   145
         Top             =   4050
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   224198659
         CurrentDate     =   36114
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date d'entrée en Invalidité"
         Height          =   255
         Index           =   66
         Left            =   225
         TabIndex        =   146
         Top             =   4080
         Width           =   2040
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date du premier paiement"
         Height          =   255
         Index           =   18
         Left            =   225
         TabIndex        =   144
         Top             =   2340
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Rente Conjoint pour 1F revalorisée"
         Height          =   255
         Index           =   47
         Left            =   -74775
         TabIndex        =   142
         Top             =   2700
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Rente Education pour 1F revalorisée"
         Height          =   255
         Index           =   46
         Left            =   -74775
         TabIndex        =   141
         Top             =   3015
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date de Dernier paiement"
         Height          =   255
         Index           =   38
         Left            =   225
         TabIndex        =   138
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Début de la dernière prestation"
         Height          =   255
         Index           =   39
         Left            =   225
         TabIndex        =   137
         Top             =   2865
         Width           =   2355
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fin de la dernière prestation"
         Height          =   255
         Index           =   40
         Left            =   225
         TabIndex        =   136
         Top             =   3180
         Width           =   2040
      End
      Begin VB.Label lblLabels 
         Caption         =   "Coefficient précalculé du BCAC"
         Height          =   255
         Index           =   57
         Left            =   -74775
         TabIndex        =   131
         Top             =   4995
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Pourcentage de lissage des provisions"
         Height          =   255
         Index           =   56
         Left            =   -74775
         TabIndex        =   130
         Top             =   4665
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Majoration par enfants à charge"
         Height          =   255
         Index           =   53
         Left            =   -74775
         TabIndex        =   127
         Top             =   4350
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Taux Garantie Décès"
         Height          =   255
         Index           =   52
         Left            =   -74775
         TabIndex        =   126
         Top             =   3405
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre d'enfants"
         Height          =   255
         Index           =   51
         Left            =   -74775
         TabIndex        =   125
         Top             =   3720
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Age moyen des enfants"
         Height          =   255
         Index           =   50
         Left            =   -74775
         TabIndex        =   124
         Top             =   4035
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Régime Garantie Décès"
         Height          =   255
         Index           =   65
         Left            =   -74775
         TabIndex        =   119
         Top             =   495
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Catégorie Garantie Décès"
         Height          =   255
         Index           =   64
         Left            =   -74775
         TabIndex        =   118
         Top             =   813
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Regime Rente de Conjoint Temporaire"
         Height          =   255
         Index           =   63
         Left            =   -74775
         TabIndex        =   117
         Top             =   1215
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Categorie Rente de Conjoint Temporaire"
         Height          =   255
         Index           =   62
         Left            =   -74775
         TabIndex        =   116
         Top             =   1545
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Régime Rente Education"
         Height          =   255
         Index           =   61
         Left            =   -74775
         TabIndex        =   115
         Top             =   2670
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Régime Rente de Conjoint Viagère"
         Height          =   255
         Index           =   60
         Left            =   -74775
         TabIndex        =   114
         Top             =   1950
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Catégorie Rente Education"
         Height          =   255
         Index           =   59
         Left            =   -74775
         TabIndex        =   113
         Top             =   2985
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Catégorie Rentede Conjoint Viagère"
         Height          =   255
         Index           =   58
         Left            =   -74775
         TabIndex        =   112
         Top             =   2265
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Situation de Famille"
         Height          =   255
         Index           =   49
         Left            =   225
         TabIndex        =   101
         Top             =   1260
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         Caption         =   "Salaire Annuel"
         Height          =   255
         Index           =   48
         Left            =   225
         TabIndex        =   99
         Top             =   945
         Width           =   3030
      End
      Begin VB.Label lblLabels 
         Caption         =   "Catégorie"
         Height          =   255
         Index           =   45
         Left            =   -74865
         TabIndex        =   94
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "PSAP"
         Height          =   255
         Index           =   33
         Left            =   -74775
         TabIndex        =   93
         Top             =   3645
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Traité de réassurance"
         Height          =   255
         Index           =   44
         Left            =   -74775
         TabIndex        =   90
         Top             =   4275
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Rte Education pour 1F annuelle"
         Height          =   255
         Index           =   43
         Left            =   -74865
         TabIndex        =   88
         Top             =   4365
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Rte Conjoint pour 1F annuelle"
         Height          =   255
         Index           =   35
         Left            =   -74865
         TabIndex        =   87
         Top             =   4050
         Width           =   3030
      End
      Begin VB.Label lblLabels 
         Caption         =   "Part PM de la réassurance"
         Height          =   255
         Index           =   29
         Left            =   -74775
         TabIndex        =   84
         Top             =   3330
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Part PSAP de la réassurance"
         Height          =   255
         Index           =   30
         Left            =   -74775
         TabIndex        =   83
         Top             =   3960
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM invalidité  pour 1F annuelle revalorisée"
         Height          =   255
         Index           =   42
         Left            =   -74775
         TabIndex        =   80
         Top             =   2385
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM passage pour 1F annuelle revalorisée"
         Height          =   255
         Index           =   41
         Left            =   -74775
         TabIndex        =   79
         Top             =   2070
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM incapacité pour 1F mensuel revalorisée"
         Height          =   255
         Index           =   31
         Left            =   -74775
         TabIndex        =   78
         Top             =   1755
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Mnt PM de la revalorisation uniquement"
         DataField       =   "PM"
         Height          =   255
         Index           =   37
         Left            =   -74775
         TabIndex        =   74
         Top             =   1125
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Rte Inval avant revalorisation"
         Height          =   255
         Index           =   36
         Left            =   -74775
         TabIndex        =   73
         Top             =   810
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Montant PM non RàZ pour les sorties"
         Height          =   255
         Index           =   34
         Left            =   -74775
         TabIndex        =   72
         Top             =   495
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cotisation annuelle pour financer la revalo"
         Height          =   255
         Index           =   32
         Left            =   -74775
         TabIndex        =   71
         Top             =   1440
         Width           =   3210
      End
      Begin VB.Label Label2 
         Caption         =   "1 : Homme ou 2 : Femme"
         Height          =   240
         Left            =   -72165
         TabIndex        =   65
         Top             =   2115
         Width           =   2625
      End
      Begin VB.Label lblLabels 
         Caption         =   "Police"
         Height          =   255
         Index           =   28
         Left            =   -74865
         TabIndex        =   63
         Top             =   540
         Width           =   555
      End
      Begin VB.Label lblLabels 
         Caption         =   "Type de franchise"
         Height          =   255
         Index           =   16
         Left            =   -74865
         TabIndex        =   60
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Société *"
         Height          =   255
         Index           =   3
         Left            =   -74865
         TabIndex        =   58
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Caption         =   "RECNO:"
         Height          =   255
         Index           =   0
         Left            =   -70905
         TabIndex        =   56
         Top             =   270
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblLabels 
         Caption         =   "POGPECLE:"
         Height          =   255
         Index           =   1
         Left            =   -71040
         TabIndex        =   55
         Top             =   405
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "POPERCLE:"
         Height          =   255
         Index           =   2
         Left            =   -71085
         TabIndex        =   54
         Top             =   495
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "N° de Police *"
         Height          =   255
         Index           =   4
         Left            =   -74865
         TabIndex        =   2
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Caption         =   "Garantie *"
         Height          =   255
         Index           =   5
         Left            =   -74865
         TabIndex        =   1
         Top             =   1125
         Width           =   1905
      End
      Begin VB.Label lblLabels 
         Caption         =   "NCA"
         Height          =   255
         Index           =   6
         Left            =   -74865
         TabIndex        =   4
         Top             =   1485
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom, prénom *"
         Height          =   255
         Index           =   7
         Left            =   -74865
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sexe *"
         Height          =   255
         Index           =   8
         Left            =   -74865
         TabIndex        =   8
         Top             =   2130
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "CSP"
         Height          =   255
         Index           =   9
         Left            =   -74865
         TabIndex        =   10
         Top             =   2445
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date de naissance *"
         Height          =   255
         Index           =   10
         Left            =   -74865
         TabIndex        =   12
         Top             =   2790
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date d'effet du contrat *"
         Height          =   255
         Index           =   11
         Left            =   -74865
         TabIndex        =   14
         Top             =   3090
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date terme du contrat *"
         Height          =   255
         Index           =   12
         Left            =   -74865
         TabIndex        =   16
         Top             =   3405
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date de l'arret de travail *"
         Height          =   255
         Index           =   13
         Left            =   -74865
         TabIndex        =   18
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Date de reprise"
         Height          =   255
         Index           =   14
         Left            =   -74865
         TabIndex        =   20
         Top             =   4050
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cause"
         Height          =   255
         Index           =   15
         Left            =   -74865
         TabIndex        =   50
         Top             =   4365
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Délai de franchise en jours *"
         Height          =   255
         Index           =   17
         Left            =   -74865
         TabIndex        =   30
         Top             =   1170
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total des prestations réglées"
         Height          =   255
         Index           =   19
         Left            =   -74865
         TabIndex        =   32
         Top             =   1800
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "Montant annualisé pour la période n-1"
         Height          =   255
         Index           =   20
         Left            =   -74865
         TabIndex        =   34
         Top             =   2130
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "Montant annualisé de la prestation "
         Height          =   255
         Index           =   21
         Left            =   -74865
         TabIndex        =   36
         Top             =   2430
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "Situation de l'assuré en fin de période *"
         Height          =   255
         Index           =   22
         Left            =   -74865
         TabIndex        =   38
         Top             =   2760
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Incapacité pour 1F mensuel"
         Height          =   255
         Index           =   23
         Left            =   -74865
         TabIndex        =   40
         Top             =   3090
         Width           =   3165
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Passage pour 1F annuelle"
         Height          =   255
         Index           =   24
         Left            =   -74865
         TabIndex        =   42
         Top             =   3405
         Width           =   3030
      End
      Begin VB.Label lblLabels 
         Caption         =   "PM Invalidité  pour 1F annuelle"
         Height          =   255
         Index           =   25
         Left            =   -74865
         TabIndex        =   44
         Top             =   3720
         Width           =   3210
      End
      Begin VB.Label lblLabels 
         Caption         =   "Montant de la PM"
         Height          =   255
         Index           =   26
         Left            =   -74865
         TabIndex        =   46
         Top             =   4680
         Width           =   2940
      End
      Begin VB.Label lblLabels 
         Caption         =   "Variation de la PM"
         Height          =   255
         Index           =   27
         Left            =   -74865
         TabIndex        =   48
         Top             =   4995
         Width           =   2940
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   5730
      TabIndex        =   0
      Top             =   6270
      Width           =   5730
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Fermer"
         Height          =   345
         Left            =   4505
         TabIndex        =   28
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Enregistrer"
         Height          =   345
         Left            =   3420
         TabIndex        =   27
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Rafraichir"
         Height          =   345
         Left            =   2313
         TabIndex        =   26
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Supprimer"
         Height          =   345
         Left            =   1217
         TabIndex        =   25
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Nouveau"
         Height          =   345
         Left            =   121
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Majoration par enfants à charge"
      Height          =   255
      Index           =   55
      Left            =   270
      TabIndex        =   103
      Top             =   4275
      Width           =   3030
   End
   Begin VB.Label lblLabels 
      Caption         =   "Age moyen des enfants"
      Height          =   255
      Index           =   54
      Left            =   270
      TabIndex        =   102
      Top             =   3960
      Width           =   2940
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "(Remplir au moins les champs avec *)"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   45
      TabIndex        =   66
      Top             =   5940
      Width           =   5640
   End
   Begin VB.Label lblGroupe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Période ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      TabIndex        =   57
      Top             =   0
      Width           =   5640
   End
End
Attribute VB_Name = "frmAssure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67DF030F"
Option Explicit

'##ModelId=5C8A67E00012
Private fmAction As Integer

'##ModelId=5C8A67E00040
Private Sub cmdAdd_Click()
  fmAction = 1
  datPrimaryRS.Recordset.AddNew
  
  'DataGarantie.RecordSource = "SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle
  'DataGarantie.Refresh
  DataGarantie.RecordSource = ""
  DataGarantie.Refresh
  
  ' renseigne le n° de groupe et de période
  datPrimaryRS.Recordset.fields("POGPECLE") = GroupeCle
  datPrimaryRS.Recordset.fields("POPERCLE") = numPeriode
  
  ' mets les valeurs par défaut
  datPrimaryRS.Recordset.fields("POSEXE") = 1
  txtFields(8) = "1"
  datPrimaryRS.Recordset.fields("POTYPEF") = 0
  datPrimaryRS.Recordset.fields("PODELAI") = 0
  datPrimaryRS.Recordset.fields("POPRESTATION") = 0
  datPrimaryRS.Recordset.fields("POPRESTATION_AN_PREC") = 0
  datPrimaryRS.Recordset.fields("POPRESTATION_AN") = 0
  datPrimaryRS.Recordset.fields("POPRESTATION_AN_PASSAGE") = 0
  datPrimaryRS.Recordset.fields("POSIT") = cdPosit_IncapAvecPassage
  txtFields(22) = "1"
  datPrimaryRS.Recordset.fields("POPM_INCAP_1F") = 0
  datPrimaryRS.Recordset.fields("POPM_PASS_1F") = 0
  datPrimaryRS.Recordset.fields("POPM_INVAL_1F") = 0
  datPrimaryRS.Recordset.fields("POPM") = 0
  datPrimaryRS.Recordset.fields("POPM_INCAP_1R") = 0
  datPrimaryRS.Recordset.fields("POPM_PASS_1R") = 0
  datPrimaryRS.Recordset.fields("POPM_INVAL_1R") = 0
  datPrimaryRS.Recordset.fields("POPM_VAR") = 0
  datPrimaryRS.Recordset.fields("POPM_RI") = 0
  datPrimaryRS.Recordset.fields("POPM_REVALO") = 0
  datPrimaryRS.Recordset.fields("POCOT_REVALO") = 0
  datPrimaryRS.Recordset.fields("POPM_RASSUR") = 0
  datPrimaryRS.Recordset.fields("POPSAP_RASSUR") = 0
  
  datPrimaryRS.Recordset.fields("POIsCadre") = 0
  
  DBCombo2.Refresh
End Sub

'##ModelId=5C8A67E00050
Private Sub cmdDelete_Click()
  If datPrimaryRS.Recordset.EOF Then Exit Sub
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If Not .EOF Then .MoveLast
  End With
  
  fmAction = 0
End Sub

'##ModelId=5C8A67E00060
Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  datPrimaryRS.Refresh
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' met le bon separateur de decimal dans les textbox (',' ou '.')
'
'##ModelId=5C8A67E0006F
Private Sub RemplaceSeparateurDecimal()
  Dim i As Integer
  
  ' convertie les ',' en '.' ou inversement
  For i = 20 To 27
    If i <> 22 Then
      txtFields(i) = m_dataHelper.GetDouble2(txtFields(i))
    End If
  Next i
End Sub

'##ModelId=5C8A67E0008F
Private Sub cmdUpdate_Click()
  If Not (txtFields(8) = "1" Or txtFields(8) = "2") Then
    MsgBox "Sexe de l'assuré:" & vbLf & "Vous ne devez entrer que '1' ou '2'", vbExclamation
    SSTab1.Tab = 0
    txtFields(8).SetFocus
    Exit Sub
  End If
  
  If (txtFields(22) < 1 Or txtFields(22) > 6) Then
    MsgBox "Situation de l'assuré:" & vbLf & "Vous devez entrer 1,2,3,4,5 ou 6", vbExclamation
    SSTab1.Tab = 1
    txtFields(22).SetFocus
    Exit Sub
  End If
  
  Call RemplaceSeparateurDecimal
  
  ' sauvegarde des données
  On Error GoTo err_cmdUpdate_Click
  Dim vMark As Variant
  vMark = datPrimaryRS.Recordset.bookmark
  datPrimaryRS.Recordset.Update
  datPrimaryRS.Recordset.bookmark = vMark
  On Error GoTo 0
  
  fmAction = 0

  Exit Sub
  
err_cmdUpdate_Click:
  On Error GoTo 0
  
  MsgBox "Erreur : vérifiez que tous les champs soit bien rempli !", vbExclamation
End Sub

'##ModelId=5C8A67E000AE
Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  
  If datPrimaryRS.Recordset.EditMode <> adEditNone Then
    datPrimaryRS.Recordset.CancelUpdate
  End If
  
  Unload Me
End Sub

'##ModelId=5C8A67E000BD
Private Sub Command1_Click()
  MsgBox "Valeurs possibles : " & vbLf _
         & "1 : Incapacité avec passage" & vbLf _
         & "2 : Invalidité" & vbLf _
         & "3 : Incapacité sans passage" & vbLf _
         & "4 : Rente Conjoint" & vbLf _
         & "5 : Rente Education" & vbLf _
         & "6 : Décés", vbInformation
End Sub

'##ModelId=5C8A67E000CE
Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Errreur : " & ErrorNumber & vbLf & Description, vbExclamation
  'Response = 0  'Throw away the error
End Sub

'##ModelId=5C8A67E0015A
Private Sub datPrimaryRS_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'This will display the current record position for dynasets and snapshots
  If datPrimaryRS.Recordset.AbsolutePosition <> -1 Then
    datPrimaryRS.Caption = "Record: " & (datPrimaryRS.Recordset.AbsolutePosition + 1) & " / " & datPrimaryRS.Recordset.RecordCount
  Else
    datPrimaryRS.Caption = "La période est vide"
  End If
  
  If Not IsNull(datPrimaryRS.Recordset.fields("POSTECLE")) Then
    DataGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle & " AND GASTECLE=" & datPrimaryRS.Recordset.fields("POSTECLE"))
  Else
    DataGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle)
  End If
  DataGarantie.Refresh
End Sub

'##ModelId=5C8A67E0016A
Private Sub datPrimaryRS_Validate(Action As Integer, Save As Integer)
  Call RemplaceSeparateurDecimal
  
  'This is where you put validation code
  'This event gets called when the following actions occur
  Select Case Action
    Case vbDataActionMoveFirst, vbDataActionMovePrevious, vbDataActionMoveNext, _
         vbDataActionMoveLast, vbDataActionAddNew, vbDataActionDelete, _
         vbDataActionFind, vbDataActionBookmark, vbDataActionClose
      Save = False
      Screen.MousePointer = vbDefault
    
    Case vbDataActionUpdate
      Save = True
  End Select
  
  DataGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle)
  DataGarantie.Refresh
  
  Screen.MousePointer = vbHourglass
End Sub

'##ModelId=5C8A67E001A8
Private Sub DBCombo1_Change()
  'If fmAction = 1 And DBCombo1.BoundText <> "" Then
  If DBCombo1.BoundText <> "" Then
    DataGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle & " AND GASTECLE = " & DBCombo1.BoundText)
    DataGarantie.Refresh
  End If
End Sub

'##ModelId=5C8A67E001C7
Private Sub Form_Load()
  m_dataSource.SetDatabase DataGarantie
  DataGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Garantie WHERE GAGROUPCLE=" & GroupeCle)
  DataGarantie.Refresh
  
  m_dataSource.SetDatabase DataSociete
  DataSociete.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Societe WHERE SOGROUPE=" & GroupeCle)
  DataSociete.Refresh
  
  m_dataSource.SetDatabase dtaSituationFamille
  dtaSituationFamille.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM SituationFamille ORDER BY Libelle")
  dtaSituationFamille.Refresh
  
  m_dataSource.SetDatabase datPrimaryRS
  datPrimaryRS.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY PONUMCLE")
  datPrimaryRS.Refresh

  lblGroupe = DescriptionPeriode
  
  ' affiche la premiere page
  SSTab1.Tab = 0

  fmAction = 0
  
  datPrimaryRS.Recordset.Find "RECNO=" & RECNO
End Sub

'##ModelId=5C8A67E001E7
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

