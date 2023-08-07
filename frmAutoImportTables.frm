VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAutoImportTables 
   Caption         =   "Import des tables de paramétrages"
   ClientHeight    =   11445
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDe15 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   69
      Top             =   9720
      Width           =   330
   End
   Begin VB.CommandButton cmdFile15 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   67
      Top             =   9720
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe14 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   66
      Top             =   9240
      Width           =   330
   End
   Begin VB.CommandButton cmdFile14 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   64
      Top             =   9240
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe13 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   63
      Top             =   8760
      Width           =   330
   End
   Begin VB.CommandButton cmdFile13 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   61
      Top             =   8760
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe12 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   60
      Top             =   8280
      Width           =   330
   End
   Begin VB.CommandButton cmdFile12 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   58
      Top             =   8280
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe11 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   57
      Top             =   7800
      Width           =   330
   End
   Begin VB.CommandButton cmdFile11 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   55
      Top             =   7800
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe10 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   54
      Top             =   7320
      Width           =   330
   End
   Begin VB.CommandButton cmdFile10 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   52
      Top             =   7320
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe9 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13800
      TabIndex        =   51
      Top             =   6840
      Width           =   330
   End
   Begin VB.CommandButton cmdFile9 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   12240
      TabIndex        =   49
      Top             =   6840
      Width           =   1425
   End
   Begin VB.CommandButton cmdDe2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   35
      Top             =   7320
      Width           =   330
   End
   Begin VB.CommandButton cmdDe3 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   34
      Top             =   7800
      Width           =   330
   End
   Begin VB.CommandButton cmdDe4 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   33
      Top             =   8280
      Width           =   330
   End
   Begin VB.CommandButton cmdDe5 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   32
      Top             =   8760
      Width           =   330
   End
   Begin VB.CommandButton cmdDe6 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   31
      Top             =   9240
      Width           =   330
   End
   Begin VB.CommandButton cmdDe7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   30
      Top             =   9720
      Width           =   330
   End
   Begin VB.CommandButton cmdDe8 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   29
      Top             =   10200
      Width           =   330
   End
   Begin VB.CommandButton cmdFile8 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   28
      Top             =   10200
      Width           =   1425
   End
   Begin VB.CommandButton cmdFile3 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   27
      Top             =   7800
      Width           =   1425
   End
   Begin VB.CommandButton cmdFile4 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   26
      Top             =   8280
      Width           =   1425
   End
   Begin VB.CommandButton cmdFile5 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   25
      Top             =   8760
      Width           =   1425
   End
   Begin VB.CommandButton cmdFile6 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   24
      Top             =   9240
      Width           =   1425
   End
   Begin VB.CommandButton cmdFile7 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   23
      Top             =   9720
      Width           =   1425
   End
   Begin VB.CommandButton cmdDel1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      TabIndex        =   15
      Top             =   6840
      Width           =   330
   End
   Begin VB.CommandButton cmdFile1 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   13
      Top             =   6840
      Width           =   1425
   End
   Begin VB.CommandButton cmdLaunchImport 
      Caption         =   "Importer dans toutes les périodes sélectionnées"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   10920
      Width           =   5145
   End
   Begin VB.CommandButton cmdFile2 
      Caption         =   "Sélectionner"
      Height          =   330
      Left            =   2160
      TabIndex        =   9
      Top             =   7320
      Width           =   1425
   End
   Begin VB.ComboBox cmb2 
      Height          =   315
      ItemData        =   "frmAutoImportTables.frx":0000
      Left            =   12720
      List            =   "frmAutoImportTables.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   17280
      TabIndex        =   6
      Top             =   10920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      ItemData        =   "frmAutoImportTables.frx":0004
      Left            =   11760
      List            =   "frmAutoImportTables.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   345
      Left            =   18840
      TabIndex        =   0
      Top             =   10920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   5880
      Top             =   120
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
   Begin FPSpreadADO.fpSpread sprListe 
      Height          =   5685
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   19605
      _Version        =   524288
      _ExtentX        =   34581
      _ExtentY        =   10028
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
      SpreadDesigner  =   "frmAutoImportTables.frx":0008
      ScrollBarTrack  =   3
      AppearanceStyle =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   330
      Left            =   18360
      TabIndex        =   4
      Top             =   5640
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18240
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSComctlLib.ProgressBar progAuto 
      Height          =   330
      Left            =   14640
      TabIndex        =   5
      Top             =   10920
      Visible         =   0   'False
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblFile15 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   68
      Top             =   9765
      Width           =   5655
   End
   Begin VB.Label lblFile14 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   65
      Top             =   9285
      Width           =   5655
   End
   Begin VB.Label lblFile13 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   62
      Top             =   8805
      Width           =   5655
   End
   Begin VB.Label lblFile12 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   59
      Top             =   8325
      Width           =   5655
   End
   Begin VB.Label lblFile11 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   56
      Top             =   7845
      Width           =   5655
   End
   Begin VB.Label lblFile10 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   53
      Top             =   7365
      Width           =   5655
   End
   Begin VB.Label lblFile9 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   14280
      TabIndex        =   50
      Top             =   6885
      Width           =   5655
   End
   Begin VB.Label lblTable9 
      Caption         =   "Correspondance_CatOption"
      Height          =   255
      Left            =   9960
      TabIndex        =   48
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label lblTable10 
      Caption         =   "DonneesSociales"
      Height          =   255
      Left            =   9960
      TabIndex        =   47
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblTable11 
      Caption         =   "PassageNCA"
      Height          =   255
      Left            =   9960
      TabIndex        =   46
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblTable12 
      Caption         =   "ParamRentes"
      Height          =   255
      Left            =   9960
      TabIndex        =   45
      Top             =   8280
      Width           =   2055
   End
   Begin VB.Label lblTable13 
      Caption         =   "PM_Retenue"
      Height          =   255
      Left            =   9960
      TabIndex        =   44
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label lblTable14 
      Caption         =   "Reassurance"
      Height          =   255
      Left            =   9960
      TabIndex        =   43
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Label lblTable15 
      Caption         =   "TBQREGA"
      Height          =   255
      Left            =   9960
      TabIndex        =   42
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label lblFile3 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   41
      Top             =   7800
      Width           =   5295
   End
   Begin VB.Label lblFile4 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   40
      Top             =   8280
      Width           =   5295
   End
   Begin VB.Label lblFile5 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   39
      Top             =   8760
      Width           =   5295
   End
   Begin VB.Label lblFile6 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   38
      Top             =   9240
      Width           =   5295
   End
   Begin VB.Label lblFile7 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   37
      Top             =   9720
      Width           =   5295
   End
   Begin VB.Label lblFile8 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   10200
      Width           =   5295
   End
   Begin VB.Label lblTable8 
      Caption         =   "CoeffAmortissement"
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   10200
      Width           =   1455
   End
   Begin VB.Label lblTable2 
      Caption         =   "Capitaux_Moyens"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label lblTable3 
      Caption         =   "CATR9"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label lblTable4 
      Caption         =   "CATR9INVAL"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label lblTable5 
      Caption         =   "CDSITUAT"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label lblTable6 
      Caption         =   "CodesCat"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   9240
      Width           =   1455
   End
   Begin VB.Label lblTable7 
      Caption         =   "CodeCatInv"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   9720
      Width           =   1455
   End
   Begin VB.Label lblTable1 
      Caption         =   "AgeDepartRetraite"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblFile1 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   6885
      Width           =   5295
   End
   Begin VB.Label lblFile2 
      Caption         =   "Fichier sélectionné :"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   7320
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre des périodes déjà traité : 3/10  - période 329 en cours de traitement"
      Height          =   300
      Left            =   8760
      TabIndex        =   7
      Top             =   10920
      Visible         =   0   'False
      Width           =   5655
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmAutoImportTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6844030F"
Option Explicit

'##ModelId=5C8A68450041
Private fmAction As Integer
'##ModelId=5C8A68450062
Private fmTableDiverse As clsTablesDiverses
'##ModelId=5C8A684500DD
Private stopProcess As Boolean

'##ModelId=5C8A6845010C
Private labelFileText As String

'##ModelId=5C8A6845011B
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'************ DELETE ACTIONS *****************

'##ModelId=5C8A6845014A
Private Sub cmdDel1_Click()
  lblFile1.Caption = labelFileText
End Sub

'##ModelId=5C8A6845015A
Private Sub cmdDe2_Click()
  lblFile2.Caption = labelFileText
End Sub

'##ModelId=5C8A6845016A
Private Sub cmdDe3_Click()
  lblFile3.Caption = labelFileText
End Sub

'##ModelId=5C8A68450179
Private Sub cmdDe4_Click()
  lblFile4.Caption = labelFileText
End Sub

'##ModelId=5C8A68450198
Private Sub cmdDe5_Click()
  lblFile5.Caption = labelFileText
End Sub

'##ModelId=5C8A684501A8
Private Sub cmdDe6_Click()
  lblFile6.Caption = labelFileText
End Sub

'##ModelId=5C8A684501C7
Private Sub cmdDe7_Click()
  lblFile7.Caption = labelFileText
End Sub

'##ModelId=5C8A684501D7
Private Sub cmdDe8_Click()
  lblFile8.Caption = labelFileText
End Sub

'##ModelId=5C8A684501E6
Private Sub cmdDe9_Click()
  lblFile9.Caption = labelFileText
End Sub

'##ModelId=5C8A68450206
Private Sub cmdDe10_Click()
  lblFile10.Caption = labelFileText
End Sub

'##ModelId=5C8A68450215
Private Sub cmdDe11_Click()
  lblFile11.Caption = labelFileText
End Sub

'##ModelId=5C8A68450235
Private Sub cmdDe12_Click()
  lblFile12.Caption = labelFileText
End Sub

'##ModelId=5C8A68450254
Private Sub cmdDe13_Click()
  lblFile13.Caption = labelFileText
End Sub

'##ModelId=5C8A68450264
Private Sub cmdDe14_Click()
  lblFile14.Caption = labelFileText
End Sub

'##ModelId=5C8A68450273
Private Sub cmdDe15_Click()
  lblFile15.Caption = labelFileText
End Sub


'************ SELECT FILE ACTIONS *****************

'##ModelId=5C8A68450292
Private Sub cmdFile1_Click()
  GetFileName lblFile1
End Sub

'##ModelId=5C8A684502A2
Private Sub cmdFile2_Click()
  GetFileName lblFile2
End Sub

'##ModelId=5C8A684502B2
Private Sub cmdFile3_Click()
  GetFileName lblFile3
End Sub

'##ModelId=5C8A684502E0
Private Sub cmdFile4_Click()
  GetFileName lblFile4
End Sub

'##ModelId=5C8A684502F0
Private Sub cmdFile5_Click()
  GetFileName lblFile5
End Sub

'##ModelId=5C8A68450300
Private Sub cmdFile6_Click()
  GetFileName lblFile6
End Sub

'##ModelId=5C8A6845031F
Private Sub cmdFile7_Click()
  GetFileName lblFile7
End Sub

'##ModelId=5C8A6845032F
Private Sub cmdFile8_Click()
  GetFileName lblFile8
End Sub

'##ModelId=5C8A6845034E
Private Sub cmdFile9_Click()
  GetFileName lblFile9
End Sub

'##ModelId=5C8A6845035E
Private Sub cmdFile10_Click()
  GetFileName lblFile10
End Sub

'##ModelId=5C8A6845036D
Private Sub cmdFile11_Click()
  GetFileName lblFile11
End Sub

'##ModelId=5C8A6845038C
Private Sub cmdFile12_Click()
  GetFileName lblFile12
End Sub

'##ModelId=5C8A6845039C
Private Sub cmdFile13_Click()
  GetFileName lblFile13
End Sub

'##ModelId=5C8A684503AC
Private Sub cmdFile14_Click()
  GetFileName lblFile14
End Sub

'##ModelId=5C8A684503CB
Private Sub cmdFile15_Click()
  GetFileName lblFile15
End Sub

'##ModelId=5C8A684503DA
Private Sub GetFileName(lbl As Label)

  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Import de la table"  ' '" & nomTable & "'"
  CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
    
  If CommonDialog1.filename = "" _
    Or CommonDialog1.filename = "*.xls" _
    Or CommonDialog1.filename = "*.xlsx" _
    Or CommonDialog1.filename = "*.*" Then
    
    lbl.Caption = labelFileText
    
    MsgBox "Vous devez sélectionner un fichier d'import valable !", vbExclamation
    Exit Sub
  Else
    lbl.Caption = CommonDialog1.filename
  End If
  
End Sub

'##ModelId=5C8A68460022
Private Function ValidateExcelFiles(filename As String, nomTable As String) As Boolean

Dim xlSource As DataAccess
Dim connString As String
Dim rsExcel As ADODB.Recordset
Dim rsTable As ADODB.Recordset
Dim idField As Integer

  On Error GoTo ValidationError
  
  connString = Replace(ConnectionStringXls, "%1", filename)
  If UCase(Right(filename, 5)) = ".XLSX" Then
    connString = Replace(ConnectionStringXlsx, "%1", filename)
  End If
  
  Set xlSource = New DataAccess
  xlSource.Connect connString
  
  Set rsTable = m_dataSource.OpenRecordset("" & nomTable, table)
  idField = IIf(nomTable = "ParamRentes", 4, 2)
  
  Set rsExcel = New ADODB.Recordset
  rsExcel.Open "SELECT * FROM " & nomTable & " WHERE " & rsTable.fields(idField).Name & " IS NOT NULL", xlSource.Connection, adOpenStatic, adLockOptimistic
  
  ValidateExcelFiles = True
  
  xlSource.Disconnect
  Set xlSource = Nothing
  Set rsTable = Nothing
  Set rsExcel = Nothing
  
  Exit Function
  
ValidationError:

  xlSource.Disconnect
  Set xlSource = Nothing
  Set rsTable = Nothing
  Set rsExcel = Nothing

  ValidateExcelFiles = False
  MsgBox "Le fichier d'import : " & filename & " ne correspond pas à la table : " & nomTable & " sur laquelle vous voulez importer des données !", vbExclamation

End Function

'##ModelId=5C8A6846006F
Private Sub cmdLaunchImport_Click()

  Dim nomTable As String
  Dim currentPeriode As Long
  
  If DroitAdmin = False Then Exit Sub
      
  Dim i As Integer
  Dim numbItemsChecked As Integer
  Dim checkboxSelected As Boolean
  Dim numPeriode As Integer
  Dim colPeriods As New Collection
  Dim automationSuccess As Boolean
  Dim statusMessage As String
  
  automationSuccess = False
  numbItemsChecked = 0
  
  sprListe.VirtualMode = False
  sprListe.DataRefresh
  sprListe.Refresh
   
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 9
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    numPeriode = CInt(sprListe.text)
    
    If checkboxSelected Then
      numbItemsChecked = numbItemsChecked + 1
      colPeriods.Add (numPeriode)
    End If
  Next i
  
  If numbItemsChecked = 0 Then
     MsgBox "Aucune période a été sélectionnée ! Sélectionnez au moins une période et cliquez le bouton 'Importer dans toutes les périodes sélectionnées'.", vbExclamation
     GoTo Cleanup
  End If
  
  'validate the matching of import files and tables
  If lblFile1.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile1.Caption, lblTable1.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile2.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile2.Caption, lblTable2.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile3.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile3.Caption, lblTable3.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile4.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile4.Caption, lblTable4.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile5.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile5.Caption, lblTable5.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile6.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile6.Caption, lblTable6.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile7.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile7.Caption, lblTable7.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile8.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile8.Caption, lblTable8.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile9.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile9.Caption, lblTable9.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile10.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile10.Caption, lblTable10.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile11.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile11.Caption, lblTable11.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile12.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile12.Caption, lblTable12.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile13.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile13.Caption, lblTable13.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile14.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile14.Caption, lblTable14.Caption) = False Then
      Exit Sub
    End If
  End If
  If lblFile15.Caption <> labelFileText Then
    If ValidateExcelFiles(lblFile15.Caption, lblTable15.Caption) = False Then
      Exit Sub
    End If
  End If
  
  
  
  If MsgBox("Est-ce que vous est sur de vouloir lancer la procédure de l'import pour les périodes sélectionnées ?", vbYesNo) = vbYes Then
      
    'we start the automation procedure
    Screen.MousePointer = vbHourglass
       
    'cmdImport.Enabled = False
    btnStop.Visible = True
    lblStatus.Visible = True
    progAuto.Visible = True
           
    progAuto.Min = 0
    progAuto.Max = colPeriods.Count + 1
    progAuto.Value = progAuto.Min + 1
    'lblStatus.Caption = "Procédure en cours..."
    
    'we create the log file
    Dim m_Logger As New clsLogger
    m_Logger.FichierLog = m_logPathAuto & "\" & GetWinUser & "_ErreurImportTables.log"
    'm_Logger.CreateLog "Import " & CommonDialog1.filename & " dans la table '" & nomTable & "'"
    m_Logger.CreateLog ""
    
    For i = 1 To colPeriods.Count
     DoEvents
     currentPeriode = colPeriods(i)
     
     m_Logger.EcritTraceDansLog "Import dans la période numéro : " & colPeriods(i)
     m_Logger.EcritTraceDansLog ""
       
     lblStatus.Caption = "Nombre des périodes déjà traité : " & i - 1 & "/" & colPeriods.Count & " - période " & colPeriods(i) & " en cours de traitement."
           
     '### add log for table that will be imported
     
     'verify if tables have been selected and if we have a valid path on the label
     
     If lblFile1.Caption <> labelFileText Then 'And cmb1.ListIndex <> -1
      'nomTable = lblTable1.Caption ' cmb1.List(cmb1.ListIndex)
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable1.Caption & " depuis le fichier : " & lblFile1.Caption
      ImportGeneriqueAuto lblFile1.Caption, lblTable1.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile2.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable2.Caption & " depuis le fichier : " & lblFile2.Caption
      ImportGeneriqueAuto lblFile2.Caption, lblTable2.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile3.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable3.Caption & " depuis le fichier : " & lblFile3.Caption
      ImportGeneriqueAuto lblFile3.Caption, lblTable3.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile4.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable4.Caption & " depuis le fichier : " & lblFile4.Caption
      ImportGeneriqueAuto lblFile4.Caption, lblTable4.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile5.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable5.Caption & " depuis le fichier : " & lblFile5.Caption
      ImportGeneriqueAuto lblFile5.Caption, lblTable5.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile6.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable6.Caption & " depuis le fichier : " & lblFile6.Caption
      ImportGeneriqueAuto lblFile6.Caption, lblTable6.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile7.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable7.Caption & " depuis le fichier : " & lblFile7.Caption
      ImportGeneriqueAuto lblFile7.Caption, lblTable7.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile8.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable8.Caption & " depuis le fichier : " & lblFile8.Caption
      ImportGeneriqueAuto lblFile8.Caption, lblTable8.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     
     'second column
     If lblFile9.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable9.Caption & " depuis le fichier : " & lblFile9.Caption
      ImportGeneriqueAuto lblFile9.Caption, lblTable9.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile10.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable10.Caption & " depuis le fichier : " & lblFile10.Caption
      ImportGeneriqueAuto lblFile10.Caption, lblTable10.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile11.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable11.Caption & " depuis le fichier : " & lblFile11.Caption
      ImportGeneriqueAuto lblFile11.Caption, lblTable11.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile12.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable12.Caption & " depuis le fichier : " & lblFile12.Caption
      ImportGeneriqueAuto lblFile12.Caption, lblTable12.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile13.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable13.Caption & " depuis le fichier : " & lblFile13.Caption
      ImportGeneriqueAuto lblFile13.Caption, lblTable13.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile14.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable14.Caption & " depuis le fichier : " & lblFile14.Caption
      ImportGeneriqueAuto lblFile14.Caption, lblTable14.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     If lblFile15.Caption <> labelFileText Then
      m_Logger.EcritTraceDansLog "Import dans la table : " & lblTable15.Caption & " depuis le fichier : " & lblFile15.Caption
      ImportGeneriqueAuto lblFile15.Caption, lblTable15.Caption, currentPeriode, m_Logger
      m_Logger.EcritTraceDansLog ""
     End If
     
     
     m_Logger.EcritTraceDansLog ""
     m_Logger.EcritTraceDansLog "******************************************************************************"
     m_Logger.EcritTraceDansLog ""
     
     lblStatus.Caption = "Nombre des périodes déjà traité : " & i & "/" & colPeriods.Count & " - période " & colPeriods(i) & " en cours de traitement."
     
     If stopProcess Then
      If MsgBox("Est-ce que vous est sur de vouloir arrêter l'import ?", vbYesNo) = vbYes Then
          stopProcess = False
          GoTo Cleanup
      Else
          stopProcess = False
      End If
     End If
     
     progAuto.Value = i + 1
           
    Next i
  
  End If
  
  Screen.MousePointer = vbDefault
  btnStop.Visible = False
    
  If MsgBox("L'import est terminé. Voulez-vous consultez le fichier log ?", vbInformation + vbYesNo) = vbNo Then
    GoTo Cleanup
  End If
  
  Dim frm As New frmDisplayLog
  frm.FichierLog = m_logPathAuto & "\" & GetWinUser & "_ErreurImportTables.log"
  frm.Show vbModal
  Set frm = Nothing
  
Cleanup:
  
  'cmdImport.Enabled = True
  lblStatus.Visible = False
  progAuto.Visible = False
  progAuto.Value = progAuto.Min
  lblStatus.Caption = ""
  Set m_Logger = Nothing
 
  Screen.MousePointer = vbDefault
  
End Sub

'##ModelId=5C8A6846007F
Private Sub Form_Load()

  Screen.MousePointer = vbHourglass
  ProgressBar1.Visible = False
  
  labelFileText = "Fichier sélectionné :"
  
  FillGrid

  ' liste des tables diverses
  InitTableDiverse
  fmTableDiverse.FillCombo cmb1
  fmTableDiverse.FillCombo cmb2
    
  Screen.MousePointer = vbDefault

End Sub

'##ModelId=5C8A6846008F
Private Sub btnStop_Click()
  stopProcess = True
End Sub

'##ModelId=5C8A684600AE
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

'##ModelId=5C8A684600CD
Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub


'##ModelId=5C8A684600DD
Private Sub Form_Unload(Cancel As Integer)
  Set fmTableDiverse = Nothing
  Screen.MousePointer = vbDefault
End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'  If SSTab1.Tab = 1 Then
'    fmTableDiverse.FillCombo cmb1
'  End If
'End Sub

'##ModelId=5C8A6846010C
Private Sub cmb1_Click()
'  Dim rq As String
'
'  If cmb1.ListIndex <> -1 Then
'    'sprCATR9.ReDraw = False
'
'    m_dataSource.SetDatabase dtaCATR9
'
'    On Error Resume Next
'
'    Dim defTable As defTableDiverse
'
'    Set defTable = fmTableDiverse.TableInfo(cmb1.ListIndex)
'    If defTable.champs = "" Then
'      defTable.champs = "*"
'    End If
'
'    rq = "SELECT " & defTable.champs & " FROM " & defTable.nomTable & " WHERE NumPeriode=" & NumPeriode & " And GroupeCle=" & GroupeCle & " ORDER BY " & defTable.orderBy
'
'    dtaCATR9.RecordSource = m_dataHelper.ValidateSQL(rq)
'    dtaCATR9.Refresh
'
'    Set sprCATR9.DataSource = dtaCATR9
'
'    If Not dtaCATR9.Recordset.EOF Then
'      dtaCATR9.Recordset.MoveLast
'      dtaCATR9.Recordset.MoveFirst
'
'      sprCATR9.Refresh
'
'      sprCATR9.MaxRows = dtaCATR9.Recordset.RecordCount
'
'      Dim i As Integer
'
'      For i = 2 To sprCATR9.MaxCols
'        sprCATR9.Col = i
'        sprCATR9.DataColCnt = True
'      Next
'
'      dtaCATR9.Refresh
'    Else
'      sprCATR9.MaxRows = 0
'    End If
'
'    sprCATR9.Refresh
'
'    ' largeur des colonnes
'    LargeurMaxColonneSpread sprCATR9
'
'    sprCATR9.ReDraw = True
'
'    On Error GoTo 0
'  End If
End Sub

'Private Sub sprCATR9_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
'
'  If DataField = "CoeffAmortissement" Then
'    sprCATR9.BlockMode = True
'    sprCATR9.Col = Col
'    sprCATR9.Row = -1
'    sprCATR9.Col2 = Col
'    sprCATR9.Row2 = -1
'
'    sprCATR9.TypeNumberDecPlaces = 4
'    sprCATR9.BlockMode = False
'  End If
'
'End Sub



'******************************************************************************************************************************
'********************************************* FILL THE GRID WITH LIST OF PERIODES ********************************************
'******************************************************************************************************************************

'##ModelId=5C8A6846011B
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

'##ModelId=5C8A6846012B
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
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT P.RECNO, P.PENUMCLE as [Numéro Période], " _
        & "CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + TP.Libelle as [Type], " _
        & "CAST(P.IdTypeCalcul as VARCHAR) + ' - ' + TC.Libelle as [Type Calcul], " _
        & "P.PEDATEDEB as [Début], " _
        & "P.PEDATEFIN as [Fin], " _
        & "P.PEDATEEXT as [Date Arrêté], " _
        & "P.PECOMMENTAIRE as Commentaire " _
        & "FROM P3IUser.Periode P LEFT JOIN P3IUser.TypePeriode TP ON TP.IdTypePeriode=P.PETYPEPERIODE " _
        & "LEFT JOIN P3IUser.TypeCalcul TC ON TC.IdTypeCalcul=P.IdTypeCalcul " _
        & "WHERE P.PEGPECLE = " & GroupeCle
        
  rq = rq & " ORDER BY P.PENUMCLE DESC "
  
  ' rafraichie le spread
  sprListe.Visible = False
  'sprListe.Visible = True

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
  SetColonneDataFill 10, True
  
  sprListe.ColWidth(2) = 10
  sprListe.ColWidth(3) = 20
  sprListe.ColWidth(4) = 10
  sprListe.ColWidth(5) = 10
  sprListe.ColWidth(6) = 10
  sprListe.ColWidth(7) = 10
  sprListe.ColWidth(8) = 65
  
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = 7
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = 3
  sprListe.Col2 = 3
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.Col = 4
  sprListe.Col2 = 4
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.Col = 8
  sprListe.Col2 = 8
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
  
  'manually add a column to spread: Selection Checkbox
  'Me.Width = 24000
  sprListe.ActiveCellHighlightStyle = ActiveCellHighlightStyleOff 'switch off rectangle around highlighted cell
  
  On Error Resume Next
     
  sprListe.OperationMode = OperationModeNormal ' OperationModeSingle ' OperationModeNormal
  sprListe.EditMode = True
  sprListe.Enabled = True
  
  'Add a checkbox column
'  sprListe.MaxCols = 8
'  sprListe.Col = 0
'  sprListe.Row = 0
'  sprListe.ColWidth(0) = 8
'  sprListe.text = "Sélection"
'
'  sprListe.Row = -1
'  sprListe.BlockMode = False
'
'  sprListe.CellType = CellTypeCheckBox
'  sprListe.TypeCheckCenter = True
'  sprListe.TypeCheckType = TypeCheckTypeNormal
'  sprListe.text = 0
  
  sprListe.MaxCols = 9
  sprListe.Col = 9
  sprListe.Row = 0
  sprListe.ColWidth(9) = 8
  sprListe.text = "Sélection"
  
  sprListe.Row = -1
  sprListe.BlockMode = False
  
  sprListe.CellType = CellTypeCheckBox
  sprListe.TypeCheckCenter = True
  sprListe.TypeCheckType = TypeCheckTypeNormal
  sprListe.text = 0

    
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True

  Me.SetFocus
  sprListe.SetFocus
  
End Sub

'##ModelId=5C8A6846014A
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  sprListe.sheet = sprListe.ActiveSheet
  sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
End Sub



'##ModelId=5C8A68460179
Private Sub Label4_Click()

End Sub

'##ModelId=5C8A68460198
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

'##ModelId=5C8A68460215
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







