VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAutomatisation 
   Caption         =   "Automatisation"
   ClientHeight    =   10335
   ClientLeft      =   -5850
   ClientTop       =   -1515
   ClientWidth     =   19080
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStat 
      Caption         =   "STAT"
      Height          =   375
      Left            =   18000
      TabIndex        =   39
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   13935
      Left            =   19440
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fermer"
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
      Left            =   17400
      TabIndex        =   35
      Top             =   12720
      Width           =   1695
   End
   Begin VB.CommandButton cmdImportTables 
      Caption         =   "Import des tables de paramétrages"
      Height          =   375
      Left            =   600
      TabIndex        =   27
      Top             =   6360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
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
      Left            =   14520
      TabIndex        =   24
      Top             =   12720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdDeleteTTProvcoll 
      BackColor       =   &H000000FF&
      Caption         =   "Ecraser les données présentes dans les tables TTLOGTRAIT et TTPROVCOLL"
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
      Left            =   11400
      MaskColor       =   &H000000FF&
      TabIndex        =   20
      Top             =   12720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdConfigInfocentre 
      Caption         =   "Configuration des paramètres pour l'export "
      Height          =   375
      Left            =   14400
      TabIndex        =   19
      Top             =   12720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Export vers l'infocentre"
      Height          =   975
      Left            =   360
      TabIndex        =   17
      Top             =   11520
      Width           =   18735
      Begin VB.CheckBox chkSignalisation 
         Caption         =   "Créer le fichier de signalisation à la fin de l'export "
         Height          =   375
         Left            =   10800
         TabIndex        =   23
         Top             =   360
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkDeleteTTProv 
         Caption         =   "Ecraser les données dans TTLOGTRAIT et TTPROVCOLL avant l'export"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   360
         Value           =   1  'Checked
         Width           =   5655
      End
      Begin VB.CommandButton cmdCreateSignalFile 
         BackColor       =   &H000000FF&
         Caption         =   "Créer le fichier de signalisation"
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
         Left            =   15360
         TabIndex        =   21
         Top             =   360
         Width           =   2895
      End
      Begin VB.CheckBox chkExportInfo 
         Caption         =   "Exports des résultats vers l'infocentre"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   3255
      End
   End
   Begin FPSpreadADO.fpSpread sprListe 
      Height          =   6045
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   18165
      _Version        =   524288
      _ExtentX        =   32041
      _ExtentY        =   10663
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
      SpreadDesigner  =   "frmAutomatisation.frx":0000
      ScrollBarTrack  =   3
      AppearanceStyle =   0
   End
   Begin VB.CommandButton cmdLaunch 
      BackColor       =   &H0000FFFF&
      Caption         =   "Lancer la procédure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaskColor       =   &H0000FFFF&
      TabIndex        =   0
      Top             =   12840
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sélectionner les périodes destinataires"
      Height          =   6615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   18735
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   0
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
   Begin VB.Frame Frame3 
      Caption         =   "Import"
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   7200
      Width           =   18735
      Begin VB.CommandButton cmdImportParams 
         Caption         =   "Modifier les paramètres d'import"
         Height          =   375
         Left            =   2520
         TabIndex        =   29
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkImport 
         Caption         =   "Import des données"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton radImportLot 
         Caption         =   "Importer les données à partir d'un lot"
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CommandButton cmdSelectLot 
         Caption         =   "Sélectionner le lot"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelectExcel 
         Caption         =   "Sélectionner le fichier Excel"
         Height          =   375
         Left            =   9120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton radImportExcel 
         Caption         =   "Importer les données à partir d'une fichier Excel"
         Height          =   735
         Left            =   6960
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lblExcelPath 
         Caption         =   "C:/Directory/FichierExcel.xls"
         Height          =   375
         Left            =   11520
         TabIndex        =   6
         Top             =   1020
         Width           =   6615
      End
      Begin VB.Label lblLot 
         Caption         =   "Importer le lot numéro 257"
         Height          =   255
         Left            =   4320
         TabIndex        =   3
         Top             =   1020
         Width           =   2535
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Export"
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   10320
      Width           =   18735
      Begin VB.CheckBox chkAvantEx 
         Caption         =   "Avant réforme retraite"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7680
         TabIndex        =   34
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkApresEx 
         Caption         =   "Après réforme retraite"
         Enabled         =   0   'False
         Height          =   375
         Left            =   10200
         TabIndex        =   33
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkEcart 
         Caption         =   "Ecart"
         Height          =   375
         Left            =   12840
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkDejaAm 
         Caption         =   "Déjà amorti"
         Height          =   375
         Left            =   14400
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkResteAm 
         Caption         =   "Restant à amortir"
         Height          =   375
         Left            =   16320
         TabIndex        =   30
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox chkExport 
         Caption         =   "Exports des résultats dans Excel"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Calculs"
      Height          =   975
      Left            =   360
      TabIndex        =   7
      Top             =   9120
      Width           =   18735
      Begin VB.CheckBox chkApres 
         Caption         =   "Après reforme retraite"
         Height          =   375
         Left            =   10200
         TabIndex        =   38
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkAvant 
         Caption         =   "Avant reforme retraite"
         Height          =   375
         Left            =   7680
         TabIndex        =   37
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkCalcul 
         Caption         =   "Lancement des calculs"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton radCalcul1 
         Caption         =   "Calculer"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton radCalcul2 
         Caption         =   "Calculer && Revaloriser"
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComctlLib.ProgressBar progAuto 
      Height          =   300
      Left            =   360
      TabIndex        =   26
      Top             =   13440
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin VB.Frame Frame6 
      Caption         =   "Import des tables"
      Height          =   975
      Left            =   360
      TabIndex        =   28
      Top             =   6000
      Visible         =   0   'False
      Width           =   18735
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre des périodes déjà traité : 3/10  - période 329 en cours de traitement"
      Height          =   300
      Left            =   6360
      TabIndex        =   25
      Top             =   13440
      Visible         =   0   'False
      Width           =   12735
   End
End
Attribute VB_Name = "frmAutomatisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A683F004F"
Option Explicit
Option Base 0

'##ModelId=5C8A683F0168
Dim oldPos As Integer

'##ModelId=5C8A683F0197
Private m_Logger As New clsLogger
'##ModelId=5C8A683F0198
Private currentLogFile As String
'##ModelId=5C8A683F01B6
Private periodCount As Long
'##ModelId=5C8A683F01D5
Private currentIndex As Long

'action types
'##ModelId=5C8A683F01F5
Private actImportLot As Boolean
'##ModelId=5C8A683F0214
Private actImportExcel As Boolean
'##ModelId=5C8A683F0224
Private actCalcul As Boolean
'##ModelId=5C8A683F0243
Private actExportExcel As Boolean
'##ModelId=5C8A683F0262
Private actExportInfocentre As Boolean

'configuration parameters for Import
'##ModelId=5C8A683F0281
Private selectedLot As Long
'##ModelId=5C8A683F02A1
Private selectedExcelFile As String
'##ModelId=5C8A683F02D0
Private dateArrete As Date
'##ModelId=5C8A683F02F1
Private typeImport As eTypeImport
'##ModelId=5C8A683F02FE
Private typeDelaiInactivite As eTypeDelaiInactivite
'##ModelId=5C8A683F0310
Private typeCalculAnnualisation As eTypeCalculAnnualisation

'##ModelId=5C8A683F0311
Private typeCalculCurrentPeriode As Integer
'##ModelId=5C8A683F032D
Private lotCurrentPeriode As Integer
'##ModelId=5C8A683F035C
Private isFirstPeriode As Boolean
'##ModelId=5C8A683F036C
Private isLastPeriode As Boolean

'configuration parameters for Infocentre export
'##ModelId=5C8A683F039B
Private exportInfoTypeProvision As Integer
'##ModelId=5C8A683F03BA
Private exportInfoDelExistant As Boolean
'##ModelId=5C8A683F03D9
Private exportInfoCreateSignalisation As Boolean

'##ModelId=5C8A68400010
Private stopProcess As Boolean

'##ModelId=5C8A6840006E
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'NOT REQUIRED - MAYBE LATER
'##ModelId=5C8A68400030
Private PeriodeType As String ' not really needed - only for periods of type statutaire



'##ModelId=5C8A684000AD
Private Sub chkApres_Click()

  If chkApres.Value = 1 Then
    chkApresEx.Value = 1
  Else
    chkApresEx.Value = 0
        
    If chkAvant.Value = 0 Then
      chkAvant.Value = 1
    End If
  End If
  
End Sub

'##ModelId=5C8A684000CC
Private Sub chkAvant_Click()
  
  If chkAvant.Value = 1 Then
    chkAvantEx.Value = 1
  Else
    chkAvantEx.Value = 0
    chkEcart.Value = 0
    chkDejaAm.Value = 0
    chkResteAm.Value = 0
    
    If chkApres.Value = 0 Then
      chkApres.Value = 1
    End If
  End If
  
End Sub

'##ModelId=5C8A684000EB
Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

'##ModelId=5C8A6840010A
Private Sub cmdStat_Click()
  
  Dim selectedPeriods As New Collection
  Dim numbItemsChecked As Integer
  Dim i As Integer
  Dim checkboxSelected As Boolean
  
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 9
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    
    If checkboxSelected Then
      numPeriode = CInt(sprListe.text)
      numbItemsChecked = numbItemsChecked + 1
      selectedPeriods.Add (numPeriode)
    End If
  Next i
  
  If numbItemsChecked = 0 Then
     MsgBox "Aucune période a été sélectionnée ! Sélectionnez au moins une période et cliquez le bouton 'Lancer la procédure'.", vbExclamation
     Exit Sub
  End If
  
  Dim frm As New frmAutoStat
  frm.selectedPeriods = selectedPeriods
  frm.Show vbModal
  
  'get params
  
  
  Set frm = Nothing
  
End Sub

'##ModelId=5C8A6840012A
Private Sub VScroll1_Change()
  Call pScrollForm
End Sub

'##ModelId=5C8A68400139
Private Sub VScroll1_Scroll()
  Call pScrollForm
End Sub

'##ModelId=5C8A68400149
Private Sub Form_Resize()
  
  If Me.Height > 14400 Then
    Me.Height = 14400
  End If
  If Me.Height < 10000 Then
    Me.Height = 10000
  End If
  If Me.Height < 14400 Then
    VScroll1.Visible = True
    With VScroll1
      .Height = Me.ScaleHeight
      .Min = 0
      '.Max = iFullFormHeigth - iDisplayHeight
      .Max = 14400 - Me.Height
      .SmallChange = Screen.TwipsPerPixelY * 50
      .LargeChange = .SmallChange
    End With
  Else
    VScroll1.Visible = False
  End If
  If Me.Width <> 19845 Then
    Me.Width = 19845
  End If
  
  VScroll1.Height = Me.ScaleHeight
  
End Sub

'##ModelId=5C8A68400168
Private Sub Form_Load()

  Dim iFullFormHeigth As Integer
  Dim iDisplayHeight As Integer
   
  iFullFormHeigth = 14400
  iDisplayHeight = 10000

  'Me.Height = iDisplayHeight

'  With VScroll1
'      .Height = Me.ScaleHeight
'      .Min = 0
'      .Max = iFullFormHeigth - iDisplayHeight
'      .SmallChange = Screen.TwipsPerPixelY * 50
'      .LargeChange = .SmallChange
'  End With
  
  DisableImportFields
  lblLot = ""
  lblExcelPath = ""
  currentLogFile = ""
  radCalcul1.Enabled = False
  radCalcul2.Enabled = False
  
  chkEcart.Enabled = False
  chkDejaAm.Enabled = False
  chkResteAm.Enabled = False
  chkAvant.Enabled = False
  chkApres.Enabled = False
  chkSignalisation.Enabled = False
  chkDeleteTTProv.Enabled = False
    
  actImportLot = False
  actImportExcel = False
  actCalcul = False
  actExportExcel = False
  actExportInfocentre = False
  
  selectedLot = 0
  lblLot = ""
  
  Left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  
  'default import params
  typeDelaiInactivite = eDateFinPeriodePaiement
  typeCalculAnnualisation = eDernierPaiement
  dateArrete = Format(Now, "dd/mm/yyyy")
  
  FillGrid
  
End Sub

'##ModelId=5C8A68400178
Private Sub pScrollForm()

   Dim ctl As Control

   For Each ctl In Me.Controls
      If Not (TypeOf ctl Is VScrollBar) And _
         Not (TypeOf ctl Is Label) And _
         Not (TypeOf ctl Is CommandButton) And _
         Not (TypeOf ctl Is CheckBox) And _
         Not (TypeOf ctl Is Label) And _
         Not (TypeOf ctl Is OptionButton) And _
         Not (TypeOf ctl Is Label) And _
         Not (TypeOf ctl Is Label) And _
         Not (TypeOf ctl Is CommonDialog) Then
            ctl.top = ctl.top + oldPos - VScroll1.Value
      End If
   Next
   
   For Each ctl In Me.Controls
      If ctl.Name = "cmdLaunch" Or ctl.Name = "btnStop" Or ctl.Name = "cmdClose" Or ctl.Name = "lblStatus" Then
        ctl.top = ctl.top + oldPos - VScroll1.Value
      End If
   Next

   VScroll1.Height = Me.ScaleHeight

   oldPos = VScroll1.Value
End Sub


' ******************************************* USER INTERFACE

'##ModelId=5C8A68400187
Private Sub chkImport_Click()

  If chkImport.Value = 1 Then
    EnableImportFields
    actImportLot = True
    actImportExcel = False
    cmdImportParams.Visible = True
    GetImportParams
  Else
    DisableImportFields
    actImportLot = False
    actImportExcel = False
    cmdImportParams.Visible = False
  End If

End Sub

'##ModelId=5C8A684001A7
Private Sub cmdImportParams_Click()
  GetImportParams
End Sub

'##ModelId=5C8A684001B6
Private Sub radImportExcel_Click()
  lblLot = ""
  selectedLot = 0
  cmdSelectExcel.Enabled = True
  cmdSelectLot.Enabled = False
  
  actImportLot = False
  actImportExcel = True
End Sub

'##ModelId=5C8A684001D5
Private Sub radImportLot_Click()
  lblExcelPath = ""
  selectedExcelFile = ""
  cmdSelectExcel.Enabled = False
  cmdSelectLot.Enabled = True
  
  actImportLot = True
  actImportExcel = False
End Sub

'##ModelId=5C8A684001F5
Private Sub cmdSelectExcel_Click()
  
  'Allow user to select a file
  
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "InputPath", App.Path)
  CommonDialog1.filename = "*.xls"
  CommonDialog1.filter = "Fichier Excel|*.xls|Base de données MS Access|*.mdb|"
  CommonDialog1.ShowOpen
  
  If Right(UCase(CommonDialog1.filename), 5) = ".XLSX" Then
    lblExcelPath = ""
    selectedExcelFile = ""
    MsgBox "Le format de fichier Excel 2007 n'est pas compatible avec ce logiciel." & vbLf & "Veuillez enregistrer votre fichier au format Excel 2003 !", vbExclamation + vbOKOnly, "Import des données"
    Exit Sub
  End If
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Or CommonDialog1.filename = "*.xlsx" Then
    lblExcelPath = ""
    selectedExcelFile = ""
  Else
    selectedExcelFile = CommonDialog1.filename
    lblExcelPath = CommonDialog1.filename
  End If
  
End Sub

'##ModelId=5C8A68400214
Private Sub cmdSelectLot_Click()
  
  'Display list of lots
  frmAutoSelectLot.Show vbModal
  selectedLot = frmAutoSelectLot.SelectedLotNumber
  
  If selectedLot <> 0 Then
    lblLot = "Importer le lot numéro " & selectedLot
  Else
    lblLot = ""
  End If
  
End Sub

'##ModelId=5C8A68400224
Private Sub chkExport_Click()
  If chkExport.Value = 1 Then
    actExportExcel = True
    chkEcart.Enabled = True
    chkDejaAm.Enabled = True
    chkResteAm.Enabled = True
  Else
    actExportExcel = False
    chkEcart.Enabled = False
    chkDejaAm.Enabled = False
    chkResteAm.Enabled = False
  End If
End Sub

'##ModelId=5C8A68400243
Private Sub chkExportInfo_Click()
  If chkExportInfo.Value = 1 Then
    actExportInfocentre = True
    chkSignalisation.Enabled = True
    chkDeleteTTProv.Enabled = True
  Else
    actExportInfocentre = False
    chkSignalisation.Enabled = False
    chkDeleteTTProv.Enabled = False
  End If
End Sub

'##ModelId=5C8A68400262
Private Sub chkCalcul_Click()

  If chkCalcul.Value = 1 Then
    radCalcul1.Enabled = True
    radCalcul2.Enabled = True
    actCalcul = True
    chkAvant.Enabled = True
    chkApres.Enabled = True
  Else
    radCalcul1.Enabled = False
    radCalcul2.Enabled = False
    actCalcul = False
    chkAvant.Enabled = False
    chkApres.Enabled = False
  End If
  
End Sub

'##ModelId=5C8A68400281
Private Sub cmdCreateSignalFile_Click()

  If MsgBox("La création du fichier de signalisation sera effectuer immédiatement. Est-ce que vous êtes sur de vouloir continuer ?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If
  
  CreationFichierSignalisation

End Sub

'##ModelId=5C8A68400291
Private Sub cmdDeleteTTProvcoll_Click()

  If MsgBox("La suppression des données dans les tables TTLOGTRAIT et TTPROVCOLL sera effectuer immédiatement. Est-ce que vous êtes sur de vouloir continuer ?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If

End Sub

'##ModelId=5C8A684002A1
Private Sub cmdLaunch_Click()

  Dim i As Integer
  Dim numbItemsChecked As Integer
  Dim checkboxSelected As Boolean
  Dim numPeriode As Long
  Dim currentPeriode As Long
  Dim colPeriods As New Collection
  Dim automationSuccess As Boolean
  Dim statusMessage As String
  
  Dim rsPeriode As ADODB.Recordset
  Dim sqlStr As String
  
  automationSuccess = False
  'connStrProd = DatabaseFileName
  
  
  'Perform user input validation
  'at least 1 periode needs to be selected
  numbItemsChecked = 0
    
  sprListe.VirtualMode = False
  sprListe.DataRefresh
  sprListe.Refresh
  
  numPeriode = 0
   
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 9
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    
    If checkboxSelected Then
      numPeriode = CInt(sprListe.text)
      numbItemsChecked = numbItemsChecked + 1
      colPeriods.Add (numPeriode)
    End If
  Next i
  
  If numbItemsChecked = 0 Then
     MsgBox "Aucune période a été sélectionnée ! Sélectionnez au moins une période et cliquez le bouton 'Lancer la procédure'.", vbExclamation
     GoTo Cleanup
  End If
    
  'at least one option must be checked
  If chkImport.Value = 0 And chkCalcul.Value = 0 And chkExport.Value = 0 And chkExportInfo.Value = 0 Then
    MsgBox "Au moins une action (import, calcul, export vers Excel ou export vers l'infocentre) doit etre sélectionné !", vbExclamation
    Exit Sub
  End If
  
  'if import is selected, either a lot or an import file must have been chosen - the file must exist
  '### can we test if the file is valid
  If chkImport.Value = 1 Then
    If radImportLot.Value = True Then
      If selectedLot = 0 Then
        MsgBox "Cliquez le bouton 'Sélectionner le lot' pour sélectionner le lot à importer !", vbExclamation
        Exit Sub
      End If
    End If
    
    If radImportExcel.Value = True Then
      If selectedExcelFile = "" Then
        MsgBox "Cliquez le bouton 'Sélectionner le fichier Excel' pour sélectionner le fichier à partir de laquelle vous voulez importer les données !", vbExclamation
        Exit Sub
      End If
    End If
  End If
  
  
  If MsgBox("Est-ce que vous est sur de vouloir lancer la procédure de l'automatisation pour les périodes sélectionnées ?", vbYesNo) = vbYes Then
  
    btnStop.Visible = True
    'we start the automation procedure
    Screen.MousePointer = vbHourglass
       
    cmdLaunch.Enabled = False
    lblStatus.Visible = True
    progAuto.Visible = True
           
    progAuto.Min = 0
    progAuto.Max = colPeriods.Count + 1
    progAuto.Value = progAuto.Min + 1
    'lblStatus.Caption = "Procédure en cours..."
    
    'we create the log file
    'Dim m_Logger As New clsLogger
    currentLogFile = m_logPathAuto & "\" & GetWinUser & "_automatisation_" & Format(Now, "dd-mm-yyyy--hh-mm") & ".log"
    m_Logger.FichierLog = currentLogFile
    m_Logger.CreateLog ""
    m_Logger.EcritTraceDansLog ""
    m_Logger.EcritTraceDansLog "********** PROCEDURE AUTOMATION **********"
    m_Logger.EcritTraceDansLog ""
    
    periodCount = colPeriods.Count
    
    For i = 1 To periodCount
      DoEvents
      
      If i = 1 Then
        isFirstPeriode = True
      Else
        isFirstPeriode = False
      End If
      
      If i = periodCount Then
        isLastPeriode = True
      Else
        isLastPeriode = False
      End If
      
      numPeriode = colPeriods(i)
          
     currentPeriode = colPeriods(i)
     currentIndex = i
     
     m_Logger.EcritTraceDansLog ""
     m_Logger.EcritTraceDansLog "******************************************************************************"
     m_Logger.EcritTraceDansLog "### Traitement de la période numéro : " & colPeriods(i)
     m_Logger.EcritTraceDansLog "******************************************************************************"
     m_Logger.EcritTraceDansLog ""
       
     lblStatus.Caption = "Nombre des périodes déjà traité : " & i - 1 & "/" & periodCount & " - période " & colPeriods(i) & " en cours de traitement."
           
    
     '***************************************** Automation Actions START ***********************************
     '******************************************************************************************************
     
     '### perform all required actions
     
     'IMPORT
     If actImportLot Or actImportExcel Then
      ImportExcelOrLot currentPeriode
     End If
     
     If stopProcess Then
         If MsgBox("Est-ce que vous est sur de vouloir arrêter la procédure ?", vbYesNo) = vbYes Then
             stopProcess = False
             GoTo Cleanup
         Else
             stopProcess = False
         End If
     End If
     
     'Calcul
     If actCalcul Then
      Calculate currentPeriode
     End If
     
     If stopProcess Then
         If MsgBox("Est-ce que vous est sur de vouloir arrêter la procédure ?", vbYesNo) = vbYes Then
             stopProcess = False
             GoTo Cleanup
         Else
             stopProcess = False
         End If
     End If
     
     
     'we need a few props for the current periode for the Export
      sqlStr = "SELECT * FROM P3IUser.Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle
      Set rsPeriode = m_dataSource.OpenRecordset(sqlStr, Snapshot)
            
      If rsPeriode.RecordCount > 0 Then
        rsPeriode.MoveFirst
        typeCalculCurrentPeriode = rsPeriode("IdTypeCalcul")
        lotCurrentPeriode = IIf(IsNull(rsPeriode("NUTRAITP3I")), 0, rsPeriode("NUTRAITP3I"))
      End If
     
     'Export Excel
     If actExportExcel Then
      ExportToExcel currentPeriode
     End If
     
     If stopProcess Then
         If MsgBox("Est-ce que vous est sur de vouloir arrêter la procédure ?", vbYesNo) = vbYes Then
             stopProcess = False
             GoTo Cleanup
         Else
             stopProcess = False
         End If
     End If
     
     'Export Infocentre
     If actExportInfocentre Then
      ExportToInfocentre currentPeriode
     End If
     
     
     '******************************************************************************************************
     '***************************************** Automation Actions END *************************************

     If stopProcess Then
         If MsgBox("Est-ce que vous est sur de vouloir arrêter la procédure ?", vbYesNo) = vbYes Then
             stopProcess = False
             GoTo Cleanup
         Else
             stopProcess = False
         End If
     End If
     
     progAuto.Value = i + 1
     lblStatus.Caption = "Nombre des périodes déjà traité : " & i & "/" & colPeriods.Count & " - période " & colPeriods(i) & " en cours de traitement."
           
    Next i
    
  Else
    Exit Sub
  End If
  
  Screen.MousePointer = vbDefault
  
  lblStatus.Caption = "La procédure est terminée ! Nombre des périodes traité : " & colPeriods.Count
  
  If MsgBox("La procédure est terminé. Voulez-vous consultez le fichier log ?", vbInformation + vbYesNo) = vbNo Then
    GoTo Cleanup
  End If
  
  Dim frm As New frmDisplayLog
  frm.FichierLog = currentLogFile
  frm.Show vbModal
  Set frm = Nothing
  
Cleanup:

  'sprListe.VirtualMode = True
  'sprListe.DataRefresh
  'sprListe.Refresh
  
  cmdLaunch.Enabled = True
  lblStatus.Visible = False
  progAuto.Visible = False
  progAuto.Value = progAuto.Min
  lblStatus.Caption = ""
  btnStop.Visible = False
  Set m_Logger = Nothing
  
  'uncheck all checkboxes
'  sprListe.VirtualMode = False
'  sprListe.VirtualMaxRows = -1
'  sprListe.DataRefresh
'  sprListe.Refresh
'
'  For i = 1 To sprListe.DataRowCnt
'      sprListe.Row = i
'      sprListe.Col = 9
'      sprListe.text = 0
'  Next i
'
'  sprListe.VirtualMode = True
 
  Screen.MousePointer = vbDefault
  RefreshListe
    
End Sub

'##ModelId=5C8A684002C0
Private Sub GetImportParams()
  
  Dim frm As New frmAutoSelectImportParams
  frm.DTPicker2.Value = dateArrete
  
  'Dim dateImport As Date
  'dateImport = Format(Now, "dd/mm/yyyy")
  'dateArrete = Format(Now, "dd/mm/yyyy")
  'frm.DTPicker2.MaxDate = dateArrete
  'frm.gDateDebut = DateDebut
  'frm.gDateFin = DateFin
  'frm.rdoImportComplet.Enabled = False
  'frm.rdoImportDonneesSeules.Value = True
  'frm.rdoImportTableParametre.Enabled = False
  'frm.lblDate2.Caption = Replace(frm.lblDate2.Caption, "888", nbJourMax)
    
  If typeDelaiInactivite = eDateFinPeriodePaiement Then
    frm.rdoDateFinPeriode.Value = True
  Else
    frm.rdoDatePaiement = True
  End If
  
  If typeCalculAnnualisation = eDernierPaiement Then
    frm.rdoDernierPaiement.Value = True
  Else
    frm.rdoEnsemblePaiement = True
  End If
 
  frm.Show vbModal
  
  'set all import params
  typeImport = eImportDonneesSeules 'this is a constant
  dateArrete = frm.dateArrete
  
  ' type de calcul de delai
  If frm.calcDateIsFinPeriode = True Then
    typeDelaiInactivite = eDateFinPeriodePaiement
  Else
    typeDelaiInactivite = eDatePaiement
  End If
  
  ' type de calcul de l'annualisation
  If frm.calcAnnualisatIsLastPaiment = True Then
    typeCalculAnnualisation = eDernierPaiement
  Else
    typeCalculAnnualisation = eEnsemblePaiement
  End If
  
  Set frm = Nothing

End Sub

'##ModelId=5C8A684002CF
Private Sub btnStop_Click()
  stopProcess = True
End Sub

'##ModelId=5C8A684002EF
Private Sub EnableImportFields()
  radImportLot.Value = True
  radImportExcel.Value = False
  cmdSelectLot.Enabled = True
  cmdSelectExcel.Enabled = False
  radImportLot.Enabled = True
  radImportExcel.Enabled = True
  lblLot.Enabled = True
  lblExcelPath.Enabled = True
End Sub

'##ModelId=5C8A6840030E
Private Sub DisableImportFields()
  cmdSelectLot.Enabled = False
  cmdSelectExcel.Enabled = False
  radImportLot.Enabled = False
  radImportExcel.Enabled = False
  lblLot.Enabled = False
  lblExcelPath.Enabled = False
End Sub

'##ModelId=5C8A6840031E
Private Sub cmdConfigInfocentre_Click()
  'Display configuration form
  'frmAutoExportInfocentre.SetModeAuto = True
  frmAutoExportInfocentre.Show vbModal
  exportInfoTypeProvision = frmAutoExportInfocentre.frmTypeProvision
  exportInfoDelExistant = frmAutoExportInfocentre.frmDelExistant
  exportInfoCreateSignalisation = frmAutoExportInfocentre.frmCreateSignalisation

End Sub

'##ModelId=5C8A6840033D
Private Sub ExportToInfocentre(numPeriode As Long)
  
  Dim module_calcul As iP3ICalcul, nomModuleCalcul As String
  Dim avecRevalo As Boolean
  Dim sTypeProvision As String
  
  On Error GoTo errExport
  
  If lotCurrentPeriode = 0 Then
    m_Logger.EcritTraceDansLog "Aucun lot n'est attache à la période numéro : " & numPeriode & " donc il n'y aura pas d'export des données pour cette période !"
    Exit Sub
  End If
    
  DoEvents
  lblStatus.Caption = "Nombre des périodes déjà traité : " & currentIndex - 1 & "/" & periodCount & " - période " & numPeriode & _
    " en cours de traitement. - Export vers l'infocentre"
        
  DoEvents
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
  m_Logger.EcritTraceDansLog "### Export de la période numéro : " & numPeriode & " vers l'infocentre"
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      
  'perform the export
  Select Case typeCalculCurrentPeriode
    Case 1
      sTypeProvision = "BILAN"
    Case 2
      sTypeProvision = "CLIENT"
    Case 3
      sTypeProvision = "SIMUL"
  End Select
  
  'we delete data only for the first period
  If isFirstPeriode Then
    If chkDeleteTTProv.Value = 1 Then
      m_dataSource.Execute "DELETE FROM TTPROVCOLL"
      'm_dataSource.Execute "TRUNCATE TABLE TTPROVCOLL"
      DoEvents
      
      m_dataSource.Execute "DELETE FROM TTLOGTRAIT"
      'm_dataSource.Execute "TRUNCATE TABLE TTLOGTRAIT"
      DoEvents
    End If
  End If
  
  
  CopyLot "P3ILOGTRAIT", "TTLOGTRAIT", "", ""
  DoEvents
  
  CopyLot "P3IPROVCOLL", "TTPROVCOLL", " AND DataVersion=0", sTypeProvision
  DoEvents
  
  m_dataSource.Execute "UPDATE TTLOGTRAIT SET NBLIGTRAIT=(SELECT count(*) FROM TTPROVCOLL), MTTRAIT=0"
  DoEvents
  
    
  'Creation du fichier top de signalisation when after treating the last period
  If isLastPeriode And chkSignalisation.Value = 1 Then
    CreationFichierSignalisation
  End If
  
  Exit Sub

errExport:
  
  m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Exit Sub
  Resume Next
  
End Sub

'##ModelId=5C8A6840036C
Private Sub ExportToExcel(numPeriode As Long)
  
  Dim module_calcul As iP3ICalcul, nomModuleCalcul As String
  Dim avecRevalo As Boolean
  
  On Error GoTo errExport
  
  Dim exportFile As String
  
  ' "\" & GetWinUser & "_Automation_" & Format(Now, "dd-mm-yyyy--hh-mm")
  exportFile = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path) & "\Periode-" & numPeriode & ".xls"
  
  DoEvents
  lblStatus.Caption = "Nombre des périodes déjà traité : " & currentIndex - 1 & "/" & periodCount & " - période " & numPeriode & _
    " en cours de traitement. - Export dans Excel "
        
  DoEvents
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
  m_Logger.EcritTraceDansLog "### Export de la période numéro : " & numPeriode & " dans le fichier : " & exportFile
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      
  'perform the export
  
  Dim frm As New frmEditPeriode
  'frm.Hide
  frm.autoMode = True
  frm.autoPeriode = numPeriode
  frm.exportFilename = exportFile
  Set frm.autoLogger = m_Logger
  frm.FormLoad
  
  '###
  frm.m_DetailAffichagePeriode.Ecart = IIf(chkEcart.Value = 1, True, False)
  frm.m_DetailAffichagePeriode.DejaAmorti = IIf(chkDejaAm.Value = 1, True, False)
  frm.m_DetailAffichagePeriode.ResteAAmortir = IIf(chkResteAm.Value = 1, True, False)
  
  frm.FormActivate
  
  frm.ExportToExcel
  
  Unload frm
  Set frm = Nothing
  
  Exit Sub

errExport:
  
  m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Exit Sub
  Resume Next
  
End Sub

'##ModelId=5C8A6840039B
Private Sub Calculate(numPeriode As Long)
  
  Dim module_calcul As iP3ICalcul, nomModuleCalcul As String
  Dim avecRevalo As Boolean
  
  On Error GoTo errCalcul
  
  ' charge l'object de calcul
  Set module_calcul = New P3ICalcul_Generali
    
  'If module_calcul Is Nothing Then Exit Sub
    
  DoEvents
  lblStatus.Caption = "Nombre des périodes déjà traité : " & currentIndex - 1 & "/" & periodCount & " - période " & numPeriode & _
    " en cours de traitement. - Calcul "
        
  DoEvents
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
  m_Logger.EcritTraceDansLog "### Calcul de la période numéro : " & numPeriode
  m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      
  ' effectue le calcul
  avecRevalo = False
  If radCalcul2.Value = True Then
    avecRevalo = True
  End If
  
  Dim av As Boolean
  Dim ap As Boolean
  
  If chkApres.Value = 1 Then
    ap = True
  Else
    ap = False
  End If
  
  If chkAvant.Value = 1 Then
    av = True
  Else
    av = False
  End If
  
  module_calcul.CalculProvisionsAssures avecRevalo, numPeriode, CLng(GroupeCle), m_Logger, True, av, ap
  
  Exit Sub

errCalcul:
  'MsgBox "Erreur durant le calcul : " & Err & vbLf & Err.Description & vbLf & "Objet = " & nomModuleCalcul, vbCritical
  m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Exit Sub
  Resume Next
  
End Sub

'Unified import for Lot and Excel
'##ModelId=5C8A684003C9
Private Sub ImportExcelOrLot(numPeriode As Long)

  Dim objImport As iP3IGeneraliImport
  Dim CleGroupe As Long
  Dim txtObjetImport As String
  Dim sectionObjectImport As String
  Dim type_periode As Integer
  Dim codeRetour As Boolean
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim m_bP3I_Individuel As Boolean
  
  CleGroupe = GroupeCle ' en dur
  
  type_periode = m_dataHelper.GetParameterAsLong("SELECT PETYPEPERIODE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
  
  If type_periode = eRevalo Or type_periode = eProvisionRetraiteRevalo Then
    sectionObjectImport = "ObjetImportRevalo"
  Else
    sectionObjectImport = "ObjetImport"
  End If
  
  If radImportLot.Value = True Then
    txtObjetImport = GetSettingIni(CompanyName, SectionName, "ObjetImportSASP3I", "#")
  Else
    txtObjetImport = GetSettingIni(CompanyName, SectionName, sectionObjectImport, "#")
  End If
  
  On Error GoTo errImport
  
  Set objImport = CreateObject(txtObjetImport)
 
  rq = "SELECT PEDATEDEB, PEDATEFIN, PENBJOURMAX, PENBJOURDC, 65 as PEAGERETRAITE, PEDATEEXT, PETYPEPERIODE " _
      & " FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode
  
  Set rs = m_dataSource.OpenRecordset(rq, Snapshot)
  
  If Not rs.EOF() Then
    DoEvents
    
    If radImportLot.Value = True Then
      'Import Lot
      
      DoEvents
      lblStatus.Caption = "Nombre des périodes déjà traité : " & currentIndex - 1 & "/" & periodCount & " - période " & numPeriode & _
        " en cours de traitement. - Import depuis le lot numéro " & selectedLot
            
      DoEvents
      m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      m_Logger.EcritTraceDansLog "### Import dans la période numéro : " & numPeriode & " depuis le lot numéro " & selectedLot
      m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      'm_Logger.EcritTraceDansLog ""
      
      '### CAREFUL - we may need to change this
      'currently we do only a standard import (no exlatement pour type statutaire)
      
      'Import of type Standard
      codeRetour = objImport.DoImportSASP3I(selectedLot, m_logPathAuto, m_dataSource, CleGroupe, numPeriode, _
        Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
        rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, False, _
        True, dateArrete, typeDelaiInactivite, typeCalculAnnualisation, m_Logger.FichierLog)
      
      
      '### we may need to activate this
      '********************************* CURRENTLY NOT USED START *************************************
      If False Then
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
            
            codeRetour = objImport.DoImportSASP3I(selectedLot, m_logPath, m_dataSource, CleGroupe, numPeriode, _
                     Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                     rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, False)
          
          Else
          
            SetCategoryCodeStatVariable
            
            If CategoryCodeSTAT <> "" Then
              objImport.SetStatutaireVariables NumPeriodeStat, NumPeriodeNonStat, PathSexFileExcel, CategoryCodeSTAT, SexAllMale, TwoLotImport
              
              codeRetour = objImport.DoImportSASP3I(selectedLot, m_logPath, m_dataSource, CleGroupe, numPeriode, _
                     Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                     rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), sFichierIni, False)
            Else
              MsgBox "Le code catégorie pour les assurées du type Statutaire n'est pas renseigné", vbOKOnly, "Code Category Manquant"
              codeRetour = False
            End If
          
          End If
          
        Else
          'we did not launch the import
          codeRetour = False
        End If
      
      End If ' if False then
      '********************************* CURRENTLY NOT USED END *************************************
    
    Else
      'Import Excel
      DoEvents
      lblStatus.Caption = "Nombre des périodes déjà traité : " & currentIndex - 1 & "/" & periodCount & " - période " & numPeriode & _
        " en cours de traitement. - Import depuis le fichier : " & selectedExcelFile
      
      DoEvents
      m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      m_Logger.EcritTraceDansLog "### Import dans la période numéro : " & numPeriode & " depuis le fichier Excel " & selectedExcelFile
      m_Logger.EcritTraceDansLog "-----------------------------------------------------------------------------"
      'm_Logger.EcritTraceDansLog ""
      
      m_bP3I_Individuel = CBool(m_dataHelper.GetParameterAsDouble("SELECT PEP3I_INDIVIDUEL FROM Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle))
  
      codeRetour = objImport.DoImport(CommonDialog1, m_dataSource, CleGroupe, numPeriode, _
                 Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
                 rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), _
                 sFichierIni, m_bP3I_Individuel, True, dateArrete, typeDelaiInactivite, _
                 typeCalculAnnualisation, selectedExcelFile, m_Logger.FichierLog)

'  codeRetour = objImport.DoImport(CommonDialog1, m_dataSource, CleGroupe, numPeriode, _
'                 Format(rs.fields("PEDATEDEB"), "dd/mm/yyyy"), Format(rs.fields("PEDATEFIN"), "dd/mm/yyyy"), _
'                 rs.fields("PENBJOURMAX"), rs.fields("PENBJOURDC"), rs.fields("PEAGERETRAITE"), rs.fields("PEDATEEXT"), _
'                 sFichierIni, m_bP3I_Individuel)
    End If
    
  End If
  
  rs.Close
  Set objImport = Nothing
 
  If codeRetour = False Then
    '###MsgBox "L'opération d'import a été INTERROMPUE !" & vbLf & "Aucun article n'a été ajouté à la période n°." & frmNumPeriode, vbExclamation
  End If
    
  Exit Sub
  
errImport:
  '###MsgBox "Erreur : " & Err.Description & vbLf & "Objet = " & txtObjetImport, vbCritical
  m_Logger.EcritTraceDansLog "Erreur pendant l'import : " & Err.Number & " : " & Err.Description
  Resume Next
  
End Sub


'##ModelId=5C8A68410010
Private Sub CopyLot(sTable As String, sOut As String, sWhereIn As String, sTypeProvision As String)
  
  If lotCurrentPeriode = 0 Then Exit Sub
  
  On Error GoTo err_CopyLot
  
  Dim rsIn As ADODB.Recordset, rsOut As ADODB.Recordset, i As Integer, f As ADODB.field
  
  Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM " & sTable & " WHERE NUTRAITP3I=" & lotCurrentPeriode & sWhereIn, Snapshot)
  Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM " & sOut, Dynamic)
  
  Do Until rsIn.EOF
    rsOut.AddNew
    
    For i = 0 To rsIn.fields.Count - 1
      Set f = rsIn.fields(i)
  
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
      rsOut.fields("TYPEPROVISION").Value = sTypeProvision
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
  If Err.Number <> -2147217873 And Err.Number <> 3219 Then
    m_Logger.EcritTraceDansLog "Erreur CopyLot " & Err & vbLf & Err.Description
  End If
  Resume Next
End Sub



'******************************************************************************************************************************
'********************************************* FILL THE GRID WITH LIST OF PERIODES ********************************************
'******************************************************************************************************************************

'##ModelId=5C8A6841007E
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

'##ModelId=5C8A6841009D
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

'##ModelId=5C8A684100AD
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  sprListe.sheet = sprListe.ActiveSheet
  sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
End Sub

'##ModelId=5C8A684100FB
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
  
  sprListe.ColWidth(3) = 20
  
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

'##ModelId=5C8A68410187
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
'    Dim statut As String
'    sprListe.Col = 10
'    sprListe.Row = Row
'    statut = sprListe.text
'
'    sprListe.Col = -1
'    sprListe.ForeColor = noir
'
'    If Len(statut) > 0 Then
'      If Left$(LCase(statut), 4) = "arch" Then
'        sprListe.BackColor = LTRED
'      End If
'    End If
    
    'change background color to black for the row that receives the focus
    sprListe.Row = Row
    sprListe.ForeColor = noir
    
    sprListe.Row = NewRow
    sprListe.BackColor = noir
    sprListe.ForeColor = blanc
    
End Sub


