VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReassurance 
   Caption         =   "Préparation de la Réassurance"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   Icon            =   "frmReassurance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
   Begin FPSpreadADO.fpSpread sprListe 
      Bindings        =   "frmReassurance.frx":1BB2
      Height          =   5550
      Left            =   0
      TabIndex        =   6
      Top             =   450
      Width           =   11085
      _Version        =   524288
      _ExtentX        =   19553
      _ExtentY        =   9790
      _StockProps     =   64
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
      SpreadDesigner  =   "frmReassurance.frx":1BCB
      AppearanceStyle =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11100
      _ExtentX        =   19579
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.TextBox lblFilter 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "fdgfhgf"
         Top             =   90
         Width           =   5010
      End
      Begin VB.ComboBox cboTraite 
         Height          =   315
         Left            =   3825
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   45
         Width           =   2040
      End
      Begin VB.CommandButton btnClose 
         Caption         =   "&Fermer"
         Height          =   285
         Left            =   2565
         TabIndex        =   3
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton btnPrint 
         Caption         =   "&Imprimer"
         Height          =   285
         Left            =   1305
         TabIndex        =   2
         Top             =   45
         Width           =   1215
      End
      Begin VB.CommandButton btnExport 
         Caption         =   "&Exporter"
         Height          =   285
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   945
      Top             =   5850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   8775
      Top             =   6030
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
End
Attribute VB_Name = "frmReassurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E60068"
Option Explicit

'##ModelId=5C8A67E60152
Public fmFilter As clsFilter
'

'##ModelId=5C8A67E60153
Private Sub btnClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A67E60162
Private Sub btnExport_Click()
  On Error GoTo err_export
  
  CommonDialog1.filename = "*.xls"
  CommonDialog1.filter = "Fichier Excel|*.xls|"
  
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
  CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
  
  CommonDialog1.ShowSave
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Then
    Exit Sub
  End If
  
  If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Then
'    ExportTableToExcelFile "Assure periode " & NumPeriode & ".xls", _
'                           "Periode " & NumPeriode & IIf(lblFilter <> "", " avec filtre", ""), _
'                           "", sprListe, CommonDialog1, lblFilter, False
    
    Screen.MousePointer = vbHourglass
    
    ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, CommonDialog1.filename, "Periode " & numPeriode & IIf(lblFilter <> "", " avec filtre", ""), sprListe
    
    Screen.MousePointer = vbDefault
  End If

  Exit Sub
  
err_export:
  
  If Err <> cdlCancel Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  CommonDialog1.CancelError = False
End Sub

'##ModelId=5C8A67E60171
Private Sub btnPrint_Click()
  Dim bUsePrintDlg As Integer
  Dim rs As Recordset
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
  
  Dim dd As String, df As String
  
  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")
  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")

  DescriptionPeriode = " période " & numPeriode & " ( " & dd & " au " & df & " ) " & " du Groupe " & NomGroupe
 
  With sprListe
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
    .PrintJobName = "Assurés de la période " & numPeriode
    .PrintHeader = "/c Préparation de la liste de réassurance : Assurés de la " & DescriptionPeriode & "/n  "
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

'##ModelId=5C8A67E60190
Private Sub cboTraite_Click()
  RefreshListe
End Sub

'##ModelId=5C8A67E601A0
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A67E601DF
Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
  
  m_dataSource.SetDatabase dtaPeriode
  
  FillComboTraite
  
  RefreshListe
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67E601EE
Private Sub FillComboTraite()
  Dim rq As String
  
  rq = "SELECT DISTINCT 1, Assure.POTRAITE_RASSUR2 "
  
  rq = rq & " FROM Assure Assure "
  
  rq = rq & " Where (POPERCLE = " & numPeriode & " AND POGPECLE = " & GroupeCle & ") " _
      & " And (Not Assure.POTRAITE_RASSUR2 Is Null And Assure.POTRAITE_RASSUR2 <> '')" _
      & " AND Assure.POTRAITE_RASSUR_TAUX <> 0"
      
  rq = rq & fmFilter.GetSelectionSQLString & " ORDER BY Assure.POTRAITE_RASSUR2"
  
  m_dataHelper.FillCombo cboTraite, rq, -1
  If cboTraite.ListCount > 0 Then
    cboTraite.ListIndex = 0
  End If
End Sub

'##ModelId=5C8A67E601FE
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar1.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.ScaleWidth - sprListe.Left - 30
  sprListe.Height = Me.ScaleHeight - sprListe.top - 30
End Sub

'##ModelId=5C8A67E6021D
Private Sub SetColonneDataFill(numCol As Integer)
  sprListe.Col = numCol
  sprListe.Col2 = numCol
  sprListe.Row = 1
  sprListe.Row2 = sprListe.MaxRows
  sprListe.BlockMode = True
  sprListe.DataFillEvent = True
  sprListe.BlockMode = False
End Sub

'##ModelId=5C8A67E6023C
Private Sub RefreshListe()
  'Dim rs As Recordset
  
  Dim rq As String
  Dim filter As String
  Dim i As Integer
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  Me.Caption = "Préparation de la liste de réassurance : Assurés de la période " & numPeriode & " du Groupe '" & NomGroupe & "'"
  
  lblFilter = fmFilter.SelectionString
  
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
        
  sprListe.MaxRows = 0
    
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  ' FORMAT(Assure.POTRAITE_RASSUR_TAUX, ""##0%"") as [Taux] "
  ' Format(Assure.POCONVENTION, ""0 00 000000 00 00"") as NCA
  rq = "SELECT Garantie.GALIB as Risques, Assure.POCONVENTION as NCA, ' ' as [Raison Sociale], " _
      & " Assure.PONUMCLE as [N° SS], Assure.PONOM as [Nom, Prénom], " _
      & " Assure.PONAIS as [Date de Naissance], " _
      & " Assure.POARRET as [Date du Sinistre],  " _
      & " Assure.POPREMIER_PAIEMENT as [Date du 1er Paiement], " _
      & " Assure.PODERNIERPAIEMENT as [Date du Dernier Paiement], " _
      & " Assure.POPRESTATION_AN as [Montant Annualisé], Assure.POPRESTATION as [Total des Prestations Réglées], " _
      & " Assure.POPSAP as [PSAP Décés],  Assure.PODEBUT as [Début Derniere Prestation], " _
      & " Assure.POFIN as [Fin Derniere Prestation], Assure.POTRAITE_RASSUR_TAUX as [Taux] "
  
  rq = rq & " FROM Assure Assure INNER JOIN Garantie Garantie ON (Garantie.GAGROUPCLE=Assure.POGPECLE) AND (Garantie.GAGARCLE=Assure.POGARCLE)"
  
  rq = rq & " Where (POPERCLE = " & numPeriode & " AND POGPECLE = " & GroupeCle & ") " _
      & " And (Not Assure.POTRAITE_RASSUR Is Null And Assure.POTRAITE_RASSUR <> '')" _
      & " AND Assure.POTRAITE_RASSUR_TAUX <> 0 "
        
  ' bad query to configure spread and fire DataFill event
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq & " AND Assure.POTRAITE_RASSUR='bidon_donc_vide'")
  dtaPeriode.Refresh
  
  Set sprListe.DataSource = dtaPeriode
    
  ' datafill event pour formater les données
  SetColonneDataFill 2 ' NCA
  SetColonneDataFill 15 ' taux
  
  If cboTraite.ListIndex <> -1 Then
    rq = rq & fmFilter.GetSelectionSQLString & " AND Assure.POTRAITE_RASSUR2='" & cboTraite.List(cboTraite.ListIndex) & "' ORDER BY Garantie.GALIB, Assure.POTRAITE_RASSUR_TAUX, Assure.PONUMCLE"
  Else
    rq = rq & fmFilter.GetSelectionSQLString & " AND Assure.POTRAITE_RASSUR2='bidon_donc_vide' ORDER BY Garantie.GALIB, Assure.POTRAITE_RASSUR_TAUX, Assure.PONUMCLE"
  End If
  
  ' real refresh
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
  
  ' mets à jours les n° de ligne dans le spread
  If Not dtaPeriode.Recordset.EOF Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
  
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
  
    dtaPeriode.Recordset.MoveFirst
  Else
    sprListe.MaxRows = 0
    sprListe.Visible = True

    Screen.MousePointer = vbDefault
    Exit Sub
  End If
      
  sprListe.BlockMode = True
  
  ' tour de la liste des assurés en noir
  sprListe.Row = 1
  sprListe.Col = 1
  sprListe.Row2 = sprListe.MaxRows
  sprListe.Col2 = sprListe.MaxCols
  sprListe.CellBorderColor = vbBlack
  sprListe.CellBorderType = SS_BORDER_TYPE_OUTLINE
  sprListe.CellBorderStyle = SS_BORDER_STYLE_SOLID
  sprListe.Action = SS_ACTION_SET_CELL_BORDER
    
  sprListe.BlockMode = False
    
  For i = 1 To sprListe.MaxCols - 1
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i)
  Next i
  
  ' affiche le spread (vitesse)
  sprListe.ReDraw = True
  sprListe.Visible = True

  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67E6025C
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  ' Taux
  If Col = 15 Then
    Dim Taux As Variant
    
    sprListe.GetDataFillData Taux, vbString
    
    sprListe.Col = Col
    sprListe.Row = Row
    sprListe.CellType = CellTypePercent
    sprListe.TypePercentDecPlaces = 0
    
    sprListe.SetText Col, Row, Format(Taux, "##0%")
    
    Cancel = True
  End If
  
  ' NCA
  If Col = 2 Then
    Dim NCA As Variant
    
    sprListe.GetDataFillData NCA, vbString
    NCA = String(13 - Len(NCA), "0") & NCA
    
    sprListe.SetText Col, Row, Format(NCA, m_FormatNCA)
    
    Cancel = True
  End If
End Sub

