VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVoirTrace 
   Caption         =   "Trace du Jeux de données ..."
   ClientHeight    =   7995
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   13170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   5625
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12600
      Top             =   5535
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   6435
      Top             =   5625
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
      RecordSource    =   "SELECT * FROM P3IUser.P3IUser.Assure WHERE RECNO=-1"
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
      Height          =   4965
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   11040
      _Version        =   524288
      _ExtentX        =   19473
      _ExtentY        =   8758
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   10
      MaxRows         =   1000000
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmVoirTrace.frx":0000
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   10
      VisibleRows     =   100
      AppearanceStyle =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11970
      Top             =   5445
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":060A
            Key             =   "openCahier"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0714
            Key             =   "openPeriode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":081E
            Key             =   "About"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0928
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0A32
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0B3C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0C96
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0DF0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":0F4A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":2B0C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":2C66
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":2DC0
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVoirTrace.frx":2F1A
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      Begin VB.OptionButton rdoSuppr 
         Caption         =   "Suppression"
         Height          =   240
         Left            =   2385
         TabIndex        =   7
         Top             =   90
         Width           =   1275
      End
      Begin VB.OptionButton rdoModif 
         Caption         =   "Modification"
         Height          =   240
         Left            =   990
         TabIndex        =   6
         Top             =   90
         Width           =   1275
      End
      Begin VB.OptionButton rdoTout 
         Caption         =   "Tout"
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   1050
      End
      Begin VB.CommandButton btnExport 
         Caption         =   "E&xporter"
         Height          =   285
         Left            =   3780
         TabIndex        =   4
         Top             =   45
         Width           =   1215
      End
      Begin VB.TextBox lblFillTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   12555
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "----"
         Top             =   90
         Visible         =   0   'False
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmVoirTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A68160073"

Option Explicit

'##ModelId=5C8A6816017C
Public frmNumeroLot As Long

'##ModelId=5C8A6816019C
Private Const colTypeChangement = 3

'##ModelId=5C8A681601AB
Public Property Let NumeroLot(n As Long)
  frmNumeroLot = n
End Property


'##ModelId=5C8A681601CA
Private Sub btnClose_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A681601EA
Private Sub btnExport_Click()
  On Error GoTo err_export
  
  CommonDialog1.filename = "Trace_Lot_" & frmNumeroLot & ".xls"
  CommonDialog1.filter = "Fichier Excel|*.xls|"
  
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
  CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
  CommonDialog1.CancelError = True
  
  CommonDialog1.ShowSave
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Then
    Exit Sub
  End If
  
  If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Then
'    ExportTableToExcelFile CommonDialog1.filename, _
'                           "Trace_Lot" & frmNumeroLot, _
'                           "Trace", sprListe, CommonDialog1, "", False, False
    Screen.MousePointer = vbHourglass
    
    ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, CommonDialog1.filename, "Trace_Lot" & frmNumeroLot, sprListe
    
    Screen.MousePointer = vbDefault
  End If

  Exit Sub
  
err_export:
  
  If Err <> cdlCancel Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  CommonDialog1.CancelError = False

End Sub


'##ModelId=5C8A681601FA
Private Sub Form_Activate()
  RefreshListe
  
  sprListe.SetFocus
End Sub

'##ModelId=5C8A68160209
Private Sub Form_Load()
  ' chargement du masque du spread
  sprListe.LoadFromFile App.Path & "\VoirTrace.ss7"
  
  m_dataSource.SetDatabase dtaPeriode
  
  rdoTout.Value = True
End Sub

'##ModelId=5C8A68160228
Private Sub SetColonneDataFill(numCol As Integer)
'  Dim i As Integer
  
  'For i = 2 To sprListe.MaxCols
  '  sprListe.Col = i
    sprListe.Col = numCol
    sprListe.DataFillEvent = True
  'Next
'  sprListe.Col = numCol
'  sprListe.Col2 = numCol
'  sprListe.Row = 1
'  sprListe.Row2 = sprListe.MaxRows
'  sprListe.BlockMode = True
'  sprListe.DataFillEvent = True
'  sprListe.BlockMode = False
End Sub

'##ModelId=5C8A68160247
Private Sub RefreshListe()
  
  If frmNumeroLot = 0 Then Exit Sub
  
  Dim rq As String, rs As ADODB.Recordset
  Dim filter As String
  Dim i As Integer
  
  Dim debut As Date, fin As Date
  
  debut = Now
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre
  Me.Caption = "Trace du jeux de données n°" & frmNumeroLot
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = True
  sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  'sprListe.VScrollSpecial = True
  'sprListe.VScrollSpecialType = 0
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT P.NUENRP3I, convert(varchar(12),P.DateModif, 103) + ' ' + convert(varchar(12),P.DateModif, 108) as DateModification, P.TypeChangement, P.Commentaire, P.CDCOMPAGNIE, P.CDAPPLI, P.CDPRODUIT, P.NUCONTRA, P.NUPRET, P.NUDOSSIERPREST, " _
       & " P.NUSOUSDOSSIERPREST, P.NUCONTRATGESTDELEG, P.CDGARAN, P.NUREFMVT, P.CDTYPMVT, P.NUTRAITE, P.CDCATADHESION, P.LBCATADHESION, " _
       & " P.CDOPTION, P.LBOPTIONCON, P.LBSOUSCR, P.IDASSUREAGI, P.IDASSURE, P.LBASSURE, P.DTNAISSASS, P.CDSEXASSURE, P.IDRENTIERAGI, " _
       & " P.IDRENTIER, P.LBRENTIER, P.DTNAISSREN, P.IDCORENTIERAGI, P.IDCORENTIER, P.LBCORENTIER, P.DTNAISSCOR, P.DTSURVSIN, P.AGESURVSIN, " _
       & " P.DURPOSARR, P.ANCARRTRA, P.CDPERIODICITE, P.CDTYPTERME, P.DTEFFREN, P.DTLIMPRO, P.DTDERREG, P.DTDEBPER, P.DTFINPER, P.CDPROEVA, " _
       & " P.ANNEADHE, P.CDSINCON, P.CDMISINV, P.DTMISINV, P.NUETABLI, P.TXREVERSION, P.DTCALCULPROV, P.DTTRAITPROV, P.DTCREATI, P.DTSIGBIA, " _
       & " P.DTDECSIN, P.CDRISQUE, P.LBRISQUE, P.CDSITUATSIN, P.DTSITUATSIN, P.CDPRETRATTSIN, P.CDCTGPRT, P.LBCTGPRT, P.DTPREECH, P.DTDERECH, " _
       & " P.CDPERIODICITEECH, P.MTECHEANCE1, P.DTDEBPERECH1, P.DTFINPERECH1, P.MTECHEANCE2, P.DTDEBPERECH2, P.DTFINPERECH2, P.MTECHEANCE3, " _
       & " P.DTDEBPERECH3, P.DTFINPERECH3, P.CDTYPAMO, P.LBTYPAMO, P.TXINVPEC, P.DTDEBPERPIP, P.DTFINPERPIP, P.DTSAISIEPERJUSTIF, P.DTDEBPERJUSTIF, " _
       & " P.DTFINPERJUSTIF, P.DTDEBDERPERRGLTADA, P.DTFINDERPERRGLTADA, P.DTDERPERRGLTADA, P.MTDERPERRGLTADA, P.DTDEBDERPERRGLTADC, P.DTFINDERPERRGLTADC, " _
       & " P.DTDERPERRGLTADC, P.MTDERPERRGLTADC, P.CDSINPREPROV, P.MTTOTREGLEICIV, P.DTDEBPROV, P.DTFINPROV, P.INDBASREV, P.MTPREANN, P.MTPREREV, " _
       & " P.MTPREMAJ , P.MTPRIREG, P.MTPRIRE1, P.MTPRIRE2, P.CDMONNAIE, P.CDPAYS, P.CDAPPLISOURCE, "
       
  rq = rq & " P.CDCATINV, P.LBCATINV,  P.CDCONTENTIEUX, P.NUSINISTRE, " _
          & " P.CDCHOIXPREST, P.LBCHOIXPREST, P.MTCAPSSRISQ, P.FLAMORTISSABLE, " _
          & " P.LBCOMLIG " _
          & " FROM P3ITRACE P " _
          & " WHERE P.NUTRAITP3I = " & frmNumeroLot
       
  If rdoModif.Value = True Then
    rq = rq & " AND (P.TypeChangement='M' OR P.TypeChangement='A')"
  ElseIf rdoSuppr.Value = True Then
    rq = rq & " AND (P.TypeChangement='S' OR P.TypeChangement='D')"
  End If
       
  rq = rq & " ORDER BY NUENRP3I, DateModif"
          
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
  
  SetColonneDataFill colTypeChangement
  
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If dtaPeriode.Recordset.EOF = False Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
  
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
    sprListe.VirtualMaxRows = dtaPeriode.Recordset.RecordCount
  
    dtaPeriode.Recordset.MoveFirst
  Else
    sprListe.MaxRows = 0
    sprListe.VirtualMaxRows = 0
    sprListe.ColWidth(1) = 0
    sprListe.Visible = True
    sprListe.ReDraw = True

    Screen.MousePointer = vbDefault
    
    GoTo pas_de_donnee
  End If
  
  ' largeur des colonnes
  LargeurMaxColonneSpread sprListe
  
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = sprListe.MaxCols - 1
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = sprListe.MaxCols
  sprListe.Col2 = sprListe.MaxCols
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
pas_de_donnee:

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True
  
  Screen.MousePointer = vbDefault

  fin = Now
  
  lblFillTime.text = "Remplissage : " & DateDiff("s", debut, fin) & " s"

  Exit Sub

err_RefreshListe:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A68160257
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar1.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 160
 
  topbtn = Me.ScaleHeight - btnHeight
  
  sprListe.Height = Maximum(topbtn - Toolbar1.Height - 100, 0)
  
  PlacePremierBoutton btnClose, topbtn
End Sub

'##ModelId=5C8A68160267
Private Sub rdoModif_Click()
  RefreshListe
End Sub

'##ModelId=5C8A68160276
Private Sub rdoSuppr_Click()
  RefreshListe
End Sub

'##ModelId=5C8A68160296
Private Sub rdoTout_Click()
  RefreshListe
End Sub

'##ModelId=5C8A681602A5
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  If Col = colTypeChangement Then
    Dim version As Variant, txt As String, color As Long
    
    sprListe.GetDataFillData version, vbString
    
    Cancel = True
    
    Select Case version
      Case "I"
        txt = "Initiale"
        color = vbWhite
        
      Case "M"
        txt = "Modifiée"
        color = LTYELLOW
        
      Case "A"
        txt = "Ajoutée"
        color = LTCYAN
        
      Case "S"
        txt = "Supprimée"
        color = LTRED
        
      Case "D"
        txt = "Doublon"
        color = PINK
             
      Case Else
        txt = "Inconnue"
        color = LTGRAY
    End Select
    
    ' changement du type de données
    sprListe.Row = Row
    sprListe.Col = Col
    sprListe.CellType = CellTypeStaticText
    sprListe.TypeHAlign = TypeHAlignCenter
    
    ' couleur de fond
    sprListe.Row = Row
    sprListe.Col = -1
    sprListe.Row2 = Row
    sprListe.Col2 = -1
    sprListe.BackColor = color
    
    ' texte
    sprListe.SetText Col, Row, txt
  End If
End Sub

'##ModelId=5C8A68160322
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A68160351
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "Commentaire" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  ElseIf dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "TypeChangement" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeStaticText
    sprListe.TypeMaxEditLen = 15
'  ElseIf dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "DateModif" Then
'    sprListe.Col = Col
'    sprListe.Row = -1
'    sprListe.CellType = CellTypeDate
'    sprListe.TypeDateFormat = TypeDateFormatDDMMYY
  End If
End Sub

