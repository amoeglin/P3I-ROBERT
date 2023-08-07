VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAutoSelectLot 
   Caption         =   "Lots de données disponibles"
   ClientHeight    =   7995
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   13170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   9045
      TabIndex        =   0
      Top             =   5085
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   5400
      Top             =   5130
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
      TabIndex        =   2
      Top             =   0
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
      SpreadDesigner  =   "frmAutoSelectLot.frx":0000
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   10
      VisibleRows     =   100
      AppearanceStyle =   0
   End
   Begin VB.TextBox lblFillTime 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "----"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2625
   End
End
Attribute VB_Name = "frmAutoSelectLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A68480283"
Option Explicit

'##ModelId=5C8A6848039C
Public ImportFromExcel As Boolean
'##ModelId=5C8A684803BB
Public ExcelFilePath As String
'##ModelId=5C8A684803CB
Public SelectedLotNumber As Integer

'##ModelId=5C8A68490002
Private frmNumPeriode As Long
'##ModelId=5C8A68490021
Private frmNumeroLot As Long

'##ModelId=5C8A68490041
Private Sub Form_Load()

  frmNumPeriode = numPeriode
   
  m_dataSource.SetDatabase dtaPeriode
  sprListe.LoadFromFile App.Path & "\ListeJeuxDonnees.ss7"
  RefreshListe
  
End Sub

'##ModelId=5C8A68490060
Private Sub SetNumeroLot()

  frmNumeroLot = 0
  If sprListe.ActiveRow < 0 Then Exit Sub
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 1
  frmNumeroLot = CLng(sprListe.text)
  SelectedLotNumber = frmNumeroLot
  
End Sub

'##ModelId=5C8A68490070
Private Sub cmdOk_Click()
  SetNumeroLot
  Unload Me
End Sub

'##ModelId=5C8A6849007F
Private Sub btnClose_Click()
  'ret_code = -1
  'SetNumeroLot
  Unload Me
End Sub

'##ModelId=5C8A6849008F
Private Sub RefreshListe()
  
  Dim rq As String, rs As ADODB.Recordset
  Dim filter As String
  Dim i As Integer
  
  Dim debut As Date, fin As Date
  
  debut = Now
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre
  Me.Caption = "Jeux de données disponibles"
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = False
  'sprListe.VirtualMode = True
  'sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  'sprListe.VScrollSpecial = True
  'sprListe.VScrollSpecialType = 0
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT NUTRAITP3I, DTTRAIT as [Date], CONVERT(CHAR(8), HHTRAIT, 8) as [Heure], NUTRAITP3I as [Identifiant du lot], RTRIM(NUTRAIT) as [N° traitement AGI], " _
       & " DTDEBPER as [Debut], DTFINPER as [Fin], NBLIGTRAIT as [nb lignes], MTTRAIT as [Montant], NUTRAIT, IDTABLESAS, Commentaire" _
       & " FROM P3ILOGTRAIT ORDER BY NUTRAITP3I DESC"
          
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
    
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If dtaPeriode.Recordset.EOF = False Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
  
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
    'sprListe.VirtualMaxRows = dtaPeriode.Recordset.RecordCount
  
    dtaPeriode.Recordset.MoveFirst
  Else
    sprListe.MaxRows = 0
    'sprListe.VirtualMaxRows = 0
    sprListe.ColWidth(1) = 0
    sprListe.Visible = True
    sprListe.ReDraw = True

    Screen.MousePointer = vbDefault
    
    GoTo pas_de_donnee
  End If
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
     
  For i = 2 To sprListe.MaxCols
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 2
  Next i
 
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


'##ModelId=5C8A684900AE
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  sprListe.top = 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 130
  topbtn = Me.ScaleHeight - btnHeight
  sprListe.Height = Maximum(topbtn - 100, 0)
  
  PlacePremierBoutton cmdOk, topbtn
  'PlaceBoutton btnImportSASP3I, btnImportXLS, topbtn
  'PlaceBoutton btnEdit, btnImportSASP3I, topbtn
  'PlaceBoutton btnUtiliser, btnEdit, topbtn
  'PlaceBoutton btnExporter, btnUtiliser, topbtn
  PlaceBoutton btnClose, cmdOk, topbtn

End Sub

'##ModelId=5C8A684900BE
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A684900FC
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "Commentaire" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  End If
End Sub

