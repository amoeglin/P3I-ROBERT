VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPeriode 
   Caption         =   " Période pour le groupe ..."
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15510
   Icon            =   "frmPeriode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15510
   Begin VB.CommandButton Command1 
      Caption         =   "ALL"
      Height          =   375
      Left            =   11160
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnConsultArchive 
      Caption         =   "Consult Archive"
      Height          =   375
      Left            =   12360
      TabIndex        =   18
      Top             =   45
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboArchiveRestore 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   50
      Width           =   2760
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Statut de période :"
      Top             =   120
      Width           =   1275
   End
   Begin MSComctlLib.ProgressBar progArchive 
      Height          =   300
      Left            =   9000
      TabIndex        =   14
      Top             =   8280
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Top             =   7320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnRestaurer 
      Caption         =   "&Restaurer"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton btnArchiver 
      Caption         =   "&Archiver"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "E&xporter"
      Height          =   375
      Left            =   8415
      TabIndex        =   8
      Top             =   5985
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   6390
      TabIndex        =   7
      Top             =   5985
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   90
      Top             =   6075
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
      Height          =   4965
      Left            =   0
      TabIndex        =   6
      Top             =   450
      Width           =   11690
      _Version        =   524288
      _ExtentX        =   20620
      _ExtentY        =   8758
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
      SpreadDesigner  =   "frmPeriode.frx":1BB2
      ScrollBarTrack  =   3
      AppearanceStyle =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15510
      _ExtentX        =   27358
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "openPeriode"
            Description     =   "Période"
            Object.ToolTipText     =   "Caractéristiques de la période"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyPeriode"
            Description     =   "Copie période"
            Object.ToolTipText     =   "Copie les caractéristique de la période"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Description     =   "Impression"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultArchive"
            Description     =   "Consult Archive"
            Object.ToolTipText     =   "Consulter la base Archives"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.ComboBox cboTypePeriode 
         Height          =   315
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   45
         Width           =   2760
      End
      Begin VB.TextBox lblFilter 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Type de période :"
         Top             =   90
         Width           =   1275
      End
   End
   Begin VB.CommandButton btnEdition 
      Caption         =   "Choix des &Editions..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4275
      TabIndex        =   2
      Top             =   5535
      Width           =   1695
   End
   Begin VB.CommandButton btnDel 
      Caption         =   "&Supprimer"
      Height          =   375
      Left            =   6390
      TabIndex        =   4
      Top             =   5490
      Width           =   1935
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Imprimer"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5985
      Width           =   1935
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Consulter, Importer, Calculer"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&Nouvelle"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   5490
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8775
      Top             =   5625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":219C
            Key             =   "openCahier"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":22A6
            Key             =   "openPeriode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":23B0
            Key             =   "About"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":24BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":25C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":271E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":2878
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":2BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPeriode.frx":2F1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2970
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre des périodes déjà archivées : 3/10  - période 329 en cours d'archivage."
      Height          =   300
      Left            =   2400
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   6015
   End
End
Attribute VB_Name = "frmPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67A903B9"
Option Explicit

'##ModelId=5C8A67AA00EA
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'##ModelId=5C8A67AA009C
Private fInitDone As Boolean
'##ModelId=5C8A67AA00BB
Private stopArchiving As Boolean

'##ModelId=5C8A67AA0109
Private Sub btnConsultArchive_Click()

  'backcolor : &H8000000F&  -default -- red: &H000000FF&
  
  If archiveMode Then 'back to normal mode
    dtaPeriode.ConnectionString = DatabaseFileName
    archiveMode = False
    btnConsultArchive.Caption = "Archive Mode"
    Me.BackColor = &H8000000F
    
    'disable controls
    Toolbar1.Enabled = True
    cboTypePeriode.Enabled = True
    cboArchiveRestore.Enabled = True
    
    EnableButtons "btnEdition"
    
  Else ' we are in archive mode
    dtaPeriode.ConnectionString = DatabaseFileNameArchive
    archiveMode = True
    btnConsultArchive.Caption = "Normal Mode"
    Me.BackColor = &HFF&
    
    'disable controls
    Toolbar1.Enabled = False
    cboTypePeriode.Enabled = False
    cboArchiveRestore.Enabled = False
    
    DisableButtons "btnDel, btnRestaurer, btnConsultArchive, btnClose, btnEdit, btnClose"
     
  End If
  
  dtaPeriode.Refresh
  RefreshListe
  
End Sub

'##ModelId=5C8A67AA0119
Private Sub Command1_Click()

Dim i As Integer

sprListe.VirtualMode = False
sprListe.VirtualMaxRows = -1
sprListe.DataRefresh
sprListe.Refresh

For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 12
    sprListe.text = 1
Next i

sprListe.VirtualMode = True

End Sub

'##ModelId=5C8A67AA0128
Private Sub Form_Activate()
  Dim r As Long, tr As Long
  
  Screen.MousePointer = vbHourglass
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
  
  RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  sprListe.TopRow = tr
  sprListe.SetActiveCell 2, r
  
  sprListe.Row = r
  sprListe.SelModeSelected = True
  
  sprListe.SetFocus
  
  If fInitDone = False Then
  
    fInitDone = True
  
    Me.top = 0
    Me.Left = 0
    
    If Screen.Width > 24000 Then
      Me.Width = 24000 ' 17500
    End If
    
    If Screen.Height > 7000 Then
      Me.Height = 7000
    End If
  
  End If
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67AA0148
Private Sub Form_Load()
  
  If DroitAdmin Then
    btnDel.Enabled = True
  Else
    btnDel.Enabled = False
  End If
  
  frmMain.mnuAnnexes.Enabled = False
    
  m_dataSource.SetDatabase dtaPeriode
  
  ' chargement du masque du spread
  'sprListe.LoadFromFile App.Path & "\Periode.ss6"
  sprListe.DataSource = dtaPeriode
  
  fInitDone = False
  stopArchiving = False
  archiveMode = False
  
  ' liste des types de période
  cboTypePeriode.Clear
  
  cboTypePeriode.AddItem "Tous"
  cboTypePeriode.ItemData(cboTypePeriode.ListCount - 1) = 0
  
  m_dataHelper.FillCombo cboTypePeriode, "SELECT IdTypePeriode, Libelle FROM TypePeriode ORDER BY IdTypePeriode", 0, False, False
  
  cboTypePeriode.ListIndex = 0
  
  'Archive - Restore Combo
  cboArchiveRestore.Clear
  
  cboArchiveRestore.AddItem "Tous"
  cboArchiveRestore.ItemData(cboArchiveRestore.NewIndex) = 0
  cboArchiveRestore.AddItem "Archivée"
  cboArchiveRestore.ItemData(cboArchiveRestore.NewIndex) = 1
  cboArchiveRestore.AddItem "Restaurée"
  cboArchiveRestore.ItemData(cboArchiveRestore.NewIndex) = 2
  
  cboArchiveRestore.ListIndex = 0
  
  'Allow access to archive functions only if there is a DB connection string
  If DatabaseFileNameArchive = "" Then
    Text1.Visible = False
    btnArchiver.Visible = False
    btnRestaurer.Visible = False
    Toolbar1.Buttons(7).Visible = False
    cboArchiveRestore.Visible = False
  End If

End Sub

'##ModelId=5C8A67AA0157
Private Sub RefreshListe()
  Dim rq As String
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  If archiveMode Then
    Me.Caption = "Attention - vous consultez actuellement la base Archives !"
  Else
    Me.Caption = "Périodes du Groupe '" & m_dataHelper.GetParameter("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle) & "'"
  End If
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = True
  sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  
  DoEvents
  
  '& "P.SelForArchRest as [Sélection] " _

  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT P.RECNO, P.PENUMCLE as [Numéro Période], " _
        & "CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + TP.Libelle as [Type], " _
        & "CAST(P.IdTypeCalcul as VARCHAR) + ' - ' + TC.Libelle as [Type Calcul], " _
        & "P.PEDATEDEB as [Début], " _
        & "P.PEDATEFIN as [Fin], " _
        & "P.PEDATEEXT as [Date Arrêté], " _
        & "P.PECOMMENTAIRE as Commentaire," _
        & "P.PELOCKED as [Période Verrouillée], " _
        & "P.StatutArchRest as [Statut], " _
        & "P.DateArchRest as [Date Archivage ou Restauration] " _
                & "FROM P3IUser.Periode P LEFT JOIN P3IUser.TypePeriode TP ON TP.IdTypePeriode=P.PETYPEPERIODE" _
        & "    LEFT JOIN P3IUser.TypeCalcul TC ON TC.IdTypeCalcul=P.IdTypeCalcul " _
        & "WHERE P.PEGPECLE = " & GroupeCle
        
  If archiveMode Then
    rq = rq & " AND P.StatutArchRest = 'Archivée'"
  Else
    If cboTypePeriode.ListIndex > 0 Then
      rq = rq & " AND P.PETYPEPERIODE=" & cboTypePeriode.ItemData(cboTypePeriode.ListIndex)
    End If
    
    If cboArchiveRestore.ListIndex = 1 Then
      rq = rq & " AND P.StatutArchRest = 'Archivée'"
    End If
    If cboArchiveRestore.ListIndex = 2 Then
      rq = rq & " AND P.StatutArchRest = 'Restaurée'"
    End If
  End If
        
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
  sprListe.ColWidth(3) = 41
  sprListe.ColWidth(4) = 10
  sprListe.ColWidth(5) = 10
  sprListe.ColWidth(6) = 10
  sprListe.ColWidth(7) = 10
  sprListe.ColWidth(8) = 50
  sprListe.ColWidth(9) = 10
  sprListe.ColWidth(10) = 9 '12 Statut
  sprListe.ColWidth(11) = 13 '12 Date
  'sprListe.ColWidth(12) = 9 '10 Sélection
  
  
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = 7
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = 3
  sprListe.Col2 = 3
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.Col = 8
  sprListe.Col2 = 8
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.Col = 9
  sprListe.Col2 = 9
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = 10
  sprListe.Col2 = 10
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = 11
  sprListe.Col2 = 11
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.BlockMode = False
  
  
  'manually add a column to spread: Selection Checkbox
  'Me.Width = 24000
  sprListe.ActiveCellHighlightStyle = ActiveCellHighlightStyleOff 'switch off rectangle around highlighted cell
  
  On Error Resume Next
     
  sprListe.OperationMode = OperationModeNormal ' OperationModeSingle ' OperationModeNormal
  sprListe.EditMode = True
  sprListe.Enabled = True
  
  'Add a checkbox column
  sprListe.MaxCols = 12
  sprListe.Col = 12
  sprListe.Row = 0
  sprListe.ColWidth(12) = 8
  sprListe.text = "Sélection"
  
  sprListe.Row = -1
  sprListe.BlockMode = False
  
  sprListe.CellType = CellTypeCheckBox
  sprListe.TypeCheckCenter = True
  sprListe.TypeCheckType = TypeCheckTypeNormal
  sprListe.text = 0
  
  'Unblock checkboxes on Col12
  sprListe.BlockMode = True
  sprListe.Col = 12
  sprListe.Row = -1
  sprListe.Protect = False
  sprListe.Lock = False
  sprListe.BlockMode = False
  
  'Block checkboxes on Col9
  sprListe.BlockMode = True
  sprListe.Col = 9
  sprListe.Row = -1
  sprListe.Protect = True
  sprListe.Lock = True
  sprListe.BlockMode = False
    
    
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True

  Me.SetFocus
  sprListe.SetFocus
  
End Sub


'##ModelId=5C8A67AA0167
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = 1 Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar1.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 250
  
  lblStatus.Left = 0
  lblStatus.Width = 8300
  progArchive.Left = lblStatus.Left + lblStatus.Width + 50
 
  If Me.Width > 9 * btnWidth Then
'    topbtn = Me.Height - Toolbar1.Height - btnHeight
    If progArchive.Visible Then
      topbtn = Me.ScaleHeight - btnHeight - btnHeight - 50
    Else
      topbtn = Me.ScaleHeight - btnHeight
    End If
    
    ' boutton 'nouvelle'
    PlacePremierBoutton btnNew, topbtn
    
    ' boutton 'Consulter'
    PlaceBoutton btnEdit, btnNew, topbtn
    
    ' boutton 'Choix des Editions'
    PlaceBoutton btnEdition, btnEdit, topbtn
    
    ' boutton 'Imprimer'
    PlaceBoutton btnPrint, btnEdition, topbtn
    
    ' boutton 'Supprimer'
    PlaceBoutton btnDel, btnPrint, topbtn
    
    ' boutton 'Exporter'
    PlaceBoutton btnExport, btnDel, topbtn
    
    If DatabaseFileNameArchive <> "" Then
      ' boutton 'Archiver'
      PlaceBoutton btnArchiver, btnExport, topbtn
      PlaceBoutton btnStop, btnExport, topbtn
      
      ' boutton 'Restaurer'
      PlaceBoutton btnRestaurer, btnArchiver, topbtn
      
      ' boutton 'Fermer'
      PlaceBoutton btnClose, btnRestaurer, topbtn
    Else
      ' boutton 'Fermer'
      PlaceBoutton btnClose, btnExport, topbtn
    End If
    
    lblStatus.top = topbtn + btnHeight + 50
    
  Else
'    topbtn = Me.Height - Toolbar1.Height - btnHeight
    
    If progArchive.Visible Then
      topbtn = Me.ScaleHeight - 3 * btnHeight - 50
    Else
      topbtn = Me.ScaleHeight - 2 * btnHeight
    End If
    
    ' boutton 'nouvelle'
    PlacePremierBoutton btnNew, topbtn
    
    ' boutton 'Consulter'
    PlaceBoutton btnEdit, btnNew, topbtn
    
    ' boutton 'Choix des Editions'
    PlaceBoutton btnEdition, btnEdit, topbtn
    
    ' boutton 'Imprimer'
    'PlacePremierBoutton btnPrint, topbtn + btnHeight + 30
    PlaceBoutton btnPrint, btnEdition, topbtn
        
    ' boutton 'Supprimer'
    'PlaceBoutton btnDel, btnPrint, topbtn + btnHeight + 30
    PlaceBoutton btnDel, btnPrint, topbtn
    
    ' boutton 'Exporter'
    'PlaceBoutton btnExport, btnDel, topbtn + btnHeight + 30
    PlacePremierBoutton btnExport, topbtn + btnHeight + 30
    
    If DatabaseFileNameArchive <> "" Then
      ' boutton 'Archiver'
      PlaceBoutton btnArchiver, btnExport, topbtn + btnHeight + 30
      PlaceBoutton btnStop, btnExport, topbtn + btnHeight + 30
      
      ' boutton 'Restaurer'
      PlaceBoutton btnRestaurer, btnArchiver, topbtn + btnHeight + 30
      
      ' boutton 'Fermer'
      PlaceBoutton btnClose, btnRestaurer, topbtn + btnHeight + 30
    Else
      ' boutton 'Fermer'
      PlaceBoutton btnClose, btnExport, topbtn + btnHeight + 30
    End If
    
    lblStatus.top = topbtn + btnHeight + btnHeight + 30
    
  End If
  
  progArchive.top = lblStatus.top
  
  ' liste
  sprListe.Height = Maximum(topbtn - btnHeight - 100, 0)
  
End Sub

'##ModelId=5C8A67AA0177
Private Sub Form_Unload(Cancel As Integer)
  If DroitAdmin Then
    frmMain.mnuAnnexes.Enabled = True
  End If
End Sub

'****************************************************************
'******************************* Spread *************************
'****************************************************************

'##ModelId=5C8A67AA0196
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
      
      
    ' vitesse de l'export
'    If frmInExport = False And Row >= sprListe.VirtualCurTop And Row < sprListe.VirtualCurTop + 20 Then
'
'      If (Row Mod 9) = 0 Then
'        ' largeur des colonnes
'        For i = 2 To sprListe.MaxCols - 1
'          sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 5
'        Next i
'        sprListe.ColWidth(sprListe.MaxCols) = 50
'      End If
'
'    End If
    
    
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
  If dtaPeriode.Recordset.fields(Col - 1).Name = "Statut" Then
      
    sprListe.GetDataFillData archive, vbString
  
    If Len(archive) > 0 Then
      If Left$(LCase(archive), 4) = "arch" Then
        sprListe.BlockMode = True
        sprListe.Col = -1
        sprListe.Row = Row
        sprListe.Col2 = -1
        sprListe.Row2 = Row
        sprListe.BackColor = LTRED
        
        sprListe.ForeColor = noir
          
        sprListe.BlockMode = False
      End If
    End If
  
  End If
  
End Sub

'##ModelId=5C8A67AA01F4
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
    'sprListe.Row = Row
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
    'sprListe.BackColor = noir
    sprListe.ForeColor = noir
    
    sprListe.Row = NewRow
    sprListe.BackColor = noir
    sprListe.ForeColor = blanc
    
    
End Sub


'##ModelId=5C8A67AA0242
Private Sub SetNumPeriode()
  numPeriode = 0
  
  If sprListe.ActiveRow < 0 Then Exit Sub
  
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 2
  
  numPeriode = CLng(sprListe.text)
End Sub

'##ModelId=5C8A67AA0261
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  sprListe.sheet = sprListe.ActiveSheet
  sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
End Sub

'##ModelId=5C8A67AA0280
Private Sub btnClose_Click()
  fInitDone = False
  
  Unload Me
End Sub

'##ModelId=5C8A67AA029F
Private Sub btnDel_Click_DELETESELECTED()
  
  Dim rs As ADODB.Recordset
  Dim i As Integer
  Dim numbItemsChecked As Integer
  Dim checkboxSelected As Boolean
  Dim numPeriode As Long
  Dim currentPeriode As Long
  Dim colPeriods As New Collection
  Dim periodCount As Long
  
  If Not DroitAdmin Then
    btnDel.Enabled = False
    Exit Sub
  End If
  
  'Call SetNumPeriode
  
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 12
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    
    If checkboxSelected Then
      numPeriode = CInt(sprListe.text)
      numbItemsChecked = numbItemsChecked + 1
      colPeriods.Add (numPeriode)
    End If
  Next i
  
  If archiveMode Then
    'verify if the periode still exists in the Prod DB - the user can delete a Periode in ArchiveMode
    'only once it has been deleted from the prod DB
    
    If m_dataHelper.PeriodeExists(numPeriode, GroupeCle) Then
      MsgBox "Avant de supprimer la période dans la base Archives vous devrait la supprimer dabord dans la base de Production !", vbCritical
      Exit Sub
    End If
    
  End If
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  Set rs = m_dataSource.OpenRecordset("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle, Snapshot)
  
  If Not rs.EOF Then
    DescriptionPeriode = "Période n°" & numPeriode & " du Groupe " & rs.fields("Nom")
  Else
    DescriptionPeriode = "Erreur ... "
  End If
  rs.Close
      
  
  periodCount = colPeriods.Count
  
  If MsgBox("Tous les ASSURES des périodes sélectionnées vont être supprimés." & vbLf & "Voulez-vous vraiment continuer ?", vbQuestion Or vbYesNo) = vbYes Then
   
    Screen.MousePointer = vbHourglass
      
    For i = 1 To periodCount
      DoEvents
      'currentPeriode = colPeriods(i)
      numPeriode = colPeriods(i)
      
      'If MsgBox("Tous les ASSURES de cette période vont être supprimés." & vbLf & "Voulez-vous vraiment supprimer la " & DescriptionPeriode, vbQuestion Or vbYesNo) = vbYes Then
            
      If archiveMode Then
      
        If CreateArchiveConnection Then
          
          On Error GoTo errSupPeriodeArchive
      
          m_dataSourceArchive.BeginTrans
         
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM TBQREGA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CodesCat WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CodeCatInv WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Correspondance_CatOption WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Capitaux_Moyens WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CATR9 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CATR9INVAL WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM PassageNCA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM PM_Retenue WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Reassurance WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CDSITUAT WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ParamRentes WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AgeDepartRetraite WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AgeDepartRetraiteInval1 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CoeffAmortissement WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
          
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Editions WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionsTemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
          
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & numPeriode)
          
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_RetraiteTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_Retraite WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AssureTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
          m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
          
          m_dataSourceArchive.CommitTrans
          
          CloseArchiveConnection
        Else
          'problem creating connection to Archive DB
          CloseArchiveConnection
          MsgBox "Impossible d'ouvrir la base de données Archive!" & vbLf & "Source: frmPeriode.btnDel_Click" _
          & vbLf & "Connection : " & DatabaseFileNameArchive, vbCritical
        End If
        
      Else 'delete from Production
        
        On Error GoTo errSupPeriode
        
        'If CBool(m_dataHelper.GetParameter("SELECT PELOCKED FROM Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle)) = True Then
        '  MsgBox "Cette période est vérrouillée et ne peut pas être supprimée!", vbCritical
        '  Exit Sub
        'End If
        
        m_dataSource.BeginTrans
        
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM TBQREGA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        'm_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM REGA01 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & NumPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CodesCat WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CodeCatInv WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Correspondance_CatOption WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        'm_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Produits WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & NumPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Capitaux_Moyens WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CATR9 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CATR9INVAL WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM PassageNCA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM PM_Retenue WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Reassurance WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CDSITUAT WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ParamRentes WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AgeDepartRetraite WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AgeDepartRetraiteInval1 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CoeffAmortissement WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Editions WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionsTemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & numPeriode)
        
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_RetraiteTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_Retraite WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AssureTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
        m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
        
        m_dataSource.CommitTrans
              
      End If
      
    Next i
    
    RefreshListe
    Screen.MousePointer = vbDefault
    CloseArchiveConnection
  End If
  
  On Error GoTo 0
  Exit Sub
  
errSupPeriode:
  Screen.MousePointer = vbDefault
  MsgBox "impossible de supprimer la période" & vbLf & Err.Description
  Resume Next
  
errSupPeriodeArchive:
  m_dataSourceArchive.RollbackTrans
  CloseArchiveConnection
  Screen.MousePointer = vbDefault
  MsgBox "impossible de supprimer la période dans la base Archives" & vbLf & Err.Description
  
End Sub

'##ModelId=5C8A67AA02BF
Private Sub btnDel_Click()
  
  Dim rs As ADODB.Recordset
  
  If Not DroitAdmin Then
    btnDel.Enabled = False
    Exit Sub
  End If
  
  Call SetNumPeriode
  
  If archiveMode Then
    'verify if the periode still exists in the Prod DB - the user can delete a Periode in ArchiveMode
    'only once it has been deleted from the prod DB
    
    If m_dataHelper.PeriodeExists(numPeriode, GroupeCle) Then
      MsgBox "Avant de supprimer la période dans la base Archives vous devrait la supprimer dabord dans la base de Production !", vbCritical
      Exit Sub
    End If
    
  End If
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  Set rs = m_dataSource.OpenRecordset("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle, Snapshot)
  
  If Not rs.EOF Then
    DescriptionPeriode = "Période n°" & numPeriode & " du Groupe " & rs.fields("Nom")
  Else
    DescriptionPeriode = "Erreur ... "
  End If
  rs.Close
    
    
  If MsgBox("Tous les ASSURES de cette période vont être supprimés." & vbLf & "Voulez-vous vraiment supprimer la " & DescriptionPeriode, vbQuestion Or vbYesNo) = vbYes Then
    
    Screen.MousePointer = vbHourglass
    
    If archiveMode Then
    
      If CreateArchiveConnection Then
        
        On Error GoTo errSupPeriodeArchive
    
        m_dataSourceArchive.BeginTrans
       
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM TBQREGA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CodesCat WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CodeCatInv WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Correspondance_CatOption WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Capitaux_Moyens WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CATR9 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CATR9INVAL WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM PassageNCA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM PM_Retenue WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Reassurance WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CDSITUAT WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ParamRentes WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AgeDepartRetraite WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AgeDepartRetraiteInval1 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM CoeffAmortissement WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
        
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Editions WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionsTemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionIIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
        
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & numPeriode)
        
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_RetraiteTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure_Retraite WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM AssureTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Assure WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
        m_dataSourceArchive.Execute m_dataHelperArchive.ValidateSQL("DELETE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
        
        m_dataSourceArchive.CommitTrans
        
        CloseArchiveConnection
      Else
        'problem creating connection to Archive DB
        CloseArchiveConnection
        MsgBox "Impossible d'ouvrir la base de données Archive!" & vbLf & "Source: frmPeriode.btnDel_Click" _
        & vbLf & "Connection : " & DatabaseFileNameArchive, vbCritical
      End If
      
    Else 'delete from Production
      
      On Error GoTo errSupPeriode
      
      If CBool(m_dataHelper.GetParameter("SELECT PELOCKED FROM Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle)) = True Then
        MsgBox "Cette période est vérrouillée et ne peut pas être supprimée!", vbCritical
        Exit Sub
      End If
      
      m_dataSource.BeginTrans
      
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM TBQREGA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      'm_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM REGA01 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & NumPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CodesCat WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CodeCatInv WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Correspondance_CatOption WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      'm_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Produits WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & NumPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Capitaux_Moyens WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CATR9 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CATR9INVAL WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM PassageNCA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM PM_Retenue WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Reassurance WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CDSITUAT WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ParamRentes WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AgeDepartRetraite WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AgeDepartRetraiteInval1 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM CoeffAmortissement WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode)
      
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Editions WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionsTemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIII WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionIIITemp WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & numPeriode)
      
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & numPeriode)
      
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_RetraiteTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_Retraite WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AssureTemp WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure WHERE POGPECLE = " & GroupeCle & " AND POPERCLE = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
      m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode)
      
      m_dataSource.CommitTrans
            
    End If
    
    RefreshListe
    Screen.MousePointer = vbDefault
    CloseArchiveConnection
  End If
  
  On Error GoTo 0
  Exit Sub
  
errSupPeriode:
  Screen.MousePointer = vbDefault
  MsgBox "impossible de supprimer la période" & vbLf & Err.Description
  Resume Next
  
errSupPeriodeArchive:
  m_dataSourceArchive.RollbackTrans
  CloseArchiveConnection
  Screen.MousePointer = vbDefault
  MsgBox "impossible de supprimer la période dans la base Archives" & vbLf & Err.Description
  
End Sub

'##ModelId=5C8A67AA02DE
Private Sub btnEdit_Click()
  
  Dim frm As New frmEditPeriode
  Dim rs As ADODB.Recordset
  
  SetNumPeriode
  
  If numPeriode = 0 Then Exit Sub
  
  'lock this periode for the current user
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM LockedPeriods WHERE Periode = " & numPeriode, Snapshot)
  
  If rs.RecordCount > 0 Then
    If user_name <> rs.fields("UserName") Then
'      MsgBox "La période que vous avez sélectionné est actuellement en traitement par l'utilisateur : " & rs.fields("UserName") & " et ne peut pas être consulté." _
'      & vbLf & vbLf & "Pour consulter la période sélectionné attendez jusqu'à l'utilisateur a fini le traitement de cette période ou demandez à l'utilisateur de fermer la fenêtre ""Assurés de la Période " & NumPeriode & """ !"
      
      MsgBox "La période " & numPeriode & "  est actuellement utilisée par " & rs.fields("UserName") _
      & vbLf & vbLf & "Demandez à l'utilisateur de quitter cette période. "

      rs.Close
      Exit Sub
    End If
  Else
    m_dataSource.Execute "Insert Into LockedPeriods (Periode, UserName) Values ('" & numPeriode & "', '" & user_name & "')"
  End If
  
  rs.Close
    
  frm.Show
End Sub

'##ModelId=5C8A67AA02EE
Private Sub btnEdition_Click()
  Dim FM As frmChoixEdition, fl As clsFilter
  
  Call SetNumPeriode
  
  If numPeriode = 0 Then Exit Sub
  
  Dim rs As ADODB.Recordset
  SoCle = 0
  frmWait.Show vbModeless
  
  frmWait.Caption = "Chargement en cours en cours..."
  
  frmWait.ProgressBar1.Min = 0
  frmWait.ProgressBar1.Value = 0
  frmWait.ProgressBar1.Max = 1
    
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  Set rs = m_dataSource.OpenRecordset("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle, Snapshot)
  
  If Not rs.EOF Then
    Dim dd As String, df As String
    
    dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")
    df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")

    DescriptionPeriode = " Période " & numPeriode & " ( " & dd & " au " & df & " ) " & vbLf & " du Groupe " & rs.fields("Nom")
  Else
    DescriptionPeriode = "  Erreur ... "
  End If
  
  rs.Close
  
  ' no filter from this button
  Set FM = New frmChoixEdition
  Set fl = New clsFilter
  
  Set FM.fmFilter = fl
  FM.frmNumPeriode = numPeriode
  
  FM.Show vbModal
  
  Set FM = Nothing
  Set fl = Nothing
End Sub

'##ModelId=5C8A67AA030D
Private Sub btnExport_Click()
  On Error GoTo err_export
  
  SetNumPeriode
  
  If numPeriode = 0 Then Exit Sub
  
  CommonDialog1.filename = "P3I Periode " & numPeriode & ".mdb"
  CommonDialog1.filter = "Fichier Excel|*.xls|Base de données MS Access|*.mdb|"
  CommonDialog1.FilterIndex = 2
  
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
  CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
  
  CommonDialog1.CancelError = True
  
  CommonDialog1.ShowSave
  
  CommonDialog1.CancelError = False
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.mdb" Or CommonDialog1.filename = "*.xls" Then
    Exit Sub
  End If
  
  If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Then
    ExportTableToExcelFile "Assure periode " & numPeriode & ".xls", _
                           "Periode " & numPeriode, _
                           "Assure", sprListe, CommonDialog1, "", False
  Else
    Dim exportModule As P3IExport.iExport
    
    Set exportModule = New P3IExport.iExport
    
    exportModule.ExportDBAccess CommonDialog1, m_dataSource, GroupeCle, numPeriode
    
    Set exportModule = Nothing
  End If
  
  Exit Sub
  
err_export:
  CommonDialog1.CancelError = False
End Sub

'##ModelId=5C8A67AA031C
Private Sub btnNew_Click()
  numPeriode = -1
  
  frmDetailPeriode.Show vbModal
  
  RefreshListe
End Sub

'##ModelId=5C8A67AA032C
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
    .PrintJobName = "Périodes du groupe " & NomGroupe
    .PrintHeader = "/c Périodes du groupe " & NomGroupe & "/n  "
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

'##ModelId=5C8A67AA033C
Private Sub btnStop_Click()
    stopArchiving = True
End Sub

'##ModelId=5C8A67AA034B
Private Sub cboArchiveRestore_Click()
  If fInitDone = True Then RefreshListe
End Sub

'##ModelId=5C8A67AA035B
Private Sub cboTypePeriode_Click()
  If fInitDone = True Then RefreshListe
End Sub

'##ModelId=5C8A67AA036B
Private Sub rdoTout_Click()
  RefreshListe
End Sub

'##ModelId=5C8A67AA037A
Private Sub rdoProvision_Click()
  RefreshListe
End Sub

'##ModelId=5C8A67AA038A
Private Sub rdoCapital_Click()
  RefreshListe
End Sub

'##ModelId=5C8A67AA03A9
Private Sub sprListe_Click(ByVal Col As Long, ByVal Row As Long)
    
    'sprListe.Col = -1
    'sprListe.Row = Row
    'sprListe.BackColor = bleu
    
End Sub

'##ModelId=5C8A67AA03D8
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim r As Long, tr As Long
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
  
  SetNumPeriode
  frmDetailPeriode.Show vbModal
  
  Screen.MousePointer = vbHourglass
  
  RefreshListe
  
  sprListe.TopRow = tr
  sprListe.SetActiveCell 2, r
  
  sprListe.Row = r
  sprListe.SelModeSelected = True
  
  sprListe.SetFocus
  
  Screen.MousePointer = vbDefault
  
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie d'une table en changeant son n° de période
'
'##ModelId=5C8A67AB002E
Private Sub CopieTableParam(nomTable As String, idPeriodeDest As Long, NomChampClePeriode As String, NomChampCleGroupe As String)
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset, f As ADODB.field

  Set rs = m_dataSource.OpenRecordset("SELECT * FROM " & nomTable & " WHERE " & NomChampCleGroupe & "=" & GroupeCle & " AND " & NomChampClePeriode & "=" & numPeriode, Snapshot)
  Set rs2 = m_dataSource.OpenRecordset(nomTable, table)
  
  Do Until rs.EOF
    rs2.AddNew
    
    For Each f In rs.fields
      Select Case f.Name
        Case "RECNO"
        
        'Case "PEP3I_INDIVIDUEL"
        '  rs2.fields("PEP3I_INDIVIDUEL") = False
  
        Case "PELOCKED"
          rs2.fields("PELOCKED") = False
          
        Case NomChampClePeriode
          rs2.fields(NomChampClePeriode) = idPeriodeDest
          
        Case Else
          rs2.fields(f.Name).Value = f.Value
      End Select
    Next f
    
    rs2.Update
    
    rs.MoveNext
  Loop
  
  rs2.Close
  rs.Close

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie d'une période sans les données
'
'##ModelId=5C8A67AB007D
Private Sub CopyPeriode()
  SetNumPeriode
    
  If MsgBox("Voulez-vous vraiment dupliquer la période n°" & numPeriode & " ?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
  End If
    
  Screen.MousePointer = vbHourglass
  
  sprListe.Row = sprListe.ActiveRow
  
  ' récupère le nouveau n° d'étude
  Dim idPeriode As Long
  
  idPeriode = m_dataHelper.GetParameterAsDouble("SELECT max(PENUMCLE) FROM Periode WHERE PEGPECLE=" & GroupeCle)
  idPeriode = idPeriode + 1
  
  
  ' copie la période
  CopieTableParam "Periode", idPeriode, "PENUMCLE", "PEGPECLE"
  
  ' copie les params
  CopieTableParam "ParamCalcul", idPeriode, "PENUMCLE", "PEGPECLE"

  ' Copie des tables de paramétrage
  CopieTableParam "TBQREGA", idPeriode, "NumPeriode", "GroupeCle"
'  CopieTableParam "REGA01", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CodesCat", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CodeCatInv", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "Correspondance_CatOption", idPeriode, "NumPeriode", "GroupeCle"
'  CopieTableParam "Produits", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "Capitaux_Moyens", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CATR9", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CATR9INVAL", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "PassageNCA", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "PM_Retenue", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "Reassurance", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CDSITUAT", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "ParamRentes", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "AgeDepartRetraite", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "AgeDepartRetraiteInval1", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "CoeffAmortissement", idPeriode, "NumPeriode", "GroupeCle"
  CopieTableParam "DonneesSociales", idPeriode, "NumPeriode", "GroupeCle"
  
 
  ' affiche la nouvelle période
  RefreshListe

  Screen.MousePointer = vbDefault

End Sub

'##ModelId=5C8A67AB009C
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  ' Utilise la propriété Key avec l'instruction SelectCase pour   spécifier une action.
  Select Case Button.key
    Case "copyPeriode"
      If sprListe.MaxRows > 0 Then
        CopyPeriode
      End If
    
    Case "print"
      btnPrint_Click
      
    Case "openPeriode"
      SetNumPeriode
      frmDetailPeriode.Show vbModal, frmMain
      
    Case "consultArchive"
      'If Button.Image = 8 Then 'back to normal mode
      If Button.ToolTipText = "Quitter la base Archives" Then 'back to normal mode
        archiveMode = False
        'Button.Image = 7
        Button.ToolTipText = "Consulter la base Archives"
        dtaPeriode.ConnectionString = DatabaseFileName
        Me.BackColor = &H8000000F
        
        'enable controls
        Toolbar1.Buttons(3).Enabled = True
        cboTypePeriode.Enabled = True
        cboArchiveRestore.Enabled = True
        
        EnableButtons "btnEdition"
        dtaPeriode.Refresh
    
      Else    'switch to archive mode
        archiveMode = True
        'Button.Image = 8
        Button.ToolTipText = "Quitter la base Archives"
        dtaPeriode.ConnectionString = DatabaseFileNameArchive
        Me.BackColor = &HFF&
        
        'disable controls
        Toolbar1.Buttons(3).Enabled = False
        cboTypePeriode.Enabled = False
        cboArchiveRestore.Enabled = False
        
        DisableButtons "btnDel, btnRestaurer, btnConsultArchive, btnClose, btnEdit, btnClose"
        dtaPeriode.Refresh
      End If
      
  End Select
  
  RefreshListe
End Sub

'##ModelId=5C8A67AB00BB
Private Sub DisableButtons(exceptions As String)
   
  Dim ctrl As Control
  
  On Error Resume Next
  
  For Each ctrl In frmPeriode.Controls
    If TypeOf ctrl Is CommandButton Then
      'Debug.Print ctrl.Name
      If InStr(exceptions, ctrl.Name) = 0 Then
        ctrl.Enabled = False
      End If
    End If
  Next ctrl
  
End Sub

'##ModelId=5C8A67AB00DA
Private Sub EnableButtons(exceptions As String)
   
  Dim ctrl As Control
  
  On Error Resume Next
  
  For Each ctrl In frmPeriode.Controls
    If TypeOf ctrl Is CommandButton Then
      'Debug.Print ctrl.Name
      If InStr(exceptions, ctrl.Name) = 0 Then
        ctrl.Enabled = True
      End If
    End If
  Next ctrl
  
  If DroitAdmin Then
    btnDel.Enabled = True
  Else
    btnDel.Enabled = False
  End If

End Sub


'****************************************************************
'************************* Archive - Restore ********************
'****************************************************************


'##ModelId=5C8A67AB00FA
Private Sub btnRestaurer_Click()

  Dim connStrProd As String
  Dim connStrArch As String
      
  Dim archRes As ArchiveRestore
  
  Dim i As Integer
  Dim numbItemsChecked As Integer
  Dim numbIgnore As Integer
  Dim numbToRestore As Integer
  Dim checkboxSelected As Boolean
  Dim status As String
  Dim numPeriode As Integer
  Dim itemList As String
  Dim colIgnore As New Collection
  Dim colToRestore As New Collection
  Dim restoreSuccess As Boolean
  Dim statusMessage As String
  
  'get connection strings from .INI
  restoreSuccess = False
  connStrProd = DatabaseFileName
  connStrArch = DatabaseFileNameArchive
  
  numbItemsChecked = 0
  numbIgnore = 0
  numbToRestore = 0
  
  sprListe.VirtualMode = False
  sprListe.DataRefresh
  sprListe.Refresh
   
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 12
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    numPeriode = CInt(sprListe.text)
    sprListe.Col = 10
    status = sprListe.text
    
    If checkboxSelected Then
        numbItemsChecked = numbItemsChecked + 1
        If Left$(LCase(status), 4) <> "arch" Then
            colIgnore.Add (numPeriode)
            numbIgnore = numbIgnore + 1
        Else
            colToRestore.Add (numPeriode)
            numbToRestore = numbToRestore + 1
        End If
    End If
  Next i
  
  If numbItemsChecked = 0 Then
     MsgBox "Aucune période a été sélectionnée ! Sélectionnez au moins une période et cliquez le bouton Restaurer.", vbExclamation
     GoTo Cleanup
  End If
  
  If numbToRestore = 0 And numbIgnore = 1 Then
     MsgBox "La période sélectionnée n'est pas une période archivée et ne peut pas être restaurée !", vbExclamation
     GoTo Cleanup
  End If
  
  If numbToRestore = 0 And numbIgnore > 1 Then
     MsgBox "Les périodes sélectionnées ne sont pas des périodes archivées et ne peuvent pas être restaurées !", vbExclamation
     GoTo Cleanup
  End If
  
  If numbIgnore = 1 Then
     MsgBox "La période " & colIgnore(1) & " n'est pas une période archivée et ne peut pas être restaurée !", vbExclamation
  End If
  
  If numbIgnore > 1 Then
     For i = 1 To colIgnore.Count
        itemList = itemList & colIgnore(i) & ", "
     Next i
     itemList = Trim$(itemList)
     itemList = Left$(itemList, Len(itemList) - 1)
  
     MsgBox "Les périodes suivantes ne sont pas des périodes archivées et ne peuvent pas être restaurées : " & itemList, vbExclamation
  End If
  
  If MsgBox("Est-ce que vous est sur de vouloir restaurer toutes les périodes sélectionnées ?", vbYesNo, "Confirmation") = vbYes Then
     'Start restoring
          
     Screen.MousePointer = vbHourglass
     
     btnArchiver.Visible = False
     btnStop.Visible = True
     btnStop.Enabled = True
     Toolbar1.Enabled = False
     cboTypePeriode.Enabled = False
     cboArchiveRestore.Enabled = False
     
     lblStatus.Visible = True
     progArchive.Visible = True
     Form_Resize
     
     progArchive.Min = 0
     progArchive.Max = colToRestore.Count + 1
     progArchive.Value = progArchive.Min + 1
     lblStatus.Caption = "Restauration en cours..."
     
     DisableButtons "btnStop"
     
     Set archRes = New ArchiveRestore
     
     For i = 1 To colToRestore.Count
      DoEvents
       
      statusMessage = "Nombre des périodes déjà restaurées : " & i & "/" & colToRestore.Count & " - période " & colToRestore(i) & " en cours de restauration."
         
      restoreSuccess = archRes.RestorePeriode(connStrProd, connStrArch, colToRestore(i), statusMessage)
        
      If Not restoreSuccess Then
          If i < colToRestore.Count Then
            If MsgBox("Il y a eu un problème avec la restauration du période " & colToRestore(i) & vbLf & _
            " Voulez-vous continuer avec la restauration des périodes restant ?", vbYesNo, "Confirmation") = vbNo Then
                GoTo Cleanup
            End If
          Else
            MsgBox "Il y a eu un problème avec la restauration du période " & colToRestore(i) & " !", vbCritical
          End If
      End If
            

      If stopArchiving Then
          If MsgBox("Est-ce que vous est sur de vouloir arrêter la restauration des périodes restant ?", vbYesNo, "Confirmation") = vbYes Then
              stopArchiving = False
              GoTo Cleanup
          Else
              stopArchiving = False
          End If
      End If
      
      progArchive.Value = i + 1
      lblStatus.Caption = "Nombre des périodes déjà restaurées : " & i & "/" & colToRestore.Count & " - période " & colToRestore(i) & " en cours de restauration."
            
     Next i
        
  Else
    GoTo Cleanup
  End If
  
  RefreshListe
  Screen.MousePointer = vbDefault
  'MsgBox "The restore process has finished successfully...", vbOKOnly, "Finished"
  
Cleanup:

  Set archRes = Nothing
  
  sprListe.VirtualMode = True
  sprListe.DataRefresh
  sprListe.Refresh
  
  btnArchiver.Visible = True
  btnStop.Visible = False
  btnStop.Enabled = False
  Toolbar1.Enabled = True
  cboTypePeriode.Enabled = True
  cboArchiveRestore.Enabled = True
  
  EnableButtons "btnStop"
  
  lblStatus.Visible = False
  progArchive.Visible = False
  progArchive.Value = progArchive.Min
  lblStatus.Caption = ""
  RefreshListe
  
  Form_Resize
  
  Screen.MousePointer = vbDefault
  
End Sub

'##ModelId=5C8A67AB0119
Private Sub btnArchiver_Click()
  
  Dim connStrProd As String
  Dim connStrArch As String
      
  Dim archRes As ArchiveRestore
  
  Dim i As Integer
  Dim numbItemsChecked As Integer
  Dim numbIgnore As Integer
  Dim numbToArchive As Integer
  Dim checkboxSelected As Boolean
  Dim status As String
  Dim numPeriode As Integer
  Dim itemList As String
  Dim colIgnore As New Collection
  Dim colToArchive As New Collection
  Dim archiveSuccess As Boolean
  Dim statusMessage As String
  
  archiveSuccess = False
  connStrProd = DatabaseFileName
  connStrArch = DatabaseFileNameArchive
  
  numbItemsChecked = 0
  numbIgnore = 0
  numbToArchive = 0
  
  sprListe.VirtualMode = False
  sprListe.DataRefresh
  sprListe.Refresh
   
  For i = 1 To sprListe.DataRowCnt
    sprListe.Row = i
    sprListe.Col = 12
    checkboxSelected = CBool(sprListe.text)
    sprListe.Col = 2
    numPeriode = CInt(sprListe.text)
    sprListe.Col = 10
    status = sprListe.text
    
    If checkboxSelected Then
        numbItemsChecked = numbItemsChecked + 1
        If Left$(LCase(status), 4) = "arch" Then
            colIgnore.Add (numPeriode)
            numbIgnore = numbIgnore + 1
        Else
            colToArchive.Add (numPeriode)
            numbToArchive = numbToArchive + 1
        End If
    End If
  Next i
  
  If numbItemsChecked = 0 Then
     MsgBox "Aucune période a été sélectionnée ! Sélectionnez au moins une période et cliquez le bouton Restaurer.", vbExclamation
     GoTo Cleanup
  End If
  
  If numbToArchive = 0 And numbIgnore = 1 Then
     MsgBox "La période sélectionnée est déjà archivée et sera donc ignorée !", vbExclamation
     GoTo Cleanup
  End If
  
  If numbToArchive = 0 And numbIgnore > 1 Then
     MsgBox "Les périodes sélectionnées sont déjà archivées et seront donc ignorées !", vbExclamation
     GoTo Cleanup
  End If
  
  If numbIgnore = 1 Then
     MsgBox "La période " & colIgnore(1) & " été déjà archivée et sera donc ignorée !", vbExclamation
  End If
  
  If numbIgnore > 1 Then
     For i = 1 To colIgnore.Count
        itemList = itemList & colIgnore(i) & ", "
     Next i
     itemList = Trim$(itemList)
     itemList = Left$(itemList, Len(itemList) - 1)
  
     MsgBox "Les périodes suivantes ont été déjà archivées et seront donc ignorées : " & itemList, vbExclamation
  End If
  
  If MsgBox("Est-ce que vous est sur de vouloir archiver toutes les périodes sélectionnées ?", vbYesNo) = vbYes Then
     'Start archiving
     'archive one element after another - handle statusbar
     
     'If CreateArchiveConnection Then
          
       Screen.MousePointer = vbHourglass
       
       btnArchiver.Visible = False
       btnStop.Visible = True
       btnStop.Enabled = True
       Toolbar1.Enabled = False
       cboTypePeriode.Enabled = False
       cboArchiveRestore.Enabled = False
       
       lblStatus.Visible = True
       progArchive.Visible = True
       Form_Resize
       
       progArchive.Min = 0
       progArchive.Max = colToArchive.Count + 1
       progArchive.Value = progArchive.Min + 1
       lblStatus.Caption = "Archivage en cours..."
       
       DisableButtons "btnStop"
       
       Set archRes = New ArchiveRestore
       
       For i = 1 To colToArchive.Count
        DoEvents
          
        statusMessage = "Nombre des périodes déjà archivées : " & i & "/" & colToArchive.Count & " - période " & colToArchive(i) & " en cours d'archivage."
              
        archiveSuccess = archRes.ArchivePeriode(connStrProd, connStrArch, colToArchive(i), statusMessage)
        
        If Not archiveSuccess Then
          If i < colToArchive.Count Then
            If MsgBox("Il y a eu un problème avec l'archivage de la période " & colToArchive(i) & vbLf & _
            "Voulez-vous continuer avec l'archivage des périodes restant ?", vbYesNo, "Confirmation") = vbNo Then
                GoTo Cleanup
            End If
          Else
            MsgBox "Il y a eu un problème avec l'archivage de la période " & colToArchive(i), vbCritical
          End If
        End If
  
        If stopArchiving Then
            If MsgBox("Est-ce que vous est sur de vouloir arrêter l'archivage des périodes restant ?", vbYesNo) = vbYes Then
                stopArchiving = False
                GoTo Cleanup
            Else
                stopArchiving = False
            End If
        End If
        
        progArchive.Value = i + 1
        lblStatus.Caption = "Nombre des périodes déjà archivées : " & i & "/" & colToArchive.Count & " - période " & colToArchive(i) & " en cours d'archivage."
              
       Next i
       
       'finished - close the connection
       'CloseArchiveConnection
     
     'There is no Archive Connection
'     Else
'        CloseArchiveConnection
'        GoTo Cleanup
'     End If
     
        
  Else
    GoTo Cleanup
  End If
  
  RefreshListe
  Screen.MousePointer = vbDefault
  'MsgBox "The archiving process has finished successfully...", vbOKOnly, "Finished"
  
Cleanup:

  Set archRes = Nothing
  CloseArchiveConnection
  
  sprListe.VirtualMode = True
  sprListe.DataRefresh
  sprListe.Refresh
  
  btnArchiver.Visible = True
  btnStop.Visible = False
  btnStop.Enabled = False
  Toolbar1.Enabled = True
  cboTypePeriode.Enabled = True
  cboArchiveRestore.Enabled = True
  
  EnableButtons "btnStop"
  
  lblStatus.Visible = False
  progArchive.Visible = False
  progArchive.Value = progArchive.Min
  lblStatus.Caption = ""
  RefreshListe
  
  'uncheck all checkboxes
'  sprListe.VirtualMode = False
'  sprListe.VirtualMaxRows = -1
'  sprListe.DataRefresh
'  sprListe.Refresh
'
'  For i = 1 To sprListe.DataRowCnt
'      sprListe.Row = i
'      sprListe.Col = 12
'      sprListe.text = 0
'  Next i
'
'  sprListe.VirtualMode = True
  
  Form_Resize
 
  Screen.MousePointer = vbDefault

End Sub

