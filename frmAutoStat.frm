VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAutoStat 
   Caption         =   "Auto"
   ClientHeight    =   11070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   19290
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   19290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Height          =   345
      Left            =   16200
      TabIndex        =   12
      Top             =   10440
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fermer"
      Height          =   345
      Left            =   17760
      TabIndex        =   11
      Top             =   10440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sélectionner les périodes destinataires"
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   18735
      Begin FPSpreadADO.fpSpread sprListe 
         Height          =   4965
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   18165
         _Version        =   524288
         _ExtentX        =   32041
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
         SpreadDesigner  =   "frmAutoStat.frx":0000
         ScrollBarTrack  =   3
         AppearanceStyle =   0
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   18735
      Begin VB.CommandButton cmdDelSexe 
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
         Index           =   0
         Left            =   10320
         TabIndex        =   10
         Top             =   300
         Width           =   330
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Sélectionner le fichier sexe"
         Height          =   330
         Index           =   0
         Left            =   7800
         TabIndex        =   8
         Top             =   300
         Width           =   2295
      End
      Begin VB.CommandButton cmdDel 
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
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   300
         Width           =   330
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Sélectionner la période statutaire"
         Height          =   330
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   300
         Width           =   2775
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3615
         Left            =   18360
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblFile 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   11280
         TabIndex        =   9
         Top             =   360
         Width           =   6615
      End
      Begin VB.Label lblPeriodeStat 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblPeriode 
         Caption         =   "Periode No. 1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
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
End
Attribute VB_Name = "frmAutoStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A684D013B"
Option Explicit

'##ModelId=5C8A684D0235
Public selectedPeriods As Collection

'##ModelId=5C8A684D0254
Private oldPos As Integer
'##ModelId=5C8A684D0264
Private numPeriode As Integer

'##ModelId=5C8A684D0283
Private Sub cmdClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A684D0293
Private Sub cmdDel_Click(Index As Integer)
  lblPeriodeStat(Index).Caption = "..."
End Sub

'##ModelId=5C8A684D02C2
Private Sub cmdDelSexe_Click(Index As Integer)
  lblFile(Index).Caption = "..."
End Sub

'##ModelId=5C8A684D02F0
Private Sub cmdFile_Click(Index As Integer)
  Dim fName As String

  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Sélectionner un fichier Excel qui contienne l'information concernant du sexe de l'assurée"
  'CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.filter = "Fichiers Excel|*.xls"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
  
  fName = CommonDialog1.filename
  
  If InStr(fName, ".xls") <= 0 Then
    MsgBox "Le fichier sélectionné n'est pas valable pour cette opération. S'il vous plait sélectionnez un fichier du type Excel !", vbOKOnly, "Mauvaise type du fichier !"
    lblFile(Index).Caption = "..."
    Exit Sub
  Else
    lblFile(Index).Caption = fName
  End If
End Sub

'##ModelId=5C8A684D0310
Private Sub cmdSelect_Click(Index As Integer)
  lblPeriodeStat(Index).Caption = "Periode No. " & numPeriode
End Sub

'##ModelId=5C8A684D033F
Private Sub Form_Load()

  Dim i As Integer
  
  numPeriode = 0
 
  For i = 0 To selectedPeriods.Count - 1
    If i <> 0 Then
      Load lblPeriode(i)
    End If
    
    lblPeriode(i).Caption = "Periode " & i + 1
    lblPeriode(i).top = lblPeriode(i - 1).top + lblPeriode(i).Height + 125
    lblPeriode(i).Visible = True
    
    Load lblPeriodeStat(i)
    lblPeriodeStat(i).Caption = "..."
    lblPeriodeStat(i).top = lblPeriodeStat(i - 1).top + lblPeriodeStat(i).Height + 125
    lblPeriodeStat(i).Visible = True
    
    Load lblFile(i)
    lblFile(i).Caption = "..."
    lblFile(i).top = lblFile(i - 1).top + lblFile(i).Height + 125
    lblFile(i).Visible = True
    
    Load cmdDel(i)
    cmdDel(i).Caption = "X"
    cmdDel(i).top = cmdDel(i - 1).top + cmdDel(i).Height + 50
    cmdDel(i).Visible = True
    
    Load cmdFile(i)
    'cmdFile(i).Caption = "X"
    cmdFile(i).top = cmdFile(i - 1).top + cmdFile(i).Height + 50
    cmdFile(i).Visible = True
    
    Load cmdSelect(i)
    'cmdSelect(i).Caption = "X"
    cmdSelect(i).top = cmdSelect(i - 1).top + cmdSelect(i).Height + 50
    cmdSelect(i).Visible = True
    
    Load cmdDelSexe(i)
    'cmdDelSexe(i).Caption = "X"
    cmdDelSexe(i).top = cmdDelSexe(i - 1).top + cmdDelSexe(i).Height + 50
    cmdDelSexe(i).Visible = True
  Next i
  
  
  With VScroll1
      '.Height = Frame1.Height
      .Min = 0
      .Max = selectedPeriods.Count * (lblPeriode(0).Height + 145) - Frame1.Height
      .SmallChange = Screen.TwipsPerPixelY * 20
      .LargeChange = .SmallChange
  End With
  
  FillGrid
 
End Sub

'##ModelId=5C8A684D034E
Private Sub pScrollForm()
   Dim ctl As Control
   Dim top As Integer
   Dim isScrollDown As Boolean
   
   If VScroll1.Value - oldPos > 0 Then
    isScrollDown = True
   Else
    isScrollDown = False
   End If

   For Each ctl In Me.Controls
      If Not (TypeOf ctl Is VScrollBar) And _
         Not (TypeOf ctl Is Frame) And _
         Not (TypeOf ctl Is fpSpread) And _
         Not (TypeOf ctl Is CommandButton) Then
          top = ctl.top + oldPos - VScroll1.Value
          If top < 100 Then
            ctl.Visible = False
          Else
            ctl.Visible = True
          End If
          ctl.top = top
      End If
   Next
   
   For Each ctl In Me.Controls
      If ctl.Name = "cmdClose" Then
        ctl.top = ctl.top + oldPos - VScroll1.Value
      End If
   Next

   oldPos = VScroll1.Value
End Sub

'##ModelId=5C8A684D036D
Private Sub VScroll1_Change()
  Call pScrollForm
End Sub

'##ModelId=5C8A684D037D
Private Sub VScroll1_Scroll()
  Call pScrollForm
End Sub


'******************************************************************************************************************************
'********************************************* FILL THE GRID WITH LIST OF PERIODES ********************************************
'******************************************************************************************************************************

'##ModelId=5C8A684D038D
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

'##ModelId=5C8A684D03AC
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

'##ModelId=5C8A684D03BB
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  sprListe.sheet = sprListe.ActiveSheet
  sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
End Sub

'##ModelId=5C8A684E0022
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

'##ModelId=5C8A684E009F
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

'##ModelId=5C8A684E011C
Private Sub sprListe_Click(ByVal Col As Long, ByVal Row As Long)
  SetNumPeriode
End Sub

'##ModelId=5C8A684E015A
Private Sub SetNumPeriode()
  numPeriode = 0
  
  If sprListe.ActiveRow < 0 Then Exit Sub
  
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 2
  
  numPeriode = CLng(sprListe.text)
End Sub
