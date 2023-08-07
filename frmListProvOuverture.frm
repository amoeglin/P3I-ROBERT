VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListProvOuverture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provision à l'ouverture"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "frmListProvOuverture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread sprProv 
      Bindings        =   "frmListProvOuverture.frx":1BB2
      Height          =   2400
      Left            =   0
      TabIndex        =   4
      Top             =   315
      Width           =   9780
      _Version        =   524288
      _ExtentX        =   17251
      _ExtentY        =   4233
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
      SpreadDesigner  =   "frmListProvOuverture.frx":1BD1
      AppearanceStyle =   0
   End
   Begin MSAdodcLib.Adodc dtaProvOuverture 
      Height          =   330
      Left            =   3285
      Top             =   2880
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
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
      Caption         =   "dtaProvOuverture"
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
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Modifier"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   2745
      Width           =   9780
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   8325
      TabIndex        =   0
      Top             =   2880
      Width           =   1440
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9780
   End
End
Attribute VB_Name = "frmListProvOuverture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E1033E"
Option Explicit

'##ModelId=5C8A67E20060
Private Sub btnClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A67E2006F
Private Sub RefreshListe()
  Dim df As Integer, rq As String
  
  sprProv.Visible = False
  
  df = Year(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode))
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  m_dataSource.SetDatabase dtaProvOuverture
  rq = "SELECT RECNO, Societe.SONOM as [Société], PROV_ANn as [" & df & "], " _
        & " PROV_ANn1 as [" & df - 1 & "], PROV_ANn2 as [" & df - 2 & "], PROV_ANn3 as [" & df - 3 & "], " _
        & " PROV_ANn4 as [" & df - 4 & "], PROV_ANn5 as [" & df - 5 & " et -] " _
        & " FROM ProvisionsOuverture ProvisionsOuverture INNER JOIN Societe Societe ON (ProvisionsOuverture.POSTECLE = Societe.SOCLE) AND (ProvisionsOuverture.GPECLE = Societe.SOGROUPE) " _
        & " WHERE NUMCLE = " & numPeriode & " And GPECLE = " & GroupeCle _
        & " ORDER BY Societe.SONOM"
  dtaProvOuverture.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaProvOuverture.Refresh
  
  ' mets à jours les n° de ligne dans le spread
  If Not dtaProvOuverture.Recordset.EOF Then
    dtaProvOuverture.Recordset.MoveLast
    dtaProvOuverture.Recordset.MoveFirst
    
    sprProv.MaxRows = dtaProvOuverture.Recordset.RecordCount
  Else
    sprProv.MaxRows = 0
  End If

  sprProv.Refresh
  
  ' cache la colonne RECNO
  sprProv.ColWidth(1) = 0
  
  sprProv.Visible = True
End Sub

'##ModelId=5C8A67E2007F
Private Sub btnEdit_Click()
  Dim fpo As New frmProvOuverture
  Dim RECNO As Long
  
  sprProv.Row = sprProv.ActiveRow
  sprProv.Col = 1
  RECNO = CLng(sprProv.text)
  
  fpo.NumSte = m_dataHelper.GetParameter("SELECT POSTECLE FROM ProvisionsOuverture WHERE RECNO = " & RECNO)
  
  fpo.Show vbModal
  
  Call RefreshListe
End Sub

'##ModelId=5C8A67E2008F
Private Sub Form_Load()
  lblGroupe = "Période n° " & numPeriode & " du Groupe '" & NomGroupe & "'"

  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
      
  If archiveMode Then
    btnEdit.Enabled = False
  End If
  
  ' teste et crée les enregistrements par societe
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT SOCLE From Societe WHERE SOGROUPE = " & GroupeCle, Snapshot)
    
  Do Until rs.EOF
    Set rs2 = m_dataSource.OpenRecordset("SELECT * FROM ProvisionsOuverture WHERE NUMCLE = " & numPeriode & " And GPECLE = " & GroupeCle & " AND POSTECLE = " & rs.fields(0), Dynamic)
    
    If rs2.EOF Then
      ' crée l'enregistrement
      rs2.AddNew
      
      rs2.fields("NUMCLE").Value = numPeriode
      rs2.fields("GPECLE").Value = GroupeCle
      rs2.fields("POSTECLE").Value = rs.fields(0).Value
      rs2.fields("PROV_ANn").Value = 0
      rs2.fields("PROV_ANn1").Value = 0
      rs2.fields("PROV_ANn2").Value = 0
      rs2.fields("PROV_ANn3").Value = 0
      rs2.fields("PROV_ANn4").Value = 0
      rs2.fields("PROV_ANn5").Value = 0
      
      rs2.Update
    End If
    
    rs2.Close
    
    rs.MoveNext
  Loop
  rs.Close
  
  Call RefreshListe
End Sub


'##ModelId=5C8A67E200AE
Private Sub sprProv_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' ga
End Sub
