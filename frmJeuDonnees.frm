VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmJeuxDonnees 
   Caption         =   "Données du Lot  ..."
   ClientHeight    =   7995
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   13170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   13170
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnDoublon 
      Caption         =   "&Doublon"
      Height          =   375
      Left            =   2610
      TabIndex        =   7
      Top             =   6075
      Width           =   1170
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Modifier"
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   6075
      Width           =   1170
   End
   Begin VB.CommandButton btnDel 
      Caption         =   "&Supprimer"
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   6075
      Width           =   1170
   End
   Begin VB.CommandButton btnControle 
      Caption         =   "&Contrôles"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   6075
      Width           =   1935
   End
   Begin VB.CommandButton btnVoirTrace 
      Caption         =   "&Afficher la Trace"
      Height          =   375
      Left            =   4050
      TabIndex        =   1
      Top             =   6075
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   9045
      TabIndex        =   0
      Top             =   6075
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12600
      Top             =   5985
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   11025
      Top             =   2970
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      Top             =   990
      Width           =   11040
      _Version        =   524288
      _ExtentX        =   19473
      _ExtentY        =   8758
      _StockProps     =   64
      BackColorStyle  =   1
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
      SpreadDesigner  =   "frmJeuDonnees.frx":0000
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   10
      VisibleRows     =   100
      AppearanceStyle =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11970
      Top             =   5895
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
            Picture         =   "frmJeuDonnees.frx":060F
            Key             =   "openCahier"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0719
            Key             =   "openPeriode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0823
            Key             =   "About"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":092D
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0A37
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0B41
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0C9B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0DF5
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":0F4F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":2B11
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":2C6B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":2DC5
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJeuDonnees.frx":2F1F
            Key             =   "IMG13"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      Begin VB.CommandButton cmdExportStat 
         Caption         =   "Exporter Statutaire"
         Height          =   285
         Left            =   8880
         TabIndex        =   24
         Top             =   90
         Width           =   1875
      End
      Begin VB.CommandButton btnFiltre 
         Caption         =   "Filtre"
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   90
         Width           =   765
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   8010
         TabIndex        =   21
         Top             =   45
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton btnImport 
         Caption         =   "&Importer"
         Height          =   285
         Left            =   11970
         TabIndex        =   20
         Top             =   90
         Width           =   1035
      End
      Begin VB.CommandButton btnExport 
         Caption         =   "E&xporter"
         Height          =   285
         Left            =   10890
         TabIndex        =   19
         Top             =   90
         Width           =   1035
      End
      Begin VB.TextBox lblNbSelected 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3735
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "XX lignes sélectionnées       NUENP3I max = 888888"
         Top             =   135
         Width           =   4155
      End
      Begin VB.CommandButton btnSelectAll 
         Caption         =   "Tous"
         Height          =   285
         Left            =   2025
         TabIndex        =   17
         Top             =   90
         Width           =   765
      End
      Begin VB.CommandButton btnSelectNone 
         Caption         =   "Aucun"
         Height          =   285
         Left            =   1170
         TabIndex        =   16
         Top             =   90
         Width           =   765
      End
      Begin VB.CheckBox chkSelect 
         Height          =   285
         Left            =   855
         TabIndex        =   15
         Top             =   90
         Width           =   240
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Données Initiales"
         Height          =   285
         Left            =   11475
         TabIndex        =   8
         Top             =   1035
         Width           =   1635
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      Begin VB.CommandButton btnNumEnrP3I 
         Caption         =   "&NUENRP3I"
         Height          =   285
         Left            =   11970
         TabIndex        =   22
         Top             =   90
         Width           =   1035
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Données Initiales"
         Height          =   285
         Left            =   11475
         TabIndex        =   14
         Top             =   1035
         Width           =   1635
      End
      Begin VB.CommandButton btnEnregistrer 
         Caption         =   "&Enregistrer"
         Height          =   285
         Left            =   10890
         TabIndex        =   13
         Top             =   90
         Width           =   1035
      End
      Begin VB.TextBox txtComment 
         Height          =   330
         Left            =   4230
         TabIndex        =   12
         Text            =   "j ç j g ,;"
         Top             =   45
         Width           =   6585
      End
      Begin VB.CheckBox chkDonneesInitiales 
         Caption         =   "Données Initiales"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   45
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Commentaire"
         Top             =   90
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmJeuxDonnees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A681301CB"

Option Explicit
Option Base 0

'##ModelId=5C8A681302C4
Private frmNumeroLot As Long
'##ModelId=5C8A681302E4
Private frmNumPeriode As Long

'##ModelId=5C8A68130303
Private frmOrdreDeTri As String

'##ModelId=5C8A68130322
Private Const colDataVersion = 4 ' 3 avant ajout de la colonne 'Selected'


'##ModelId=5C8A68130332
Private bInMultiDeleteRecord As Boolean
'##ModelId=5C8A68130361
Private bInMultiCheckRecord As Boolean
'


'##ModelId=5C8A68130380
Public Property Let NumeroLot(n As Long)
  frmNumeroLot = n
End Property


'##ModelId=5C8A6813039F
Public Property Let numPeriode(n As Long)
  frmNumPeriode = n
End Property


'##ModelId=5C8A681303BE
Private Sub btnClose_Click()
  ret_code = -1
  Unload Me
End Sub


'##ModelId=5C8A681303DE
Private Sub btnControle_Click()
  
  ' charge l'object d'import
  Dim txtObjetImport As String

  txtObjetImport = GetSettingIni(CompanyName, SectionName, "ObjetImportSASP3I", "#")

  If txtObjetImport = "#" Then
    MsgBox "La section ObjetImport n'est pas présente," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
    Exit Sub
  End If

  On Error GoTo errImport
  
  Dim objImport As iP3IGeneraliImport
  
  Set objImport = CreateObject(txtObjetImport)

  objImport.DoControle CommonDialog1, m_dataSource, GroupeCle, frmNumeroLot, frmNumPeriode, sFichierIni

  Set objImport = Nothing

  Exit Sub

errImport:
  MsgBox "Erreur : " & Err & vbLf & Err.Description & vbLf & "(Objet Import= " & txtObjetImport & ")", vbCritical
  Resume Next
End Sub


'##ModelId=5C8A68140005
Private Sub MultiDeleteRecord(Action As Long)
  If sprListe.SelectionCount > 0 Then
    
    Dim i As Long, c As Variant, r As Variant, c2 As Variant, r2 As Variant
    Dim frm As frmDetailFlux
    Dim sel() As Long
    
    sprListe.ReDraw = False
    
    ReDim sel(sprListe.SelectionCount) As Long
    Dim nbSel As Long
    nbSel = sprListe.SelectionCount - 1
    
    For i = 0 To nbSel
      sprListe.GetSelection i, c, r, c2, r2
      sel(i) = r
    Next
    
    
    bInMultiDeleteRecord = True

    
    For i = 0 To nbSel
      r = sel(i)
      
      sprListe.SetActiveCell 2, r
      
      Set frm = New frmDetailFlux
      
      Load frm
            
      sprListe.Col = 1
      sprListe.Row = r
      frm.NumeroLot = CLng(sprListe.text)
      
      sprListe.Col = 3
      frm.numeroEnregistrement = CLng(sprListe.text)
      
      frm.frmAutomate = True
      frm.frmAction = Action
      
      frm.Show vbModal
          
      Set frm = Nothing
      
    Next
  
    Erase sel
  
    
    bInMultiDeleteRecord = False
    
    sprListe.ReDraw = True
    
    RefreshListe
  End If
End Sub


'##ModelId=5C8A68140024
Private Sub btnDel_Click()
  
  MultiDeleteRecord 2 ' supprimer

End Sub


'##ModelId=5C8A68140044
Private Sub btnDoublon_Click()
  
  MultiDeleteRecord 3 ' doublon

End Sub


'##ModelId=5C8A68140053
Private Sub btnEdit_Click()
  If sprListe.ActiveRow <= 0 Then Exit Sub
  
  ret_code = 0
  
  Dim frm As New frmDetailFlux
  
  sprListe.Col = 1
  sprListe.Row = sprListe.ActiveRow
  frm.NumeroLot = CLng(sprListe.text)
  
  sprListe.Col = 3 ' 2 avec ajout colonne 'Selected'
  frm.numeroEnregistrement = CLng(sprListe.text)
  
  frm.Show vbModal
  
  If ret_code = 0 Then
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
    
    Screen.MousePointer = vbDefault
  End If

End Sub


'##ModelId=5C8A68140063
Private Sub btnEnregistrer_Click()
  Dim rs As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT Commentaire From P3ILOGTRAIT WHERE NUTRAITP3I=" & frmNumeroLot, Dynamic)
  
  If Not rs.EOF Then
    
    rs.fields(0) = Trim(txtComment)
    rs.Update
    
  End If
  
  rs.Close

End Sub


'##ModelId=5C8A68140073
Private Sub btnExport_Click()
  On Error GoTo err_export
  
  CommonDialog1.filename = "Lot" & frmNumeroLot & ".xls"
  'CommonDialog1.filename = "*.xls"
  CommonDialog1.filter = "Fichier Excel|*.xls|"
  
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
  CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
  
  CommonDialog1.ShowSave
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Then
    Exit Sub
  End If
  
  Dim rq As String
  
  rq = dtaPeriode.RecordSource
  
  rq = Replace(rq, ", P.Selected as '#'", "")
  
  rq = Replace(rq, ", P.DataVersion", ", DV.Code as DataVersion")
  rq = Replace(rq, " P3IPROVCOLL P ", " P3IPROVCOLL P INNER JOIN DataVersion DV ON DV.IdDataVersion=P.DataVersion ")

  If MsgBox("Voulez-vous exporter l'ensemble des lignes ou seulement celles sélectionnées ?" & vbLf & vbLf _
            & "Choisir Oui pour l'ENSEMBLE des lignes" & vbLf _
            & "Choisir Non pour les lignes SELECTIONNEES seules" & vbLf, vbQuestion + vbYesNo) = vbNo Then
    rq = Replace(rq, "WHERE ", "WHERE P.Selected<>0 AND ")
  End If
  
  If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Or Right(UCase(CommonDialog1.filename), 5) = ".XLSX" Then
'    ExportTableToExcelFile CommonDialog1.filename, _
'                           "Lot" & frmNumeroLot, _
'                           "Donnees", sprListe, CommonDialog1, "", False, False
    
    Screen.MousePointer = vbHourglass
    
'    ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, CommonDialog1.filename, "Lot" & frmNumeroLot, sprListe
    ExportQueryResultToExcel m_dataSource, rq, CommonDialog1.filename, "Lot" & frmNumeroLot, sprListe, "DONNEES_LOT"
    
    Screen.MousePointer = vbDefault
  End If

  Exit Sub
  
err_export:
  
  If Err <> cdlCancel Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  CommonDialog1.CancelError = False
End Sub

'##ModelId=5C8A68140092
Private Sub cmdExportStat_Click()
  On Error GoTo err_export
  
  CommonDialog1.filename = "Lot_Statutaire" & frmNumeroLot & ".xls"
  'CommonDialog1.filename = "*.xls"
  CommonDialog1.filter = "Fichier Excel|*.xls|"
  CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
  CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
  CommonDialog1.ShowSave
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Then
    Exit Sub
  End If
  
  Dim rq As String
  
  rq = dtaPeriode.RecordSource
  
  rq = Replace(rq, ", P.Selected as '#'", "")
  rq = Replace(rq, ", P.DataVersion", ", DV.Code as DataVersion")
  rq = Replace(rq, " P3IPROVCOLL P ", " P3IPROVCOLL P INNER JOIN DataVersion DV ON DV.IdDataVersion=P.DataVersion ")
  'rq = Replace(rq, "WHERE ", "WHERE P.Selected<>0 AND ")
  
  SetCategoryCodeStatVariable
  
  If CategoryCodeSTAT <> "" Then
    rq = Replace(rq, "WHERE ", "WHERE P.CDPRODUIT IN (" & CategoryCodeSTAT & ") AND ")
  End If
  
  If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Or Right(UCase(CommonDialog1.filename), 5) = ".XLSX" Then
'    ExportTableToExcelFile CommonDialog1.filename, _
'                           "Lot" & frmNumeroLot, _
'                           "Donnees", sprListe, CommonDialog1, "", False, False
    
    Screen.MousePointer = vbHourglass
    
'    ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, CommonDialog1.filename, "Lot" & frmNumeroLot, sprListe
    ExportQueryResultToExcel m_dataSource, rq, CommonDialog1.filename, "Lot" & frmNumeroLot, sprListe, "DONNEES_LOT"
    
    Screen.MousePointer = vbDefault
  End If

  Exit Sub
  
err_export:
  
  If Err <> cdlCancel Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  CommonDialog1.CancelError = False
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie une ligne en lui donnant le statut modifiee
'
'##ModelId=5C8A681400A1
Private Sub CopyLigne(NUMENRP3I As Long, etat As Integer)
  If NUMENRP3I = 0 Or frmNumeroLot = 0 Then Exit Sub
  
  On Error GoTo err_CopyLigne
    
  ' copy manuelle, Execute volide la transaction
  Dim rsIn As ADODB.Recordset, rsOut As ADODB.Recordset, f As ADODB.field
  
  Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & NUMENRP3I & " ORDER BY NUENRP3I, DataVersion DESC", Snapshot)
  
  If rsIn.EOF = False Then
    Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & NUMENRP3I, Dynamic)
    
    rsOut.AddNew
    
    For Each f In rsIn.fields
      If UCase(f.Name) = "DATAVERSION" Then
        rsOut.fields(f.Name).Value = etat
      Else
        rsOut.fields(f.Name).Value = f.Value
      End If
    Next
    
    rsOut.Update
    
    rsOut.Close
  End If
  
  rsIn.Close
  
  Exit Sub
  
err_CopyLigne:
  MsgBox "CopyLigne() : Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A681400E0
Private Function EcritTrace(rsXL As ADODB.Recordset, rsCR As ADODB.Recordset, etat As tagDataVersion) As Boolean
  ' ajoute une ligne dans la trace avec les données
  Dim rsTrace As ADODB.Recordset
  Dim comment As String, txtChamps As String
  Dim depart As Integer, i As Integer
  
  On Error GoTo err_EcritTrace
  
  Set rsTrace = m_dataSource.OpenRecordset("SELECT * FROM P3ITRACE WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & rsXL.fields("NUENRP3I"), Dynamic)
  
  rsTrace.AddNew
  
  
  rsTrace.fields("NUTRAITP3I") = frmNumeroLot
  rsTrace.fields("NUENRP3I") = rsXL.fields("NUENRP3I")
  rsTrace.fields("DateModif") = Now
  
  ' copie les anciennes données (appeler EcritTrace() avant d'effectuer les changements)
  If etat <> eAjouter Then
    For i = 3 To rsCR.fields.Count - 1
      If UCase(rsCR.fields(i).Name) <> "SELECTED" Then
        rsTrace.fields(rsCR.fields(i).Name).Value = rsCR.fields(i).Value
      End If
    Next
  End If
  
  ' test le code action ou test si il y a eu un changement de valeur
  Select Case etat
    Case eUndelete ' Undelete
      ' modifie simplement le statut de la ligne
      rsTrace.fields("TypeChangement").Value = "M"
      rsTrace.fields("Commentaire").Value = "Récupération de la ligne"
    
    Case eModifie ' Modifier
      rsTrace.fields("TypeChangement").Value = "M"
      comment = "Modification de la ligne"
      
      txtChamps = vbNullString
      depart = 3 ' on ignore les cles
      
      ' changements ?
      Dim bChange As Boolean, FieldName As String
      Dim fCR As ADODB.field, fXL As ADODB.field
      
      For i = 3 To rsCR.fields.Count - 1
        bChange = False
        FieldName = rsCR.fields(i).Name
        
        If UCase(FieldName) <> "SELECTED" And UCase(FieldName) <> "DATAVERSION" Then
          Set fCR = rsCR.fields(i)
          Set fXL = rsXL.fields(FieldName)
          
          Select Case fCR.Type
            Case adDate, adDBDate, adDBTimeStamp
              If fXL.Value = vbNullString And Not IsNull(fCR.Value) Then
                bChange = True
              ElseIf fXL.Value = vbNullString And IsNull(fCR.Value) Then
                bChange = False ' à cause du CDate
              ElseIf fXL.Value <> vbNullString And IsNull(fCR.Value) Then
                bChange = True
              ElseIf CDate(fXL.Value) <> fCR.Value Then
                bChange = True
              End If
            
            Case adChar, adVarChar, adVarWChar
              If Trim(fXL.Value) <> Trim(fCR.Value) _
                Or ((IsEmpty(fCR.Value) Or IsNull(fCR.Value)) And Trim(fXL.Value) <> vbNullString) Then
                bChange = True
              End If
            
            Case adNumeric
              If IsNull(fXL.Value) And IsNull(fCR.Value) Then
                bChange = False
              ElseIf fXL.Value <> "" And IsNull(fCR.Value) Then
                bChange = True
              ElseIf fXL.Value = "" And IsNull(fCR.Value) Then
                bChange = False
              ElseIf IsNull(fXL.Value) And Not IsNull(fCR.Value) Then
                bChange = True
              ElseIf CStr(CDbl(fXL.Value)) <> CStr(CDbl(fCR.Value)) Then
                bChange = True
              End If
            
            Case Else
              If fXL.Value <> fCR.Value Or ((IsEmpty(fCR.Value) Or IsNull(fCR.Value)) And fXL.Value <> vbNullString) Then
                bChange = True
              End If
          End Select
        End If
        
        If bChange = True Then
          txtChamps = txtChamps & FieldName & ", "
        End If
      Next
    
      If txtChamps <> vbNullString Then
        If Right(txtChamps, 2) = ", " Then
          txtChamps = Left(txtChamps, Len(txtChamps) - 2)
        End If
        comment = comment & " - " & txtChamps
      End If
      rsTrace.fields("Commentaire").Value = comment
    
    Case eSupprimer ' Supprimer
      rsTrace.fields("TypeChangement").Value = "S"
      rsTrace.fields("Commentaire").Value = "Suppression de la ligne"
      
    Case eDoublon ' Doublon
      rsTrace.fields("TypeChangement").Value = "D"
      rsTrace.fields("Commentaire").Value = "Suppression (Doublon) de la ligne"
  
    Case eAjouter ' Ajout
      rsTrace.fields("TypeChangement").Value = "A"
      rsTrace.fields("Commentaire").Value = "Ajout de la ligne"
  
  End Select
  
  
  If txtChamps <> vbNullString Or rsTrace.fields("TypeChangement").Value = "S" _
     Or rsTrace.fields("TypeChangement").Value = "D" Or rsTrace.fields("TypeChangement").Value = "A" Then
    rsTrace.Update
    EcritTrace = True
  Else
    rsTrace.CancelUpdate
    EcritTrace = False
  End If
  
  rsTrace.Close
  
  Exit Function
  
err_EcritTrace:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
  
End Function


'##ModelId=5C8A6814012E
Private Sub btnImport_Click()
  Dim Connexion As String
  Dim xlSource As DataAccess
  Dim rsXL As ADODB.Recordset
  Dim rsCR As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean, bChange As Boolean
  Dim f As ADODB.field, FieldName As String
  Dim n As Integer, bookmark As Variant, i As Integer
  Dim xlDataVersion As String
  Dim RECNO As Long
    
  ' demande le nom du fichier xls
  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Import de données dans le lot n°" & frmNumeroLot
  CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
  
  If CommonDialog1.filename = "" _
     Or CommonDialog1.filename = "*.xls" _
     Or CommonDialog1.filename = "*.xlsx" _
     Or CommonDialog1.filename = "*.*" Then
    Exit Sub
  End If

  On Error GoTo GestionErreur
  
  cmdExportStat.Visible = False
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Value = 0
  ProgressBar1.Max = 100
  ProgressBar1.Refresh
  
  Screen.MousePointer = vbHourglass
 
  Dim m_Logger As New clsLogger
  
  m_Logger.FichierLog = m_logPath & GetWinUser & "_ErreurImport.log"
  m_Logger.CreateLog "Import " & CommonDialog1.filename & " dans le lot n°" & frmNumeroLot

  ' ouvre la feuille excel
  Set xlSource = New DataAccess
  
  ' chaine de connexion ADO pour Excel
  'Connexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.filename & ";Extended Properties=" & cdExcelExtendedProperties & ";Persist Security Info=False"
  
  ' chaine de connexion
  
  If UCase(Right(CommonDialog1.filename, 4)) = ".XLS" Then
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  ElseIf UCase(Right(CommonDialog1.filename, 5)) = ".XLSX" Then
    Connexion = Replace(ConnectionStringXlsx, "%1", CommonDialog1.filename)
  Else
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  End If
  
  lblNbSelected.Alignment = 1
  
  xlSource.Connect Connexion

  Set rsXL = xlSource.OpenRecordset("SELECT * FROM DONNEES_LOT WHERE not IsNull(NUENRP3I)", Snapshot)
  
  If Not rsXL Is Nothing Then
    If rsXL.EOF Then
      ProgressBar1.Max = 1
      m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
    Else
      rsXL.MoveLast
      rsXL.MoveFirst
    
      ' cree une transaction
      fFoundError = False
      m_dataSource.BeginTrans
      
      ProgressBar1.Max = rsXL.RecordCount + 1
      
      ' lit les enregistrements
      Do Until rsXL.EOF
        bRejet = False
        If (rsXL.AbsolutePosition Mod 9) = 0 Then
          ' affiche la position
          ProgressBar1.Value = rsXL.AbsolutePosition
          ProgressBar1.Refresh
          
          lblNbSelected.text = "Import en cours : " & rsXL.AbsolutePosition & " / " & rsXL.RecordCount
          DoEvents
        End If
        
        If Not IsNull(rsXL.fields("NUENRP3I")) Then
          
          'recherche de l'enregistrement
          Set rsCR = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & rsXL.fields("NUENRP3I") & " ORDER BY NUENRP3I, DataVersion DESC", Dynamic)
          
          xlDataVersion = Trim(UCase(rsXL.fields("DataVersion")))
          
          If rsCR.EOF Then
            
            ' Ajout ?
            If xlDataVersion <> "A" Then
              ' Erreur sur DataVersion
              Call m_Logger.EcritTraceDansLog("   Erreur  : DataVersion '" & rsXL.fields("DataVersion") & "' non correcte pour nouveau NUENRP3I " & rsXL.fields("NUENRP3I") & ", devrait être 'A'")
            Else
              ' NUENRP3I
              If (rsXL.fields("NUENRP3I") = -1) Then
                RECNO = m_dataHelper.GetParameterAsDouble("SELECT Max(P.NUENRP3I) FROM P3IPROVCOLL P WHERE P.NUTRAITP3I = " & frmNumeroLot) + 1
              Else
                RECNO = rsXL.fields("NUENRP3I")
              End If
              
              rsCR.AddNew
              
              EcritTrace rsXL, rsCR, eAjouter
            End If
          
          Else
            
            ' Ajout ?
            If xlDataVersion = "A" Then
              ' Erreur sur NUENRP3I (ne devrait pas exister)
              RECNO = m_dataHelper.GetParameterAsDouble("SELECT Max(P.NUENRP3I) FROM P3IPROVCOLL P WHERE P.NUTRAITP3I = " & frmNumeroLot) + 1
              
              Call m_Logger.EcritTraceDansLog("   Erreur  : DataVersion '" & rsXL.fields("DataVersion") & "' non correcte pour NUENRP3I déjà existant " & rsCR.fields("NUENRP3I") & ", affecté à " & RECNO)
              
              rsCR.AddNew
              
              EcritTrace rsXL, rsCR, eAjouter
            End If
          
          End If
          
          
          Select Case xlDataVersion
            Case "R"
              ' modifie simplement le statut de la ligne
              If rsCR.fields("DataVersion").Value > eModifie Then
                EcritTrace rsXL, rsCR, eUndelete
                
                rsCR.fields("DataVersion").Value = eModifie
                rsCR.Update
              End If

            Case "I", "M", "A"
              If rsCR.fields("DataVersion").Value <> eInitiale Or xlDataVersion = "A" Then
                
                '
                ' Modifiée : modifie simplement le statut de la ligne
                '
                
                If xlDataVersion <> "A" Then
                  EcritTrace rsXL, rsCR, eModifie
                End If
                
                ' changements ?
                For i = 3 To rsCR.fields.Count - 1
                  FieldName = rsCR.fields(i).Name
                  
                  If UCase(FieldName) <> "SELECTED" Then
                    If rsXL.fields(FieldName).Value <> "" And IsNull(rsCR.fields(i).Value) Then
                      rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                    ElseIf rsCR.fields(i).Type = adNumeric Then
                      If rsXL.fields(FieldName).Value <> rsCR.fields(i).Value Then
                        rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                      End If
                    ElseIf Trim(rsXL.fields(FieldName).Value) <> Trim(rsCR.fields(i).Value) Then
                      rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                    End If
                  End If
                Next
                  
                rsCR.fields("NUTRAITP3I").Value = frmNumeroLot
                rsCR.fields("NUENRP3I").Value = rsXL.fields("NUENRP3I").Value
                rsCR.fields("DataVersion").Value = eModifie
                
                
                If xlDataVersion = "A" Then
                  rsCR.fields("NUENRP3I").Value = RECNO
                  rsCR.fields("LBCOMLIG").Value = "Ligne ajoutee"
                End If
                
                rsCR.Update
              
              Else
                
                '
                ' Initiale : copie la ligne et change son statut
                '
                
                ' changements ?
                bChange = EcritTrace(rsXL, rsCR, eModifie)
                
                If bChange = True Then
                  rsCR.Close
                  
                  ' copie la ligne et change son statut
                  CopyLigne rsXL.fields("NUENRP3I").Value, eModifie
                  
                  Set rsCR = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & rsXL.fields("NUENRP3I") & " ORDER BY NUENRP3I, DataVersion DESC", Dynamic)
                  
                  For i = 3 To rsCR.fields.Count - 1
                    FieldName = rsCR.fields(i).Name
                    
                    If UCase(FieldName) <> "SELECTED" Then
                      If rsXL.fields(FieldName).Value <> "" And IsNull(rsCR.fields(i).Value) Then
                        rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                      ElseIf rsCR.fields(i).Type = adNumeric Then
                        If rsXL.fields(FieldName).Value <> rsCR.fields(i).Value Then
                          rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                        End If
                      ElseIf Trim(rsXL.fields(FieldName).Value) <> Trim(rsCR.fields(i).Value) Then
                        rsCR.fields(i).Value = rsXL.fields(FieldName).Value
                      End If
                    End If
                  Next
                  
                  rsCR.fields("NUTRAITP3I").Value = frmNumeroLot
                  rsCR.fields("NUENRP3I").Value = rsXL.fields("NUENRP3I").Value
                  rsCR.fields("DataVersion").Value = eModifie
                  
                  rsCR.Update
                End If
              
              End If
            
            Case "S"
              If rsCR.fields("DataVersion").Value <> eInitiale Then
                ' modifie simplement le statut de la ligne
                EcritTrace rsXL, rsCR, eSupprimer
                
                rsCR.fields("DataVersion").Value = eSupprimer
                rsCR.Update
              Else
                rsCR.Close
                
                ' copie la ligne et change son statut
                CopyLigne rsXL.fields("NUENRP3I").Value, eSupprimer
                
                Set rsCR = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & rsXL.fields("NUENRP3I") & " ORDER BY NUENRP3I, DataVersion DESC", Dynamic)
        
                EcritTrace rsXL, rsCR, eSupprimer
              End If
              
            Case "D"
              If rsCR.fields("DataVersion").Value <> eInitiale Then
                ' modifie simplement le statut de la ligne
                EcritTrace rsXL, rsCR, eDoublon
                
                rsCR.fields("DataVersion").Value = eDoublon
                rsCR.Update
              Else
                rsCR.Close
                
                ' copie la ligne et change son statut
                CopyLigne rsXL.fields("NUENRP3I").Value, eDoublon
                
                Set rsCR = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot & " AND NUENRP3I=" & rsXL.fields("NUENRP3I") & " ORDER BY NUENRP3I, DataVersion DESC", Dynamic)
        
                EcritTrace rsXL, rsCR, eDoublon
              End If
            
            Case Else
              fFoundError = True
              NbRejet = NbRejet + 1
              Call m_Logger.EcritTraceDansLog("   Erreur : DataVersion invalide '" & xlDataVersion & "'.")
            
          End Select
          
          rsCR.Close
            
        End If
        
        rsXL.MoveNext
      Loop
    End If
  End If
  
  Call m_Logger.EcritTraceDansLog(rsXL.RecordCount & " lignes dans le fichier " & CommonDialog1.filename)
  
  If fFoundError Then
    Call m_Logger.EcritTraceDansLog(">>>>> Fichier rejetté à cause des erreurs durant l'import !")
    
    m_dataSource.RollbackTrans
  Else
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
  
    m_dataSource.Execute "UPDATE P3IPROVCOLL SET Selected=0 WHERE NUTRAITP3I=" & frmNumeroLot
    
    m_dataSource.CommitTrans
  End If
  
  
  m_Logger.EcritTraceDansLog "Fin Import"
  
  If Not rsXL Is Nothing Then
    rsXL.Close
    Set rsXL = Nothing
  End If
  
  Set rsCR = Nothing
  
  xlSource.Disconnect
  
  Set xlSource = Nothing
  
  Screen.MousePointer = vbDefault
  
  ProgressBar1.Visible = False
  cmdExportStat.Visible = True
  
  ' affichage des erreurs
  m_Logger.AfficheErreurLog
  
  
  lblNbSelected.text = ""
  lblNbSelected.Alignment = 0
  
  RefreshListe
  
  
  Exit Sub
  
GestionErreur:
  If rsXL Is Nothing Then
    If Err = -2147217865 Then
      m_Logger.EcritTraceDansLog "   Format Incorrect : le fichier " & CommonDialog1.filename & " ne correspond pas au format de la table 'P3IPROVCOLL'. " & Err.Description
    Else
      m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
    End If
  Else
    Select Case Err
      Case 3265
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import ")
      
      Case 3421
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rsXL.AbsolutePosition)
      
      Case -2147217887
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' une valeur doit être précisée (NULL interdit) - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  fFoundError = True
  Resume Next
End Sub


'##ModelId=5C8A6814014D
Private Sub btnNumEnrP3I_Click()
  Dim NUMENRP3I As String
  
  NUMENRP3I = InputBox("Veuillez saisir le NUENRP3I voulu :", "Accès à NUENRP3I")
  If NUMENRP3I = "" Then Exit Sub
  
  If IsNumeric(NUMENRP3I) = False Then
    MsgBox "Valeur saisie '" & NUMENRP3I & "' incorrecte !", vbCritical
    Exit Sub
  End If
  
  ' recherche de NUMENRP3I dans le spread
  Dim r As Long, bOk As Boolean
  
  Screen.MousePointer = vbHourglass
    
  bOk = False
  dtaPeriode.Recordset.MoveFirst
  Do Until dtaPeriode.Recordset.EOF
    If dtaPeriode.Recordset.fields("NUENRP3I").Value = NUMENRP3I Then
      r = dtaPeriode.Recordset.AbsolutePosition
      sprListe.TopRow = r
      sprListe.SetActiveCell 1, r
      sprListe.Row = r
      sprListe.SelModeSelected = True
      sprListe.SetFocus
      bOk = True
      Exit Do
    End If
    dtaPeriode.Recordset.MoveNext
  Loop
  
  If bOk = False Then
    MsgBox "NUENRP3I=" & NUMENRP3I & " introuvable !", vbCritical
  End If
  
  Screen.MousePointer = vbDefault
End Sub


'##ModelId=5C8A6814015D
Private Sub btnSelectAll_Click()
  
  Dim sSql As String
  
  If chkDonneesInitiales.Value = vbChecked Then Exit Sub

  ' mise à jour de la DB
  sSql = "UPDATE P3IPROVCOLL SET Selected = 1 FROM P3IPROVCOLL P WHERE P.NUTRAITP3I = " & frmNumeroLot
  sSql = sSql & "  AND ( " _
         & "          P.DataVersion >= 1 " _
         & "          OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IPROVCOLL P2 " _
         & "                                             WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
         & "         ) "
  
  m_dataSource.Execute sSql
  
  ' rafraichissement
  RefreshListe
End Sub


'##ModelId=5C8A6814016D
Private Sub btnSelectNone_Click()
  
  Dim sSql As String
  
  If chkDonneesInitiales.Value = vbChecked Then Exit Sub

  ' mise à jour de la DB
  sSql = "UPDATE P3IPROVCOLL SET Selected = 0 WHERE NUTRAITP3I = " & frmNumeroLot
  
  m_dataSource.Execute sSql
  
  ' rafraichissement
  RefreshListe
End Sub


'##ModelId=5C8A6814018C
Private Sub btnFiltre_Click()
  
  On Error GoTo err_filtre
  
  If chkDonneesInitiales.Value = vbChecked Then Exit Sub

  Dim sSql As String, sFile As String
  
  ' Suppression des selections existantes
  sSql = "UPDATE P3IPROVCOLL SET Selected = 0 WHERE NUTRAITP3I = " & frmNumeroLot
  
  m_dataSource.Execute sSql

  ' Selection des lignes d'après le filtre
  sFile = sReadIniFile("P3I", "ScriptSelectionDonnee", App.Path & "\Sql\P3I_Selection_Donnees_lot.sql", 255, sFichierIni)

  sSql = FileToString(sFile)
  
  If sSql = "" Then
    MsgBox "Le fichier '" & sFile & "' est introuvable !", vbCritical
    Exit Sub
  End If
  
  'sSql = Replace(sSql, "@NumeroLot", NumeroLot)
  sSql = Replace(sSql, "<NUMEROLOT>", frmNumeroLot)
  
  m_dataSource.Execute sSql
  
  RefreshListe
  
  Exit Sub
  
err_filtre:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
End Sub


'##ModelId=5C8A6814019B
Private Sub btnVoirTrace_Click()
  Dim frm As frmVoirTrace
  
  Set frm = New frmVoirTrace
  
  frm.NumeroLot = frmNumeroLot
  
  frm.Show vbModal
End Sub


'##ModelId=5C8A681401BB
Private Sub chkDonneesInitiales_Click()
  RefreshListe
  
  sprListe.SetFocus
End Sub


'##ModelId=5C8A681401CA
Private Sub chkSelect_Click()
  If bInMultiCheckRecord = True Then Exit Sub
  bInMultiCheckRecord = True

  If chkDonneesInitiales.Value = vbChecked Then Exit Sub
  
  chkSelect.Value = vbUnchecked

  If sprListe.SelectionCount = 0 Then Exit Sub

  Screen.MousePointer = vbHourglass
    
  sprListe.ReDraw = False
  
  Dim tr As Long, r As Long
  
  tr = sprListe.TopRow
  r = sprListe.ActiveRow
  
  Dim s As Integer, i As Integer
  Dim tabSel() As Long
  
  i = 0
  ReDim tabSel(sprListe.SelectionCount)
  
  ' sauvegarde la selection courante
  Do
    s = sprListe.GetMultiSelItem(s)
    If s = -1 Then Exit Do
    
    tabSel(i) = s
    i = i + 1
  Loop
  
  ' clique les lignes
  For s = 0 To i - 1
    CheckRow 2, tabSel(s)
  Next
  
  'RefreshListe ' pas utile
  sprListe.VirtualRefresh
  
  ' retabli la position du scroll vertical
  sprListe.TopRow = tr
  sprListe.SetActiveCell 1, r
  
  ' retabli la selection
  For s = 0 To i - 1
    sprListe.Row = tabSel(s)
    sprListe.SelModeSelected = True
  Next
  
  Erase tabSel
  
  sprListe.SetFocus
  
  RefreshCounter
  
  Screen.MousePointer = vbDefault
    
  sprListe.ReDraw = True
  
  bInMultiCheckRecord = False
End Sub



'##ModelId=5C8A681401DA
Private Sub Form_Activate()
  If bInMultiDeleteRecord = True Then
    Exit Sub
  End If

  frmOrdreDeTri = "NUENRP3I"
  
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
  
  Screen.MousePointer = vbDefault
End Sub


'##ModelId=5C8A681401F9
Private Sub Form_Load()
  ' chargement du masque du spread
  sprListe.LoadFromFile App.Path & "\JeuxDonnees.ss7"
  
  sprListe.OperationMode = OperationModeExtended
  
  m_dataSource.SetDatabase dtaPeriode
  
  bInMultiDeleteRecord = False
  bInMultiCheckRecord = False
End Sub


'##ModelId=5C8A68140209
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


'##ModelId=5C8A68140228
Private Sub RefreshListe()
  
  Dim rq As String, rs As ADODB.Recordset
  Dim filter As String
  Dim i As Integer
  
  'Dim debut As Date, fin As Date
  
  'debut = Now
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre
  Me.Caption = "Données du Lot n°" & frmNumeroLot
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  DoEvents
  
  'sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = True
  sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  'sprListe.VScrollSpecial = True
  'sprListe.VScrollSpecialType = 0
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  rq = "SELECT P.NUTRAITP3I, P.Selected as '#', P.NUENRP3I, P.DataVersion, P.CDCOMPAGNIE, P.CDAPPLI, RTRIM(P.CDPRODUIT) as CDPRODUIT, RTRIM(P.NUCONTRA) as NUCONTRA, P.NUPRET, RTRIM(P.NUDOSSIERPREST) as NUDOSSIERPREST, " _
       & " RTRIM(P.NUSOUSDOSSIERPREST) as NUSOUSDOSSIERPREST, RTRIM(P.NUCONTRATGESTDELEG) as NUCONTRATGESTDELEG, P.CDGARAN, P.NUREFMVT, P.CDTYPMVT, P.NUTRAITE, P.CDCATADHESION, RTRIM(P.LBCATADHESION) as LBCATADHESION, " _
       & " RTRIM(P.CDOPTION) as CDOPTION, RTRIM(P.LBOPTIONCON) as LBOPTIONCON, RTRIM(P.LBSOUSCR) as LBSOUSCR, RTRIM(P.IDASSUREAGI) as IDASSUREAGI, P.IDASSURE, RTRIM(P.LBASSURE) as LBASSURE, P.DTNAISSASS, P.CDSEXASSURE, RTRIM(P.IDRENTIERAGI) as IDRENTIERAGI, " _
       & " P.IDRENTIER, RTRIM(P.LBRENTIER) as LBRENTIER, P.DTNAISSREN, RTRIM(P.IDCORENTIERAGI) as IDCORENTIERAGI, P.IDCORENTIER, RTRIM(P.LBCORENTIER) as LBCORENTIER, P.DTNAISSCOR, P.DTSURVSIN, P.AGESURVSIN, " _
       & " P.DURPOSARR, P.ANCARRTRA, P.CDPERIODICITE, P.CDTYPTERME, P.DTEFFREN, P.DTLIMPRO, P.DTDERREG, P.DTDEBPER, P.DTFINPER, P.CDPROEVA, " _
       & " P.ANNEADHE, P.CDSINCON, P.CDMISINV, P.DTMISINV, P.NUETABLI, P.TXREVERSION, P.DTCALCULPROV, P.DTTRAITPROV, P.DTCREATI, P.DTSIGBIA, " _
       & " P.DTDECSIN, P.CDRISQUE, RTRIM(P.LBRISQUE) as LBRISQUE, P.CDSITUATSIN, P.DTSITUATSIN, P.CDPRETRATTSIN, P.CDCTGPRT, RTRIM(P.LBCTGPRT) as LBCTGPRT, P.DTPREECH, P.DTDERECH, " _
       & " P.CDPERIODICITEECH, P.MTECHEANCE1, P.DTDEBPERECH1, P.DTFINPERECH1, P.MTECHEANCE2, P.DTDEBPERECH2, P.DTFINPERECH2, P.MTECHEANCE3, " _
       & " P.DTDEBPERECH3, P.DTFINPERECH3, P.CDTYPAMO, RTRIM(P.LBTYPAMO) as LBTYPAMO, P.TXINVPEC, P.DTDEBPERPIP, P.DTFINPERPIP, P.DTSAISIEPERJUSTIF, P.DTDEBPERJUSTIF, " _
       & " P.DTFINPERJUSTIF, P.DTDEBDERPERRGLTADA, P.DTFINDERPERRGLTADA, P.DTDERPERRGLTADA, P.MTDERPERRGLTADA, P.DTDEBDERPERRGLTADC, P.DTFINDERPERRGLTADC, " _
       & " P.DTDERPERRGLTADC, P.MTDERPERRGLTADC, P.CDSINPREPROV, P.MTTOTREGLEICIV, P.DTDEBPROV, P.DTFINPROV, P.INDBASREV, P.MTPREANN, P.MTPREREV, " _
       & " P.MTPREMAJ , P.MTPRIREG, P.MTPRIRE1, P.MTPRIRE2, P.CDMONNAIE, P.CDPAYS, P.CDAPPLISOURCE, "
       
  rq = rq & " RTRIM(P.CDCATINV) as CDCATINV, RTRIM(P.LBCATINV) as LBCATINV,  P.CDCONTENTIEUX, RTRIM(P.NUSINISTRE) as NUSINISTRE, " _
          & " RTRIM(P.CDCHOIXPREST) as CDCHOIXPREST, RTRIM(P.LBCHOIXPREST) as LBCHOIXPREST, P.MTCAPSSRISQ, RTRIM(P.FLAMORTISSABLE) as FLAMORTISSABLE, " _
          & " RTRIM(P.LBCOMLIG) as LBCOMLIG, " _
          & " RTRIM(P.Commentaire) as COMMENTAIRE " _
          & " FROM P3IPROVCOLL P " _
          & " WHERE P.NUTRAITP3I = " & frmNumeroLot
       
  If chkDonneesInitiales.Value = vbChecked Then
    
    rq = rq & " AND P.DataVersion=0 "
    
    btnEdit.Enabled = False
    
  Else
  
    rq = rq & "  AND ( " _
         & "          P.DataVersion >= 1 " _
         & "          OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
         & "                                             WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
         & "         ) "
    
'    rq = rq & " EXCEPT "
'
'  rq = rq & "SELECT P.NUTRAITP3I, P.NUENRP3I, P.DataVersion, P.CDCOMPAGNIE, P.CDAPPLI, RTRIM(P.CDPRODUIT) as CDPRODUIT, RTRIM(P.NUCONTRA) as NUCONTRA, P.NUPRET, RTRIM(P.NUDOSSIERPREST) as NUDOSSIERPREST, " _
'          & " RTRIM(P.NUSOUSDOSSIERPREST) as NUSOUSDOSSIERPREST, RTRIM(P.NUCONTRATGESTDELEG) as NUCONTRATGESTDELEG, P.CDGARAN, P.NUREFMVT, P.CDTYPMVT, P.NUTRAITE, P.CDCATADHESION, RTRIM(P.LBCATADHESION) as LBCATADHESION, " _
'          & " RTRIM(P.CDOPTION) as CDOPTION, RTRIM(P.LBOPTIONCON) as LBOPTIONCON, RTRIM(P.LBSOUSCR) as LBSOUSCR, RTRIM(P.IDASSUREAGI) as IDASSUREAGI, P.IDASSURE, RTRIM(P.LBASSURE) as LBASSURE, P.DTNAISSASS, P.CDSEXASSURE, RTRIM(P.IDRENTIERAGI) as IDRENTIERAGI, " _
'          & " P.IDRENTIER, RTRIM(P.LBRENTIER) as LBRENTIER, P.DTNAISSREN, RTRIM(P.IDCORENTIERAGI) as IDCORENTIERAGI, P.IDCORENTIER, RTRIM(P.LBCORENTIER) as LBCORENTIER, P.DTNAISSCOR, P.DTSURVSIN, P.AGESURVSIN, " _
'          & " P.DURPOSARR, P.ANCARRTRA, P.CDPERIODICITE, P.CDTYPTERME, P.DTEFFREN, P.DTLIMPRO, P.DTDERREG, P.DTDEBPER, P.DTFINPER, P.CDPROEVA, " _
'          & " P.ANNEADHE, P.CDSINCON, P.CDMISINV, P.DTMISINV, P.NUETABLI, P.TXREVERSION, P.DTCALCULPROV, P.DTTRAITPROV, P.DTCREATI, P.DTSIGBIA, " _
'          & " P.DTDECSIN, P.CDRISQUE, RTRIM(P.LBRISQUE) as LBRISQUE, P.CDSITUATSIN, P.DTSITUATSIN, P.CDPRETRATTSIN, P.CDCTGPRT, RTRIM(P.LBCTGPRT) as LBCTGPRT, P.DTPREECH, P.DTDERECH, " _
'          & " P.CDPERIODICITEECH, P.MTECHEANCE1, P.DTDEBPERECH1, P.DTFINPERECH1, P.MTECHEANCE2, P.DTDEBPERECH2, P.DTFINPERECH2, P.MTECHEANCE3, " _
'          & " P.DTDEBPERECH3, P.DTFINPERECH3, P.CDTYPAMO, RTRIM(P.LBTYPAMO) as LBTYPAMO, P.TXINVPEC, P.DTDEBPERPIP, P.DTFINPERPIP, P.DTSAISIEPERJUSTIF, P.DTDEBPERJUSTIF, " _
'          & " P.DTFINPERJUSTIF, P.DTDEBDERPERRGLTADA, P.DTFINDERPERRGLTADA, P.DTDERPERRGLTADA, P.MTDERPERRGLTADA, P.DTDEBDERPERRGLTADC, P.DTFINDERPERRGLTADC, " _
'          & " P.DTDERPERRGLTADC, P.MTDERPERRGLTADC, P.CDSINPREPROV, P.MTTOTREGLEICIV, P.DTDEBPROV, P.DTFINPROV, P.INDBASREV, P.MTPREANN, P.MTPREREV, " _
'          & " p.MTPREMAJ , p.MTPRIREG, p.MTPRIRE1, p.MTPRIRE2, p.CDMONNAIE, p.CDPAYS, p.CDAPPLISOURCE, RTRIM(P.LBCOMLIG) as LBCOMLIG " _
'          & " FROM P3IPROVCOLL P " _
'          & " WHERE P.NUTRAITP3I = " & frmNumeroLot & " AND P.DataVersion=0 " _
'          & "       AND EXISTS(SELECT 1 FROM P3IPROVCOLL P2 " _
'          & "                  WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>0)"
  
    btnEdit.Enabled = True
    
  End If
  
  rq = rq & " ORDER BY " & frmOrdreDeTri
          
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
  
  SetColonneDataFill colDataVersion
  
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
  
  txtComment.text = m_dataHelper.GetParameterAsStringCRW("SELECT Commentaire From P3ILOGTRAIT WHERE NUTRAITP3I=" & frmNumeroLot)
  If txtComment.text = "0" Then
    txtComment.text = vbNullString
  End If
  
  ' largeur des colonnes
  LargeurMaxColonneSpread sprListe
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
     
'  For i = 2 To sprListe.MaxCols
'    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 2
'  Next i
 
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row2 = -1
  
  sprListe.Col = 1
  sprListe.Col2 = sprListe.MaxCols - 1
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = sprListe.MaxCols
  sprListe.Col2 = sprListe.MaxCols
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
  sprListe.ColsFrozen = 2
  
pas_de_donnee:

  ' nb sde ligne 'Selected' et max  NUENRP3I
  RefreshCounter

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True
    
  Screen.MousePointer = vbDefault

  'fin = Now
  
  'lblFillTime.text = "Remplissage : " & DateDiff("s", debut, fin) & " s"

  Exit Sub

err_RefreshListe:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A68140247
Private Sub RefreshCounter()
  ' nb sde ligne 'Selected' et max  NUENRP3I
  Dim rq As String
  
  rq = m_dataHelper.GetParameter("SELECT count(*) FROM P3IPROVCOLL P WHERE Selected<>0 AND P.NUTRAITP3I = " & frmNumeroLot) & " lignes sélectionnées"
  rq = rq & "            NUENP3I max = " & m_dataHelper.GetParameter("SELECT max(P.NUENRP3I) FROM P3IPROVCOLL P WHERE P.NUTRAITP3I = " & frmNumeroLot)
  lblNbSelected.text = rq

End Sub


'##ModelId=5C8A68140257
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar2.top + Toolbar2.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 160
 
  topbtn = Me.ScaleHeight - btnHeight
  
  sprListe.Height = Maximum(topbtn - 2 * Toolbar1.Height - 100, 0)
  
  PlacePremierBoutton btnEdit, topbtn
  
  PlaceBoutton btnDel, btnEdit, topbtn
  
  PlaceBoutton btnDoublon, btnDel, topbtn
  
  PlaceBoutton btnVoirTrace, btnDoublon, topbtn
  
  PlaceBoutton btnControle, btnVoirTrace, topbtn
  
  PlaceBoutton btnClose, btnControle, topbtn

End Sub


'##ModelId=5C8A68140276
Private Sub sprListe_Click(ByVal Col As Long, ByVal Row As Long)
  ' tri ?
  If Col <> 0 And Row = 0 Then
#If TRI_SPREAD Then
    ' on utilise les fonction de tri du spread (tres lent)
    Screen.MousePointer = vbHourglass
    sprListe.Visible = False
    
    sprListe.Col = 0
    sprListe.Col2 = sprListe.MaxCols
    
    sprListe.Row = 0
    sprListe.Row2 = sprListe.MaxRows
    
    sprListe.SortBy = SS_SORT_BY_ROW
    sprListe.SortKey(1) = Col
    sprListe.SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
    
    sprListe.Action = SS_ACTION_SORT
    
    sprListe.Visible = True
    Screen.MousePointer = vbDefault
#Else
    ' on change la requete SQL (ORDER BY)
    If frmOrdreDeTri = dtaPeriode.Recordset.fields(Col - 1).Name Then
      If dtaPeriode.Recordset.fields(Col - 1).Name = "POCONVENTION" Then
        frmOrdreDeTri = "CAST(POCONVENTION as bigint) DESC"
      Else
        frmOrdreDeTri = dtaPeriode.Recordset.fields(Col - 1).Name & " DESC"  ' nom du champs de la table
      End If
    Else
      If dtaPeriode.Recordset.fields(Col - 1).Name = "POCONVENTION" Then
        frmOrdreDeTri = "CAST(POCONVENTION as bigint)"
      Else
        frmOrdreDeTri = dtaPeriode.Recordset.fields(Col - 1).Name   ' nom du champs de la table
      End If
    End If
    
    RefreshListe
#End If
  End If
End Sub


'##ModelId=5C8A681402A5
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  If Col = colDataVersion Then
    Dim version As Variant, txt As String, color As Long
    
    sprListe.GetDataFillData version, vbLong
    
    Cancel = True
    
    Select Case version
      Case eInitiale
        txt = "Initiale"
        color = vbWhite
        
      Case eModifie ' et Ajouter aussi)
        txt = "Modifiée"
        color = LTYELLOW
        
      Case eSupprimer
        txt = "Supprimée"
        color = LTRED
        
      Case eDoublon
        txt = "Doublon"
        color = PINK
             
      Case eAjouter
        txt = "Ajoutée"
        color = LTCYAN
             
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
    
    sprListe.Row = Row
    sprListe.Col = 2
    sprListe.Row2 = Row
    sprListe.Col2 = 2
    sprListe.BackColor = lavande_clair
    
    ' texte
    sprListe.SetText Col, Row, txt
  End If
End Sub


'##ModelId=5C8A68140332
Private Sub CheckRow(ByVal Col As Long, ByVal Row As Long)
    
  '
  ' bascule la sélection de la ligne
  '
  
  Dim sSql As String, numEnr As String, version As String
  
  sprListe.Row = Row
  
  sprListe.Col = 3
  numEnr = sprListe.text
    
  sprListe.Col = 4
  version = sprListe.text
    
  Select Case version
    Case "Initiale"
      version = eInitiale
      
    Case "Modifiée"
      version = eModifie
      
    Case "Supprimée"
      version = eSupprimer
      
    Case "Doublon"
      version = eDoublon
           
    Case "Ajoutée"
      version = eAjouter
           
    Case Else
      MsgBox "Erreur DataVersion dans CheckRow()", vbCritical
      Exit Sub
  End Select
  
  
  ' mise à jour de la DB
'  sSql = "UPDATE P3IPROVCOLL SET Selected = ~Selected WHERE NUTRAITP3I = " & frmNumeroLot & " AND NUENRP3I=" & numEnr & " AND Dataversion=" & version
'  m_dataSource.Execute sSql
  
  If dtaPeriode.Recordset.fields("NUENRP3I").Value = numEnr And dtaPeriode.Recordset.fields("DataVersion").Value = version Then
    
    ' modification directe
    dtaPeriode.Recordset.fields("#").Value = Not dtaPeriode.Recordset.fields("#").Value
    dtaPeriode.Recordset.Update
  
  Else
    
    ' filtre puis modification
    Dim pos As Variant
    
    pos = dtaPeriode.Recordset.bookmark
    
    m_dataHelper.Multi_Find dtaPeriode.Recordset, "NUENRP3I=" & numEnr & " AND Dataversion=" & version
    If Not dtaPeriode.Recordset.EOF Then
      dtaPeriode.Recordset.fields("#").Value = Not dtaPeriode.Recordset.fields("#").Value
      dtaPeriode.Recordset.Update
    End If
    
    dtaPeriode.Recordset.filter = adFilterNone
    dtaPeriode.Recordset.bookmark = pos
  End If
  
End Sub


'##ModelId=5C8A68140361
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
    
  If chkDonneesInitiales.Value = vbChecked Then Exit Sub
    
  Screen.MousePointer = vbHourglass
    
  sprListe.ReDraw = False
    
  Dim r As Long, tr As Long
  
  tr = sprListe.TopRow
  r = Row
  
  If r <> 0 Then
    ' bascule la sélection de la ligne
    CheckRow Col, Row
    
    ' retablit l'affichage
    sprListe.VirtualRefresh
    RefreshCounter
  '  RefreshListe
  
    sprListe.TopRow = tr
    sprListe.SetActiveCell 2, r
  
    sprListe.Row = r
    sprListe.SelModeSelected = True
  End If
  
  sprListe.ReDraw = True
  
  Screen.MousePointer = vbDefault
  
  sprListe.SetFocus
End Sub


'##ModelId=5C8A6814039F
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "Commentaire" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  ElseIf dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "DataVersion" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeStaticText
    sprListe.TypeMaxEditLen = 15
  End If
End Sub

