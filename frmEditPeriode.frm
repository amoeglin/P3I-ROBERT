VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditPeriode 
   Caption         =   "Periode du ..."
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   855
   ClientWidth     =   13725
   Icon            =   "frmEditPeriode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   13725
   Begin VB.CommandButton btnExportSAS 
      Caption         =   "&Exporter"
      Height          =   375
      Left            =   45
      TabIndex        =   12
      Top             =   5940
      Width           =   1935
   End
   Begin VB.CommandButton btnCalcRevalo 
      Caption         =   "Calculer && &Revaloriser"
      Height          =   375
      Left            =   2070
      TabIndex        =   7
      Top             =   5940
      Width           =   1935
   End
   Begin VB.CommandButton btnPurge 
      Caption         =   "&Purger"
      Height          =   375
      Left            =   4095
      TabIndex        =   6
      Top             =   5940
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   5940
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10575
      Top             =   5895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "&Importer"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   5520
      Width           =   1935
   End
   Begin VB.CommandButton btnEdition 
      Caption         =   "Choix des &Editions..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      Top             =   5535
      Width           =   1935
   End
   Begin VB.CommandButton btnCalc 
      Caption         =   "&Calculer"
      Height          =   375
      Left            =   2070
      TabIndex        =   1
      Top             =   5490
      Width           =   1935
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "Im&primer"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5535
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9900
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":1BB2
            Key             =   "openCahier"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":1CBC
            Key             =   "openPeriode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":1DC6
            Key             =   "About"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":1ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":1FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":20E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":223E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":2398
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":24F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":40B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":420E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":4368
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":44C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditPeriode.frx":461C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   8325
      Top             =   5535
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
      TabIndex        =   10
      Top             =   495
      Width           =   11040
      _Version        =   524288
      _ExtentX        =   19473
      _ExtentY        =   8758
      _StockProps     =   64
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OperationMode   =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmEditPeriode.frx":56AE
      AppearanceStyle =   0
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   13725
      _ExtentX        =   24209
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Période"
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
            Key             =   "editTable"
            Description     =   "Saisie manuelle des assurés"
            Object.ToolTipText     =   "Saisie manuelle des assurés"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Description     =   "Impression"
            Object.ToolTipText     =   "Imprimer"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sousproduit"
            Description     =   "sous-produits"
            Object.ToolTipText     =   "Préparation de la liste par sous-produits"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reass"
            Description     =   "Réassurance"
            Object.ToolTipText     =   "Préparation de la liste de réassurance"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CR"
            Description     =   "ComptesRésultats"
            Object.ToolTipText     =   "Export ComptesRésultats"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "check"
            Object.ToolTipText     =   "Valide les tables de paramètres"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "regimes"
            Object.ToolTipText     =   "Sommes par régime"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filter_name"
            Object.ToolTipText     =   "Tous les flux de l'assuré sélectionné"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "filter"
            Object.ToolTipText     =   "Filtre"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "manageDisplays"
            Object.ToolTipText     =   "Gestion d'affichages"
            ImageIndex      =   14
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   11280
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton btnDetail 
         Caption         =   "&Détails"
         Height          =   285
         Left            =   8460
         TabIndex        =   13
         Top             =   45
         Width           =   1215
      End
      Begin VB.TextBox lblFillTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   11025
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "----"
         Top             =   90
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.CommandButton btnExport 
         Caption         =   "&Exporter"
         Height          =   285
         Left            =   9765
         TabIndex        =   9
         Top             =   45
         Width           =   1215
      End
      Begin VB.TextBox lblFilter 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4365
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "fdgfhgf"
         Top             =   90
         Width           =   3840
      End
   End
End
Attribute VB_Name = "frmEditPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67A202EC"
Option Explicit

'##ModelId=5C8A67A203B7
Public autoMode As Boolean
'##ModelId=5C8A67A203C7
Public autoPeriode As Long
'##ModelId=5C8A67A203E6
Public exportFilename As String
'##ModelId=5C8A67A3001D
Public autoLogger As clsLogger

'##ModelId=5C8A67A3001E
Private fInitDone As Boolean

'##ModelId=5C8A67A3002D
Private frmNumPeriode As Long
'##ModelId=5C8A67A3004C
Private frmOrdreDeTri As String

'##ModelId=5C8A67A3005C
Private frmTypePeriode As Integer

'##ModelId=5C8A67A3007B
Private frmInExport As Boolean

'##ModelId=5C8A67A3009C
Public m_DetailAffichagePeriode As clsDetailAffichagePeriode

' filtre
'##ModelId=5C8A67A300AC
Private fmFilter As clsFilter            ' filtre en cours
'##ModelId=5C8A67A300B9
Private fmFilter_Precedent As clsFilter  ' filtre précédent (svg si filtre par nom)

'##ModelId=5C8A67A300BA
Private Sub RefreshListe()
  'Dim rs As Recordset
  
  Dim rs As ADODB.Recordset
  Dim rq As String, filter As String, sWhere As String, sResultingQuery As String, sFrom As String
  Dim i As Integer
  Dim debut As Date, fin As Date
  Dim disp As AssureDisplay
  Dim field As AssureField
  Dim colNumber As Integer
  
'  Dim dateInv As Date
'  Dim semaine As Integer
'
'  dateInv = m_dataHelper.GetParameter("Select PEDATEEXT From P3IUser.Periode Where PENUMCLE = " & frmNumPeriode)
  
  'frmOrdreDeTri = ""
  'UpdateOrderByString ("Assure.PONUMCLE")
  'UpdateOrderByString ("Assure.POARRET")
  
  Set disp = AssureDisplays.CurrentlySelectedDisplay
  
  If disp Is Nothing Then
    If Not autoMode Then
      MsgBox "Error: No default display is defined!"
      Unload Me
    Else
      autoLogger.EcritTraceDansLog "Error: No default display is defined!"
    End If
    
    Exit Sub
  End If
  
  debut = Now
  
  sFrom = "FROM Societe INNER JOIN Assure ON Societe.SOCLE = Assure.POSTECLE " _
            & " INNER JOIN Garantie ON Assure.POGARCLE = Garantie.GAGARCLE  " _
            & " LEFT JOIN CodesCat ON Assure.POGPECLE = CodesCat.GroupeCle AND Assure.POPERCLE = CodesCat.NumPeriode AND Assure.POCATEGORIE = CodesCat.Code_Cat_Contrat AND Assure.POCompagnie=CodesCat.Code_Cie AND Assure.POAppli=CodesCat.Code_APP " _
            & " LEFT JOIN CodePosition ON Assure.POSIT = CodePosition.Position " _
            & " LEFT JOIN CodeProvision ON Assure.POCATEGORIE_NEW = CodeProvision.CodeProv " _
            & " LEFT JOIN TypeTermeEchu ON Assure.POECHU = TypeTermeEchu.IdTypeTermeEchu " _
            & " LEFT JOIN TypeFractionnement ON Assure.POFRACT = TypeFractionnement.IdTypeFractionnement " _
            & " LEFT JOIN ParamCalcul ON Assure.POGPECLE = ParamCalcul.PEGPECLE AND Assure.POPERCLE = ParamCalcul.PENUMCLE AND Assure.PONumParamCalcul = ParamCalcul.PENUMPARAMCALCUL " _
            & " LEFT JOIN SituationFamille ON Assure.POCleSituationFamille = SituationFamille.CleSituationFamille "
   
  sWhere = " WHERE Assure.POPERCLE = " & frmNumPeriode & " AND Assure.POGPECLE = " & GroupeCle
  
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
'  Me.Caption = "Assurés de la période " & frmNumPeriode
'  Me.Caption = Me.Caption & " (" & m_dataHelper.GetParameterAsStringCRW("SELECT 'Type ' + CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + T.Libelle FROM Periode P LEFT JOIN TypePeriode T ON T.IdTypePeriode=P.PETYPEPERIODE WHERE P.PEGPECLE = " & GroupeCle & " AND P.PENUMCLE = " & frmNumPeriode)
'  Me.Caption = Me.Caption & ") du Groupe '" & NomGroupe & "' : " & m_dataHelper.GetParameterAsStringCRW("SELECT CASE WHEN LEN(PECOMMENTAIRE)>40 THEN left(PECOMMENTAIRE, 40)+'...' ELSE PECOMMENTAIRE END as COMMENTAIRE FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle)
  
  lblFilter = fmFilter.SelectionString
  filter = fmFilter.BuildFilterString
  
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
  
    
  sprListe.DAutoCellTypes = True
  sprListe.DAutoSizeCols = True
  
  ' Type de période
  Dim typePeriode As Integer
  typePeriode = m_dataHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE=" & frmNumPeriode)
  frmTypePeriode = typePeriode
    
  'create query string
  rq = "SELECT Assure.RECNO, "
  For Each field In disp.AssureFields
     rq = rq & field.DBQuery & " as " & field.DispalyField & ", "
  Next
  
  'delete last colon and add a space
  rq = Trim(rq)
  rq = Left(rq, Len(rq) - 1)
  rq = rq + " "
  
  
'  Round(Assure.POPM_INCAP_1F*Assure.POPRESTATION_AN/365,2) as [PM Incap],
'  Round(Assure.POPM_PASS_1F*Assure.POPRESTATION_AN_PASSAGE/365,2) as [PM Pass],
'  Round(Assure.POPM_INVAL_1F*Assure.POPRESTATION_AN/365,2) as [PM Inval],

'  PM MO/AT/MAT
'  PM PASS CLM/CLM
'  PM PASS CLD/CLD
  
  'some modifications are required for statutaire
  If typePeriode = 6 Then
  
    'semaine = Int(DateDiff("d", Assure.POARRET, dateInv) / 7) + 1
    'DATEDIFF(day,'2014-08-05','2014-06-05')
    'rq = Replace$(rq, "Assure.POPM_ANC as Anc", "(DATEDIFF(day, Assure.POARRET, dateInv) / 7) +1 as Anc")
    
    rq = Replace$(rq, "POPM_INCAP_1F*Assure.POPRESTATION_AN/12,2) as [PM Incap]", "POPM_INCAP_1F*Assure.POPRESTATION_AN/365,2) as [PM MO/AT/MAT]")
    rq = Replace$(rq, "POPM_PASS_1F*Assure.POPRESTATION_AN_PASSAGE,2) as [PM Pass]", "POPM_PASS_1F*Assure.POPRESTATION_AN/365,2) as [PM PASS CLM/CLM]")
    rq = Replace$(rq, "POPM_INVAL_1F*Assure.POPRESTATION_AN,2) as [PM Inval]", "POPM_INVAL_1F*Assure.POPRESTATION_AN/365,2) as [PM PASS CLD/CLD]")
  End If
  
  'finalise the query string
  If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
    If m_DetailAffichagePeriode.Avant Then
      sResultingQuery = ProcessQuery(rq, sFrom, sWhere, filter, "1 - Avant")
    End If
    
    If m_DetailAffichagePeriode.Apres Then
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "2 - Après")
    End If
    
    If m_DetailAffichagePeriode.Ecart Then
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "3 - Ecart")
    End If
    
    If m_DetailAffichagePeriode.DejaAmorti Then
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "4 - Amorti")
    End If
    
    If m_DetailAffichagePeriode.ResteAAmortir Then
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "5 - A Amortir")
    End If
    
    ' Tri
    Dim moreOrderBy As String
    moreOrderBy = CheckOrderByField(sResultingQuery, "NUENRP3I", False)
    moreOrderBy = moreOrderBy & CheckOrderByField(sResultingQuery, "Garantie", True)
    moreOrderBy = moreOrderBy & CheckOrderByField(sResultingQuery, "[Base/Revalo]", True)
    moreOrderBy = moreOrderBy & CheckOrderByField(sResultingQuery, "TypeLigne", True)
    
    If frmOrdreDeTri = "" Then
      moreOrderBy = Trim(moreOrderBy)
      If InStr(moreOrderBy, ",") = 1 Then
        moreOrderBy = mID(moreOrderBy, 2)
      End If
    End If
    
    If frmOrdreDeTri = "" And frmOrdreDeTri = "" Then
      sResultingQuery = sResultingQuery
    Else
      sResultingQuery = sResultingQuery & " ORDER BY " & frmOrdreDeTri & moreOrderBy
    End If
    
    'sResultingQuery = sResultingQuery & " ORDER BY " & frmOrdreDeTri & IIf(InStr(1, frmOrdreDeTri, "NUENRP3I") = 0, ", NUENRP3I", "") & IIf(InStr(1, frmOrdreDeTri, "Garantie") = 0, ", Garantie", "") & IIf(InStr(1, frmOrdreDeTri, "[Base/Revalo]") = 0, ", [Base/Revalo]", "") & IIf(InStr(1, frmOrdreDeTri, "TypeLigne") = 0, ", TypeLigne", "")
  Else
    If frmOrdreDeTri = "" Then
      sResultingQuery = rq & sFrom & sWhere & filter
    Else
      sResultingQuery = rq & sFrom & sWhere & filter & " ORDER BY " & frmOrdreDeTri
    End If
    
  End If
  
  sResultingQuery = m_dataHelper.ValidateSQL(sResultingQuery)


  'Set dataFill event for certain columns
  If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
    SetColonneDataFill GetSpreadColNumber(sResultingQuery, "TypeLigne"), True
    SetColonneDataFill GetSpreadColNumber(sResultingQuery, "Garantie"), True
  Else
    SetColonneDataFill GetSpreadColNumber(sResultingQuery, "Garantie"), True
  End If
  
  
'####
'starttest:
'Dim start As Single
'start = Timer
  
  dtaPeriode.RecordSource = sResultingQuery
  dtaPeriode.Refresh
  

'####
'Debug.Print "Duration: " & Format(Timer - start, "0.000")
'MsgBox "Duration " & NumPeriode & " : " & Format(Timer - start, "0.000")
    
    
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If Not dtaPeriode.Recordset.EOF Then
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
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
  
  
  '### freeze columns: RECNO , NUENRP3I , TypeLigne
  If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
    'sprListe.ColsFrozen = 3
  Else
    'sprListe.ColsFrozen = 2
  End If
  
  
  ' Couleurs des colonnes
  SetColumnColors sResultingQuery
  
  
  'change le format des colonnes pour trier correctement les dates
#If TRI_SPREAD Then
  SetColumnDate sResultingQuery
#End If
   
  
  ' largeur des colonnes
  For i = 2 To sprListe.MaxCols
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 5
  Next i
  
  'only for comment cols
  colNumber = GetSpreadColNumber(sResultingQuery, "[Raison Annulation]")
  If colNumber <> -1 Then sprListe.ColWidth(colNumber) = 50
  colNumber = GetSpreadColNumber(sResultingQuery, "Commentaire")
  If colNumber <> -1 Then sprListe.ColWidth(colNumber) = 50
  
  
pas_de_donnee:

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True

  If Not autoMode Then
    Me.SetFocus
    sprListe.SetFocus
    
    Dim bLocked As Boolean
    
    bLocked = CBool(m_dataHelper.GetParameterAsDouble("SELECT PELOCKED FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle))
    
    If bLocked = True Then
      btnCalc.Enabled = False
      btnCalcRevalo.Enabled = False
      btnImport.Enabled = False
      btnPurge.Enabled = False
    Else
      If archiveMode Then
        btnCalcRevalo.Enabled = False
        btnCalc.Enabled = False
        btnExportSAS.Enabled = False
        btnImport.Enabled = False
        
        'disable edit button on Toolbar
        Toolbar1.Buttons(3).Enabled = False
      Else
        btnCalc.Enabled = True
        btnCalcRevalo.Enabled = True
        btnImport.Enabled = True
        btnPurge.Enabled = True
      End If
    End If
    
    Screen.MousePointer = vbDefault
  
    fin = Now
    
    lblFillTime.text = "Remplissage : " & DateDiff("s", debut, fin) & " s"
  End If
  
  Exit Sub

err_RefreshListe:

  If Not autoMode Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    autoLogger.EcritTraceDansLog "Erreur " & Err & vbLf & Err.Description
  End If
  
  Resume Next
  
End Sub

'##ModelId=5C8A67A300C9
Private Sub SetColumnDate(sql As String)

  Dim disp As AssureDisplay
  Dim field As AssureField
    
  Set disp = AssureDisplays.CurrentlySelectedDisplay
  
  For Each field In disp.AssureFields
    If field.IsDateColumn Then
      SetColonneDeTypeDate GetSpreadColNumber(sql, field.DispalyField)
    End If
  Next
  
End Sub

'##ModelId=5C8A67A300F8
Private Sub SetColumnColors(sql As String)

  Dim disp As AssureDisplay
  Dim field As AssureField
  Dim Col As OLE_COLOR
  
  Set disp = AssureDisplays.CurrentlySelectedDisplay
  
  For Each field In disp.AssureFields
    Select Case field.SpreadColor
      Case "bleu_clair"
        Col = bleu_clair
      Case "jaune_clair"
        Col = jaune_clair
      Case "orange_clair"
        Col = orange_clair
      Case "vert_clair"
        Col = vert_clair
      Case "lavande_clair"
        Col = lavande_clair
        
      Case Else
        Col = vbWindowBackground
    End Select
    
    SetColBackColor GetSpreadColNumber(sql, field.DispalyField), 1, Col
  Next
  
End Sub

'##ModelId=5C8A67A30117
Private Function GetSpreadColNumber(sql As String, field As String) As Integer

  Dim arrCols() As String
  Dim sqlSelect As String
  Dim colName As String
  Dim i As Integer
  Dim fromPos As Integer
  Dim asPos As Integer
  
  fromPos = InStr(1, sql, "FROM")
  If fromPos > 0 Then
    sqlSelect = Left(sql, fromPos - 1)
  End If
  
  arrCols = Split(sqlSelect, ",")
  
  GetSpreadColNumber = -1
  For i = 0 To UBound(arrCols)
    asPos = InStr(1, arrCols(i), " as ")
    colName = arrCols(i)
    If asPos > 0 Then
      colName = mID$(arrCols(i), asPos + 4)
    End If
    
    If InStr(1, colName, field) > 0 Then
      GetSpreadColNumber = i + 1
      Exit For
    End If
  Next
  
End Function

'##ModelId=5C8A67A30156
Private Sub SetColumnDataFillEvent(field As String, fActive As Boolean)
  
  sprListe.sheet = sprListe.ActiveSheet
  'sprListe.Col = numCol
  sprListe.DataFillEvent = fActive
 
End Sub

'##ModelId=5C8A67A30175
Private Function CheckOrderByField(sqlQuery As String, field As String, addComma As Boolean) As String

  Dim sqlField As String
  Dim orderField As String
  
  CheckOrderByField = ""
  sqlField = field
  orderField = field
  
  'there are several items that contain 'Garantie', so we need to do an additional test
  If field = "Garantie" Then
    sqlField = "Garantie.GALIB"
    orderField = "Garantie"
  End If
  If field = "[Base/Revalo]" Then
    sqlField = "Base/Revalo"
    orderField = "Base/Revalo"
  End If
  
'  fld = Replace(fld, "[", "")
'  fld = Replace(fld, "]", "")
  
  If InStr(1, sqlQuery, sqlField) > 0 And InStr(1, frmOrdreDeTri, orderField) = 0 Then
    If frmOrdreDeTri <> "" Or addComma Then
      CheckOrderByField = " ," & field
    Else
      CheckOrderByField = field
    End If
  End If
  
End Function

'##ModelId=5C8A67A301C3
Private Sub UpdateOrderByString(field As String)

  Dim disp As AssureDisplay
  
  Set disp = AssureDisplays.CurrentlySelectedDisplay
  
  If Not disp Is Nothing Then
    If InStr(1, frmOrdreDeTri, field) = 0 Then
      If disp.ContainsFieldName(field) Then
        If frmOrdreDeTri <> "" Then
          frmOrdreDeTri = frmOrdreDeTri & " ," & field
        Else
          frmOrdreDeTri = field
        End If
      End If
    End If
  End If
  
End Sub




'**********************************






'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Charge le moteur de calcul
'
'##ModelId=5C8A67A301E2
Private Function LoadModuleCalcul(ByRef nomModuleCalcul As String) As iP3ICalcul
  'Dim nomModuleCalcul As String
  
  On Error GoTo errLoadModuleCalcul
  
  ' charge l'object de calcul
  nomModuleCalcul = GetSettingIni(CompanyName, SectionName, "ModuleCalcul", "#")
  
  If nomModuleCalcul = "#" Then
    MsgBox "Le paramètre ModuleCalcul n'est pas présent dans " & sFichierIni & "," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
    Exit Function
  End If
  
  Select Case UCase(nomModuleCalcul)
    Case "P3ICALCUL_GENERALI"
      Set LoadModuleCalcul = New P3ICalcul_Generali
    
    Case "P3ICALCUL_ISICA"
      MsgBox "Le module de calcul " & nomModuleCalcul & " n'est pas supporté dans cette version de P3I." & vbLf & "Veuillez corriger " & sFichierIni & "," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
      Exit Function
      
    Case Else
      MsgBox "Le module de calcul " & nomModuleCalcul & " est inconnu." & vbLf & "Veuillez corriger " & sFichierIni & "," & vbLf & "Le programme n'a pas été correctement installé :" & vbLf & VeuillezContacterMoeglin, vbCritical
      Exit Function
  End Select
  
  Exit Function

errLoadModuleCalcul:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Set LoadModuleCalcul = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' calcul des provisions
'
'##ModelId=5C8A67A30202
Private Sub DoCalculProvision(avecRevalo As Boolean)
  Dim module_calcul As iP3ICalcul, nomModuleCalcul As String
  Dim Logger As clsLogger
  
  On Error GoTo errCalcul
  
  ' charge l'object de calcul
  Set module_calcul = LoadModuleCalcul(nomModuleCalcul)
  If module_calcul Is Nothing Then Exit Sub
  
  ' prépare le log
  Set Logger = New clsLogger
  
  Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurCalcul.log"
  
  ' effectue le calcul
  frmMain.Enabled = False
  module_calcul.CalculProvisionsAssures avecRevalo, frmNumPeriode, CLng(GroupeCle), Logger ' appel de la fonction de calcul des provisions pour les assurés
  frmMain.Enabled = True
  
  ' rafraichi la liste
  RefreshListe
  
  Screen.MousePointer = vbDefault
  
  
  ' affiche les erreurs
  Logger.AfficheErreurLog
  
  Exit Sub

errCalcul:
  MsgBox "Erreur durant le calcul : " & Err & vbLf & Err.Description & vbLf & "Objet = " & nomModuleCalcul, vbCritical
  Exit Sub
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' calcul des provisions sans revalorisation
'
'##ModelId=5C8A67A30221
Private Sub btnCalc_Click()
  DoCalculProvision False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' calcul des provisions avec revalorisation
'
'##ModelId=5C8A67A30231
Private Sub btnCalcRevalo_Click()
  DoCalculProvision True
End Sub

'##ModelId=5C8A67A30240
Private Sub btnClose_Click()
  
  'Unlock Periode
  m_dataSource.Execute "Delete From LockedPeriods Where UserName Like '" & user_name & "'"
  
  Unload Me
End Sub

'##ModelId=5C8A67A30250
Private Sub BuildDescription()
   ' fabrique le titre de la fenetre en fonction du groupe en cours
  Dim rs As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle, Snapshot)
  
  If Not rs.EOF Then
    Dim dd As String, df As String
    
    dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & frmNumPeriode), "dd/mm/yyyy")
    df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & frmNumPeriode), "dd/mm/yyyy")

    DescriptionPeriode = " Période " & frmNumPeriode & " ( " & dd & " au " & df & " ) " & vbLf & " du Groupe " & rs.fields("Nom")
  Else
    DescriptionPeriode = "  Erreur ... "
  End If
  rs.Close
End Sub

'##ModelId=5C8A67A3025F
Private Sub btnDetail_Click()
  Dim frm As frmDetailAffichagePeriode
  
  Set frm = New frmDetailAffichagePeriode
  
  Set frm.m_DetailAffichagePeriode = m_DetailAffichagePeriode
  frm.m_TypePeriode = frmTypePeriode
  
  frm.Show vbModal
  If frm.ret_code = 1 Then
    ' initialise l'order de tri : n° de police/date d'arret
    'frmOrdreDeTri = "Assure.PONUMCLE, Assure.POARRET"
    frmOrdreDeTri = ""
    UpdateOrderByString ("Assure.PONUMCLE")
    UpdateOrderByString ("Assure.POARRET")
    
    RefreshListe
  End If
  
  Set frm = Nothing
End Sub

'##ModelId=5C8A67A3027F
Private Sub btnEdition_Click()
  Dim FM As frmChoixEdition
  
  'SoCle = cboSociete.ItemData(cboSociete.ListIndex)
  SoCle = 0
  frmWait.Show vbModeless
  
  frmWait.Caption = "Chargement en cours en cours..."
  
  frmWait.ProgressBar1.Min = 0
  frmWait.ProgressBar1.Value = 0
  frmWait.ProgressBar1.Max = 1
   
   ' fabrique le titre de la fenetre en fonction du groupe en cours
  Call BuildDescription
  
  Set FM = New frmChoixEdition
  
  Set FM.fmFilter = fmFilter
  FM.frmNumPeriode = frmNumPeriode
  
  FM.Show vbModal
  
  Set FM = Nothing
End Sub

'##ModelId=5C8A67A3028E
Private Sub btnExport_Click()
  ExportToExcel
End Sub

'##ModelId=5C8A67A3029E
Public Sub ExportToExcel()

  On Error GoTo err_export
  
  If Not autoMode Then
    CommonDialog1.filename = "Periode" & frmNumPeriode & ".xls"
    CommonDialog1.filter = "Fichier Excel|*.xls|Base de données MS Access|*.mdb|"
    
    CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
    CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    CommonDialog1.filename = ""
    
    CommonDialog1.ShowSave
    
    If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.mdb" Or CommonDialog1.filename = "*.xls" Then
      Exit Sub
    End If
    
    frmInExport = True
    
    If Right(UCase(CommonDialog1.filename), 4) = ".XLS" Then
      
      Screen.MousePointer = vbHourglass
      Dim fWait As frmWait
      Set fWait = New frmWait
      fWait.Caption = "Export vers en cours..."
      fWait.Label1(0).Caption = "Export vers " & CommonDialog1.filename & " en cours..."
      fWait.ProgressBar1.Min = 0
      fWait.ProgressBar1.Value = 0
      fWait.ProgressBar1.Max = 100
      fWait.ProgressBar1.Value = 50
      fWait.Show vbModeless
      fWait.Refresh
      
      ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, CommonDialog1.filename, "Assure", sprListe
      
      If Not autoMode Then
        fWait.Hide
        Unload fWait
        Set fWait = Nothing
      End If
    
    Else
      Dim exportModule As P3IExport.iExport
      Set exportModule = New P3IExport.iExport
      exportModule.ExportDBAccess CommonDialog1, m_dataSource, GroupeCle, numPeriode
      Set exportModule = Nothing
    End If
    
    frmInExport = False
    
  Else
    'auto mode
    ExportQueryResultToExcel m_dataSource, dtaPeriode.RecordSource, exportFilename, "Assure", sprListe, "", autoMode, autoLogger
  
  End If
  
    Exit Sub
  
err_export:
  
  If Not autoMode Then
    If Err <> cdlCancel Then
      MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
    End If
    CommonDialog1.CancelError = False
  Else
    autoLogger.EcritTraceDansLog "Erreur " & Err & vbLf & Err.Description
  End If
  
End Sub


'##ModelId=5C8A67A302AD
Private Sub btnImport_Click()
  numPeriode = frmNumPeriode
  
  frmListeJeuxDonnees.Show vbModal
  
  RefreshListe
End Sub


'##ModelId=5C8A67A302BD
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
  
  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & frmNumPeriode), "dd/mm/yyyy")
  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & frmNumPeriode), "dd/mm/yyyy")

  DescriptionPeriode = " période " & frmNumPeriode & " ( " & dd & " au " & df & " ) " & " du Groupe " & NomGroupe
 
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
    .PrintJobName = "Assurés de la période " & frmNumPeriode
    .PrintHeader = "/c Assurés de la " & DescriptionPeriode & "/n  "
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

'##ModelId=5C8A67A302CD
Private Sub btnPurge_Click()

' PURGE DESACTIVEE
'  Dim rs As ADODB.Recordset
'
'  If sprListe.MaxRows = 0 Then Exit Sub
'
'  Load frmSelectPurge
'
'  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT PODATEIMPORT FROM Assure WHERE POPERCLE = " & frmNumPeriode & " AND POGPECLE = " & GroupeCle, Snapshot)
'
'  Do Until rs.EOF
'    If Not IsNull(rs.fields(0)) Then
'      frmSelectPurge.lstDate.AddItem Format(rs.fields(0), "dd/mm/yyyy hh:nn")
'    End If
'
'    rs.MoveNext
'  Loop
'
'  rs.Close
'
'  If frmSelectPurge.lstDate.ListCount = 0 Then
'    MsgBox "Il n'y a aucun salarié à purger !", vbInformation
'    Unload frmSelectPurge
'    Exit Sub
'  End If
'
'  frmSelectPurge.Show vbModal, frmMain
'  If ret_code = 0 Then
'    Dim d As Date
'
'    d = CDate(frmSelectPurge.lstDate.List(frmSelectPurge.lstDate.ListIndex))
'    m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE * FROM Assure WHERE POPERCLE = " & frmNumPeriode & " AND POGPECLE = " & GroupeCle _
'                  & " AND PODATEIMPORT = '" & Format(d, "yyyy-mm-dd") & " " & Format(d, "hh:nn") & "'")
'
'    Unload frmSelectPurge
'
'    RefreshListe
'  End If
End Sub

'##ModelId=5C8A67A302DC
Private Sub cboSociete_Click()
  If fInitDone Then
    RefreshListe
  End If
End Sub

'##ModelId=5C8A67A302FC
Private Sub chkDonneesBrutes_Click()
  ' initialise l'order de tri : n° de police/date d'arret
  'frmOrdreDeTri = "Assure.PONUMCLE, Assure.POARRET"
  frmOrdreDeTri = "Police, [Date du sinistre]"
  
  RefreshListe
End Sub

'##ModelId=5C8A67A3030B
Private Function BuildFilter() As clsFilter
  Dim theFilter As clsFilter
  
  Set theFilter = New clsFilter
  
  theFilter.numPeriode = frmNumPeriode
  theFilter.NumGroupe = GroupeCle
  
  ' définition des filtres disponibles
  theFilter.AddFilterElem Number, "Société", "Assure.POSTECLE", "{Assure.POSTECLE}", "SELECT SONOM FROM Societe WHERE SOCLE=<Value> AND SOGROUPE=<CleGroupe>", False, False, False, True
  theFilter.AddFilterElem Number, "Régime", "Assure.POGARCLE", "{Assure.POGARCLE}", "SELECT CleGarantie=CASE WHEN GAGARCLE>90 THEN GAGARCLE WHEN GAGARCLE>50 AND GAGARCLE<90 THEN GAGARCLE-50 END FROM Garantie WHERE GAGARCLE=<Value>", False, False, False, True
  theFilter.AddFilterElem Char, "Catégorie", "Assure.POCATEGORIE", "{Assure.POCATEGORIE}", "", False, False, False, True
  theFilter.AddFilterElem Char, "Nom", "Assure.PONOM", "{Assure.PONOM}", "", False, False, True, True
  theFilter.AddFilterElem Char, "Contrat", "Assure.POCONVENTION", "{Assure.POCONVENTION}", "", False, False, False, True
  theFilter.AddFilterElem Char, "Police", "Assure.PONUMCLE", "{Assure.PONUMCLE}", "", False, False, True, True
  theFilter.AddFilterElem Char, "Code GE", "Assure.POGARCLE_NEW", "{Assure.POGARCLE_NEW}", "", True, False, False, True
  theFilter.AddFilterElem Char, "Code Provision", "Assure.POCATEGORIE_NEW", "{Assure.POCATEGORIE_NEW}", "", True, False, False, True
  theFilter.AddFilterElem Char, "Code Position", "Assure.POSIT", "{Assure.POSIT}", "", True, False, False, True
  theFilter.AddFilterElem Char, "RegroupAnnexe", "Assure.POREGROUPEMENT", "{Assure.POREGROUPEMENT}", "", True, False, False, True
  theFilter.AddFilterElem Char, "RegroupStat", "Assure.POCODENATURE", "{Assure.POCODENATURE}", "", True, False, False, True
  theFilter.AddFilterElem Number, "CCN", "Assure.POCCN", "{Assure.POCCN}", "", True, False, False, True
  theFilter.AddFilterElem Char, "NUENRP3I", "Assure.NUENRP3I", "{Assure.NUENRP3I}", "", False, False, False, True
  
  Set BuildFilter = theFilter
End Function

'##ModelId=5C8A67A3032A
Private Sub Command1_Click()

'sprListe.ExportToTextFile "c:\aaa.txt", "", ";", "", ExportToTextFileCreateNewFile, "c:\log.txt"
'sprListe.SaveToFile "c:\aaa.txt", True
'sprListe.SaveTabFile "c:\aaa.txt"

End Sub

'##ModelId=5C8A67A3033A
Private Sub Form_Activate()
  FormActivate
End Sub

'##ModelId=5C8A67A3034A
Public Sub FormActivate()

  If autoMode Then
    frmNumPeriode = autoPeriode
  Else
    frmNumPeriode = numPeriode
  End If
  
  ' Init du filtre
  If fmFilter Is Nothing Then
    Set fmFilter = BuildFilter
    
    ' recharge le dernier filtre
    fmFilter.Load frmNumPeriode
  End If
  
  'frmOrdreDeTri = "Assure.PONUMCLE,Assure.POARRET"
  frmOrdreDeTri = ""
  UpdateOrderByString ("Assure.PONUMCLE")
  UpdateOrderByString ("Assure.POARRET")
  
  RefreshListe
End Sub


'##ModelId=5C8A67A30369
Private Sub Form_Load()
  FormLoad
End Sub

'##ModelId=5C8A67A30379
Public Sub FormLoad()
  Dim rs As ADODB.Recordset
    
  'required for display management
  Set AssureDisplays = New AssureDisplays
  AssureDisplays.InitDisplayObjects
  
  btnCalcRevalo.Enabled = False
  btnCalc.Enabled = False
  ' test de branchement de l'ancienne methode AM et AG le 27/10/2023
 ' btnExportSAS.Enabled = False
  btnImport.Enabled = False
  
  'disable edit button on Toolbar
  Toolbar1.Buttons(3).Enabled = False
  
  Set m_DetailAffichagePeriode = New clsDetailAffichagePeriode

  If autoMode Then
    Me.WindowState = vbMinimized
    frmNumPeriode = autoPeriode
  Else
    frmNumPeriode = numPeriode
  End If
  
  Me.Caption = "Assurés de la période " & frmNumPeriode
  Me.Caption = Me.Caption & " (" & m_dataHelper.GetParameterAsStringCRW("SELECT 'Type ' + CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + T.Libelle FROM Periode P LEFT JOIN TypePeriode T ON T.IdTypePeriode=P.PETYPEPERIODE WHERE P.PEGPECLE = " & GroupeCle & " AND P.PENUMCLE = " & frmNumPeriode)
  Me.Caption = Me.Caption & ") du Groupe '" & NomGroupe & "' : " & m_dataHelper.GetParameterAsStringCRW("SELECT CASE WHEN LEN(PECOMMENTAIRE)>40 THEN left(PECOMMENTAIRE, 40)+'...' ELSE PECOMMENTAIRE END as COMMENTAIRE FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle)
  

  frmInExport = False
  fInitDone = False
  
  ' initialise l'order de tri : n° de police/date d'arret
  'frmOrdreDeTri = "Assure.PONUMCLE,Assure.POARRET"
  frmOrdreDeTri = ""
  UpdateOrderByString ("Assure.PONUMCLE")
  UpdateOrderByString ("Assure.POARRET")
  
  ' init des controles
  If archiveMode Then
    CreateArchiveConnection
    m_dataSourceArchive.SetDatabase dtaPeriode
  Else
    m_dataSource.SetDatabase dtaPeriode
  End If
  
  ' chargement du masque du spread
  sprListe.LoadFromFile App.Path & "\EditPeriode.ss6"

  Set fmFilter_Precedent = Nothing

  fInitDone = True
End Sub

'##ModelId=5C8A67A30388
Private Sub SetColonneDataFill(numCol As Integer, fActive As Boolean)
  
  Dim i As Integer
  
  If numCol = -1 Then Exit Sub
  
  'For i = 2 To sprListe.MaxCols
  '  sprListe.Col = i
    sprListe.sheet = sprListe.ActiveSheet
    sprListe.Col = numCol
    sprListe.DataFillEvent = fActive
  'Next
'  sprListe.Col = numCol
'  sprListe.Col2 = numCol
'  sprListe.Row = 1
'  sprListe.Row2 = sprListe.MaxRows
'  sprListe.BlockMode = True
'  sprListe.DataFillEvent = True
'  sprListe.BlockMode = False
End Sub

'##ModelId=5C8A67A303B7
Private Sub SetColonneDeTypeDate(numCol As Integer)

  If numCol = -1 Then Exit Sub
  
  sprListe.Col = numCol
  sprListe.Col2 = numCol
  sprListe.Row = 1
  sprListe.Row2 = sprListe.MaxRows
  sprListe.BlockMode = True
  sprListe.CellType = SS_CELL_TYPE_DATE
  sprListe.TypeDateCentury = True
  sprListe.TypeDateFormat = SS_CELL_DATE_FORMAT_DDMMYY

  sprListe.BlockMode = False
End Sub

'##ModelId=5C8A67A303D6
Private Sub SetColBackColor(numCol As Integer, nbCol As Integer, color As OLE_COLOR)
  
  If numCol = -1 Then Exit Sub
  
  sprListe.BlockMode = True
  sprListe.Col = numCol
  sprListe.Col2 = sprListe.Col + nbCol - 1
  sprListe.Row = -1
  sprListe.Row2 = -1
  sprListe.BackColor = color
  sprListe.BackColorStyle = BackColorStyleUnderGrid
  sprListe.BlockMode = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Accepted only 0-9, A-Z, a-z and '_'
'
'##ModelId=5C8A67A4003C
Private Function CheckChar(c As Integer) As Boolean
  CheckChar = False
  
  ' If ((c >= Asc("0") And c <= Asc("9")) Or (c >= Asc("A") And c <= Asc("Z")) Or (c >= Asc("a") And c <= Asc("z")) Or (c = Asc("_"))) Then
  If ((c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c = 95)) Then
    CheckChar = True
  End If

End Function


'##ModelId=5C8A67A4005C
Private Function FindNextSeparator(s As String, Start As Integer) As Integer
  Dim i As Integer
  
  For i = Start To Len(s)
    If CheckChar(Asc(mID(s, i, 1))) = False Then
      FindNextSeparator = i
      Exit Function
    End If
  Next
  
  FindNextSeparator = 0
End Function


'##ModelId=5C8A67A4009A
Private Function ProcessThisField(sName As String, TypeLigne As Integer) As Boolean
  Dim i As Integer
  Dim astrSplitItems() As String
  Dim sListFields As String
  
  
  ' Liste des champs rejetés
  If TypeLigne = 3 Then ' Ecart
    sListFields = "POPM_AVECCORRECTIF,POPM_SANSCORRECTIF,POPMREASSAVECCORRECTIF,POPMANNULEE,POPSAPANNULEE,PODATEPAIEMENTESTIMEE," _
                & "POPSAPCAPMOYEN,PODOSSIERCLOS,POPMReassAvecCorrectif,POPMAvecCorrectif,POCaptive,POTopAmortissable,POIdAssureMDC"
  Else
    sListFields = "POPM_AVECCORRECTIF,POPM_SANSCORRECTIF,POPMREASSAVECCORRECTIF,POPMANNULEE,POPSAPANNULEE,PODATEPAIEMENTESTIMEE," _
                & "POPSAPCAPMOYEN,PODOSSIERCLOS,POPMReassAvecCorrectif,POPMAvecCorrectif,POCaptive,POTopAmortissable,POIdAssureMDC" _
                & "POPM_X,POPM_XTERME,POPM_ANC,POPM_DUREE"
  End If
  
  astrSplitItems = Split(UCase(sListFields), ",")
  For i = 0 To UBound(astrSplitItems)
    If Left(sName, Len(astrSplitItems(i))) = astrSplitItems(i) Then
      ProcessThisField = False
      Exit Function
    End If
  Next

  
  ' Liste des champs acceptésPONbJourIndemn
  If TypeLigne = 3 Then ' Ecart
    sListFields = "POPREST,POPM,POPSAP,POCOT,PONbJourIndemn,POMontantBase,POMontantRevalo"
  Else
    sListFields = "POPREST,POPM,POPSAP,POCOT"
  End If
  
  astrSplitItems = Split(UCase(sListFields), ",")
  For i = 0 To UBound(astrSplitItems)
    If Left(sName, Len(astrSplitItems(i))) = astrSplitItems(i) Then
      ProcessThisField = True
      Exit Function
    End If
  Next
  
  ProcessThisField = False

End Function


'##ModelId=5C8A67A400D9
Private Function ProcessQuery(rq As String, sFrom As String, sWhere As String, filter As String, sType As String) As String
  Dim pos As Integer, i As Integer
  Dim sTemp As String, sAvant As String, sApres As String, sField As String, sChange As String, sFieldAvant As String
  
  On Error GoTo err_ProcessQuery
  
  Select Case sType
    Case "1 - Avant"
      ' Ligne Avant
'      ProcessQuery = Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '1 - Avant' as TypeLigne,"), "Assure.POCOMMENTANNUL AS Commentaire", "Null as [Top Amortissable], Null as [Coeff Amortissement], Null as [Age Mini Départ Retraite], Null as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS Commentaire ") _
'                     & sFrom & sWhere & filter
'      ProcessQuery = Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, (CASE WHEN Assure.POSIT IN(1,2,3,90) THEN '1 - Avant' ELSE '0 - Hors' END) as TypeLigne,"), "Assure.POCOMMENTANNUL AS Commentaire", "Null as [Top Amortissable], Null as [Coeff Amortissement], Null as [Age Mini Départ Retraite], Null as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS Commentaire ") _
'                     & sFrom & sWhere & filter
      ProcessQuery = Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, (CASE WHEN Assure.POGARCLE=59 OR Assure.POGARCLE=96 THEN '1 - Avant' ELSE '0 - Hors' END) as TypeLigne,"), "Assure.POCOMMENTANNUL AS [Raison Annulation]", "Null as [Top Amortissable], Null as [Coeff Amortissement], Null as [Age Mini Départ Retraite], Null as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS [Raison Annulation] ") _
                     & sFrom & sWhere & filter

    Case "2 - Après"
      ' - Remplacement de Assure par Assure_Retraite
      ' - Ajout des champs après réforme
      ' - Ajout de la jointure
      ProcessQuery = Replace(Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '2 - Après' as TypeLigne,"), "Assure.POCOMMENTANNUL AS [Raison Annulation]", "Assure.POTopAmortissable as [Top Amortissable], 100*Assure.POCoeffAmortissement as [Coeff Amortissement], Assure.POAgeMiniDepartRetraite as [Age Mini Départ Retraite], Assure.POAgeDepartRetraiteTauxPlein as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS [Raison Annulation] "), "Assure", "Assure_Retraite") _
                   & Replace(sFrom, "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO AND Assure_Retraite.POGARCLE=Assure.POGARCLE " _
                   & Replace(sWhere & filter, "Assure", "Assure_Retraite")
  
    Case "3 - Ecart"
      ' - Remplacement de Assure.xxx par (Assure_Retraite.xxx-Assure.xxx)
      ' - Ajout des champs après réforme
      ' - Ajout de la jointure
      
      sTemp = Replace(Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '3 - Ecart' as TypeLigne,"), "Assure.POCOMMENTANNUL AS [Raison Annulation]", "Null as [Top Amortissable], Null as [Coeff Amortissement], Null as [Age Mini Départ Retraite], Null as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS [Raison Annulation] "), "Assure", "Assure_Retraite")
      
      ' Traitement des champs POPREST...,POPM...,POPSAP,POCOT...
      pos = 1
      Do
        pos = InStr(pos, sTemp, "Assure_Retraite.")
        If pos <> 0 Then
          ' pos = pointe sur "Assure.PO..."
          ' Recherche de la fin du champ
          If ProcessThisField(UCase(mID(sTemp, pos + 16, 30)), 3) Then
            i = FindNextSeparator(sTemp, pos + 16)
            If i <> 0 Then
              sFieldAvant = mID(sTemp, pos - 1, i - pos + 1)
              If sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN" And sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN_PASSAGE" Then
                sAvant = Left(sTemp, pos - 1)
                sApres = mID(sTemp, i)
                sField = mID(sTemp, pos + 16, i - pos - 16)
                sFieldAvant = mID(sTemp, pos - 1, i - pos + 1)
                sChange = "(Assure_Retraite." & sField & "-Assure." & sField & ")"
                
                sTemp = sAvant & sChange & sApres
                
                pos = pos + Len(sChange)
              Else
                pos = pos + (i - pos)
              End If
            Else
              Exit Do
            End If
          Else
            pos = pos + 1
          End If
        End If
      Loop Until pos = 0
    
      sTemp = sTemp & Replace(sFrom, "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO AND Assure_Retraite.POGARCLE=Assure.POGARCLE " _
                    & Replace(sWhere & filter, "Assure", "Assure_Retraite")
    
      ProcessQuery = sTemp
    
    Case "4 - Amorti"
      ' - Remplacement de Assure.xxx par ((Assure_Retraite.POCoeffAmortissement*(Assure_Retraite.xxx-Assure.xxx))+Assure.xxx)
      ' - Ajout des champs après réforme
      ' - Ajout de la jointure
      
      sTemp = Replace(Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '4 - Amorti' as TypeLigne,"), "Assure.POCOMMENTANNUL AS [Raison Annulation]", "Assure.POTopAmortissable as [Top Amortissable], 100*Assure.POCoeffAmortissement as [Coeff Amortissement], Assure.POAgeMiniDepartRetraite as [Age Mini Départ Retraite], Assure.POAgeDepartRetraiteTauxPlein as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS [Raison Annulation] "), "Assure", "Assure_Retraite")
      
      ' Traitement des champs POPREST...,POPM...,POPSAP,POCOT...
      pos = 1
      Do
        pos = InStr(pos, sTemp, "Assure_Retraite.") ' Len("Assure_Retraite.")=16
        If pos <> 0 Then
          ' pos = pointe sur "Assure.PO..."
          ' Recherche de la fin du champ
          If ProcessThisField(UCase(mID(sTemp, pos + 16, 30)), 4) Then
            i = FindNextSeparator(sTemp, pos + 16)
            If i <> 0 Then
              sFieldAvant = mID(sTemp, pos - 1, i - pos + 1)
              If sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN" And sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN_PASSAGE" Then
                sAvant = Left(sTemp, pos - 1)
                sApres = mID(sTemp, i)
                sField = mID(sTemp, pos + 16, i - pos - 16)
  '              sChange = "((Assure_Retraite.POCoeffAmortissement*(Assure_Retraite." & sField & "-Assure." & sField & "))+Assure." & sField & ")"
                sChange = "(Assure_Retraite.POCoeffAmortissement*(Assure_Retraite." & sField & "-Assure." & sField & "))"
                
                sTemp = sAvant & sChange & sApres
                
                pos = pos + Len(sChange)
              Else
                pos = pos + (i - pos)
              End If
            Else
              Exit Do
            End If
          Else
            pos = pos + 1
          End If
        End If
      Loop Until pos = 0
    
      sTemp = sTemp & Replace(sFrom, "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO AND Assure_Retraite.POGARCLE=Assure.POGARCLE " _
                    & Replace(sWhere & filter, "Assure", "Assure_Retraite")
    
      ProcessQuery = sTemp
    
    Case "5 - A Amortir"
      ' - Remplacement de Assure.xxx par (((1-Assure_Retraite.POCoeffAmortissement)*(Assure_Retraite.xxx-Assure.xxx))+Assure.xxx)
      ' - Ajout des champs après réforme
      ' - Ajout de la jointure
      
      sTemp = Replace(Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '5 - A Amortir' as TypeLigne,"), "Assure.POCOMMENTANNUL AS [Raison Annulation]", "Assure.POTopAmortissable as [Top Amortissable], 100*Assure.POCoeffAmortissement as [Coeff Amortissement], Assure.POAgeMiniDepartRetraite as [Age Mini Départ Retraite], Assure.POAgeDepartRetraiteTauxPlein as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS [Raison Annulation] "), "Assure", "Assure_Retraite")
      
      ' Traitement des champs POPREST...,POPM...,POPSAP,POCOT...
      pos = 1
      Do
        pos = InStr(pos, sTemp, "Assure_Retraite.") ' Len("Assure_Retraite.")=16
        If pos <> 0 Then
          ' pos = pointe sur "Assure.PO..."
          ' Recherche de la fin du champ
          If ProcessThisField(UCase(mID(sTemp, pos + 16, 30)), 5) Then
            i = FindNextSeparator(sTemp, pos + 16)
            If i <> 0 Then
              sFieldAvant = mID(sTemp, pos - 1, i - pos + 1)
              If sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN" And sFieldAvant <> "*Assure_Retraite.POPRESTATION_AN_PASSAGE" Then
                sAvant = Left(sTemp, pos - 1)
                sApres = mID(sTemp, i)
                sField = mID(sTemp, pos + 16, i - pos - 16)
  '              sChange = "(((1.0-Assure_Retraite.POCoeffAmortissement)*(Assure_Retraite." & sField & "-Assure." & sField & "))+Assure." & sField & ")"
                sChange = "((1.0-Assure_Retraite.POCoeffAmortissement)*(Assure_Retraite." & sField & "-Assure." & sField & "))"
                
                sTemp = sAvant & sChange & sApres
                
                pos = pos + Len(sChange)
              Else
                pos = pos + (i - pos)
              End If
            Else
              Exit Do
            End If
          Else
            pos = pos + 1
          End If
        End If
      Loop Until pos = 0
    
      sTemp = sTemp & Replace(sFrom, "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO AND Assure_Retraite.POGARCLE=Assure.POGARCLE " _
                    & Replace(sWhere & filter, "Assure", "Assure_Retraite")
    
      ProcessQuery = sTemp
    
    Case Else
      'Stop
      ProcessQuery = rq
  
  End Select

  Exit Function
  
err_ProcessQuery:

  If Not autoMode Then
    MsgBox "Erreur " & Err & " dans err_ProcessQuery(" & sType & ") : " & Err.Description, vbCritical
  Else
    autoLogger.EcritTraceDansLog "Erreur " & Err & " dans err_ProcessQuery(" & sType & ") : " & Err.Description
  End If
  ProcessQuery = ""
  
  Exit Function
  
  Resume Next
End Function


'##ModelId=5C8A67A40136
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar1.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 130
 
  If Me.Width > 7 * btnWidth Then
'    topbtn = Me.Height - Toolbar1.Height - btnHeight
    topbtn = Me.ScaleHeight - btnHeight
    
    ' boutton 'Importer'
    PlacePremierBoutton btnImport, topbtn
    
    ' boutton 'Calculer'
    PlaceBoutton btnCalc, btnImport, topbtn
    
    ' bouton 'Calculer & Revaloriser'
    PlaceBoutton btnCalcRevalo, btnCalc, topbtn
    
    ' boutton 'Choix des Editions'
    PlaceBoutton btnEdition, btnCalcRevalo, topbtn
    
    ' boutton 'Imprimer'
    PlaceBoutton btnPrint, btnEdition, topbtn
    
    ' boutton 'Purger'
    PlaceBoutton btnPurge, btnPrint, topbtn
    
    ' boutton 'Exporter'
    PlaceBoutton btnExportSAS, btnPrint, topbtn
    
    ' boutton 'Fermer'
    PlaceBoutton btnClose, btnExportSAS, topbtn
  Else
    'topbtn = Me.Height - Toolbar1.Height - 2 * btnHeight
    topbtn = Me.ScaleHeight - 2 * btnHeight
    
    ' boutton 'Importer'
    PlacePremierBoutton btnImport, topbtn
    
    ' boutton 'Calculer'
    PlaceBoutton btnCalc, btnImport, topbtn
    
    ' boutton 'Choix des Editions'
    PlaceBoutton btnEdition, btnCalc, topbtn
    
    ' boutton 'Imprimer'
    PlacePremierBoutton btnPrint, topbtn + btnHeight + 30
    
    ' bouton 'Calculer & Revaloriser'
    PlaceBoutton btnCalcRevalo, btnPrint, topbtn + btnHeight + 30
    
    ' boutton 'Purger'
    PlaceBoutton btnPurge, btnCalcRevalo, topbtn + btnHeight + 30
    
    ' boutton 'Exporter'
    PlaceBoutton btnExportSAS, btnCalcRevalo, topbtn + btnHeight + 30
    
    ' boutton 'Fermer'
    PlaceBoutton btnClose, btnExportSAS, topbtn + btnHeight + 30
    
  End If
  
  ' liste
  sprListe.Height = Maximum(topbtn - btnHeight - 100, 0)
End Sub

'##ModelId=5C8A67A40146
Private Sub Form_Unload(Cancel As Integer)
  
  'Unlock Periode
  m_dataSource.Execute "Delete From LockedPeriods Where UserName Like '" & user_name & "'"
  
  CloseArchiveConnection
  
  If Not fmFilter_Precedent Is Nothing Then
    Set fmFilter = Nothing
    
    ' restaure l'ancien filtre
    Set fmFilter = fmFilter_Precedent
    Set fmFilter_Precedent = Nothing
  End If
  
  ' sauvegarde du filtre
  fmFilter.Save frmNumPeriode
  
  Set fmFilter = Nothing
End Sub

'##ModelId=5C8A67A40175
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
    Dim sChamp As String, sTable As String
    
    If Not IsNull(dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value) Then
      sChamp = dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value   ' nom du champs de la table
    Else
      sChamp = "'" & dtaPeriode.Recordset.fields(Col - 1).Name & "'"
    End If
    
    If Not IsNull(dtaPeriode.Recordset.fields(Col - 1).Properties("BASETABLENAME").Value) Then
      sTable = dtaPeriode.Recordset.fields(Col - 1).Properties("BASETABLENAME").Value & "."  ' nom de la table
    Else
      sTable = ""
    End If
    
    ' on change la requete SQL (ORDER BY)
    If frmOrdreDeTri = sTable & sChamp Then
      If sChamp = "POCONVENTION" Then
        frmOrdreDeTri = "CAST(" & sTable & "POCONVENTION as bigint) DESC"
      Else
        frmOrdreDeTri = sTable & sChamp & " DESC"   ' nom du champs de la table
      End If
    Else
      If sChamp = "POCONVENTION" Then
        frmOrdreDeTri = "CAST(" & sTable & "POCONVENTION as bigint)"
      Else
        frmOrdreDeTri = sTable & sChamp
      End If
    End If
    
    RefreshListe
#End If
  End If
End Sub

'##ModelId=5C8A67A401B3
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  'If dtaPeriode.Recordset.fields(Col - 1).Name = "Date de Naissance" Then Stop
  
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "POCOMMENTANNUL" Then
    
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  
  Else
    Select Case dtaPeriode.Recordset.fields(Col - 1).Type
      Case adBoolean
        sprListe.Col = Col
        sprListe.Row = -1
        sprListe.CellType = CellTypeCheckBox
        sprListe.TypeCheckCenter = True
        sprListe.TypeCheckType = TypeCheckTypeNormal
      
      Case adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
        sprListe.Col = Col
        sprListe.Row = -1
        sprListe.CellType = CellTypeNumber
        sprListe.TypeNumberDecPlaces = 0
        If dtaPeriode.Recordset.fields(Col - 1).Name = "Dossier" Then
          sprListe.TypeNumberShowSep = False
          sprListe.TypeHAlign = TypeHAlignLeft
        Else
          sprListe.TypeNumberShowSep = True
          sprListe.TypeHAlign = TypeHAlignRight
        End If
        
      Case adDecimal, adDouble, adNumeric, adSingle, adDecimal
        sprListe.Col = Col
        sprListe.Row = -1
        sprListe.CellType = CellTypeNumber
        If dtaPeriode.Recordset.fields(Col - 1).NumericScale = 0 Then
          sprListe.TypeNumberDecPlaces = 0
        ElseIf InStr(dtaPeriode.Recordset.fields(Col - 1).Name, "1€") <> 0 Or InStr(dtaPeriode.Recordset.fields(Col - 1).Name, "12€") <> 0 _
               Or InStr(dtaPeriode.Recordset.fields(Col - 1).Name, "1R") <> 0 Or InStr(dtaPeriode.Recordset.fields(Col - 1).Name, "12R") <> 0 _
               Or dtaPeriode.Recordset.fields(Col - 1).Name = "Correctif PM Incap" Or dtaPeriode.Recordset.fields(Col - 1).Name = "Coefficient BCAC" _
               Or dtaPeriode.Recordset.fields(Col - 1).Name = "Coefficient BCAC R" Then
          sprListe.TypeNumberDecPlaces = 4
        Else
          sprListe.TypeNumberDecPlaces = 2
        End If
        sprListe.TypeNumberShowSep = True
        sprListe.TypeHAlign = TypeHAlignRight
      
      Case adBSTR, adChar, adVarChar, adVarWChar, adWChar
        sprListe.Col = Col
        sprListe.Row = -1
        sprListe.CellType = CellTypeStaticText
        sprListe.TypeHAlign = TypeHAlignLeft
        
      Case adDate, adDBDate, adDBTimeStamp
        sprListe.Col = Col
        sprListe.Row = -1
'        sprListe.CellType = CellTypeDate
'        sprListe.TypeDateCentury = True
'        sprListe.TypeDateFormat = TypeDateFormatDDMMYY
        sprListe.TypeHAlign = TypeHAlignCenter
      
      Case Else
        sprListe.TypeHAlign = TypeHAlignCenter
    End Select
  End If
End Sub

'##ModelId=5C8A67A40211
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  Dim comment As Variant, i As Integer

  If dtaPeriode.Recordset.fields(Col - 1).Name = "Garantie" Then

    sprListe.BlockMode = True
    sprListe.Col = -1
    sprListe.Row = Row
    sprListe.Col2 = -1
    sprListe.Row2 = Row
      
    If (eProvisionRetraite <> frmTypePeriode And eProvisionRetraiteRevalo <> frmTypePeriode) Then
      sprListe.ForeColor = vbBlack
    End If
    
    sprListe.GetDataFillData comment, vbString
    If comment = "MG Décès" Then
      Select Case sprListe.ForeColor
        Case vbBlack
          sprListe.ForeColor = vbBlue
        
        Case DKRED
          sprListe.ForeColor = vbRed
        
        Case vbMagenta
          sprListe.ForeColor = DKRED
          
        Case Else
          sprListe.ForeColor = vbBlue
      End Select
    End If
    
    sprListe.BlockMode = False
  
  ElseIf dtaPeriode.Recordset.fields(Col - 1).Name = "TypeLigne" Then

    sprListe.BlockMode = True
    sprListe.Col = -1
    sprListe.Row = Row
    sprListe.Col2 = -1
    sprListe.Row2 = Row
      
    sprListe.FontItalic = False
    
    sprListe.GetDataFillData comment, vbString
    Select Case comment
      Case "1 - Avant"
        sprListe.ForeColor = vbBlack
      
      Case "2 - Après"
        sprListe.ForeColor = DKRED
      
      Case "3 - Ecart"
        sprListe.ForeColor = vbMagenta
      
      Case "4 - Amorti"
        sprListe.FontItalic = True
      
      Case "5 - A Amortir"
        sprListe.FontItalic = True
      
      Case Else
        sprListe.ForeColor = vbBlack
    End Select
    
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
End Sub


'##ModelId=5C8A67A4028E
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub


'##ModelId=5C8A67A402BD
Private Sub CheckParametre()
  Dim module_calcul As iP3ICalcul, Logger As clsLogger, nomModuleCalcul As String
  
  ' charge l'object de calcul
  Set module_calcul = LoadModuleCalcul(nomModuleCalcul)
  If module_calcul Is Nothing Then Exit Sub
  
  Set Logger = New clsLogger
  Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurParametre.log"
  
  Logger.CreateLog "Vérification de la présence des paramètres"
  
  module_calcul.CheckParametresAssures frmNumPeriode, CLng(GroupeCle), Logger
  Logger.AfficheErreurLog False
  
  Set module_calcul = Nothing
  Set Logger = Nothing
End Sub


'##ModelId=5C8A67A402DC
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim f As frmSousProduit, sNom As String
  
  ' Utilise la propriété Key avec l'instruction SelectCase pour   spécifier une action.
  Select Case Button.key
    Case "manageDisplays"
      Dim FM As frmManageDisplays
      Set FM = New frmManageDisplays
      FM.Show vbModal
      Set FM = Nothing
      
      frmOrdreDeTri = ""
      UpdateOrderByString ("Assure.PONUMCLE")
      UpdateOrderByString ("Assure.POARRET")
  
      RefreshListe
    
    Case "openPeriode"
      frmDetailPeriode.Show vbModal, frmMain
      RefreshListe
  
    Case "editTable"
      If sprListe.ActiveRow <= 0 Then Exit Sub
      Call BuildDescription
      sprListe.Col = 1
      sprListe.Row = sprListe.ActiveRow
      RECNO = CLng(sprListe.text)
      frmAssure.Show vbModal, frmMain
      RefreshListe
      
    Case "sousproduit"
      Set f = New frmSousProduit
      
      Load f
      
      numPeriode = frmNumPeriode
      
      Set f.fmFilter = fmFilter
      
      f.Show vbModal, frmMain
      
      Set f = Nothing
      
    Case "regimes"
      Set f = New frmSousProduit
      
      f.fmTypePeriode = frmTypePeriode
      
      Load f
      
      f.fmMode = SommeParRegime
      numPeriode = frmNumPeriode
      
      Set f.fmFilter = fmFilter
      
      f.Show vbModal, frmMain
      
      Set f = Nothing
      
    Case "CR"
      Set f = New frmSousProduit
      
      Load f
      
      numPeriode = frmNumPeriode
      
      f.fmMode = Prestation
      
      Set f.fmFilter = fmFilter
      
      f.Show vbModal, frmMain
      
      Set f = Nothing
      
    Case "reass"
      numPeriode = frmNumPeriode
            
      Set frmReassurance.fmFilter = fmFilter
      
      frmReassurance.Show vbModal, frmMain
      
    Case "check"
      CheckParametre
    
    Case "print"
      btnPrint_Click

    Case "filter"
      ' filtre
      numPeriode = frmNumPeriode
      Set frmFilter.fmFilter = fmFilter
      
      ret_code = 0
      frmFilter.Show vbModal, frmMain
      
      fmFilter.Save frmNumPeriode
      
      If ret_code = 1 Then
        Toolbar1.Buttons("filter_name").Value = tbrUnpressed
        
        Set fmFilter_Precedent = Nothing
        
        RefreshListe
      End If
    
    Case "filter_name"
      If Not fmFilter_Precedent Is Nothing Then
        Set fmFilter = Nothing
        
        ' restaure l'ancien filtre
        Set fmFilter = fmFilter_Precedent
        Set fmFilter_Precedent = Nothing
        
        Toolbar1.Buttons("filter_name").Value = tbrUnpressed
      Else
        If sprListe.ActiveRow <= 0 Then Exit Sub
        
        ' sauvegarde l'ancien filtre
        Set fmFilter_Precedent = fmFilter
        
        ' filtre par nom de l'assuré
        Set fmFilter = BuildFilter
        
        sNom = Trim(dtaPeriode.Recordset.fields("Nom Assuré"))
        sNom = Replace(sNom, " (conjoint)", "")
        sNom = Replace(sNom, " (enfant(s))", "")
        fmFilter.SetFilterElemValue "Nom", sNom
        
        Toolbar1.Buttons("filter_name").Value = tbrPressed
      End If
        
      RefreshListe
  End Select
End Sub


'##ModelId=5C8A67A4030B
Private Function ReadDouble(f As ADODB.field) As Double
  If IsNull(f.Value) Then
    ReadDouble = 0#
  Else
    ReadDouble = f.Value
    'ReadDouble = FormatNumber(ReadDouble, 8)
    
  End If
End Function


'##ModelId=5C8A67A4032A
Private Function ConvertDateToLong(dDate As Variant) As Long
  If IsNull(dDate) Then
    ConvertDateToLong = 0
  Else
    ' YYYYMMDD = Y*10000 + M * 100 + D
    ConvertDateToLong = Year(dDate) * 10000 + Month(dDate) * 100 + Day(dDate)
  End If
End Function

'old version of:  - no Replace$("", ",", ".")
'##ModelId=5C8A67A40359
Private Sub CopieVersTTProvColl_OLD(CleGroupe As Integer, numPeriode As Long, rsPeriode As ADODB.Recordset, rsAssure As ADODB.Recordset, rsCodeCatInval As ADODB.Recordset, cLogger As clsLogger, sTypeProsision As String)
  
  On Error GoTo err_CopyLot
  
  Dim i As Integer, f As ADODB.field, nb As Long, bOk As Boolean
  Dim rsAssureP3IProvColl As ADODB.Recordset, rsAssure_Retraite As ADODB.Recordset
  
  'm_dataHelper.Multi_Find rsAssureP3IProvColl, "NUENRP3I=" & rsAssure.fields("NUENRP3I")
  Set rsAssureP3IProvColl = m_dataSource.OpenRecordset("SELECT * FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & numPeriode _
                                                        & " AND NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I") & " AND NUENRP3I=" & rsAssure.fields("NUENRP3I"), Snapshot)
  
  If rsAssureP3IProvColl.EOF Then Exit Sub
  
  
  If IsNull(rsAssure.fields("PONumParamCalcul")) Then
    MsgBox "Erreur dans CopieVersTTProvColl() : Paramètre de calcul non renseigné pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !", vbCritical
    Exit Sub
  End If
  
  
  m_dataHelper.Multi_Find rsPeriode, "PENUMPARAMCALCUL=" & rsAssure.fields("PONumParamCalcul")
    
  If rsPeriode.EOF = True Then
    MsgBox "Erreur dans CopieVersTTProvColl() : Paramètre de calcul n°" & rsAssure.fields("PONumParamCalcul") & " introuvable pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !", vbCritical
    Exit Sub
  End If
    
  If rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraite Or rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraiteRevalo Then
    Set rsAssure_Retraite = m_dataSource.OpenRecordset("SELECT * FROM Assure_Retraite WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & numPeriode & " AND POIdAssure=" & rsAssure.fields("RECNO"), Snapshot)
  Else
    Set rsAssure_Retraite = Nothing
  End If
  
  'rsTTPROVCOLL.AddNew
  Dim theTTPROVCOLL As New clsTTPROVCOLL
    
  theTTPROVCOLL.Load_OLD m_dataSource, rsAssureP3IProvColl, True
    
  theTTPROVCOLL.m_MTCPLEST = ReadDouble(rsAssure.fields("POPSAP"))
  theTTPROVCOLL.m_MTCPLREE = ReadDouble(rsAssure.fields("POPSAP"))
        
  theTTPROVCOLL.m_MTPROCAL = ReadDouble(rsAssure.fields("POPM"))
               
  theTTPROVCOLL.m_MTPROIMP = ReadDouble(rsAssure.fields("POPM")) + ReadDouble(rsAssure.fields("POPSAP"))
   'theTTPROVCOLL.m_MTPREANN2 = ReadDouble(rsAssure.fields("MTPREANN2")) ' bug 04 08 2013
  theTTPROVCOLL.m_MTPREANN2 = ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE")) '04 08 2013
               
  If Not rsAssureP3IProvColl.EOF Then
    theTTPROVCOLL.m_MTPROVIT = ReadDouble(rsAssure.fields("POPM_INVAL_1F")) * ReadDouble(rsAssureP3IProvColl.fields("MTPREANN")) ' champs calculé dans l'import et déjà présent dans AssureP3IProvColl
    'theTTPROVCOLL.m_MTPROPAS = ReadDouble(rsAssure.fields("POPM_PASS_1F")) * ReadDouble(rsAssureP3IProvColl.fields("MTPREANN")) ' bug 04 08 2013
    theTTPROVCOLL.m_MTPROPAS = ReadDouble(rsAssure.fields("POPM_PASS_1F")) * ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE")) ' 04 08 2013 champs absent dans l'import et calculé dans provisions, à charger depuis assuré
  Else
    theTTPROVCOLL.m_MTPROVIT = 0
    theTTPROVCOLL.m_MTPROPAS = 0
  End If
               
  theTTPROVCOLL.m_TXPROV = ReadDouble(rsAssure.fields("POPM_INVAL_1F")) + ReadDouble(rsAssure.fields("POPM_PASS_1F")) _
                           + ReadDouble(rsAssure.fields("POPM_INCAP_1F")) / 12# _
                           + ReadDouble(rsAssure.fields("POPM_RCJT_1F")) + ReadDouble(rsAssure.fields("POPM_REDUC_1F"))
               
  theTTPROVCOLL.m_TXPROPASS = ReadDouble(rsAssure.fields("POPM_PASS_1F"))
               
  theTTPROVCOLL.m_TXTECHN = ReadDouble(rsAssure.fields("TXTECHN"))
      
  theTTPROVCOLL.m_TXFRAIS = ReadDouble(rsAssure.fields("TXFRAIS"))
        
  theTTPROVCOLL.m_MTCAPCONORI = 0
  theTTPROVCOLL.m_MTPANCAL = 0
  theTTPROVCOLL.m_MTCMPPAN = 0
  theTTPROVCOLL.m_MTCOMPAN = 0
  theTTPROVCOLL.m_MTCOMCMP = 0
      
  If Not rsAssureP3IProvColl.EOF Then
    theTTPROVCOLL.m_CDLOTIMPORT = rsAssureP3IProvColl.fields("NUTRAITP3I")
  Else
    theTTPROVCOLL.m_CDLOTIMPORT = rsPeriode.fields("NUTRAITP3I")
  End If
      
  theTTPROVCOLL.m_PERALIM = Round(rsPeriode.fields("DTTRAIT") / 100#, 0)     ' Year(Now) * 100 + Month(Now)
     
  theTTPROVCOLL.m_TXLISSAGE = ReadDouble(rsAssure.fields("POPourcentLissage"))
          
  theTTPROVCOLL.m_TXPROSANSLISSAGE = ReadDouble(rsAssure.fields("POCoeffBCAC"))
      
  theTTPROVCOLL.m_TYPEPROVISION = sTypeProsision
      
' PHM 19/11/2009
  ' cas du MGDC : garantie élémentaire pour le maintien décès = code 6125 : « Exonération prévoyance » et mettre 0 pour le montant de la prestation annualisée ?
  If rsAssure.fields("POGARCLE") >= 90 Then
    
    If IsNull(rsAssure.fields("POGARCLE_NEW")) Then
    
      ' Code générique
      theTTPROVCOLL.m_CDGARAN = 6125
    
    Else
      
      theTTPROVCOLL.m_CDGARAN = rsAssure.fields("POGARCLE_NEW")
    
    End If
    
    theTTPROVCOLL.m_MTPREANN = 0  ' Annualisation Base
    theTTPROVCOLL.m_MTPREANN2 = 0 ' Annualisation Passage
    
' PHM 17/09/2010
    theTTPROVCOLL.m_MTPREREV = 0 ' Annualisation Revalo
' PHM 17/09/2010
  
' PHM 28/06/2011
  
  ElseIf rsAssure.fields("POGARCLE") = cdGarRente Then
    theTTPROVCOLL.m_DTFINPER = ConvertDateToLong(rsAssure.fields("POFIN"))

' PHM 28/06/2011
  
  End If
' PHM 19/11/2009
      
      
  ' Evol 2010 - Lot 2
  If Not IsNull(rsAssure.fields("POCategorieInval")) Then
    If rsAssure.fields("POCategorieInval") = 1 Or rsAssure.fields("POCategorieInval") = 2 Or rsAssure.fields("POCategorieInval") = 3 Then
      m_dataHelper.Multi_Find rsCodeCatInval, "CDCHOIXPREST='" & rsAssureP3IProvColl.fields("CDCHOIXPREST") & "'"
      
      If rsCodeCatInval.EOF = False Then
        theTTPROVCOLL.m_CDCHOIXPREST = rsCodeCatInval.fields("CDCHOIXPREST")
        theTTPROVCOLL.m_LBCHOIXPREST = rsCodeCatInval.fields("LBCHOIXPREST")
        theTTPROVCOLL.m_CDCATINV = rsCodeCatInval.fields("CDCATINV")
        theTTPROVCOLL.m_LBCATINV = rsCodeCatInval.fields("LBCATINV")
      End If
    End If
  End If
  
  theTTPROVCOLL.m_MTCAPCON = ReadDouble(rsAssure.fields("POMontantCapConstit"))
  theTTPROVCOLL.m_MTCAPSSRISQ = ReadDouble(rsAssure.fields("POMontantCapSousRisque"))
      
  ' deja copié depuis P3IPROVCOLL
  'theTTPROVCOLL.m_CDCONTENTIEUX = ReadDouble(rsAssure.fields("POCDCONTENTIEUX"))
  'theTTPROVCOLL.m_NUSINISTRE = ReadDouble(rsAssure.fields("PONUSINISTRE"))
  
      
  theTTPROVCOLL.m_FLAMORTISSABLE = rsAssure.fields("POTopAmortissable")
  
  
  ' Se place sur la ligne retraite
  bOk = False
  If (rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraite Or rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraiteRevalo) Then
    bOk = True
  End If
  
  If rsAssure_Retraite Is Nothing Then
    bOk = False
  ElseIf rsAssure_Retraite.EOF = True Then
    bOk = False
  Else
    bOk = bOk And True
  End If
  
  If bOk = True Then
    
    ' Péridoe réforme des retraites
    theTTPROVCOLL.m_TXAMORT = 100# * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    
    theTTPROVCOLL.m_DTLIMPROAPR = ConvertDateToLong(rsAssure_Retraite.fields("POTERME"))
    
    theTTPROVCOLL.m_AGELIMINC = 0
    theTTPROVCOLL.m_AGELIMINV = 0
    
    If rsAssure_Retraite.fields("POSIT") = cdPosit_Inval Then
      theTTPROVCOLL.m_AGELIMINV = DateDiff("yyyy", rsAssure_Retraite.fields("PONAIS"), rsAssure_Retraite.fields("POTERME"))
    
    ElseIf rsAssure_Retraite.fields("POSIT") = cdPosit_IncapAvecPassage Or rsAssure_Retraite.fields("POSIT") = cdPosit_IncapSansPassage Then
      theTTPROVCOLL.m_AGELIMINC = DateDiff("yyyy", rsAssure_Retraite.fields("PONAIS"), rsAssure_Retraite.fields("POTERME"))
    End If
                                       
    theTTPROVCOLL.m_MTPROIMPAVR = Round(ReadDouble(rsAssure.fields("POPM")) + ReadDouble(rsAssure.fields("POPSAP")), 2)
    theTTPROVCOLL.m_MTPROIMPAPR = Round(ReadDouble(rsAssure_Retraite.fields("POPM")) + ReadDouble(rsAssure_Retraite.fields("POPSAP")), 2)
    
    theTTPROVCOLL.m_MTPROCALAVR = ReadDouble(rsAssure.fields("POPM"))
    theTTPROVCOLL.m_MTPROCALAPR = ReadDouble(rsAssure_Retraite.fields("POPM"))
    
    theTTPROVCOLL.m_MTINDEMRES = ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")) * (theTTPROVCOLL.m_MTPROCALAPR - theTTPROVCOLL.m_MTPROCALAVR)
    
    theTTPROVCOLL.m_MTCOUTREFPROCAL = (theTTPROVCOLL.m_MTPROCALAPR - theTTPROVCOLL.m_MTPROCALAVR)
    
    theTTPROVCOLL.m_MTAMORTCAL = theTTPROVCOLL.m_MTCOUTREFPROCAL * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    theTTPROVCOLL.m_MTRESTEAMORCAL = theTTPROVCOLL.m_MTCOUTREFPROCAL * (1# - ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")))
    
    theTTPROVCOLL.m_MTPROPASAVR = Round(ReadDouble(rsAssure.fields("POPM_PASS_1F")) * ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE")), 2)
    theTTPROVCOLL.m_MTPROPASAPR = Round(ReadDouble(rsAssure_Retraite.fields("POPM_PASS_1F")) * ReadDouble(rsAssure_Retraite.fields("POPRESTATION_AN_PASSAGE")), 2)
    
    theTTPROVCOLL.m_MTCOUTREFPROPAS = theTTPROVCOLL.m_MTPROPASAPR - theTTPROVCOLL.m_MTPROPASAVR
    
    theTTPROVCOLL.m_MTAMORTPAS = theTTPROVCOLL.m_MTCOUTREFPROPAS * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    theTTPROVCOLL.m_MTRESTEAMORPAS = theTTPROVCOLL.m_MTCOUTREFPROPAS * (1# - ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")))
    
    
    rsAssure_Retraite.Close
    Set rsAssure_Retraite = Nothing
    
  Else
    
    ' Période non retraite ou pas de ligne retraite
    theTTPROVCOLL.m_TXAMORT = 0#
    theTTPROVCOLL.m_DTLIMPROAPR = 0#
    theTTPROVCOLL.m_MTPROIMPAVR = 0#
    theTTPROVCOLL.m_MTPROIMPAPR = 0#
    theTTPROVCOLL.m_MTINDEMRES = 0#
    theTTPROVCOLL.m_MTPROCALAVR = 0#
    theTTPROVCOLL.m_MTPROCALAPR = 0#
    theTTPROVCOLL.m_MTCOUTREFPROCAL = 0#
    theTTPROVCOLL.m_MTAMORTCAL = 0#
    theTTPROVCOLL.m_MTRESTEAMORCAL = 0#
    theTTPROVCOLL.m_MTPROPASAVR = 0#
    theTTPROVCOLL.m_MTPROPASAPR = 0#
    theTTPROVCOLL.m_MTCOUTREFPROPAS = 0#
    theTTPROVCOLL.m_MTAMORTPAS = 0#
    theTTPROVCOLL.m_MTRESTEAMORPAS = 0#
  
    theTTPROVCOLL.m_AGELIMINC = 0#
    theTTPROVCOLL.m_AGELIMINV = 0#
  End If
  
  theTTPROVCOLL.Save m_dataSource
      
  rsAssureP3IProvColl.Close
    
  Exit Sub
  
err_CopyLot:
  MsgBox "Erreur dans CopieVersTTProvColl() : " & Err & vbLf & Err.Description, vbCritical
   Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données de nos tables SQL Server vers TTPROVCOLL
'
' CleGroupe  : n° de groupe
' NumPeriode : n° de période
' rsPeriode  : recordset periode+paramcalcul placé sur le bon jeux de
'paramètres
' rsAssure   : recordset des assure placé sur l'assuré en cours
' rsAssureP3IProvColl : liste des assures - on doit rechercher NUENRP3I
' rsTTPROVCOLL : recordset de destination
' cLogger      : log pour les messages
'
'##ModelId=5C8A67A5000E
Private Sub CopieVersTTProvColl(CleGroupe As Integer, numPeriode As Long, rsPeriode As ADODB.Recordset, rsAssure As ADODB.Recordset, rsCodeCatInval As ADODB.Recordset, cLogger As clsLogger, sTypeProsision As String)
  
  On Error GoTo err_CopyLot
  
  Dim StartTime As Double, EndTime As Double
  
  Dim i As Integer, f As ADODB.field, nb As Long, bOk As Boolean
  Dim rsAssureP3IProvColl As ADODB.Recordset
  Dim rsAssure_Retraite As ADODB.Recordset
  
  'StartTime = Timer
  
  'm_dataHelper.Multi_Find rsAssureP3IProvColl, "NUENRP3I=" & rsAssure.fields("NUENRP3I")
  Set rsAssureP3IProvColl = m_dataSource.OpenRecordset("SELECT * FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & numPeriode _
          & " AND NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I") & " AND NUENRP3I=" & rsAssure.fields("NUENRP3I"), Snapshot)
  
  'EndTime = Timer
  'Debug.Print "OpenRecordSet : rsAssureP3IProvColl", EndTime - StartTime
   
  If rsAssureP3IProvColl.EOF Then Exit Sub
  
  
  If IsNull(rsAssure.fields("PONumParamCalcul")) Then
    If Not autoMode Then
      MsgBox "Erreur dans CopieVersTTProvColl() : Paramètre de calcul non renseigné pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !", vbCritical
    Else
      autoLogger.EcritTraceDansLog "Erreur dans CopieVersTTProvColl() : Paramètre de calcul non renseigné pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !"
    End If
    Exit Sub
  End If
  
  
  m_dataHelper.Multi_Find rsPeriode, "PENUMPARAMCALCUL=" & rsAssure.fields("PONumParamCalcul")
    
  If rsPeriode.EOF = True Then
    If Not autoMode Then
      MsgBox "Erreur dans CopieVersTTProvColl() : Paramètre de calcul n°" & rsAssure.fields("PONumParamCalcul") & " introuvable pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !", vbCritical
    Else
      autoLogger.EcritTraceDansLog "Erreur dans CopieVersTTProvColl() : Paramètre de calcul n°" & rsAssure.fields("PONumParamCalcul") & " introuvable pour l'assuré NUENRP3I=" & rsAssure.fields("NUENRP3I") & vbLf & "L'export de cet assuré vers TTPROVCOLL ne peut être éffectué !"
    End If
    Exit Sub
  End If
    
  If rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraite Or rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraiteRevalo Then
  
    'StartTime = Timer
    Set rsAssure_Retraite = m_dataSource.OpenRecordset("SELECT * FROM Assure_Retraite WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & numPeriode & " AND POIdAssure=" & rsAssure.fields("RECNO"), Snapshot)
  
    'EndTime = Timer
    'Debug.Print "OpenRecordSet : ELECT * FROM Assure_Retraite", EndTime - StartTime
   
  
  Else
    Set rsAssure_Retraite = Nothing
  End If
  
  'rsTTPROVCOLL.AddNew
  Dim theTTPROVCOLL As New clsTTPROVCOLL
  
  'StartTime = Timer
    
  theTTPROVCOLL.Load m_dataSource, rsAssureP3IProvColl, True
    
  If rsAssure.fields("Commentaire") <> Null Then
    theTTPROVCOLL.m_COMMENTAIRE = Replace$(rsAssure.fields("Commentaire"), ";", ",")
  End If
  
  theTTPROVCOLL.m_MTCPLEST = Replace$(ReadDouble(rsAssure.fields("POPSAP")), ",", ".")
  theTTPROVCOLL.m_MTCPLREE = Replace$(ReadDouble(rsAssure.fields("POPSAP")), ",", ".")
        
  theTTPROVCOLL.m_MTPROCAL = Replace$(ReadDouble(rsAssure.fields("POPM")), ",", ".")
  
  theTTPROVCOLL.m_MTPROIMP = Replace$((ReadDouble(rsAssure.fields("POPM")) + ReadDouble(rsAssure.fields("POPSAP"))), ",", ".")
  
  theTTPROVCOLL.m_MTPREANN2 = Replace$(ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE")), ",", ".")
               
  If Not rsAssureP3IProvColl.EOF Then
    theTTPROVCOLL.m_MTPROVIT = Replace$((ReadDouble(rsAssure.fields("POPM_INVAL_1F")) * ReadDouble(rsAssureP3IProvColl.fields("MTPREANN"))), ",", ".")
    
    theTTPROVCOLL.m_MTPROPAS = Replace$((ReadDouble(rsAssure.fields("POPM_PASS_1F")) * ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE"))), ",", ".")
  Else
    theTTPROVCOLL.m_MTPROVIT = 0
    theTTPROVCOLL.m_MTPROPAS = 0
  End If
               
  theTTPROVCOLL.m_TXPROV = ReadDouble(rsAssure.fields("POPM_INVAL_1F")) + ReadDouble(rsAssure.fields("POPM_PASS_1F")) _
                           + ReadDouble(rsAssure.fields("POPM_INCAP_1F")) / 12# _
                           + ReadDouble(rsAssure.fields("POPM_RCJT_1F")) + ReadDouble(rsAssure.fields("POPM_REDUC_1F"))
                           
                           
  theTTPROVCOLL.m_TXPROV = FormatNumber(theTTPROVCOLL.m_TXPROV, 5)
  theTTPROVCOLL.m_TXPROV = Replace$(theTTPROVCOLL.m_TXPROV, ",", ".")
  
  theTTPROVCOLL.m_TXPROPASS = Replace$(ReadDouble(rsAssure.fields("POPM_PASS_1F")), ",", ".")
               
  theTTPROVCOLL.m_TXTECHN = Replace$(ReadDouble(rsAssure.fields("TXTECHN")), ",", ".")
      
  theTTPROVCOLL.m_TXFRAIS = Replace$(ReadDouble(rsAssure.fields("TXFRAIS")), ",", ".")
        
  theTTPROVCOLL.m_MTCAPCONORI = 0
  theTTPROVCOLL.m_MTPANCAL = 0
  theTTPROVCOLL.m_MTCMPPAN = 0
  theTTPROVCOLL.m_MTCOMPAN = 0
  theTTPROVCOLL.m_MTCOMCMP = 0
      
  If Not rsAssureP3IProvColl.EOF Then
    theTTPROVCOLL.m_CDLOTIMPORT = Replace$(rsAssureP3IProvColl.fields("NUTRAITP3I"), ",", ".")
  Else
    theTTPROVCOLL.m_CDLOTIMPORT = Replace$(rsPeriode.fields("NUTRAITP3I"), ",", ".")
  End If
      
  theTTPROVCOLL.m_PERALIM = Round(rsPeriode.fields("DTTRAIT") / 100#, 0)
     
  theTTPROVCOLL.m_TXLISSAGE = Replace$(ReadDouble(rsAssure.fields("POPourcentLissage")), ",", ".")
          
  theTTPROVCOLL.m_TXPROSANSLISSAGE = Replace$(ReadDouble(rsAssure.fields("POCoeffBCAC")), ",", ".")
      
  theTTPROVCOLL.m_TYPEPROVISION = sTypeProsision
      
' PHM 19/11/2009
  ' cas du MGDC : garantie élémentaire pour le maintien décès = code 6125 : « Exonération prévoyance » et mettre 0 pour le montant de la prestation annualisée ?
  If rsAssure.fields("POGARCLE") >= 90 Then
    
    If IsNull(rsAssure.fields("POGARCLE_NEW")) Then
    
      ' Code générique
      theTTPROVCOLL.m_CDGARAN = 6125
    
    Else
      
      theTTPROVCOLL.m_CDGARAN = Replace$(rsAssure.fields("POGARCLE_NEW"), ",", ".")
    
    End If
    
    theTTPROVCOLL.m_MTPREANN = 0  ' Annualisation Base
    theTTPROVCOLL.m_MTPREANN2 = 0 ' Annualisation Passage
    
' PHM 17/09/2010
    theTTPROVCOLL.m_MTPREREV = 0 ' Annualisation Revalo
' PHM 17/09/2010
  
' PHM 28/06/2011
  
  ElseIf rsAssure.fields("POGARCLE") = cdGarRente Then
    theTTPROVCOLL.m_DTFINPER = ConvertDateToLong(rsAssure.fields("POFIN"))

' PHM 28/06/2011
  
  End If
' PHM 19/11/2009
      
      
  ' Evol 2010 - Lot 2
  If Not IsNull(rsAssure.fields("POCategorieInval")) Then
    If rsAssure.fields("POCategorieInval") = 1 Or rsAssure.fields("POCategorieInval") = 2 Or rsAssure.fields("POCategorieInval") = 3 Then
      m_dataHelper.Multi_Find rsCodeCatInval, "CDCHOIXPREST='" & rsAssureP3IProvColl.fields("CDCHOIXPREST") & "'"
      
      If rsCodeCatInval.EOF = False Then
        theTTPROVCOLL.m_CDCHOIXPREST = Replace$(rsCodeCatInval.fields("CDCHOIXPREST"), ",", ".")
        theTTPROVCOLL.m_LBCHOIXPREST = Replace$(rsCodeCatInval.fields("LBCHOIXPREST"), ",", ".")
        theTTPROVCOLL.m_CDCATINV = Replace$(rsCodeCatInval.fields("CDCATINV"), ",", ".")
        theTTPROVCOLL.m_LBCATINV = Replace$(rsCodeCatInval.fields("LBCATINV"), ",", ".")
      End If
    End If
  End If
  
  theTTPROVCOLL.m_MTCAPCON = Replace$(ReadDouble(rsAssure.fields("POMontantCapConstit")), ",", ".")
  theTTPROVCOLL.m_MTCAPSSRISQ = Replace$(ReadDouble(rsAssure.fields("POMontantCapSousRisque")), ",", ".")
      
  ' deja copié depuis P3IPROVCOLL
  'theTTPROVCOLL.m_CDCONTENTIEUX = ReadDouble(rsAssure.fields("POCDCONTENTIEUX"))
  'theTTPROVCOLL.m_NUSINISTRE = ReadDouble(rsAssure.fields("PONUSINISTRE"))
  
      
  theTTPROVCOLL.m_FLAMORTISSABLE = rsAssure.fields("POTopAmortissable")
  
  
  ' Se place sur la ligne retraite
  bOk = False
  If (rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraite Or rsPeriode.fields("PETYPEPERIODE") = eProvisionRetraiteRevalo) Then
    bOk = True
  End If
  
  If rsAssure_Retraite Is Nothing Then
    bOk = False
  ElseIf rsAssure_Retraite.EOF = True Then
    bOk = False
  Else
    bOk = bOk And True
  End If
  
  If bOk = True Then
    
    ' Péridoe réforme des retraites
    If rsAssure_Retraite.fields("Commentaire") <> Null Then
      theTTPROVCOLL.m_COMMENTAIRE = Replace$(rsAssure_Retraite.fields("Commentaire"), ";", ",")
    End If
    
    theTTPROVCOLL.m_TXAMORT = 100# * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    theTTPROVCOLL.m_TXAMORT = Replace$(theTTPROVCOLL.m_TXAMORT, ",", ".")
    
    theTTPROVCOLL.m_DTLIMPROAPR = ConvertDateToLong(rsAssure_Retraite.fields("POTERME"))
    
    theTTPROVCOLL.m_AGELIMINC = 0
    theTTPROVCOLL.m_AGELIMINV = 0
    
    If rsAssure_Retraite.fields("POSIT") = cdPosit_Inval Then
      theTTPROVCOLL.m_AGELIMINV = DateDiff("yyyy", rsAssure_Retraite.fields("PONAIS"), rsAssure_Retraite.fields("POTERME"))
    
    ElseIf rsAssure_Retraite.fields("POSIT") = cdPosit_IncapAvecPassage Or rsAssure_Retraite.fields("POSIT") = cdPosit_IncapSansPassage Then
      theTTPROVCOLL.m_AGELIMINC = DateDiff("yyyy", rsAssure_Retraite.fields("PONAIS"), rsAssure_Retraite.fields("POTERME"))
    End If
    
    
    'Problem Fields:
'8: m_MTINDEMRES
'11: m_MTCOUTREFPROCAL
'12: m_MTAMORTCAL
    
    theTTPROVCOLL.m_MTPROIMPAVR = Round(ReadDouble(rsAssure.fields("POPM")) + ReadDouble(rsAssure.fields("POPSAP")), 2)
    theTTPROVCOLL.m_MTPROIMPAVR = Replace$(theTTPROVCOLL.m_MTPROIMPAVR, ",", ".")
    
    theTTPROVCOLL.m_MTPROIMPAPR = Round(ReadDouble(rsAssure_Retraite.fields("POPM")) + ReadDouble(rsAssure_Retraite.fields("POPSAP")), 2)
    theTTPROVCOLL.m_MTPROIMPAPR = Replace$(theTTPROVCOLL.m_MTPROIMPAPR, ",", ".")
    
    theTTPROVCOLL.m_MTPROCALAVR = ReadDouble(rsAssure.fields("POPM"))
    theTTPROVCOLL.m_MTPROCALAPR = ReadDouble(rsAssure_Retraite.fields("POPM"))
    
    '### TEST
    'theTTPROVCOLL.m_MTPROCALAPR = 3.25
    'theTTPROVCOLL.m_MTPROCALAVR = 3.16
    
    '8
    theTTPROVCOLL.m_MTINDEMRES = ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")) * (theTTPROVCOLL.m_MTPROCALAPR - theTTPROVCOLL.m_MTPROCALAVR)
    theTTPROVCOLL.m_MTINDEMRES = CDbl(Format(theTTPROVCOLL.m_MTINDEMRES, "#0.0000000"))
    theTTPROVCOLL.m_MTINDEMRES = Replace$(theTTPROVCOLL.m_MTINDEMRES, ",", ".")
        
    '11
    theTTPROVCOLL.m_MTCOUTREFPROCAL = (theTTPROVCOLL.m_MTPROCALAPR - theTTPROVCOLL.m_MTPROCALAVR)
    theTTPROVCOLL.m_MTCOUTREFPROCAL = CDbl(Format(theTTPROVCOLL.m_MTCOUTREFPROCAL, "#0.0000000"))
    
        
    theTTPROVCOLL.m_MTPROCALAPR = Replace$(ReadDouble(rsAssure_Retraite.fields("POPM")), ",", ".")
    theTTPROVCOLL.m_MTPROCALAVR = Replace$(ReadDouble(rsAssure.fields("POPM")), ",", ".")
    
    '12
    theTTPROVCOLL.m_MTAMORTCAL = theTTPROVCOLL.m_MTCOUTREFPROCAL * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    theTTPROVCOLL.m_MTAMORTCAL = CDbl(Format(theTTPROVCOLL.m_MTAMORTCAL, "#0.0000000"))
    theTTPROVCOLL.m_MTAMORTCAL = Replace$(theTTPROVCOLL.m_MTAMORTCAL, ",", ".")
    
    '13
    theTTPROVCOLL.m_MTRESTEAMORCAL = theTTPROVCOLL.m_MTCOUTREFPROCAL * (1# - ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")))
    theTTPROVCOLL.m_MTRESTEAMORCAL = CDbl(Format(theTTPROVCOLL.m_MTRESTEAMORCAL, "#0.0000000"))
    theTTPROVCOLL.m_MTRESTEAMORCAL = Replace$(theTTPROVCOLL.m_MTRESTEAMORCAL, ",", ".")
    
  
    theTTPROVCOLL.m_MTPROPASAVR = Round(ReadDouble(rsAssure.fields("POPM_PASS_1F")) * ReadDouble(rsAssure.fields("POPRESTATION_AN_PASSAGE")), 2)
    theTTPROVCOLL.m_MTPROPASAPR = Round(ReadDouble(rsAssure_Retraite.fields("POPM_PASS_1F")) * ReadDouble(rsAssure_Retraite.fields("POPRESTATION_AN_PASSAGE")), 2)
        
    
    '### modify formating
    theTTPROVCOLL.m_MTCOUTREFPROPAS = theTTPROVCOLL.m_MTPROPASAPR - theTTPROVCOLL.m_MTPROPASAVR
    theTTPROVCOLL.m_MTCOUTREFPROPAS = CDbl(Format(theTTPROVCOLL.m_MTCOUTREFPROPAS, "#0.0000000"))
    
    
    theTTPROVCOLL.m_MTAMORTPAS = theTTPROVCOLL.m_MTCOUTREFPROPAS * ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement"))
    theTTPROVCOLL.m_MTAMORTPAS = Replace$(theTTPROVCOLL.m_MTAMORTPAS, ",", ".")
    
    theTTPROVCOLL.m_MTRESTEAMORPAS = theTTPROVCOLL.m_MTCOUTREFPROPAS * (1# - ReadDouble(rsAssure_Retraite.fields("POCoeffAmortissement")))
    theTTPROVCOLL.m_MTRESTEAMORPAS = Replace$(theTTPROVCOLL.m_MTRESTEAMORPAS, ",", ".")
    
    
    theTTPROVCOLL.m_MTCOUTREFPROCAL = Replace$(theTTPROVCOLL.m_MTCOUTREFPROCAL, ",", ".")
    theTTPROVCOLL.m_MTPROPASAPR = Replace$(theTTPROVCOLL.m_MTPROPASAPR, ",", ".")
    theTTPROVCOLL.m_MTCOUTREFPROPAS = Replace$(theTTPROVCOLL.m_MTCOUTREFPROPAS, ",", ".")
    theTTPROVCOLL.m_MTPROPASAVR = Replace$(theTTPROVCOLL.m_MTPROPASAVR, ",", ".")
        
    'le 15/10/2017 ALAIN et ALI
    theTTPROVCOLL.m_FLAMORTISSABLE = rsAssure_Retraite.fields("POTopAmortissable")
    
    rsAssure_Retraite.Close
    Set rsAssure_Retraite = Nothing
    
    'EndTime = Timer
    'Debug.Print "theTTPROVCOLL: ", EndTime - StartTime
    '2,59999999980209E-02
  
    
  Else
    
    ' Période non retraite ou pas de ligne retraite
    theTTPROVCOLL.m_TXAMORT = 0#
    theTTPROVCOLL.m_DTLIMPROAPR = 0#
    theTTPROVCOLL.m_MTPROIMPAVR = 0#
    theTTPROVCOLL.m_MTPROIMPAPR = 0#
    theTTPROVCOLL.m_MTINDEMRES = 0#
    theTTPROVCOLL.m_MTPROCALAVR = 0#
    theTTPROVCOLL.m_MTPROCALAPR = 0#
    theTTPROVCOLL.m_MTCOUTREFPROCAL = 0#
    theTTPROVCOLL.m_MTAMORTCAL = 0#
    theTTPROVCOLL.m_MTRESTEAMORCAL = 0#
    theTTPROVCOLL.m_MTPROPASAVR = 0#
    theTTPROVCOLL.m_MTPROPASAPR = 0#
    theTTPROVCOLL.m_MTCOUTREFPROPAS = 0#
    theTTPROVCOLL.m_MTAMORTPAS = 0#
    theTTPROVCOLL.m_MTRESTEAMORPAS = 0#
  
    theTTPROVCOLL.m_AGELIMINC = 0#
    theTTPROVCOLL.m_AGELIMINV = 0#
  End If
  
  
  'StartTime = Timer
  
  theTTPROVCOLL.SaveToCSV
  
  'theTTPROVCOLL.Save m_dataSource
  
  'EndTime = Timer
  'Debug.Print "theTTPROVCOLL.Save: ", EndTime - StartTime
  '0,282999999995809
  
      
  rsAssureP3IProvColl.Close
    
  Exit Sub
  
err_CopyLot:
  If Not autoMode Then
    MsgBox "Erreur dans CopieVersTTProvColl() : " & Err & vbLf & Err.Description, vbCritical
  Else
    autoLogger.EcritTraceDansLog "Erreur dans CopieVersTTProvColl() : " & Err & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'##ModelId=5C8A67A5009A
Private Sub CopieVersTTLogTrait(CleGroupe As Integer, numPeriode As Long, NumeroLot As Long, cLogger As clsLogger, bDeleteExistant As Boolean, sTypeProvision As String)
  On Error GoTo err_CopyLot
  
  Dim i As Integer, f As ADODB.field, nb As Long, bOk As Boolean
  
  Dim rsTTLogTrait As ADODB.Recordset, rsIn As ADODB.Recordset
  
  If bDeleteExistant = True Then
    
    m_dataSource.Execute "DELETE FROM TTPROVCOLL WHERE NUTRAITP3I=" & NumeroLot
    m_dataSource.Execute "DELETE FROM TTLOGTRAIT WHERE NUTRAITP3I=" & NumeroLot
  
    Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & numPeriode & " AND NUTRAITP3I=" & NumeroLot, Snapshot)
    
    Set rsTTLogTrait = m_dataSource.OpenRecordset("SELECT * FROM TTLOGTRAIT", Dynamic)
    
    rsTTLogTrait.AddNew
  
    rsTTLogTrait.fields("NUTRAITP3I") = rsIn.fields("NUTRAITP3I")
    rsTTLogTrait.fields("NUTRAIT") = rsIn.fields("NUTRAIT")
    rsTTLogTrait.fields("DTTRAIT") = rsIn.fields("DTTRAIT")
    rsTTLogTrait.fields("HHTRAIT") = rsIn.fields("HHTRAIT")
    rsTTLogTrait.fields("DTDEBPER") = rsIn.fields("DTDEBPER")
    rsTTLogTrait.fields("DTFINPER") = rsIn.fields("DTFINPER")
    rsTTLogTrait.fields("IDTABLESAS") = rsIn.fields("IDTABLESAS")
    rsTTLogTrait.fields("NBLIGTRAIT") = rsIn.fields("NBLIGTRAIT")
    rsTTLogTrait.fields("MTTRAIT") = rsIn.fields("MTTRAIT")
    
    rsTTLogTrait.Update
  
    rsTTLogTrait.Close
    
    rsIn.Close
  
  Else
  
    m_dataSource.Execute "DELETE FROM TTPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND TYPEPROVISION='" & sTypeProvision & "'"
  
    Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & numPeriode & " AND NUTRAITP3I=" & NumeroLot, Snapshot)
    
    Set rsTTLogTrait = m_dataSource.OpenRecordset("SELECT * FROM TTLOGTRAIT WHERE NUTRAITP3I=" & NumeroLot, Dynamic)
        
    If rsTTLogTrait.EOF = True Then
      rsTTLogTrait.AddNew
    
      rsTTLogTrait.fields("NUTRAITP3I") = rsIn.fields("NUTRAITP3I")
      rsTTLogTrait.fields("NUTRAIT") = rsIn.fields("NUTRAIT")
      rsTTLogTrait.fields("DTTRAIT") = rsIn.fields("DTTRAIT")
      rsTTLogTrait.fields("HHTRAIT") = rsIn.fields("HHTRAIT")
      rsTTLogTrait.fields("DTDEBPER") = rsIn.fields("DTDEBPER")
      rsTTLogTrait.fields("DTFINPER") = rsIn.fields("DTFINPER")
      rsTTLogTrait.fields("IDTABLESAS") = rsIn.fields("IDTABLESAS")
      rsTTLogTrait.fields("NBLIGTRAIT") = rsIn.fields("NBLIGTRAIT")
      rsTTLogTrait.fields("MTTRAIT") = rsIn.fields("MTTRAIT")
      
      rsTTLogTrait.Update
    End If
    
    rsTTLogTrait.Close
    
    rsIn.Close
  
  End If
  
  Exit Sub
  
err_CopyLot:
  If Not autoMode Then
    MsgBox "Erreur dans CopieVersTTLogTrait() : " & Err & vbLf & Err.Description, vbCritical
  Else
    autoLogger.EcritTraceDansLog "Erreur dans CopieVersTTLogTrait() : " & Err & vbLf & Err.Description
  End If
  
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Export vers TTLOGTRAIT et TTPROVCOLL
'
' TTLOGTRAIT : on recopie les données de 'Assure_P3ILOGTRAIT' et on mets à
'jour le nb de ligne 'NBLIGTRAIT'
' TTPROVCOLL : on recopie les données de 'Assure_P3IPROVCOLL' et on mets à
'jour à partir des données de 'Assure' calculées
'              et on ajoute une ligne pour les MGDC
'
'##ModelId=5C8A67A50127
Public Sub btnExportSAS_Click()

  On Error GoTo err_export
  
  Dim StartTime As Double, EndTime As Double

  '
  ' Selection du type de provision
  '
  Dim frm As frmTypeExport, sTypeProvision As String
  Dim bDeleteExistant As Boolean, bCreateSignalisation As Boolean, fErreur As Boolean
  Dim Logger As clsLogger
  
  Set frm = New frmTypeExport
  
  frm.frmTypeProvision = m_dataHelper.GetParameterAsLong("SELECT IdTypeCalcul FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & frmNumPeriode)
  
  ret_code = 0
  frm.Show vbModal
  If ret_code = 0 Then
    
    ' nom standardise Generali
    Select Case frm.frmTypeProvision
      Case 1
        sTypeProvision = "BILAN"
        
      Case 2
        sTypeProvision = "CLIENT"
        
      Case 3
        sTypeProvision = "SIMUL"
             
    End Select
    
    ' ajout à l'existant ?
    bDeleteExistant = frm.frmDelExistant
    
    ' creation du fichier de signalisation (pour le robot de l'infocentre)
    bCreateSignalisation = frm.frmCreateSignalisation
    
  Else
    Exit Sub
  End If
  
  'Unload frm
  Set frm = Nothing

  
  Dim fWait As frmWait
  Dim maxRecord As Long
 
  Set fWait = New frmWait
  
  fWait.Caption = "Export en cours..."
  
  fWait.ProgressBar1.Min = 0
  fWait.ProgressBar1.Value = 0
  fWait.ProgressBar1.Max = 100
  
  Screen.MousePointer = vbHourglass
 
  fWait.Show vbModeless
  fWait.Refresh
  
  
  ' prépare le log
  
  Set Logger = New clsLogger
  
  Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurExport.log"
  
  '  chargement des paramètres de la période
  Dim rsPeriode As ADODB.Recordset ' recordset Période
  Dim rq As String

  rq = "SELECT PE.*, PA.*, LG.* FROM Periode PE INNER JOIN ParamCalcul PA " _
      & " ON PA.PEGPECLE = PE.PEGPECLE AND PA.PENUMCLE = PE.PENUMCLE " _
      & " LEFT OUTER JOIN P3IUser.Assure_P3ILOGTRAIT AS LG ON PE.PEGPECLE = LG.CleGroupe AND PE.PENUMCLE = LG.NumPeriode AND PE.NUTRAITP3I = LG.NUTRAITP3I " _
      & " WHERE PE.PEGPECLE=" & GroupeCle & " AND PE.PENUMCLE=" & frmNumPeriode _
      & " ORDER BY PA.PENUMPARAMCALCUL"
  
  Set rsPeriode = m_dataSource.OpenRecordset(rq, Disconnected)
  
  If rsPeriode.EOF Then
    Screen.MousePointer = vbDefault
    
    Logger.EcritTraceDansLog "Les informations du groupe " & GroupeCle & " période " & frmNumPeriode & " sont absentes des Tables Periode et ParamCalcul"
    
    fWait.Hide
    Unload fWait
    
    Set fWait = Nothing
    
    ' affiche les erreurs
    Logger.AfficheErreurLog
      
    Exit Sub
  End If
  
  Logger.CreateLog "Export de la période " & numPeriode & " vers le lot " & rsPeriode.fields("NUTRAITP3I")
  
  If IsNull(rsPeriode.fields("NUTRAITP3I")) Then
    Logger.EcritTraceDansLog "Export impossible : aucun lot attaché (NUTRAITP3I non renseigné). Vérifier le type de période ou l'origine des données."
    Logger.EcritTraceDansLog "Export annulé !"
    
    Screen.MousePointer = vbDefault
    
    fWait.Hide
    Unload fWait
    
    Set fWait = Nothing
    
    ' affiche les erreurs
    Logger.AfficheErreurLog
      
    Exit Sub
  End If
  
  
  m_dataSource.BeginTrans
  
  'StartTime = Timer
  ' TTLOGTRAIT
  Logger.EcritTraceDansLog "TTLOGTRAIT..."
  CopieVersTTLogTrait GroupeCle, frmNumPeriode, rsPeriode.fields("NUTRAITP3I"), Logger, bDeleteExistant, sTypeProvision
  
  'EndTime = Timer
  'Debug.Print "CopieVersTTLogTrait: ", EndTime - StartTime
  
   
  ' TTPROVCOLL
  Dim rsAssure As ADODB.Recordset, rsCodeCatInval As ADODB.Recordset ', rsAssureP3IProvColl As ADODB.Recordset, rsTTPROVCOLL As ADODB.Recordset
  
  
  'deja fait dans CopieVersTTLogTrait : m_dataSource.Execute "DELETE FROM TTPROVCOLL WHERE NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I")
  
    
  Set rsCodeCatInval = m_dataSource.OpenRecordset("SELECT * FROM CODECATINV WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & frmNumPeriode, Disconnected)
  
  'StartTime = Timer
  
  Set rsAssure = m_dataSource.OpenRecordset("SELECT * FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & frmNumPeriode, Snapshot)
'  Set rsAssureP3IProvColl = m_dataSource.OpenRecordset("SELECT * FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & frmNumPeriode & " AND NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I"), Snapshot)
'  Set rsTTPROVCOLL = m_dataSource.OpenRecordset("SELECT * FROM TTPROVCOLL", Dynamic)
    
  'EndTime = Timer
  'Debug.Print "OpenRecordSet rsAssure: ", EndTime - StartTime
    
'  rsAssure.MoveLast
'  rsAssure.MoveFirst
  maxRecord = rsAssure.RecordCount + 1
  
  If maxRecord <> 0 Then
    fWait.ProgressBar1.Max = maxRecord
  Else
    fWait.ProgressBar1.Max = 1
    fWait.Hide
    MsgBox "Aucun article trouvé pour la période spécifiée", vbInformation
  End If
    
  Logger.EcritTraceDansLog "TTPROVCOLL..."
  
  
'##############################################################################

  'Close the CSV file, try to delete the csv file and then open it
  Dim fs As New FileSystemObject
  Dim fileCSV As String
  
  fileCSV = CSVUNCPath & GetWinUser & "_TTProvColl.csv"
  'fileCSV = CSVUNCPath & "_TTProvColl.csv"
  
  On Error Resume Next
  Close #1
  
  If fs.FileExists(fileCSV) Then
    fs.DeleteFile fileCSV
  End If
        
OpenCSV:
  
  Open fileCSV For Output As #1
  
  If Err.Number = 70 Then
  
    MsgBox "Il semble que le fichier " & fileCSV & _
          " est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok.", vbCritical
    
    Err.Clear
    GoTo OpenCSV
  
  End If
  
  On Error GoTo err_export
  
'##############################################################################

    
  '### Looping 66548 for Periode 891
  Do Until rsAssure.EOF
  
    'StartTime = Timer
    
    If (rsAssure.AbsolutePosition Mod 9) = 0 Then
      ' max record
      If maxRecord < rsAssure.RecordCount + 1 Then
        maxRecord = rsAssure.RecordCount + 1
        fWait.ProgressBar1.Max = maxRecord
      End If
      
      ' affiche la position
      fWait.ProgressBar1.Value = rsAssure.AbsolutePosition
      fWait.Label1(0).Caption = "Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max
      fWait.Refresh
      DoEvents
      
      If fWait.fTravailAnnule = True Then
        Exit Do
      End If
    End If
    
    '##### MODIFY
    'CopieVersTTProvColl_OLD GroupeCle, frmNumPeriode, rsPeriode, rsAssure, rsCodeCatInval, Logger, sTypeProvision
    CopieVersTTProvColl GroupeCle, frmNumPeriode, rsPeriode, rsAssure, rsCodeCatInval, Logger, sTypeProvision  ', rsAssureP3IProvColl, rsAssure_Retraite
      
    'EndTime = Timer
    'Debug.Print "CopieVersTTProvColl: ", EndTime - StartTime
    '3,19999999992433E-02
    
    DoEvents
    
    rsAssure.MoveNext
  Loop
  
  'close tthe CSV file and do a bulk insert - m_dataHelper
  Close #1
  
  '##### MODIFY
  If True Then
    If BulkInsert(m_dataSource.Connection, "TTPROVCOLL", fileCSV) = OperationStatus.efailure Then
      MsgBox "The BulkInsert operation into the table TTPROVCOLL failed!", vbCritical
    End If
  End If
    
  
'  rsTTPROVCOLL.Close
'  rsAssureP3IProvColl.Close
  rsCodeCatInval.Close
  rsAssure.Close
  
  
  If fWait.fTravailAnnule = False Then
    
    m_dataSource.CommitTrans
    
    ' Update NBLIGTRAIT et message nb lignes exportées
    maxRecord = m_dataHelper.GetParameterAsDouble("SELECT count(*) FROM TTPROVCOLL WHERE NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I"))
    Dim Cumul_MTPROIMP As Double
    Cumul_MTPROIMP = m_dataHelper.GetParameterAsDouble("SELECT SUM(MTPROIMP) FROM TTPROVCOLL WHERE NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I"))
    m_dataSource.Execute "UPDATE TTLOGTRAIT SET NBLIGTRAIT=" & maxRecord & " WHERE NUTRAITP3I=" & rsPeriode.fields("NUTRAITP3I")
    'Logger.EcritTraceDansLog "Export de " & maxRecord & " lignes pour le Lot " & rsPeriode.fields("NUTRAITP3I")
    Logger.EcritTraceDansLog "Export de " & maxRecord & " lignes et Cumul Provisions MTPROIMP = " & Format(Cumul_MTPROIMP, "# ##0.00") & " € pour le Lot " & rsPeriode.fields("NUTRAITP3I")
  
    ' Creation du fichier top de signalisation
    If bCreateSignalisation = True Then
      CreationFichierSignalisation
    End If
  
  Else
    
    m_dataSource.RollbackTrans
  
    Logger.EcritTraceDansLog "Export annulé !"
  
  End If
  
  
  rsPeriode.Close
  
  
  'MsgBox "Time: " & Now
     
      
  fWait.Hide
  Unload fWait
  
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  ' affiche les erreurs
  Logger.AfficheErreurLog
  
  Exit Sub

err_export:
  If Logger Is Nothing Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  fErreur = True
  
  Resume Next
End Sub


'****************************************************************************
'******************** OLD - TO BE DELETED ***********************************
'****************************************************************************

'##ModelId=5C8A67A50136
Private Sub RefreshListeOld()
  'Dim rs As Recordset
  
  Dim rs As ADODB.Recordset
  Dim rq As String, filter As String, sWhere As String, sResultingQuery As String, sFrom As String
  Dim i As Integer
  
  Dim debut As Date, fin As Date
  
  debut = Now
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
'  Me.Caption = "Assurés de la période " & frmNumPeriode
'  Me.Caption = Me.Caption & " (" & m_dataHelper.GetParameterAsStringCRW("SELECT 'Type ' + CAST(P.PETYPEPERIODE as VARCHAR) + ' - ' + T.Libelle FROM Periode P LEFT JOIN TypePeriode T ON T.IdTypePeriode=P.PETYPEPERIODE WHERE P.PEGPECLE = " & GroupeCle & " AND P.PENUMCLE = " & frmNumPeriode)
'  Me.Caption = Me.Caption & ") du Groupe '" & NomGroupe & "' : " & m_dataHelper.GetParameterAsStringCRW("SELECT CASE WHEN LEN(PECOMMENTAIRE)>40 THEN left(PECOMMENTAIRE, 40)+'...' ELSE PECOMMENTAIRE END as COMMENTAIRE FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle)
'
  lblFilter = fmFilter.SelectionString
  filter = fmFilter.GetSelectionSQLString
  
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
  
    
  sprListe.DAutoCellTypes = True
  sprListe.DAutoSizeCols = True
  
  ' Type de période
  Dim typePeriode As Integer
  typePeriode = m_dataHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE=" & frmNumPeriode)
  frmTypePeriode = typePeriode
    
  If Not m_DetailAffichagePeriode.DonneesBrutes Then
    
    rq = "SELECT Assure.RECNO, Assure.NUENRP3I as NUENRP3I, CAST(Assure.PONumParamCalcul AS varchar(20))+ ' - ' + ParamCalcul.PENOMPARAM AS [Param Calcul], " _
          & " Societe.SONOM AS Société, Assure.PONUMCLE AS Sinistre, Assure.PONumSinistre AS Dossier, Assure.POCONVENTION AS Contrat, " _
          & " CodesCat.RgrtPdtSup AS RegSupProduit, Assure.POGARCLE_NEW AS GE, Garantie.GALIB AS Garantie, CodeProvision.Libelle AS [Code Provision], " _
          & " CodePosition.Libelle AS Position, CAST(Assure.POCategorieInval as tinyint) as [Categorie Inval], Cast(Assure.POBaseRevalo as bit) AS [Base/Revalo], Assure.PONOM AS [Nom Assuré], " _
          & " Assure.PONAIS AS [Date de Naissance], Assure.POARRET AS [Date du sinistre], " _
          & " Assure.PODATEENTREEINVAL AS [Date Entrée Inval], Assure.POTERME AS [Extinction Garantie], Assure.PODEBUT AS DTDEBPER, " _
          & " Assure.POFIN AS DTFINPER, Assure.POPRESTATION_AN AS Annualisation, " _
          & " Assure.POPRESTATION_AN_PASSAGE AS [Annualisation Passage], " _
          & " Round(ISNULL(Assure.POPM,0) + ISNULL(Assure.POPSAP,0),2) AS [Provisions Imputées], " _
          & " Round(ISNULL(Assure.POPM,0) + ISNULL(Assure.POPSAP,0) + ISNULL(Assure.POPM_REVALO,0),2) AS [Provision Imputées avec Revalorisation], " _
          & " Assure.POPSAP as PSAP, Round(ISNULL(Assure.POPM,0), 2) AS [Provisions Calculées], " _
          & " Assure.POPM_INCAP_1F AS [Coeff Incap 12€], Assure.POPM_PASS_1F AS [Coeff Pass 1€], Assure.POPM_INVAL_1F AS [Coeff Inval 1€], " _
          & " Assure.POPM_REDUC_1F AS [Coeff Rente 1€], Assure.POCoeffBCAC AS [Coefficient BCAC], " _
          & " Round(Assure.POPM_INCAP_1F*Assure.POPRESTATION_AN/12,2) AS [PM Incap], Round(Assure.POPM_PASS_1F*Assure.POPRESTATION_AN_PASSAGE,2) AS [PM Pass], " _
          & " Round(Assure.POPM_INVAL_1F*Assure.POPRESTATION_AN,2) AS [PM Inval], Round(Assure.POPM_REDUC_1F*Assure.POPRESTATION_AN,2) AS [PM Rente], " _
          & " CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM,0),2) ELSE NULL END AS [PM MGDC], Assure.POPM_REVALO AS [Provision Revalorisation], Assure.POPM+Assure.POPM_REVALO AS [Provision avec Revalorisation], " _
          & " Assure.POCOT_REVALO AS [Cotisation Revalorisation], Assure.POPM_INCAP_1R AS [Coeff Incap 12R], Assure.POPM_PASS_1R AS [Coeff Pass 1R], " _
          & " Assure.POPM_INVAL_1R AS [Coeff Inval 1R], Assure.POPM_REDUC_1R AS [Coeff Rente 1R], Assure.POCoeffBCACR AS [Coefficient BCAC R], " _
          & " Round(Assure.POPM_INCAP_1R*Assure.POPRESTATION_AN/12,2) AS [PM Incap R], Round(Assure.POPM_PASS_1R*Assure.POPRESTATION_AN_PASSAGE,2) AS [PM Pass R], " _
          & " Round(Assure.POPM_INVAL_1R*Assure.POPRESTATION_AN,2) AS [PM Inval R], Round(Assure.POPM_REDUC_1R*Assure.POPRESTATION_AN,2) AS [PM Rente R], " _
          & " CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM+Assure.POPM_REVALO,0),2) ELSE NULL END AS [PM MGDC R], Assure.POCOMMENTANNUL AS [Raison Annulation], Assure.Commentaire AS Commentaire "
    
    sFrom = "FROM Societe INNER JOIN Assure ON Societe.SOCLE = Assure.POSTECLE " _
            & " INNER JOIN Garantie ON Assure.POGARCLE = Garantie.GAGARCLE  " _
            & " LEFT JOIN CodesCat ON Assure.POGPECLE = CodesCat.GroupeCle AND Assure.POPERCLE = CodesCat.NumPeriode AND Assure.POCATEGORIE = CodesCat.Code_Cat_Contrat AND Assure.POCompagnie=CodesCat.Code_Cie AND Assure.POAppli=CodesCat.Code_APP  " _
            & " LEFT JOIN CodePosition ON Assure.POSIT = CodePosition.Position " _
            & " LEFT JOIN CodeProvision ON Assure.POCATEGORIE_NEW = CodeProvision.CodeProv " _
            & " LEFT JOIN TypeTermeEchu ON Assure.POECHU = TypeTermeEchu.IdTypeTermeEchu " _
            & " LEFT JOIN TypeFractionnement ON Assure.POFRACT = TypeFractionnement.IdTypeFractionnement " _
            & " LEFT JOIN ParamCalcul ON Assure.POGPECLE = ParamCalcul.PEGPECLE AND Assure.POPERCLE = ParamCalcul.PENUMCLE AND Assure.PONumParamCalcul = ParamCalcul.PENUMPARAMCALCUL " _
            & " LEFT JOIN SituationFamille ON Assure.POCleSituationFamille = SituationFamille.CleSituationFamille "

    sWhere = " WHERE Assure.POPERCLE = " & frmNumPeriode & " AND Assure.POGPECLE = " & GroupeCle

  Else
    
    rq = "SELECT Assure.RECNO, Assure.NUENRP3I as NUENRP3I, CAST(Assure.PONumParamCalcul AS varchar(20)) + ' - ' + ParamCalcul.PENOMPARAM AS [Param Calcul], " _
            & "Societe.SONOM AS Société, Assure.PONUMCLE AS Sinistre, Assure.PONumSinistre AS Dossier, Assure.POCONVENTION AS Contrat, " _
            & "Assure.POCode_Option_Contrat AS CodeOption, Assure.POCATEGORIE AS Categorie, Assure.POTypeMvt AS [Type Mvt], Assure.POContractant AS Contractant, " _
            & "CodesCat.RgrtPdtSup AS RegSupProduit, CodesCat.RgrtPdt AS RegProduit, Assure.PONOM AS [Nom Assuré], Assure.POGARCLE_NEW AS GE, " _
            & "Garantie.GALIB AS Garantie, CodeProvision.Libelle AS [Code Provision], CodePosition.Libelle AS Position, CAST(Assure.POCategorieInval as char(1)) as [Categorie Inval], TypeFractionnement.Libelle AS Fractionnement, " _
            & "TypeTermeEchu.Libelle AS TermeEchu, Cast(Assure.POBaseRevalo as bit) AS [Base/Revalo], Assure.PONAIS AS [Date de Naissance], " _
            & "Assure.POARRET AS [Date du sinistre], Assure.PODATEENTREEINVAL AS [Date Entrée Inval], Assure.POEFFET AS Effet, " _
            & "Assure.POTERME AS [Extinction Garantie], Assure.POPREMIER_PAIEMENT AS [Premier paiement], Assure.PODERNIERPAIEMENT AS [Dernier paiement], " _
            & "Assure.PODEBUT AS DTDEBPER, Assure.POFIN AS DTFINPER, Assure.PODEBUTTOTAL AS [Début Total], Assure.POFINTOTAL AS [Fin Total], " _
            & "Assure.POPRESTATION_AN AS Annualisation,  " _
            & " Assure.POPRESTATION_AN_PASSAGE AS [Annualisation Passage], " _
            & "Round(ISNULL(Assure.POPM,0),2) AS [Provisions Calculées], Round(ISNULL(Assure.POPM,0) + ISNULL(Assure.POPSAP,0),2) AS [Provisions Imputées], " _
            & "Round(ISNULL(Assure.POPM,0) + ISNULL(Assure.POPSAP,0) + ISNULL(Assure.POPM_REVALO,0),2) AS [Provision Imputées avec Revalorisation], Assure.POPSAP AS PSAP, " _
            & "Round(Assure.POPM_INCAP_1F*Assure.POPRESTATION_AN/12,2) AS [PM Incap], Round(Assure.POPM_PASS_1F*Assure.POPRESTATION_AN_PASSAGE,2) AS [PM Pass], " _
            & "Round(Assure.POPM_INVAL_1F*Assure.POPRESTATION_AN,2) AS [PM Inval], Round(Assure.POPM_REDUC_1F*Assure.POPRESTATION_AN,2) AS [PM Rente], " _
            & "CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM,0),2) ELSE NULL END AS [PM MGDC], Assure.POPM_X AS [Age Arrêt], Assure.POPM_XTERME AS [Age Terme], " _
            & "Assure.POPM_ANC AS Anc, Assure.POPM_DUREE AS Durée, " _
            & "Assure.POPM_INCAP_1F AS [Coeff Incap 12€], Assure.POPM_PASS_1F AS [Coeff Pass 1€], Assure.POPM_INVAL_1F AS [Coeff Inval 1€], " _
            & "Assure.POPM_REDUC_1F AS [Coeff Rente 1€], " _
            & "Assure.POCoeffBCAC AS [Coefficient BCAC], Assure.TXTECHN as [Tx Tech], Assure.TXFRAIS as [Frais], "
              
    rq = rq & "Assure.POSEXE AS Sexe, Assure.POCSP AS CSP, Assure.POPourcentLissage AS [Pct Lissage Provision], Assure.POPM_AvecCorrectif AS [PM Incap Avec Correctif], " _
            & "Assure.POPM_SansCorrectif AS [PM Incap Sans Correctif], Assure.POCorrectif AS [Correctif PM Incap], Assure.POPMAvecCorrectif AS [PM Incap Avec Corr], " _
            & "Assure.POPMReassAvecCorrectif AS [Reass PM Incap Avec Corr], Assure.POMontantBase AS [Montant Base], Assure.POMontantRevalo AS [Montant Revalo], " _
            & "Assure.POTypeReglement AS [Type Rglt], Assure.PODebutIndemnisation AS [Debut Indemn], Assure.PONbJourIndemn AS [nb Jour Indemn], " _
            & "Assure.POREPRISE AS Reprise, Assure.POCAUSE AS Cause, Assure.POTYPEF AS [Franchise en jours], Assure.PODELAI AS [délai de carence], " _
            & "Cast(Assure.PODATEPAIEMENTESTIMEE as bit) AS [Date estimée], Assure.POPRESTATION_AN_PREC AS [Annualisation précédente], " _
            & "Assure.POPM_VAR AS [Variation PM], Assure.POPSAPCAPMOYEN AS [Capital Moyen], Assure.POPM_RI AS [Provision Relative], " _
            & "Assure.POPM_REVALO AS [Provision Revalorisation], Assure.POPM+Assure.POPM_REVALO AS [Provision avec Revalorisation], " _
            & "Assure.POCOT_REVALO AS [Cotisation Revalorisation], Assure.POPM_INCAP_1R AS [Coeff Incap 12R], Assure.POPM_PASS_1R AS [Coeff Pass 1R], " _
            & "Assure.POPM_INVAL_1R AS [Coeff Inval 1R], Assure.POPM_REDUC_1R AS [Coeff Rente 1R], Assure.POCoeffBCACR AS [Coefficient BCAC R], Assure.TXTECHNR as [Tx Tech R], Assure.TXFRAISR as [Frais R], " _
            & "Round(Assure.POPM_INCAP_1R*Assure.POPRESTATION_AN/12,2) AS [PM Incap R], Round(Assure.POPM_PASS_1R*Assure.POPRESTATION_AN_PASSAGE,2) AS [PM Pass R], " _
            & "Round(Assure.POPM_INVAL_1R*Assure.POPRESTATION_AN,2) AS [PM Inval R], Round(Assure.POPM_REDUC_1R*Assure.POPRESTATION_AN,2) AS [PM Rente R], " _
            & "CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM+Assure.POPM_REVALO,0),2) ELSE NULL END AS [PM MGDC R], Assure.POTRAITE_RASSUR AS Traité, Assure.POPRESTA_RASSUR AS [Prestation Reass], " _
            & "Assure.POPM_RASSUR AS [PM Reass], Assure.POPSAP_RASSUR AS [PSAP Reass], cast(Assure.PODOSSIERCLOS as bit) AS [Dossier Clos], " _
            & "Assure.POCCN AS CCN, Assure.POCODERISQUE AS [Code Risque], " _
            & "Cast(Assure.POIsCadre as bit) AS Cadre, Assure.POSalaireAnnuel AS [Salaire Annuel], SituationFamille.Libelle AS [Situation Famille], " _
            & "Assure.POTauxGarantieDC AS [PMGD Taux Gar Décès], Assure.PONbEnfant AS [Nb Enfants], Assure.POAgeMoyenEnfant AS [Age Moyen Enfant], " _
            & "Assure.POMajoEnfant AS [Majoration Enfant], Assure.PORegimeDeces AS [Régime MG Décès], Assure.POCategorieDeces AS [Catégorie MG Décès], "

    rq = rq & "Assure.PORegimeRenteConjointTempo AS [Régime MG Rte Conjoint Temp], Assure.POCategorieRenteConjointTempo AS [Catégorie MG Rte Conjoint Temp], " _
            & "Assure.PORegimeRenteConjointViagere AS [Régime MG Rte Conjoint Viagère], Assure.POCategorieRenteConjointViagere AS [Catégorie MG Rte Conjoint Viagère], " _
            & "Assure.PORegimeRenteEduc AS [Régime MG Rte Educ], Assure.POCategorieRenteEduc AS [Catégorie MG Rte Educ], " _
            & "Assure.POCaptive AS Captive, Assure.POSituConv AS [Situ Conv], Assure.POEffetSitu AS [Effet Situ], " _
            & "Assure.POEtablissement AS Etablissement, Assure.POCreationDossier AS [Création Dossier], Assure.PODebutDossier AS [Debut Dossier], " _
            & "Assure.POFinDossier AS [Fin Dossier], Assure.POMotifCloture AS [Motif Cloture], Assure.PODebutRefSalaire AS [Debut Ref Salaire], " _
            & "Assure.POFinRefSalaire AS [Fin Ref Salalaire], " _
            & "Assure.POFamilleComptable AS [Famille Comptable], Assure.POInspecteur AS Inspecteur, Assure.POInsp AS Insp, Assure.POApport AS Apport, " _
            & "Assure.POApport2 AS Apport2, Assure.POGestionnaire AS Gestionnaire, Assure.POIndicCC AS IndicCC, Assure.POCompagnie AS Compagnie, Assure.POAppli AS Application, " _
            & "Assure.PONbIntervenant AS NbIntervenant, Assure.POPMANNULEE AS [PM Annulée], Assure.POPSAPANNULEE AS [PSAP Annulée], " _
            & "Assure.POCDCONTENTIEUX as CDCONTENTIEUX, Assure.PONUSINISTRE as NUSINISTRE, " _
            & "Assure.POMontantCapConstit as [Capital Constitutif], Assure.POMontantCapSousRisque as [Capital Sous-Risque], " _
            & "Assure.POCOMMENTANNUL AS [Raison Annulation], Assure.Commentaire AS Commentaire "
              
'            & "Assure.POPRESTATION AS [Prestations Payées], Assure.POPRESTATIONTOTAL AS [Prestations Payées TOTAL], "

    sFrom = "FROM Societe INNER JOIN Assure ON Societe.SOCLE = Assure.POSTECLE " _
            & " INNER JOIN Garantie ON Assure.POGARCLE = Garantie.GAGARCLE  " _
            & " LEFT JOIN CodesCat ON Assure.POGPECLE = CodesCat.GroupeCle AND Assure.POPERCLE = CodesCat.NumPeriode AND Assure.POCATEGORIE = CodesCat.Code_Cat_Contrat AND Assure.POCompagnie=CodesCat.Code_Cie AND Assure.POAppli=CodesCat.Code_APP " _
            & " LEFT JOIN CodePosition ON Assure.POSIT = CodePosition.Position " _
            & " LEFT JOIN CodeProvision ON Assure.POCATEGORIE_NEW = CodeProvision.CodeProv " _
            & " LEFT JOIN TypeTermeEchu ON Assure.POECHU = TypeTermeEchu.IdTypeTermeEchu " _
            & " LEFT JOIN TypeFractionnement ON Assure.POFRACT = TypeFractionnement.IdTypeFractionnement " _
            & " LEFT JOIN ParamCalcul ON Assure.POGPECLE = ParamCalcul.PEGPECLE AND Assure.POPERCLE = ParamCalcul.PENUMCLE AND Assure.PONumParamCalcul = ParamCalcul.PENUMPARAMCALCUL " _
            & " LEFT JOIN SituationFamille ON Assure.POCleSituationFamille = SituationFamille.CleSituationFamille "
    
    sWhere = " WHERE Assure.POPERCLE = " & frmNumPeriode & " AND Assure.POGPECLE = " & GroupeCle
  
  End If
          
  'sprListe.MaxRows = 0
  'sprListe.VirtualMaxRows = 0
  
  'dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq & filter & " AND RECNO=0 " & " ORDER BY " & frmOrdreDeTri)
  'dtaPeriode.Refresh
  
  ' mise en forme des données
  If m_DetailAffichagePeriode.DonneesBrutes Then
    
    If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
      SetColonneDataFill 3, True ' TypeLigne
      SetColonneDataFill 17, True ' Garantie
    Else
      SetColonneDataFill 16, True ' Garantie
    End If
  
  Else
    
    If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
      SetColonneDataFill 3, True ' TypeLigne
      SetColonneDataFill 11, True ' Garantie
    Else
      SetColonneDataFill 10, True ' Garantie
    End If
  
  End If


  If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
    If m_DetailAffichagePeriode.Avant Then
      ' Ligne Avant
'      sResultingQuery = Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '1 - Avant' as TypeLigne,"), "Assure.POCOMMENTANNUL AS Commentaire", "Null as [Top Amortissable], Null as [Coeff Amortissement], Null as [Age Mini Départ Retraite], Null as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS Commentaire ") & sFrom & sWhere & filter
      sResultingQuery = ProcessQuery(rq, sFrom, sWhere, filter, "1 - Avant")
    End If
    
    If m_DetailAffichagePeriode.Apres Then
      ' Ligne Après
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
'      sResultingQuery = sResultingQuery & Replace(Replace(Replace(rq, "as NUENRP3I,", "as NUENRP3I, '2 - Après' as TypeLigne,"), "Assure.POCOMMENTANNUL AS Commentaire", "Assure.POTopAmortissable as [Top Amortissable], 100*Assure.POCoeffAmortissement as [Coeff Amortissement], Assure.POAgeMiniDepartRetraite as [Age Mini Départ Retraite], Assure.POAgeDepartRetraiteTauxPlein as [Age Départ Retraite Taux Plein], Assure.POCOMMENTANNUL AS Commentaire "), "Assure", "Assure_Retraite") _
'                        & Replace(sFrom, "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO " & Replace(sWhere & filter, "Assure", "Assure_Retraite")
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "2 - Après")
    End If
    
    If m_DetailAffichagePeriode.Ecart Then
      ' Ligne Ecart
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "3 - Ecart")
    End If
    
    If m_DetailAffichagePeriode.DejaAmorti Then
      ' Ligne Déjà amorti
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "4 - Amorti")
    End If
    
    If m_DetailAffichagePeriode.ResteAAmortir Then
      ' Ligne Reste à amortir
      If sResultingQuery <> "" Then
        sResultingQuery = sResultingQuery & " UNION ALL "
      End If
      
      sResultingQuery = sResultingQuery & ProcessQuery(rq, sFrom, sWhere, filter, "5 - A Amortir")
    End If
    
    ' Tri
    sResultingQuery = sResultingQuery & " ORDER BY " & frmOrdreDeTri & IIf(InStr(1, frmOrdreDeTri, "NUENRP3I") = 0, ", NUENRP3I", "") & IIf(InStr(1, frmOrdreDeTri, "Garantie") = 0, ", Garantie", "") & IIf(InStr(1, frmOrdreDeTri, "[Base/Revalo]") = 0, ", [Base/Revalo]", "") & IIf(InStr(1, frmOrdreDeTri, "TypeLigne") = 0, ", TypeLigne", "")
  Else
    sResultingQuery = rq & sFrom & sWhere & filter & " ORDER BY " & frmOrdreDeTri
  End If
  
  sResultingQuery = m_dataHelper.ValidateSQL(sResultingQuery)
  
'  Clipboard.Clear
'  Clipboard.SetText sResultingQuery
  
  dtaPeriode.RecordSource = sResultingQuery
  dtaPeriode.Refresh
    
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If Not dtaPeriode.Recordset.EOF Then
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
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
  
  Dim decal As Integer
  
  If (eProvisionRetraite = typePeriode Or eProvisionRetraiteRevalo = typePeriode) Then
    sprListe.ColsFrozen = 3
    decal = 1
  Else
    sprListe.ColsFrozen = 2
    decal = 0
  End If
  
  ' Couleurs des colonnes
  If m_DetailAffichagePeriode.DonneesBrutes Then
    
    ' mode detail
    SetColBackColor 1, 10 + decal, bleu_clair
    
    SetColBackColor 11 + decal, 4, jaune_clair
    
    SetColBackColor 15 + decal, 1, bleu_clair
  
    SetColBackColor 16 + decal, 6, jaune_clair
    
    SetColBackColor 22 + decal, 1, bleu_clair
    
    SetColBackColor 23 + decal, 11, orange_clair

'   code avant modif du 14/06/2013
    
'    SetColBackColor 34 + decal, 10, vert_clair
    
'    SetColBackColor 44 + decal, 4, orange_clair
'
'    SetColBackColor 48 + decal, 7, lavande_clair
'
'    SetColBackColor 55 + decal, 2, vbWindowBackground
'
'    SetColBackColor 57 + decal, 8, vert_clair
'
'    SetColBackColor 65 + decal, 1, vbWindowBackground
'
'    SetColBackColor 66 + decal, 1, orange_clair
'
'    SetColBackColor 67 + decal, 1, lavande_clair
'
'    SetColBackColor 68 + decal, 9, vbWindowBackground
'
'    SetColBackColor 77 + decal, 3, vert_clair
'
'    SetColBackColor 80 + decal, 7, lavande_clair
'
'    SetColBackColor 87 + decal, 5, vert_clair
'
'    SetColBackColor 92 + decal, sprListe.MaxCols - 89, vbWindowBackground
'
'    SetColBackColor 138 + decal, 2, vert_clair
'
'    SetColBackColor 140 + decal, 1, vbWindowBackground
'
'    SetColBackColor 141 + decal, 3, orange_clair
'
'    SetColBackColor 144 + decal, 1, vbWindowBackground


    SetColBackColor 34 + decal, 11, vert_clair
    
    SetColBackColor 45 + decal, 4, orange_clair
    
    SetColBackColor 49 + decal, 7, lavande_clair
    
    SetColBackColor 56 + decal, 2, vbWindowBackground
    
    SetColBackColor 58 + decal, 8, vert_clair
  
    SetColBackColor 66 + decal, 1, vbWindowBackground
  
    SetColBackColor 67 + decal, 1, orange_clair
  
    SetColBackColor 68 + decal, 1, lavande_clair
    
    SetColBackColor 69 + decal, 9, vbWindowBackground
    
    SetColBackColor 78 + decal, 3, vert_clair
    
    SetColBackColor 81 + decal, 7, lavande_clair
    
    SetColBackColor 88 + decal, 5, vert_clair
    
    SetColBackColor 93 + decal, sprListe.MaxCols - 89, vbWindowBackground
    
    SetColBackColor 139 + decal, 2, vert_clair
    
    SetColBackColor 141 + decal, 1, vbWindowBackground
    
    SetColBackColor 142 + decal, 3, orange_clair
    
    SetColBackColor 145 + decal, 1, vbWindowBackground

  Else
    
    ' mode compact
    SetColBackColor 1, 7 + decal, bleu_clair
    
    SetColBackColor 8 + decal, 1, jaune_clair
    
    SetColBackColor 9 + decal, 1, bleu_clair
  
    SetColBackColor 10 + decal, 4, jaune_clair
    
    SetColBackColor 14 + decal, 1, bleu_clair
    
    SetColBackColor 15 + decal, 1, jaune_clair
    
    SetColBackColor 16 + decal, 6, orange_clair
    
'   code avant modif du 14/06/2013
    
'    SetColBackColor 22 + decal, 5, vert_clair
'
'    SetColBackColor 27 + decal, 5, lavande_clair
'
'    SetColBackColor 32 + decal, 8, vert_clair
'
'    SetColBackColor 40 + decal, 5, lavande_clair
'
'    SetColBackColor 45 + decal, 5, vert_clair
'
'    SetColBackColor 50 + decal, 1, vbWindowBackground
'
'    SetColBackColor 51 + decal, 3, orange_clair
'
'    SetColBackColor 54 + decal, 1, vbWindowBackground
    
    SetColBackColor 22 + decal, 6, vert_clair
    
    SetColBackColor 28 + decal, 5, lavande_clair
    
    SetColBackColor 33 + decal, 8, vert_clair
    
    SetColBackColor 41 + decal, 5, lavande_clair
    
    SetColBackColor 46 + decal, 5, vert_clair
    
    SetColBackColor 51 + decal, 2, vbWindowBackground
    
    SetColBackColor 53 + decal, 3, orange_clair
  
    SetColBackColor 56 + decal, 1, vbWindowBackground
  End If
  
  
  ' change le format des colonnes pour trier correctement les dates
#If TRI_SPREAD Then
  If chkDonneesBrutes.Value = 0 Then
    SetColonneDeTypeDate 6 ' POARRET
  Else
    SetColonneDeTypeDate 6 ' POARRET
    SetColonneDeTypeDate 14 ' PONAIS
    SetColonneDeTypeDate 15 ' POEFFET
    SetColonneDeTypeDate 16 ' POTERME
    SetColonneDeTypeDate 17 ' POREPRISE
    SetColonneDeTypeDate 21 ' POPREMIER_PAIEMENT
    SetColonneDeTypeDate 22 ' PODERNIER_PAIEMENT
  End If
#End If
   
  
  ' largeur des colonnes
  For i = 2 To sprListe.MaxCols - 2
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 5
  Next i
  sprListe.ColWidth(sprListe.MaxCols - 1) = 50
  sprListe.ColWidth(sprListe.MaxCols) = 50
  
  
pas_de_donnee:

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True

  Me.SetFocus
  sprListe.SetFocus
  
  Dim bLocked As Boolean
  
  bLocked = CBool(m_dataHelper.GetParameterAsDouble("SELECT PELOCKED FROM Periode WHERE PENUMCLE = " & frmNumPeriode & " AND PEGPECLE = " & GroupeCle))
  
  If bLocked = True Then
    btnCalc.Enabled = False
    btnCalcRevalo.Enabled = False
    btnImport.Enabled = False
    btnPurge.Enabled = False
  Else
    If archiveMode Then
      btnCalcRevalo.Enabled = False
      btnCalc.Enabled = False
      btnExportSAS.Enabled = False
      btnImport.Enabled = False
      
      'disable edit button on Toolbar
      Toolbar1.Buttons(3).Enabled = False
    Else
      btnCalc.Enabled = True
      btnCalcRevalo.Enabled = True
      btnImport.Enabled = True
      btnPurge.Enabled = True
    End If
  End If
  
  Screen.MousePointer = vbDefault

  fin = Now
  
  lblFillTime.text = "Remplissage : " & DateDiff("s", debut, fin) & " s"

  Exit Sub

err_RefreshListe:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


