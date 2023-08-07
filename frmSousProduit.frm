VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSousProduit 
   Caption         =   "Préparation des Sous-Produits"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
   Icon            =   "frmSousProduit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
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
      Begin VB.OptionButton rdoMaintien 
         Caption         =   "Maintien Décès"
         Height          =   195
         Left            =   5130
         TabIndex        =   7
         Top             =   90
         Width           =   1455
      End
      Begin VB.OptionButton rdoPrestation 
         Caption         =   "Prestations"
         Height          =   195
         Left            =   3915
         TabIndex        =   6
         Top             =   90
         Width           =   1140
      End
      Begin VB.TextBox lblFilter 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   6660
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "fdgfhgf"
         Top             =   90
         Width           =   6270
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
   Begin FPSpreadADO.fpSpread sprListe 
      Bindings        =   "frmSousProduit.frx":1BB2
      Height          =   5550
      Left            =   0
      TabIndex        =   5
      Top             =   495
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
      SpreadDesigner  =   "frmSousProduit.frx":1BCB
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "frmSousProduit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E70058"
Option Explicit

Public Enum displayMode
  SousProduit = 0
  Prestation
  MaintienDeces
  SommeParRegime
End Enum

'##ModelId=5C8A67E70171
Public fmFilter As clsFilter
'##ModelId=5C8A67E70174
Public fmMode As displayMode
'##ModelId=5C8A67E70181
Public fmTypePeriode As Integer
'

'##ModelId=5C8A67E701B0
Private Sub btnClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A67E701BF
Private Sub btnExport_Click()
  Dim filename As String, nomSheet As String, nomTable As String
  
  Select Case fmMode
    Case SousProduit
      filename = "Assure Periode " & numPeriode & ".xls"
      nomSheet = "Periode " & numPeriode & IIf(lblFilter <> "", " avec filtre", "")
      nomTable = ""
  
    Case Prestation
      filename = "Prestation Periode " & numPeriode & ".xls"
      nomSheet = "Prestations"
      nomTable = "Prestation"
      
    Case MaintienDeces
      filename = "Maintien Deces Periode " & numPeriode & ".xls"
      nomSheet = "MaintiensDeces"
      nomTable = "MaintienDeces"
    
    Case SommeParRegime
      filename = "Régimes Periode " & numPeriode & ".xls"
      nomSheet = "Régimes"
      nomTable = ""
  End Select
  
  ExportTableToExcelFile filename, nomSheet, nomTable, sprListe, CommonDialog1, lblFilter, False
End Sub

'##ModelId=5C8A67E701CF
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
    .PrintHeader = "/c Préparation de la liste par sous-produits : Assurés de la " & DescriptionPeriode & "/n  "
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

'##ModelId=5C8A67E701DF
Private Sub rdoMaintien_Click()
  fmMode = MaintienDeces
  
  RefreshListe
End Sub

'##ModelId=5C8A67E701EE
Private Sub rdoPrestation_Click()
  fmMode = Prestation
  
  RefreshListe
End Sub

'##ModelId=5C8A67E701FE
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A67E7024C
Private Sub Form_Activate()
  If fmMode = SousProduit Or fmMode = SommeParRegime Then
    rdoMaintien.Visible = False
    rdoPrestation.Visible = False
    
    lblFilter.Left = rdoPrestation.Left
  End If
  
  If fmMode = Prestation And rdoPrestation = False Then
    rdoPrestation = True
  Else
    RefreshListe
  End If
  
  ' Taille et position
  Me.top = 2000
  Me.Left = 250
  
  Me.Width = Screen.Width - Me.Left * 2
  Me.Height = 7000
End Sub

'##ModelId=5C8A67E7025C
Private Sub Form_Load()
  m_dataSource.SetDatabase dtaPeriode
  
  fmMode = SousProduit
End Sub

'##ModelId=5C8A67E7026B
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = Toolbar1.Height + 30
  sprListe.Left = 30
  sprListe.Width = Me.ScaleWidth - sprListe.Left - 30
  sprListe.Height = Me.ScaleHeight - sprListe.top - 30
End Sub

'##ModelId=5C8A67E7027B
Private Sub SetColonneDataFill(numCol As Integer)
  sprListe.Col = numCol
  sprListe.Col2 = numCol
  sprListe.Row = 1
  sprListe.Row2 = sprListe.MaxRows
  sprListe.BlockMode = True
  sprListe.DataFillEvent = True
  sprListe.BlockMode = False
End Sub

'##ModelId=5C8A67E702AA
Private Sub SetColBackColor(numCol As Integer, nbCol As Integer, color As OLE_COLOR)
  sprListe.BlockMode = True
  sprListe.Col = numCol
  sprListe.Col2 = sprListe.Col + nbCol - 1
  sprListe.Row = -1
  sprListe.Row2 = -1
  sprListe.BackColor = color
  sprListe.BackColorStyle = BackColorStyleUnderGrid
  sprListe.BlockMode = False
End Sub


'##ModelId=5C8A67E70307
Private Function GetColonneRef(numCol As Integer) As String
  GetColonneRef = ""

  Dim idx As Integer

  If numCol > 26 Then
    idx = Int(numCol / 26)
    GetColonneRef = GetColonneRef & Chr(Asc("A") + idx - 1)
  End If
  
  idx = (numCol Mod 26)
  If idx = 0 Then
    GetColonneRef = GetColonneRef & "Z"
  Else
    GetColonneRef = GetColonneRef & Chr(Asc("A") + idx - 1)
  End If

End Function

'##ModelId=5C8A67E70346
Private Sub RefreshListe()
  Dim rq As String, sWhere As String
  Dim filter As String
  Dim i As Integer
  
  Screen.MousePointer = vbHourglass
  
  rq = m_dataHelper.GetParameterAsStringCRW("SELECT CASE WHEN LEN(PECOMMENTAIRE)>40 THEN left(PECOMMENTAIRE, 40)+'...' ELSE PECOMMENTAIRE END as COMMENTAIRE FROM Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle)
  
  ' fabrique le titre de la fenetre en fonction du groupe en cours
  Select Case fmMode
    Case SousProduit
      Me.Caption = "Préparation de la liste par sous-produits : Assurés de la période " & numPeriode & " du Groupe '" & NomGroupe & "': " & rq
  
    Case Prestation
      Me.Caption = "Prestation : Assurés de la période " & numPeriode & " du Groupe '" & NomGroupe & "': " & rq
      
    Case MaintienDeces
      Me.Caption = "Maintiens en Garantie Décès : Assurés de la période " & numPeriode & " du Groupe '" & NomGroupe & "': " & rq
  
    Case SommeParRegime
      Me.Caption = "Régimes : Assurés de la période " & numPeriode & " du Groupe '" & NomGroupe & "': " & rq
  End Select
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  lblFilter = fmFilter.SelectionString
  
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
    
  ' & "Format(POCONVENTION, ""0 00 000000 00 00"") as [NCA],"
  
  Select Case fmMode
    Case SousProduit
      rq = "SELECT RECNO," _
             & " POGARCLE-50 as Regime," _
             & " POCATEGORIE as Categorie, " _
             & " POCONVENTION as [Convention]," _
             & " PONUMCLE as [Police]," _
             & " POARRET as [Date de Survenance]," _
             & " PONOM as [Nom Assuré]," _
             & " POPREMIER_PAIEMENT as [Premier paiement], PODERNIERPAIEMENT as [Dernier paiement], " _
             & " POPRESTATION as [Total Payé]," _
             & " POREPRISE as [Date de reprise]," _
             & " ' ' as [Date de depart], " _
             & " ' ' as [Motif de depart], " _
             & " ' ' as [Situation de famille], " _
             & " ' ' as [Nb enfants à charge], " _
             & " ' ' as [Salaire BRUT] "
      rq = rq & " FROM Assure WHERE POPERCLE = " & numPeriode & " AND POGPECLE = " & GroupeCle _
              & " AND POGARCLE < 90 "
    
    Case Prestation
      rq = "SELECT RECNO, " _
             & " POGARCLE-50 as Regime," _
             & " POCATEGORIE as Categorie, " _
             & " POCONVENTION as [Convention]," _
             & " PONOM as [NomPrenom]," _
             & " PONUMCLE as [Police]," _
             & " PONAIS as DateNaissance, " _
             & " POARRET as DateSurvenance, " _
             & " POREPRISE as DateReprise, " _
             & " PODATEENTREEINVAL as DateEntreeInval, " _
             & " PODERNIERPAIEMENT as DateComptable, " _
             & " POPRESTATION as MontantBrut, " _
             & " POPRESTA_RASSUR as MontantReass, " _
             & " POPRESTATION_AN as Annualisation, " _
             & " POPM as Provision, " _
             & " POPSAP as PSAP, " _
             & " POPM_RASSUR as ProvisionReass, " _
             & " POPSAP_RASSUR as PSAPReass"
      rq = rq & " FROM Assure WHERE POPERCLE = " & numPeriode & " AND POGPECLE = " & GroupeCle _
              & " AND POGARCLE < 90 "

    Case MaintienDeces
      rq = "SELECT RECNO, " _
             & " POGARCLE-90 as Regime," _
             & " POCATEGORIE as Categorie, " _
             & " POCONVENTION as [Convention]," _
             & " PONOM as [NomPrenom]," _
             & " PONUMCLE as [Police]," _
             & " PONAIS as DateNaissance, " _
             & " POARRET as DateSinistre, " _
             & " POREPRISE as DateReprise, " _
             & " PODATEENTREEINVAL as DateEntreeInval, " _
             & " PODERNIERPAIEMENT as DateComptable, " _
             & " POPRESTATION as MontantBrut, " _
             & " POPRESTA_RASSUR as MontantReass, " _
             & " POPRESTATION_AN as Annualisation, " _
             & " POPM as Provision, " _
             & " POPSAP as PSAP, " _
             & " POPM_RASSUR as ProvisionReass, " _
             & " POPSAP_RASSUR as PSAPReass"
      rq = rq & " FROM Assure WHERE POPERCLE = " & numPeriode & " AND POGPECLE = " & GroupeCle _
              & " AND POGARCLE > 90 "
    
    Case SommeParRegime
      ' requete pour faire les sommes
      rq = "SELECT Régime=CASE WHEN Assure.POGARCLE>90 THEN Assure.POGARCLE WHEN Assure.POGARCLE>50 AND Assure.POGARCLE<90 THEN Assure.POGARCLE-50 END, " _
           & " G.GALIB as [Libellé], P.Libelle AS Position, CASE WHEN Assure.POBaseRevalo=1 THEN CAST(1 as bit) ELSE CAST(0 as bit) END as [Base / Revalo], " _
           & " SUM(Assure.POPRESTATION) as [Prestation], " _
           & " SUM(Assure.POPRESTATIONTOTAL) as [Prestation Totale], " _
           & " SUM(Assure.POPRESTATION_AN) as [Annualisation], " _
           & " SUM(Assure.POPRESTATION_AN_PASSAGE) as [Annualisation Passage], " _
           & " SUM(Assure.POPM)+SUM(Assure.POPSAP) as [Provision Imputée], " _
           & " SUM(Round(ISNULL(Assure.POPM,0) + ISNULL(Assure.POPSAP,0) + ISNULL(Assure.POPM_REVALO,0),2)) AS [Provision Imputées avec Revalorisation], " _
           & " SUM(Assure.POPM) as [Provision  Calculées] " _
           & ", SUM(Assure.POPSAP) as [PSAP]" _
           & ", SUM(Assure.POPRESTA_RASSUR) as [Prestation Réassurée]" _
           & ", SUM(Assure.POPM_RASSUR) as [Provision Réassurée]" _
           & ", SUM(Assure.POPSAP_RASSUR) as [PSAP Réassurée]" _
           & ", SUM(Assure.POPM_RI) as [Provision Relative]" _
           & ", SUM(Assure.POPM_REVALO) as [Provision Revalorisation]" _
           & ", SUM(Assure.POPM+Assure.POPM_REVALO) AS [Provision avec Revalorisation] " _
           & ", SUM(Assure.POCOT_REVALO) as [Cotisation Revalorisation]"
      rq = rq & ", SUM(Assure.POPM_INCAP_1F) as [Incap 12€], SUM(Assure.POPM_PASS_1F) as [Pass 1€], SUM(Assure.POPM_INVAL_1F) as [Inval 1€]  " _
           & ", SUM(Assure.POPM_REDUC_1F) as [Rente 1€] " _
           & ", SUM(Assure.POPM_INCAP_1F*Assure.POPRESTATION_AN/12) as [PM Incap], SUM(Assure.POPM_PASS_1F*Assure.POPRESTATION_AN_PASSAGE) as [PM Pass], SUM(Assure.POPM_INVAL_1F*Assure.POPRESTATION_AN) as [PM Inval]  " _
           & ", SUM(Assure.POPM_REDUC_1F*Assure.POPRESTATION_AN) as [PM Rente], SUM(CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM,0),2) ELSE 0 END) AS [PM MGDC] "
      rq = rq & ", SUM(Assure.POPM_INCAP_1R) as [Incap 12R], SUM(Assure.POPM_PASS_1R) as [Pass 1R], SUM(Assure.POPM_INVAL_1R) as [Inval 1R]  " _
           & ", SUM(Assure.POPM_REDUC_1R) as [Rente 1R] " _
           & ", SUM(Assure.POPM_INCAP_1R*Assure.POPRESTATION_AN/12) as [PM Incap R], SUM(Assure.POPM_PASS_1R*Assure.POPRESTATION_AN_PASSAGE) as [PM Pass R], SUM(Assure.POPM_INVAL_1R*Assure.POPRESTATION_AN) as [PM Inval R]  " _
           & ", SUM(Assure.POPM_REDUC_1R*Assure.POPRESTATION_AN) as [PM Rente R], SUM(CASE WHEN Assure.POSIT=90 THEN Round(ISNULL(Assure.POPM+Assure.POPM_REVALO,0),2) ELSE 0 END) AS [PM MGDC R] "
      rq = rq & " FROM Assure INNER JOIN " _
           & "         Garantie AS G ON Assure.POGARCLE = G.GAGARCLE LEFT OUTER JOIN " _
           & "         CodePosition AS P ON Assure.POSIT = P.Position "
      sWhere = " WHERE Assure.POPERCLE = " & numPeriode & " AND Assure.POGPECLE = " & GroupeCle
    
'           & ", SUM(POPM_RCJT_1F) as [Rte Conjoint 1€], SUM(POPM_REDUC_1F) as [Rte Educ 1€] "
    Case Else
      Err.Raise -1, "SousProduit", "Mode inconnu !"
  End Select
          
  sprListe.MaxRows = 0
  
  ' pour datafill
  dtaPeriode.RecordSource = m_dataHelper.ValidateSQL(rq & sWhere & " AND RECNO=0" _
                              & IIf(fmMode = SommeParRegime, _
                              " GROUP BY POGARCLE, G.GALIB, P.Libelle, POBaseRevalo", _
                              ""))
  dtaPeriode.Refresh
  
  Set sprListe.DataSource = dtaPeriode
    
  ' datafill event pour formater les données
  If fmMode = MaintienDeces Or fmMode = Prestation Or fmMode = SousProduit Then
    SetColonneDataFill 4 ' NCA
  End If
  
  ' vrai requete
  If fmMode <> SommeParRegime Then
    rq = rq & fmFilter.GetSelectionSQLString & " ORDER BY POGARCLE, POCATEGORIE, POCONVENTION, PONOM"
  Else
    rq = rq & sWhere & fmFilter.GetSelectionSQLString
    
    If (eProvisionRetraite = fmTypePeriode Or eProvisionRetraiteRevalo = fmTypePeriode) Then
      rq = Replace(rq, "Position, CASE", "Position, '1 - Avant' as Type, CASE") & " GROUP BY Assure.POGARCLE, G.GALIB, P.Libelle, Assure.POBaseRevalo " _
           & " UNION ALL " _
           & Replace(Replace(Replace(rq, "Position, CASE", "Position, '2 - Après' as Type, CASE"), sWhere, ""), "Assure", "Assure_Retraite") & " INNER JOIN Assure ON Assure_Retraite.POIdAssure=Assure.RECNO " _
           & sWhere & fmFilter.GetSelectionSQLString _
           & " GROUP BY Assure_Retraite.POGARCLE, G.GALIB, P.Libelle, Assure_Retraite.POBaseRevalo" _
           & " ORDER BY Régime, Libellé, Position, Type, [Base / Revalo]"
    Else
      rq = rq & " GROUP BY Assure.POGARCLE, G.GALIB, P.Libelle, Assure.POBaseRevalo " _
              & " ORDER BY POGARCLE, GALIB, Libelle, POBaseRevalo"
    End If
  
  End If
  
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
    
    If fmMode = SousProduit Then
      sprListe.ColWidth(1) = 0
    End If
    
    sprListe.Visible = True

    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  ' cache la colonne RECNO
  If fmMode = SousProduit Then
    sprListe.ColWidth(1) = 0
  End If
  
  ' cache la colonne RECNO
  If fmMode = SousProduit Then
    sprListe.ColWidth(1) = 0
  End If
  
  ' total
  If fmMode = SommeParRegime Then
    Dim decal As Integer
    
    decal = 0
    If (eProvisionRetraite = fmTypePeriode Or eProvisionRetraiteRevalo = fmTypePeriode) Then
      decal = 1
    End If

    
    ' colorisation
    SetColBackColor 1, 3, jaune_clair
    
    SetColBackColor 4, 1, bleu_clair
    
    SetColBackColor 5 + decal, 14, vert_clair
    
    SetColBackColor 19 + decal, 4, lavande_clair
    
    SetColBackColor 23 + decal, 5, vert_clair
    
    SetColBackColor 28 + decal, 4, lavande_clair
    
    SetColBackColor 32 + decal, 5, vert_clair
    
    
    ' Sous totaux
    sprListe.MaxRows = sprListe.MaxRows + 2
    sprListe.RowHeight(sprListe.MaxRows - 1) = sprListe.RowHeight(sprListe.MaxRows) / 3
    
    sprListe.Row = sprListe.MaxRows
    
    sprListe.BlockMode = True
    
    ' ligne de séparation en gris
    sprListe.Row = sprListe.MaxRows - 1
    sprListe.Col = -1
    sprListe.Row2 = sprListe.MaxRows - 1
    sprListe.Col2 = -1
    sprListe.CellBorderColor = sprListe.GrayAreaBackColor
    sprListe.BackColor = sprListe.GrayAreaBackColor
    sprListe.CellBorderType = SS_BORDER_TYPE_OUTLINE
    sprListe.CellBorderStyle = SS_BORDER_STYLE_SOLID
    sprListe.Action = SS_ACTION_SET_CELL_BORDER
  
    ' tour de la liste des assurés en noir
    sprListe.Row = sprListe.MaxRows - 1
    sprListe.Col = 1
    sprListe.Row2 = sprListe.MaxRows - 2
    sprListe.Col2 = sprListe.MaxCols
    sprListe.CellBorderColor = vbBlack
    sprListe.CellBorderType = SS_BORDER_TYPE_OUTLINE
    sprListe.CellBorderStyle = SS_BORDER_STYLE_SOLID
    sprListe.Action = SS_ACTION_SET_CELL_BORDER
    
    ' ligne de totalisation entouré de rouge sur fond jaune
    sprListe.Row = sprListe.MaxRows
    sprListe.Col = 1
    sprListe.Row2 = sprListe.MaxRows
    sprListe.Col2 = sprListe.MaxCols
    sprListe.CellBorderColor = vbRed
    sprListe.ForeColor = vbRed
    sprListe.BackColor = vbYellow
    sprListe.CellBorderType = SS_BORDER_TYPE_OUTLINE
    sprListe.CellBorderStyle = SS_BORDER_STYLE_SOLID
    sprListe.Action = SS_ACTION_SET_CELL_BORDER
    
    sprListe.BlockMode = False
    
    sprListe.Col = 1
    sprListe.text = ""
    
    sprListe.Col = 2
    sprListe.text = "TOTAL"
    sprListe.TypeHAlign = TypeHAlignCenter
    
    For i = 1 To sprListe.MaxCols
'      sprListe.Row = sprListe.MaxRows - 1
'      sprListe.Col = i
'      sprListe.CellType = CellTypeStaticText
'      sprListe.text = GetColonneRef(i)
      
      sprListe.Row = sprListe.MaxRows - 2
      sprListe.Col = i
      
      If sprListe.BackColor = vert_clair Then
        sprListe.Row = sprListe.MaxRows
        sprListe.Col = i
        'sprListe.Formula = "SUM(" & Chr(Asc("A") + i - 1) & "1:" & Chr(Asc("A") + i - 1) & sprListe.MaxRows - 1 & ")"
        sprListe.formula = "SUM(" & GetColonneRef(i) & "1:" & GetColonneRef(i) & sprListe.MaxRows - 1 & ")"
      ElseIf i <> 2 Then
        sprListe.Row = sprListe.MaxRows
        sprListe.Col = i
        sprListe.BackColor = sprListe.GrayAreaBackColor
      End If
    Next i
  End If
  
  
  ' tour de la liste des assurés en noir
  sprListe.BlockMode = True
  
  sprListe.Row = 1
  sprListe.Col = IIf(fmMode = SousProduit, 2, 1)
  sprListe.Row2 = sprListe.MaxRows
  sprListe.Col2 = sprListe.MaxCols
  sprListe.CellBorderColor = vbBlack
  sprListe.CellBorderType = SS_BORDER_TYPE_OUTLINE
  sprListe.CellBorderStyle = SS_BORDER_STYLE_SOLID
  sprListe.Action = SS_ACTION_SET_CELL_BORDER
    
  sprListe.BlockMode = False
    
  For i = IIf(fmMode = SousProduit, 2, 1) To sprListe.MaxCols
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i)
  Next i
  
  ' affiche le spread (vitesse)
  sprListe.ReDraw = True
  sprListe.Visible = True
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67E70365
Private Sub sprListe_DataFill(ByVal Col As Long, ByVal Row As Long, ByVal DataType As Integer, ByVal fGetData As Integer, Cancel As Integer)
  ' NCA
  If Col = 4 Then
    Dim NCA As Variant
    
    sprListe.GetDataFillData NCA, vbString
    NCA = String(13 - Len(NCA), "0") & NCA
    
    sprListe.SetText Col, Row, Format(NCA, m_FormatNCA)
    
    Cancel = True
  End If
End Sub

