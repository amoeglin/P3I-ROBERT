VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtre"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "Filter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Caption         =   "Filtre par NUENRP3I"
      Height          =   735
      Left            =   90
      TabIndex        =   29
      Top             =   4680
      Width           =   4380
      Begin VB.TextBox txtNUENRP3I 
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Filtre par Code Position"
      Height          =   735
      Left            =   4545
      TabIndex        =   27
      Top             =   1620
      Width           =   4380
      Begin VB.ComboBox cboPosition 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Filtre par CCN"
      Height          =   735
      Left            =   4545
      TabIndex        =   21
      Top             =   3915
      Width           =   4380
      Begin VB.ComboBox cboCCN 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Filtre par Regroupement Statistique"
      Height          =   735
      Left            =   4545
      TabIndex        =   20
      Top             =   3150
      Width           =   4380
      Begin VB.ComboBox cboCodeNature 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Filtre par Regroupement Annexe"
      Height          =   735
      Left            =   4545
      TabIndex        =   19
      Top             =   2385
      Width           =   4380
      Begin VB.ComboBox cboRegroupement 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Filtre par Code Provision"
      Height          =   735
      Left            =   4545
      TabIndex        =   18
      Top             =   855
      Width           =   4380
      Begin VB.ComboBox cboNewCategorie 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Filtre par Code GE"
      Height          =   735
      Left            =   4545
      TabIndex        =   17
      Top             =   90
      Width           =   4380
      Begin VB.ComboBox cboNewRegime 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Filtre par Dossier"
      Height          =   735
      Left            =   90
      TabIndex        =   15
      Top             =   3915
      Width           =   4380
      Begin VB.TextBox txtNumSS 
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Filtre par Contrat"
      Height          =   735
      Left            =   90
      TabIndex        =   12
      Top             =   3150
      Width           =   4380
      Begin VB.CheckBox chkNCA 
         Caption         =   "Check1"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   195
      End
      Begin VB.ComboBox cboNCA 
         Height          =   315
         Left            =   405
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   270
         Width           =   3840
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Filtre par Nom"
      Height          =   735
      Left            =   90
      TabIndex        =   11
      Top             =   2385
      Width           =   4380
      Begin VB.TextBox txtNom 
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filtre par Catégorie"
      Height          =   735
      Left            =   90
      TabIndex        =   8
      Top             =   1620
      Width           =   4380
      Begin VB.ComboBox cboCategorie 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8970
      TabIndex        =   4
      Top             =   5460
      Width           =   8970
      Begin VB.CommandButton btnNoFilter 
         Caption         =   "&Aucun Filtre"
         Height          =   345
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   1575
      End
      Begin VB.CommandButton btnFilter 
         Caption         =   "&Valider"
         Default         =   -1  'True
         Height          =   345
         Left            =   3458
         TabIndex        =   6
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Annuler"
         Height          =   345
         Left            =   4538
         TabIndex        =   5
         Top             =   45
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtre par Régime"
      Height          =   735
      Left            =   90
      TabIndex        =   2
      Top             =   855
      Width           =   4380
      Begin VB.ComboBox cboRegime 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtre par Société"
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4380
      Begin VB.ComboBox cboSociete 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   4110
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E50104"
Option Explicit

'##ModelId=5C8A67E501E1
Public fmFilter As clsFilter

'##ModelId=5C8A67E501E2
Private Sub btnFilter_Click()
  ' efface les anciennes valeurs de filtre
  fmFilter.ClearValue
  
  ' sauve les nouvelles valeurs de filtre
  If cboSociete.ListIndex > 0 Then
    fmFilter.SetFilterElemValue "Société", cboSociete.ItemData(cboSociete.ListIndex)
  End If
  
  If cboRegime.ListIndex > 0 Then
    fmFilter.SetFilterElemValue "Régime", cboRegime.ItemData(cboRegime.ListIndex)
  End If
  
  If cboCategorie.ListIndex > 0 Then
    fmFilter.SetFilterElemValue "Catégorie", Replace(cboCategorie.List(cboCategorie.ListIndex), "Catégorie ", "")
  End If
  
  txtNom = Trim(txtNom)
  If txtNom <> "" Then
    fmFilter.SetFilterElemValue "Nom", txtNom
  End If
  
  If chkNCA.Value = vbChecked Then
    fmFilter.SetFilterElemValue "Convention", cboNCA.List(cboNCA.ListIndex)
  End If
  
  txtNumSS = Trim(txtNumSS)
  If txtNumSS <> "" Then
    fmFilter.SetFilterElemValue "Police", txtNumSS
  End If
  
  txtNUENRP3I = Trim(txtNUENRP3I)
  If txtNUENRP3I <> "" Then
    fmFilter.SetFilterElemValue "NUENRP3I", txtNUENRP3I
  End If
  
  If cboNewRegime.ListIndex > 0 Then
    If cboNewRegime.List(cboNewRegime.ListIndex) = "(aucun)" Then
      fmFilter.SetFilterElemValue "Code GE", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "Code GE", cboNewRegime.List(cboNewRegime.ListIndex)
    End If
  End If
  
  If cboNewCategorie.ListIndex > 0 Then
    If cboNewCategorie.List(cboNewCategorie.ListIndex) = "(aucun)" Then
      fmFilter.SetFilterElemValue "Code Provision", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "Code Provision", cboNewCategorie.ItemData(cboNewCategorie.ListIndex)
    End If
  End If
  
  If cboPosition.ListIndex > 0 Then
    If cboPosition.List(cboPosition.ListIndex) = "(aucun)" Then
      fmFilter.SetFilterElemValue "Code Position", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "Code Position", cboPosition.ItemData(cboPosition.ListIndex)
    End If
  End If
  
  If cboRegroupement.ListIndex > 0 Then
    If cboRegroupement.List(cboRegroupement.ListIndex) = "(aucun)" Then
      fmFilter.SetFilterElemValue "RegroupAnnexe", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "RegroupAnnexe", cboRegroupement.List(cboRegroupement.ListIndex)
    End If
  End If
  
  If cboCodeNature.ListIndex > 0 Then
    If cboCodeNature.List(cboCodeNature.ListIndex) = "(aucun)" Then
      fmFilter.SetFilterElemValue "RegroupStat", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "RegroupStat", cboCodeNature.List(cboCodeNature.ListIndex)
    End If
  End If
  
  If cboCCN.ListIndex > 0 Then
    If cboCCN.ItemData(cboCCN.ListIndex) = -1 Then
      fmFilter.SetFilterElemValue "CCN", FILTER_VALUE_NULL
    Else
      fmFilter.SetFilterElemValue "CCN", cboCCN.List(cboCCN.ListIndex)
    End If
  End If

  ret_code = 1
  
  Unload Me
End Sub

'##ModelId=5C8A67E501FE
Private Sub btnNoFilter_Click()
  cboSociete.ListIndex = 0
  cboRegime.ListIndex = 0
  cboCategorie.ListIndex = 0
  cboNewRegime.ListIndex = 0
  cboNewCategorie.ListIndex = 0
  cboPosition.ListIndex = 0
  cboRegroupement.ListIndex = 0
  cboCodeNature.ListIndex = 0
  cboCCN.ListIndex = 0
  txtNom = ""
  chkNCA.Value = vbUnchecked
  If cboNCA.ListCount > 0 Then cboNCA.ListIndex = 0
  txtNumSS = ""
  txtNUENRP3I = ""
End Sub


'##ModelId=5C8A67E5020D
Private Sub chkNCA_Click()
  If chkNCA.Value = vbChecked Then
    cboNCA.Enabled = True
  Else
    cboNCA.Enabled = False
  End If
End Sub

'##ModelId=5C8A67E5021D
Private Sub cmdClose_Click()
  ret_code = 0
  Unload Me
End Sub

'##ModelId=5C8A67E5022D
Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Dim i As Integer
  
  On Error GoTo err_Load
  
  ' rempli le combo societe
  cboSociete.Clear
  cboSociete.AddItem "* Toutes les sociétés du groupe"
  cboSociete.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT SOCLE, SONOM FROM Societe WHERE SOGROUPE = " & GroupeCle & " ORDER BY SONOM", Snapshot)
  Do Until rs.EOF
    cboSociete.AddItem "Société '" & rs.fields(1) & "'"
    cboSociete.ItemData(cboSociete.ListCount - 1) = rs.fields(0)
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo Regime
  cboRegime.Clear
  cboRegime.AddItem "*Tous les Régimes"
  cboRegime.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT GAGARCLE, GALIB FROM Garantie WHERE GAGARCLE>50 ORDER BY GAGARCLE ", Snapshot)
  Do Until rs.EOF
    cboRegime.AddItem "Régime " & IIf(rs.fields(0) > 90, rs.fields(0), rs.fields(0) - 50) & " - " & rs.fields(1)
    cboRegime.ItemData(cboRegime.ListCount - 1) = rs.fields(0)
    
    rs.MoveNext
  Loop
  rs.Close

  ' rempli le combo Categorie
  cboCategorie.Clear
  cboCategorie.AddItem "*Toutes les Catégories"
  cboCategorie.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POCATEGORIE FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POCATEGORIE", Snapshot)
  Do Until rs.EOF
    cboCategorie.AddItem "Catégorie " & rs.fields(0)
    'cboCategorie.ItemData(cboCategorie.ListCount - 1) = rs.fields(0)
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo NCA
  cboNCA.Clear
  'cboNCA.AddItem "*Toutes les Catégories"
  'cboNCA.ListIndex = 0
  'Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT Format(POCONVENTION, ""0 00 000000 00 00"") FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & " ORDER BY Format(POCONVENTION, ""0 00 000000 00 00"")", Snapshot)
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POCONVENTION FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POCONVENTION", Snapshot)
  Do Until rs.EOF
    'cboNCA.AddItem rs.Fields(0)
    cboNCA.AddItem Format(rs.fields(0), m_FormatNCA)
'    cboNCA.ItemData(cboNCA.ListCount - 1) = rs.Fields(0)
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo Codes GE
  cboNewRegime.Clear
  cboNewRegime.AddItem "*Tous les Codes GE"
  cboNewRegime.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POGARCLE_NEW FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POGARCLE_NEW", Snapshot)
  Do Until rs.EOF
    If Not IsNull(rs.fields(0)) Then
      cboNewRegime.AddItem rs.fields(0)
      'cboNewRegime.ItemData(cboNewRegime.ListCount - 1) = rs.fields(0)
    Else
      cboNewRegime.AddItem "(aucun)"
      'cboNewRegime.ItemData(cboNewRegime.ListCount - 1) = -1
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo Code Provision
  cboNewCategorie.Clear
  cboNewCategorie.AddItem "*Tous les Codes Provision"
  cboNewCategorie.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT CP.CodeProv, CP.Libelle FROM Assure A INNER JOIN CodeProvision CP ON CP.CodeProv=A.POCATEGORIE_NEW WHERE A.POGPECLE=" & GroupeCle & " AND A.POPERCLE=" & numPeriode & " ORDER BY CP.CodeProv", Snapshot)
  Do Until rs.EOF
    If Not IsNull(rs.fields(0)) Then
      cboNewCategorie.AddItem rs.fields(1)
      cboNewCategorie.ItemData(cboNewCategorie.ListCount - 1) = rs.fields(0)
    Else
      cboNewCategorie.AddItem "(aucun)"
      cboNewCategorie.ItemData(cboNewCategorie.ListCount - 1) = -1
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo Code Position
  cboPosition.Clear
  cboPosition.AddItem "*Tous les Codes Position"
  cboPosition.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT CP.Position, CP.Libelle FROM Assure A INNER JOIN CodePosition CP ON CP.Position=A.POSIT WHERE A.POGPECLE=" & GroupeCle & " AND A.POPERCLE=" & numPeriode & " ORDER BY CP.Position", Snapshot)
  Do Until rs.EOF
    If Not IsNull(rs.fields(0)) Then
      cboPosition.AddItem rs.fields(1)
      cboPosition.ItemData(cboPosition.ListCount - 1) = rs.fields(0)
    Else
      cboPosition.AddItem "(aucun)"
      cboPosition.ItemData(cboPosition.ListCount - 1) = -1
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo Regroupement
  cboRegroupement.Clear
  cboRegroupement.AddItem "*Tous les Regroupements Annexes"
  cboRegroupement.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POREGROUPEMENT FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POREGROUPEMENT", Snapshot)
  Do Until rs.EOF
    If IsNull(rs.fields(0)) Then
      cboRegroupement.AddItem "(aucun)"
    Else
      cboRegroupement.AddItem rs.fields(0)
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo CodeNature
  cboCodeNature.Clear
  cboCodeNature.AddItem "*Tous les Regroupements Statistiques"
  cboCodeNature.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POCODENATURE FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POCODENATURE", Snapshot)
  Do Until rs.EOF
    If IsNull(rs.fields(0)) Then
      cboCodeNature.AddItem "(aucun)"
    Else
      cboCodeNature.AddItem rs.fields(0)
    End If
    rs.MoveNext
  Loop
  rs.Close
  
  ' rempli le combo CCN
  cboCCN.Clear
  cboCCN.AddItem "*Toutes les CCN"
  cboCCN.ListIndex = 0
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT POCCN FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & " ORDER BY POCCN", Snapshot)
  Do Until rs.EOF
    If IsNull(rs.fields(0)) Then
      cboCCN.AddItem "(aucune)"
      cboCCN.ItemData(cboCCN.ListCount - 1) = -1
    Else
      cboCCN.AddItem rs.fields(0)
      cboCCN.ItemData(cboCCN.ListCount - 1) = rs.fields(0)
    End If
    
    rs.MoveNext
  Loop
  rs.Close
  
  ' recherche le filtre en cours
  For i = 0 To cboSociete.ListCount - 1
    If fmFilter.GetFilterElemValue("Société") = cboSociete.ItemData(i) Then
      cboSociete.ListIndex = i
      Exit For
    End If
  Next i
  
  For i = 0 To cboRegime.ListCount - 1
    If fmFilter.GetFilterElemValue("Régime") = cboRegime.ItemData(i) Then
      cboRegime.ListIndex = i
      Exit For
    End If
  Next i
  
  For i = 0 To cboCategorie.ListCount - 1
    If fmFilter.GetFilterElemValue("Catégorie") = Replace(cboCategorie.List(i), "Catégorie ", "") Then
      cboCategorie.ListIndex = i
      Exit For
    End If
  Next i
  
  For i = 0 To cboNewRegime.ListCount - 1
    If fmFilter.GetFilterElemValue("Code GE") = cboNewRegime.List(i) Then
      cboNewRegime.ListIndex = i
      Exit For
    End If
  Next i
  
  For i = 0 To cboNewCategorie.ListCount - 1
    If fmFilter.GetFilterElemValue("Code Provision") <> "" Then
      If fmFilter.GetFilterElemValue("Code Provision") = cboNewCategorie.ItemData(i) Then
        cboNewCategorie.ListIndex = i
        Exit For
      End If
    End If
  Next i
  
  For i = 0 To cboPosition.ListCount - 1
    If fmFilter.GetFilterElemValue("Code Position") <> "" Then
      If fmFilter.GetFilterElemValue("Code Position") = cboPosition.ItemData(i) Then
        cboPosition.ListIndex = i
        Exit For
      End If
    End If
  Next i
  
  If fmFilter.GetFilterElemValue("Contrat") <> "" Then
    chkNCA.Value = vbChecked
    For i = 0 To cboNCA.ListCount - 1
      If InStr(1, fmFilter.SelectionString, cboNCA.List(i)) Then
        cboNCA.ListIndex = i
        Exit For
      End If
    Next i
  Else
    cboNCA.Enabled = False
  End If
  
  If fmFilter.GetFilterElemValue("RegroupAnnexe") <> "" Then
    For i = 0 To cboRegroupement.ListCount - 1
      If fmFilter.GetFilterElemValue("Regroup") = cboRegroupement.List(i) Then
        cboRegroupement.ListIndex = i
        Exit For
      End If
    Next i
  End If
  
  If fmFilter.GetFilterElemValue("RegroupStat") <> "" Then
    For i = 0 To cboCodeNature.ListCount - 1
      If fmFilter.GetFilterElemValue("RegroupStat") = cboCodeNature.List(i) Then
        cboCodeNature.ListIndex = i
        Exit For
      End If
    Next i
  End If
  
  For i = 0 To cboCCN.ListCount - 1
    If fmFilter.GetFilterElemValue("CCN") = FILTER_VALUE_NULL Then
      cboCCN.ListIndex = 1
    Else
      If fmFilter.GetFilterElemValue("CCN") = cboCCN.ItemData(i) Then
        cboCCN.ListIndex = i
        Exit For
      End If
    End If
  Next i
  
  txtNom = fmFilter.GetFilterElemValue("Nom")
  txtNumSS = fmFilter.GetFilterElemValue("Police")
  txtNUENRP3I = fmFilter.GetFilterElemValue("NUENRP3I")
  
  Exit Sub
  
err_Load:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub

