VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixEdition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix des Editions"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmChoixEdition.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Provisions"
      Height          =   735
      Left            =   90
      TabIndex        =   44
      Top             =   5760
      Width           =   4575
      Begin VB.CheckBox chkListeProvision 
         Caption         =   "Liste par risques"
         Height          =   360
         Left            =   270
         TabIndex        =   23
         Top             =   255
         Width           =   1605
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Réassurance"
      Height          =   735
      Left            =   3420
      TabIndex        =   42
      Top             =   4980
      Width           =   1245
      Begin VB.CheckBox Check2 
         Caption         =   "Récap"
         Height          =   360
         Left            =   2280
         TabIndex        =   43
         Top             =   255
         Width           =   885
      End
      Begin VB.CheckBox chkListReassurance 
         Caption         =   "Liste"
         Height          =   360
         Left            =   270
         TabIndex        =   22
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Provisions pour sinistres avec revalorisation"
      Height          =   735
      Left            =   60
      TabIndex        =   41
      Top             =   4980
      Width           =   3405
      Begin VB.CheckBox chkListeRevalo 
         Caption         =   "Liste"
         Height          =   360
         Left            =   1170
         TabIndex        =   20
         Top             =   255
         Width           =   750
      End
      Begin VB.CheckBox chkRecapProvRevalo 
         Caption         =   "Récap"
         Height          =   360
         Left            =   2280
         TabIndex        =   21
         Top             =   255
         Width           =   885
      End
      Begin VB.CheckBox chkDetailProvRevalo 
         Caption         =   "Détail"
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   255
         Width           =   750
      End
   End
   Begin VB.OptionButton rdoScreen 
      Caption         =   "&à l'Ecran"
      Height          =   195
      Left            =   2700
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.OptionButton rdoPrinter 
      Caption         =   "&vers l'imprimante"
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   6600
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   45
      TabIndex        =   40
      Top             =   6870
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Etats règlementaires C7"
      Enabled         =   0   'False
      Height          =   1680
      Left            =   60
      TabIndex        =   35
      Top             =   3240
      Width           =   4605
      Begin VB.CommandButton btnTousC7 
         Caption         =   "Tous"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   270
         Width           =   690
      End
      Begin VB.CheckBox chkC7A 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   16
         Top             =   675
         Width           =   240
      End
      Begin VB.CheckBox chkC7B 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   17
         Top             =   990
         Width           =   240
      End
      Begin VB.CheckBox chkC7C 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   18
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label16 
         Caption         =   "Tableau A"
         Height          =   285
         Left            =   1350
         TabIndex        =   38
         Top             =   675
         Width           =   2580
      End
      Begin VB.Label Label15 
         Caption         =   "Tableau B"
         Height          =   285
         Left            =   1350
         TabIndex        =   37
         Top             =   990
         Width           =   2625
      End
      Begin VB.Label Label14 
         Caption         =   "Tableau C"
         Height          =   285
         Left            =   1350
         TabIndex        =   36
         Top             =   1305
         Width           =   2580
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Récapitulation par année de survenance"
      Enabled         =   0   'False
      Height          =   2505
      Left            =   60
      TabIndex        =   28
      Top             =   675
      Width           =   4605
      Begin VB.CommandButton btnRecap 
         Caption         =   "Récap"
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Top             =   270
         Width           =   645
      End
      Begin VB.CommandButton btnDetail 
         Caption         =   "Détail"
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   600
      End
      Begin VB.CheckBox chkRecapEvolution 
         Caption         =   "Check1"
         Height          =   240
         Left            =   900
         TabIndex        =   14
         Top             =   2145
         Width           =   240
      End
      Begin VB.CheckBox chkDetailEvolution 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   13
         Top             =   2145
         Width           =   240
      End
      Begin VB.CheckBox chkRecapSortie 
         Caption         =   "Check1"
         Height          =   240
         Left            =   900
         TabIndex        =   12
         Top             =   1830
         Width           =   240
      End
      Begin VB.CheckBox chkDetailSortie 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   11
         Top             =   1830
         Width           =   240
      End
      Begin VB.CheckBox chkRecapEntree 
         Caption         =   "Check1"
         Height          =   240
         Left            =   900
         TabIndex        =   10
         Top             =   1515
         Width           =   240
      End
      Begin VB.CheckBox chkDetailEntree 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   9
         Top             =   1515
         Width           =   240
      End
      Begin VB.CheckBox chkRecapProvision 
         Caption         =   "Check1"
         Height          =   240
         Left            =   900
         TabIndex        =   8
         Top             =   1215
         Width           =   240
      End
      Begin VB.CheckBox chkDetailProvision 
         Caption         =   "Check1"
         Height          =   240
         Left            =   315
         TabIndex        =   7
         Top             =   1215
         Width           =   240
      End
      Begin VB.CheckBox chkRecapMontantsAnnualises 
         Caption         =   "Check1"
         Height          =   240
         Left            =   900
         TabIndex        =   6
         Top             =   945
         Width           =   240
      End
      Begin VB.CheckBox chkDetailMontantsAnnualises 
         Height          =   240
         Left            =   315
         TabIndex        =   5
         Top             =   945
         Width           =   240
      End
      Begin VB.CheckBox chkRecapMontantsPayes 
         Height          =   240
         Left            =   900
         TabIndex        =   4
         Top             =   630
         Width           =   240
      End
      Begin VB.CheckBox chkDetailMontantsPayes 
         Height          =   240
         Left            =   315
         TabIndex        =   3
         Top             =   630
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Evolutions des provisions pour sinistres"
         Height          =   285
         Left            =   1350
         TabIndex        =   34
         Top             =   2145
         Width           =   2805
      End
      Begin VB.Label Label8 
         Caption         =   "Sorties des provisions pour sinistres"
         Height          =   285
         Left            =   1350
         TabIndex        =   33
         Top             =   1830
         Width           =   2580
      End
      Begin VB.Label Label7 
         Caption         =   "Entrées des provisions pour sinistres"
         Height          =   285
         Left            =   1350
         TabIndex        =   32
         Top             =   1515
         Width           =   2625
      End
      Begin VB.Label Label6 
         Caption         =   "Provisions pour sinistres"
         Height          =   285
         Left            =   1350
         TabIndex        =   31
         Top             =   1215
         Width           =   2580
      End
      Begin VB.Label Label4 
         Caption         =   "Montants annualisés"
         Height          =   285
         Left            =   1350
         TabIndex        =   30
         Top             =   945
         Width           =   2625
      End
      Begin VB.Label Label1 
         Caption         =   "Montants payés"
         Height          =   285
         Left            =   1350
         TabIndex        =   29
         Top             =   630
         Width           =   2580
      End
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -45
      TabIndex        =   0
      Top             =   7140
      Width           =   4755
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton btnLancer 
      Caption         =   "&Lancer"
      Height          =   375
      Left            =   135
      TabIndex        =   26
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label txtDescription 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Période 1 (12/88/1999 au 12/34/1956) du groupe 'Générali'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      TabIndex        =   39
      Top             =   90
      Width           =   4590
   End
End
Attribute VB_Name = "frmChoixEdition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67AD01A6"
Option Explicit

'##ModelId=5C8A67AD0290
Private etatDetail As Boolean
'##ModelId=5C8A67AD02AF
Private etatRecap As Boolean
' variable globale pour les champs des etats
'
'##ModelId=5C8A67AD02BF
Private Soc1 As String
'##ModelId=5C8A67AD02DE
Private Soc2 As String
'##ModelId=5C8A67AD02EE
Private Soc3 As String
'##ModelId=5C8A67AD031D
Private Gar11 As String
'##ModelId=5C8A67AD033C
Private Gar12 As String
'##ModelId=5C8A67AD035B
Private Gar13 As String
'##ModelId=5C8A67AD037A
Private Gar14 As String
'##ModelId=5C8A67AD039A
Private Gar15 As String
'##ModelId=5C8A67AD03B9
Private Gar21 As String
'##ModelId=5C8A67AD03D8
Private Gar22 As String
'##ModelId=5C8A67AE000F
Private Gar23 As String
'##ModelId=5C8A67AE002F
Private Gar31 As String
'##ModelId=5C8A67AE003E
Private Gar32 As String
'##ModelId=5C8A67AE005D
Private Gar33 As String

'##ModelId=5C8A67AE007D
Private Gar51 As String
'##ModelId=5C8A67AE009C
Private Gar53 As String
'##ModelId=5C8A67AE00BB
Private Gar56 As String
'##ModelId=5C8A67AE00DA
Private Gar57 As String
'##ModelId=5C8A67AE00EA
Private Gar59 As String

' description de la periode avec les / doublés
'##ModelId=5C8A67AE0109
Private DescriptionPourEtat As String

' filtre
'##ModelId=5C8A67AE0129
Public frmNumPeriode As Long
'##ModelId=5C8A67AE013A
Public fmFilter As clsFilter
'

'##ModelId=5C8A67AE0148
Private Sub ClearFomulas()
  Dim i As Integer
  
  'CrystalReport1.Reset
  
  For i = 0 To 50
    'CrystalReport1.Formulas(i) = ""
  Next i
End Sub

'
'##ModelId=5C8A67AE0158
Private Sub btnDetail_Click()
  If etatDetail = True Then
    chkDetailEvolution.Value = 1
    chkDetailSortie.Value = 1
    chkDetailEntree.Value = 1
    chkDetailProvision.Value = 1
    chkDetailMontantsAnnualises.Value = 1
    chkDetailMontantsPayes.Value = 1
  Else
    chkDetailEvolution.Value = 0
    chkDetailSortie.Value = 0
    chkDetailEntree.Value = 0
    chkDetailProvision.Value = 0
    chkDetailMontantsAnnualises.Value = 0
    chkDetailMontantsPayes.Value = 0
  End If
  
  etatDetail = Not etatDetail
End Sub

'##ModelId=5C8A67AE0167
Private Sub BuildTableEdition(NomChamps As String, Requete As String)
  ' fabrique les 3 tables intermediaires
  ' Editions   : details
  ' editionII  : somme par garantie
  ' editionIII : somme par société
  
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset, rs3 As ADODB.Recordset, rs4 As ADODB.Recordset
  Dim Montant As Double, szSQL As String
  Dim numPolice As String
  Dim numPol As Double
  Dim n As Integer, currPos As Long, stepPct As Long
  Dim bUseDisconnected As Boolean
    
  On Error GoTo err_build
  
  ' lancement des calculs
  ProgressBar1.Min = 0
  ProgressBar1.Value = 0
  ProgressBar1.Max = 100
  
  ProgressBar1.Value = 0
  
  ' vide les tables intermédiaires
  szSQL = " WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & frmNumPeriode
  
  m_dataSource.Execute "DELETE FROM EditionsTemp" & szSQL
  m_dataSource.Execute "DELETE FROM EditionIITemp" & szSQL
  m_dataSource.Execute "DELETE FROM EditionIIITemp" & szSQL
    
  m_dataSource.Execute "DELETE FROM Editions" & szSQL
  m_dataSource.Execute "DELETE FROM EditionII" & szSQL
  m_dataSource.Execute "DELETE FROM EditionIII" & szSQL

    
  ' parcours la table Assuré
  Requete = Replace(Requete, "*", "PONUMCLE, POGARCLE, POARRET, POSIT, POPRESTATION, POPRESTATION_AN, POPM, POPM_VAR")
  Set rs = m_dataSource.OpenRecordset(Requete, Snapshot)
  
  'If MsgBox("BuildTableEdition() : RecordSets Déconnectés ?", vbYesNo) = vbYes Then
     bUseDisconnected = False ' true trop lent a la reconnexion
  'Else
    ' ça n'aporte rien au niveau performance
'     bUseDisconnected = False
  'End If
  
  If bUseDisconnected = False Then
    Set rs2 = m_dataSource.OpenRecordset("EditionsTemp", table)
    Set rs3 = m_dataSource.OpenRecordset("EditionIITemp", table)
    Set rs4 = m_dataSource.OpenRecordset("EditionIIITemp", table)
  Else
    Set rs2 = m_dataSource.OpenRecordset("select * from EditionsTemp", Disconnected)
    Set rs3 = m_dataSource.OpenRecordset("select * from EditionIITemp", Disconnected)
    Set rs4 = m_dataSource.OpenRecordset("select * from EditionIIITemp", Disconnected)
  End If
  
  ' obligatoire pour accéler la création des tables
  m_dataSource.BeginTrans
  
  If Not rs.EOF Then
    rs.MoveLast
    
    ProgressBar1.Max = rs.RecordCount
    stepPct = ProgressBar1.Max / 200
    
    If stepPct = 0 Then
      stepPct = 1
    End If
    
    rs.MoveFirst
  Else
    ProgressBar1.Max = 1
  End If
  
  Do Until rs.EOF
    If (currPos Mod stepPct) = 0 Then
      ProgressBar1.Value = currPos
    End If
    currPos = currPos + 1
    
    If IsNull(rs.fields(NomChamps)) Then
      Montant = 0
    Else
      Montant = rs.fields(NomChamps)
    End If
    
   ' nouveau record
    rs2.AddNew
    
    rs2.fields("CleGroupe") = GroupeCle
    rs2.fields("ClePeriode") = frmNumPeriode
    
    ' maj des champs Editions
    numPolice = rs.fields("PONUMCLE")
'*** UPESE
    If Not IsNumeric(numPolice) Then
      Do
        n = InStr(numPolice, ".")
        If n <> 0 Then
          numPolice = Left(numPolice, n - 1) & mID(numPolice, n + 1)
        End If
      Loop Until IsNumeric(numPolice)
    End If
'***
    numPol = CDbl(numPolice)
    While numPol > 2000000000
      numPol = numPol / 10
    Wend
    numPolice = numPol
      
    rs2.fields("NoPolice") = val(numPolice)
    
    rs2.fields("DateArret") = DateSerial(Year(rs.fields("POARRET")), 1, 1)
    
    Select Case rs.fields("POGARCLE")
      Case cdGar51
        rs2.fields("tExo1") = Montant
      
      Case cdGar53
        rs2.fields("tIncap1") = Montant
      
      Case cdGar56
        rs2.fields("tRente1") = Montant
      
      Case cdGar57
        rs2.fields("tExo2") = Montant
      
      Case cdGar59
        Select Case rs.fields("POSIT")
          Case cdPosit_IncapAvecPassage
            rs2.fields("tRente2") = Montant ' incap avec passage
      
          Case cdPosit_Inval
            rs2.fields("tExo3") = Montant ' inval
          
          Case cdPosit_IncapSansPassage
            rs2.fields("tIncap2") = Montant ' incap seul
        End Select
    End Select
    
    rs2.Update
    
    ' champs EditionII
    rs3.AddNew
    
    rs3.fields("CleGroupe") = GroupeCle
    rs3.fields("ClePeriode") = frmNumPeriode
    
    rs3.fields("NoPolice") = val(numPolice)
    
    rs3.fields("DateArret") = DateSerial(Year(rs.fields("POARRET")), 1, 1)
    
    Select Case rs.fields("POGARCLE")
      Case cdGar59
        Select Case rs.fields("POSIT")
          Case cdPosit_IncapAvecPassage
            rs3.fields("tIncap1") = Montant ' incap avec passage
      
          Case cdPosit_Inval
            rs3.fields("tRente1") = Montant ' inval
          
          Case cdPosit_IncapSansPassage
            rs3.fields("tExo1") = Montant ' incap seul
        End Select
      
      Case cdGar59
        rs3.fields("tExo1") = Montant ' incap avec passage
      
      Case cdGar51, cdGar53, cdGar56, cdGar57
        rs3.fields("tRente1") = Montant ' inval
    End Select
    
    rs3.Update
    
    ' champs EditionIII
    rs4.AddNew
    
    rs4.fields("CleGroupe") = GroupeCle
    rs4.fields("ClePeriode") = frmNumPeriode
    
    rs4.fields("NoPolice") = val(numPolice)
    
    rs4.fields("DateArret") = DateSerial(Year(rs.fields("POARRET")), 1, 1)
    Select Case rs.fields("POGARCLE")
      Case cdGar51, cdGar53, cdGar56, cdGar57, cdGar59
        rs4.fields("tSoc1") = Montant
      
      Case cdGar21, cdGar22, cdGar23
        rs4.fields("tSoc2") = Montant
      
      Case cdGar31, cdGar32, cdGar33
        rs4.fields("tSoc3") = Montant
    End Select
    
    ' validation
    rs4.Update
    
    ' next record
    rs.MoveNext
  Loop
    
  If bUseDisconnected = True Then
    ' sauvegarde les enregistrements dans la base de données
    Set rs2.ActiveConnection = m_dataSource.Connection
    Set rs3.ActiveConnection = m_dataSource.Connection
    Set rs4.ActiveConnection = m_dataSource.Connection
    
    rs2.UpdateBatch
    rs3.UpdateBatch
    rs4.UpdateBatch
  End If
  
  m_dataSource.CommitTrans
  
  rs.Close
  
  rs2.Close
  rs3.Close
  rs4.Close
  
  ' regroupe les paiments de la meme annee d'arret pour la meme police
'  theDB.Execute "SELECT NoPolice, DateArret, SUM(tExo1) as Exo1, SUM(tIncap1) as Incap1, " _
'                & " SUM(tRente1) as Rente1, SUM(tExo2) as Exo2, SUM(tIncap2) as Incap2, SUM(tRente2) as Rente2, " _
'                & " SUM(tExo3) as Exo3, SUM(tIncap3) as Incap3, SUM(tRente3) as Rente3 " _
'                & " INTO Editions FROM EditionsTemp GROUP BY NoPolice, DateArret"
  m_dataSource.Execute "INSERT INTO Editions(CleGroupe, ClePeriode, NoPolice, DateArret, Exo1, Incap1, " _
                & " Rente1, Exo2, Incap2, Rente2, Exo3, Incap3, Rente3) " _
                & "SELECT CleGroupe, ClePEriode, NoPolice, DateArret, SUM(tExo1) , SUM(tIncap1), " _
                & " SUM(tRente1), SUM(tExo2), SUM(tIncap2), SUM(tRente2), " _
                & " SUM(tExo3), SUM(tIncap3), SUM(tRente3) " _
                & " FROM EditionsTemp " & szSQL & " GROUP BY CleGroupe, ClePeriode, NoPolice, DateArret"

'  theDB.Execute "SELECT NoPolice, DateArret, SUM(tExo1) as Exo1, SUM(tIncap1) as Incap1, SUM(tRente1) as Rente1 INTO EditionII FROM EditionIITemp GROUP BY NoPolice, DateArret"
  m_dataSource.Execute "INSERT INTO EditionII(CleGroupe, ClePeriode, NoPolice, DateArret, Exo1, Incap1, Rente1) " _
                & " SELECT CleGroupe, ClePEriode, NoPolice, DateArret, SUM(tExo1), SUM(tIncap1), SUM(tRente1) " _
                & " FROM EditionIITemp " & szSQL & " GROUP BY CleGroupe, ClePeriode, NoPolice, DateArret"

'  theDB.Execute "SELECT NoPolice, DateArret, SUM(tSoc1) as Soc1, SUM(tSoc2) as Soc2, SUM(tSoc3) as Soc3 INTO EditionIII FROM EditionIIITemp GROUP BY NoPolice, DateArret"
  m_dataSource.Execute "INSERT INTO EditionIII(CleGroupe, ClePeriode, NoPolice, DateArret, Soc1, Soc2, Soc3) " _
                & " SELECT CleGroupe, ClePeriode, NoPolice, DateArret, SUM(tSoc1), SUM(tSoc2), SUM(tSoc3) " _
                & " FROM EditionIIITemp " & szSQL & " GROUP BY CleGroupe, ClePeriode, NoPolice, DateArret"
  
  ProgressBar1.Value = 0
  
  Exit Sub
  
err_build:
  DisplayError
  Resume Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-
''''''''''''''''''''''''''''
' lance le tableau A de l'etat C7
'
'##ModelId=5C8A67AE0196
Private Sub LanceEtatC7A(An As Integer, Condition As String, Description As String)
'  Dim numFormule As Integer
'
'  Dim dateCalcul As String
'  Dim fraisGestionIncap As Double, fraisGestionInval As Double, FraisGestionC7 As Double, RevenuFinancierC7 As Double
'
'  dateCalcul = Format(m_dataHelper.GetParameter("SELECT PEDATEEXT FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  fraisGestionIncap = m_dataHelper.GetParameter("SELECT PEGESTIONINCAP FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  fraisGestionInval = m_dataHelper.GetParameter("SELECT PEGESTIONINVAL FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  FraisGestionC7 = 1 + (0.01 * m_dataHelper.GetParameter("SELECT PEGESTIONC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'  RevenuFinancierC7 = (0.01 * m_dataHelper.GetParameter("SELECT PETXREVENUC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'
'  ' date de la période
'  Dim dd As String, df As String
'
'  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'
'  Call ClearFomulas
'
'  CrystalReport1.ReportFileName = App.Path & "\etatc7_a.rpt"
'
'  numFormule = 0
'
'  CrystalReport1.Formulas(numFormule) = "Description=""" & Description & """"
'  numFormule = numFormule + 1
'
'  ' totaux uniquement
'  ' Provision à l'ouvertue
'  CrystalReport1.Formulas(numFormule) = "TA_1=" & m_dataHelper.GetParameterAsStringCRW("SELECT sum(PROV_ANn)+sum(PROV_ANn1)+sum(PROV_ANn2)+sum(PROV_ANn3)+sum(PROV_ANn4)+sum(PROV_ANn5) AS Champs1 " _
'                                                              & " From ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & NumPeriode & Condition)
'  numFormule = numFormule + 1
'  ' capitaux entrés
'  CrystalReport1.Formulas(numFormule) = "TA_2=0"
'  'CrystalReport1.Formulas(numFormule) = "TA_2=" & m_dataHelper.GetParameterAsStringCRW("SELECT " & fraisGestionC7 & " * SUM(POPM) " _
'  '                                                            & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode _
'  '                                                            & " AND POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POPREMIER_PAIEMENT " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
'  ' autres ressources
'  CrystalReport1.Formulas(numFormule) = "TA_3=" & 0
'  numFormule = numFormule + 1
'  ' prestations payées
''  CrystalReport1.Formulas(numFormule) = "TA_5=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPRESTATION) " _
''                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
''                                                              & " AND POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) _
''                                                              & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  CrystalReport1.Formulas(numFormule) = "TA_5=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPRESTATION) " _
'                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
'                                                              & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
'  ' Capitaux sorties
'  CrystalReport1.Formulas(numFormule) = "TA_6=0"
'  'CrystalReport1.Formulas(numFormule) = "TA_6=" & m_dataHelper.GetParameterAsStringCRW("SELECT " & fraisGestionC7 & " * SUM(POPM) " _
'  '                                                            & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode _
'  '                                                            & " AND POTERME " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
'  ' provisions à la cloture
'  CrystalReport1.Formulas(numFormule) = "TA_7=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPM) " _
'                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
'                                                              & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                              & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
'  ' Charges de gestion
'  CrystalReport1.Formulas(numFormule) = "TA_8=" & m_dataHelper.GetParameterAsStringCRW("SELECT " & (FraisGestionC7 - 1) & " * SUM(POPRESTATION_AN) " _
'                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
'                                                              & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                              & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
'  ' produits financiers
'  CrystalReport1.Formulas(numFormule) = "TA_4=(({@TA_1}+{@TA_7})/2)*" & m_dataHelper.VirguleVersPoint(CStr(RevenuFinancierC7))
'  numFormule = numFormule + 1
'
'  ' param date
'  CrystalReport1.Formulas(numFormule) = "DateCalcul=""" & dateCalcul & """"
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "DatePeriode=" & DescriptionPourEtat & ""
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionIncap=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionIncap / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionInval=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionInval / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.ReportTitle = "Etat C7 - Tableau A"
'  CrystalReport1.WindowState = crptMaximized
'  CrystalReport1.WindowTitle = "Etat C7 - Tableau A"
'
'  If rdoPrinter.Value = True Then
'    CrystalReport1.Destination = crptToPrinter
'  Else
'    CrystalReport1.Destination = crptToWindow
'  End If
'
'  SetConnectionString
'
'  CrystalReport1.MarginLeft = 400
'
'  CrystalReport1.Action = 1
'
'  CrystalReport1.ReportFileName = ""
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-
''''''''''''''''''''''''''''
' rempli une colonne de l'année Annee dans le tableau B de l'etat C7
'
'##ModelId=5C8A67AE01E4
Private Sub FillColumnInEtatC7B(Suffixe As String, Condition As String, Annee As String, ByRef numFormule As Integer, fraisGestionIncap As Double, fraisGestionInval As Double, FraisGestionC7 As Double, RevenuFinancierC7 As Double, dd As String, df As String, ConditionNumSte As String)
'  ' Provision à l'ouvertue
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_1=" & m_dataHelper.GetParameterAsStringCRW("SELECT PROV_ANn" & Suffixe & "" _
'                                                                 & " From ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & NumPeriode & ConditionNumSte)
'  numFormule = numFormule + 1
'
'  ' capitaux entrés
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_2=0"
'  'CrystalReport1.Formulas(numFormule) = "AN_2=" & m_dataHelper.GetParameter("SELECT " & fraisGestionC7 & " * SUM(POPM) " _
'  '                                                            & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode _
'  '                                                            & " AND POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POPREMIER_PAIEMENT " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POSIT = " & cdTypeInvalidite _
'  '                                                            & " AND Year(POARRET)" & Condition & Annee) ' invalidité
'  numFormule = numFormule + 1
'
'  ' autres ressources
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_3=" & 0
'  numFormule = numFormule + 1
'
'  ' prestations payées
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_5=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPRESTATION) " _
'                                                                                      & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & ConditionNumSte _
'                                                                                      & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                                                      & " AND POSIT = " & cdTypeInvalidite _
'                                                                                      & " AND Year(POARRET)" & Condition & Annee) ' invalidité
'  numFormule = numFormule + 1
'
'  ' Capitaux sorties
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_6=0"
'  'CrystalReport1.Formulas(numFormule) = "TA_6=" & m_dataHelper.GetParameter("SELECT " & fraisGestionC7 & " * SUM(POPM) " _
'  '                                                            & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode _
'  '                                                            & " AND POTERME " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                                            & " AND POSIT = " & cdTypeInvalidite _
'  '                                                            & " AND Year(POARRET)" & Condition & Annee) ' invalidité
'  numFormule = numFormule + 1
'
'  ' provisions à la cloture
'  Dim provIncap As Double, provInval As Double
'
'  provIncap = 0
'  'provIncap = m_dataHelper.GetParameterAsDouble("SELECT SUM(POPM) " _
'  '                                  & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & ConditionNumSte _
'  '                                  & " AND POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) _
'  '                                  & " AND POSIT = " & cdTypeIncapacite _
'  '                                  & " AND Year(POARRET)" & Condition & Annee) ' incapacite
'  provInval = m_dataHelper.GetParameterAsDouble("SELECT SUM(POPM) " _
'                                    & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & ConditionNumSte _
'                                    & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                    & " AND POSIT = " & cdTypeInvalidite _
'                                    & " AND Year(POARRET)" & Condition & Annee) ' invalidité
'
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_7=" & m_dataHelper.VirguleVersPoint(provInval + provIncap)
'  numFormule = numFormule + 1
'
'  ' Charges de gestion
'  'CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_8=" & VirguleVersPoint((provInval * (1 - 1 / (1 + (0.01 * fraisGestionInval)))) _
'  '                                                               + (provIncap * (1 - 1 / (1 + (0.01 * fraisGestionIncap)))))
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_8=" & m_dataHelper.GetParameterAsStringCRW("SELECT " & (FraisGestionC7 - 1) & " * SUM(POPRESTATION_AN) " _
'                                                               & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode _
'                                                               & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                               & " AND POSIT = " & cdTypeInvalidite _
'                                                               & " AND Year(POARRET)" & Condition & Annee) ' invalidité
'  numFormule = numFormule + 1
'
'  ' produits financiers
'  CrystalReport1.Formulas(numFormule) = "AN" & Suffixe & "_4=(({@AN" & Suffixe & "_1}+{@AN" & Suffixe & "_7})/2)*" & m_dataHelper.VirguleVersPoint(CStr(RevenuFinancierC7))
'  numFormule = numFormule + 1
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-
''''''''''''''''''''''''''''
' lance le tableau B de l'etat C7
'
'##ModelId=5C8A67AE02BF
Private Sub LanceEtatC7B(An As Integer, Condition As String, Description As String)
'  Dim numFormule As Integer
'
'  Dim dateCalcul As String
'  Dim fraisGestionIncap As Double, fraisGestionInval As Double, FraisGestionC7 As Double, RevenuFinancierC7 As Double
'
'  dateCalcul = Format(m_dataHelper.GetParameter("SELECT PEDATEEXT FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  fraisGestionIncap = m_dataHelper.GetParameter("SELECT PEGESTIONINCAP FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  fraisGestionInval = m_dataHelper.GetParameter("SELECT PEGESTIONINVAL FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  FraisGestionC7 = 1 + (0.01 * m_dataHelper.GetParameter("SELECT PEGESTIONC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'  RevenuFinancierC7 = (0.01 * m_dataHelper.GetParameter("SELECT PETXREVENUC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'
'  ' date de la période
'  Dim dd As String, df As String
'
'  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'
'  Call ClearFomulas
'
'  CrystalReport1.ReportFileName = App.Path & "\etatc7_b.rpt"
'
'  numFormule = 0
'
'  CrystalReport1.Formulas(numFormule) = "Description=""" & Description & """"
'  numFormule = numFormule + 1
'
'  ' Année en entete de colonne
'  CrystalReport1.Formulas(numFormule) = "AnN=" & An
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN1=" & An - 1
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN2=" & An - 2
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN3=" & An - 3
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN4=" & An - 4
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN5=""" & An - 5 & " et ant."""
'  numFormule = numFormule + 1
'
'  '''''''''
'  '
'  ' Reprendre les formules de l'etat A avec une selection par année
'  ' Faire une fonction pour remplir les cellules en fct du nom et la condition en plus
'  '
'  '''''''''
'
'  ' Année N
'  Call FillColumnInEtatC7B("", "=", Year(df), numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  ' Année N-1
'  Call FillColumnInEtatC7B("1", "=", Year(df) - 1, numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  ' Année N-2
'  Call FillColumnInEtatC7B("2", "=", Year(df) - 2, numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  ' Année N-3
'  Call FillColumnInEtatC7B("3", "=", Year(df) - 3, numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  ' Année N-4
'  Call FillColumnInEtatC7B("4", "=", Year(df) - 4, numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  ' Année N-5 et antérieur
'  Call FillColumnInEtatC7B("5", "<=", Year(df) - 5, numFormule, fraisGestionIncap, _
'                           fraisGestionInval, FraisGestionC7, RevenuFinancierC7, _
'                           dd, df, Condition)
'
'  CrystalReport1.Formulas(numFormule) = "DateCalcul=""" & dateCalcul & """"
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "DatePeriode=" & DescriptionPourEtat & ""
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionIncap=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionIncap / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionInval=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionInval / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.ReportTitle = "Etat C7 - Tableau B"
'  CrystalReport1.WindowState = crptMaximized
'  CrystalReport1.WindowTitle = "Etat C7 - Tableau B"
'
'  If rdoPrinter.Value = True Then
'    CrystalReport1.Destination = crptToPrinter
'  Else
'    CrystalReport1.Destination = crptToWindow
'  End If
'
'  SetConnectionString
'
'  CrystalReport1.MarginLeft = 400
'
'  CrystalReport1.Action = 1
'
'  CrystalReport1.ReportFileName = ""
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''-
''''''''''''''''''''''''''''
' lance le tableau C de l'etat C7
'
'##ModelId=5C8A67AE030D
Private Sub SendFormulasToEtatC7B(CellName1 As String, CellName2 As String, test As String, Annee As String, ByRef numFormule As Integer, dd As String, df As String, Condition As String)
'  CrystalReport1.Formulas(numFormule) = CellName1 & "=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPRESTATION) " _
'                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
'                                                              & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                              & " AND Year(POARRET) " & test & Annee _
'                                                              & " AND POSIT = " & cdTypeIncapacite) ' incapacité
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = CellName2 & "=" & m_dataHelper.GetParameterAsStringCRW("SELECT SUM(POPRESTATION) " _
'                                                              & " FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & NumPeriode & Condition _
'                                                              & " AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) _
'                                                              & " AND Year(POARRET) " & test & Annee _
'                                                              & " AND POSIT = " & cdTypeInvalidite) ' invalidité
'  numFormule = numFormule + 1
End Sub

'##ModelId=5C8A67AE03B9
Private Sub LanceEtatC7C(An As Integer, Condition As String, Description As String)
'  Dim numFormule As Integer
'
'  Dim dateCalcul As String
'  Dim fraisGestionIncap As Double, fraisGestionInval As Double, FraisGestionC7 As Double, RevenuFinancierC7 As Double
'
'  dateCalcul = Format(m_dataHelper.GetParameter("SELECT PEDATEEXT FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  fraisGestionIncap = m_dataHelper.GetParameter("SELECT PEGESTIONINCAP FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  fraisGestionInval = m_dataHelper.GetParameter("SELECT PEGESTIONINVAL FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode)
'  FraisGestionC7 = 1 + (0.01 * m_dataHelper.GetParameter("SELECT PEGESTIONC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'  RevenuFinancierC7 = (0.01 * m_dataHelper.GetParameter("SELECT PETXREVENUC7 FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode))
'
'  ' date de la période
'  Dim dd As String, df As String
'
'  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & NumPeriode), "dd/mm/yyyy")
'
'  Call ClearFomulas
'
'  CrystalReport1.ReportFileName = App.Path & "\etatc7_c.rpt"
'
'  numFormule = 0
'
'  CrystalReport1.Formulas(numFormule) = "Description=""" & Description & """"
'  numFormule = numFormule + 1
'
'  ' Année pour les entetes de colonne
'  CrystalReport1.Formulas(numFormule) = "AnN=" & An
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN1=" & An - 1
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN2=" & An - 2
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN3=" & An - 3
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN4=" & An - 4
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN5=" & An - 5
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN6=" & An - 6
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN7=" & An - 7
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN8=" & An - 8
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN9=" & An - 9
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN10=" & An - 10
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN11=" & An - 11
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "AnN12=""" & An - 12 & " et ant."""
'  numFormule = numFormule + 1
'
'  ''''''''''''''''''''''
'  '
'  'Verifier l 'incrementation de numFormule
'  '
'  ''''''''''''''''''''''
'
'  ' année n-12
'  Call SendFormulasToEtatC7B("AN5_1", "AN5_2", "<=", Year(df) - 12, numFormule, dd, df, Condition)
'
'  ' année n-11
'  Call SendFormulasToEtatC7B("AN4_1", "AN4_2", "=", Year(df) - 11, numFormule, dd, df, Condition)
'
'  ' année n-10
'  Call SendFormulasToEtatC7B("AN3_1", "AN3_2", "=", Year(df) - 10, numFormule, dd, df, Condition)
'
'  ' année n-9
'  Call SendFormulasToEtatC7B("AN2_1", "AN2_2", "=", Year(df) - 9, numFormule, dd, df, Condition)
'
'  ' année n-8
'  Call SendFormulasToEtatC7B("AN1_1", "AN1_2", "=", Year(df) - 8, numFormule, dd, df, Condition)
'
'  ' année n-7
'  Call SendFormulasToEtatC7B("AN_1", "AN_2", "=", Year(df) - 7, numFormule, dd, df, Condition)
'
'  ' année n-6
'  Call SendFormulasToEtatC7B("AN6_1", "AN6_2", "=", Year(df) - 6, numFormule, dd, df, Condition)
'
'  ' année n-5
'  Call SendFormulasToEtatC7B("AN5_3", "AN5_4", "=", Year(df) - 5, numFormule, dd, df, Condition)
'
'  ' année n-4
'  Call SendFormulasToEtatC7B("AN4_3", "AN4_4", "=", Year(df) - 4, numFormule, dd, df, Condition)
'
'  ' année n-3
'  Call SendFormulasToEtatC7B("AN3_3", "AN3_4", "=", Year(df) - 3, numFormule, dd, df, Condition)
'
'  ' année n-2
'  Call SendFormulasToEtatC7B("AN2_3", "AN2_4", "=", Year(df) - 2, numFormule, dd, df, Condition)
'
'  ' année n-1
'  Call SendFormulasToEtatC7B("AN1_3", "AN1_4", "=", Year(df) - 1, numFormule, dd, df, Condition)
'
'  ' année n
'  Call SendFormulasToEtatC7B("AN_3", "AN_4", "=", Year(df), numFormule, dd, df, Condition)
'
'  ' sdfqsdf dfjh ds hh
'  CrystalReport1.Formulas(numFormule) = "DateCalcul=""" & dateCalcul & """"
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "DatePeriode=" & DescriptionPourEtat & ""
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionIncap=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionIncap / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionInval=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionInval / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.ReportTitle = "Etat C7 - Tableau A"
'  CrystalReport1.WindowState = crptMaximized
'  CrystalReport1.WindowTitle = "Etat C7 - Tableau A"
'
'  If rdoPrinter.Value = True Then
'    CrystalReport1.Destination = crptToPrinter
'  Else
'    CrystalReport1.Destination = crptToWindow
'  End If
'
'  SetConnectionString
'
'  CrystalReport1.MarginLeft = 400
'
'  CrystalReport1.Action = 1
'
'  CrystalReport1.ReportFileName = ""
End Sub

'##ModelId=5C8A67AF001F
Private Sub LanceEtat(titre As String, NomFichier As String, dateCalcul As String, fraisGestionIncap As Double, fraisGestionInval As Double)
'  Dim numFormule As Integer
'
'  On Error GoTo err_LanceEtat
'
'  Call ClearFomulas
'
'  numFormule = 0
'
'  CrystalReport1.ReportFileName = App.Path & "\" & NomFichier
'
'  CrystalReport1.Formulas(numFormule) = "Soc1=""" & Soc1 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Soc2=""" & Soc2 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Soc3=""" & Soc3 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar11=""" & Gar51 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar12=""" & Gar53 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar13=""" & Gar56 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar21=""" & Gar57 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar22=""" & Gar59 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar23=""" & Gar59 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar31=""" & Gar59 & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar32="""""
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Gar33="""""
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "Titre=""" & titre & """"
'  numFormule = numFormule + 1
'  CrystalReport1.Formulas(numFormule) = "DateCalcul=""" & dateCalcul & """"
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "DatePeriode=" & DescriptionPourEtat & ""
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionIncap=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionIncap / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.Formulas(numFormule) = "FraisGestionInval=" & m_dataHelper.VirguleVersPoint(Format(fraisGestionInval / 100, "###0.0####"))
'  numFormule = numFormule + 1
'
'  CrystalReport1.ReportTitle = titre
'  CrystalReport1.WindowState = crptMaximized
'  CrystalReport1.WindowTitle = titre
'
'  If rdoPrinter.Value = True Then
'    CrystalReport1.Destination = crptToPrinter
'  Else
'    CrystalReport1.Destination = crptToWindow
'  End If
'
'
'  SetConnectionString
'
'
'
'  ' record selection formula
'  Dim recordSelection As String
'
'  Select Case NomFichier
'    Case "type1.rpt", "type1bis.rpt", "type4.rpt", "type4bis.rpt", "détail.rpt"
'      recordSelection = "Editions"
'
'    Case "type2.rpt", "type2bis.rpt", "type5.rpt", "type5bis.rpt"
'      recordSelection = "EditionII"
'
'    Case "type3.rpt", "type3bis.rpt", "type6.rpt", "type6bis.rpt"
'      recordSelection = "EditionIII"
'
'    Case Else
'      Err.Raise -1, "LanceEtat()", "Rapport inconnu pour RecordSelectionFormula : " & NomFichier
'  End Select
'  CrystalReport1.ReplaceSelectionFormula "{" & recordSelection & ".CleGroupe}=" & GroupeCle & " AND {" & recordSelection & ".ClePeriode}=" & frmNumPeriode
'
'  CrystalReport1.MarginLeft = 400
'
'  CrystalReport1.Action = 1
'
'  CrystalReport1.ReportFileName = ""
'
'  Exit Sub
'
'err_LanceEtat:
'  DisplayError
'  Resume Next
End Sub

'##ModelId=5C8A67AF007D
Private Sub BuildAndLaunchReport(chkDetail As CheckBox, chkRecap As CheckBox, typeTotaux As Boolean, FieldName As String, Filtre As String, titre As String)
  If chkRecap.Value = 1 Or chkDetail.Value = 1 Then
    Dim dateCalcul As String
    Dim fraisGestionIncap As Double, fraisGestionInval As Double
    
    Call BuildTableEdition(FieldName, "SELECT * FROM Assure WHERE POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & IIf(Filtre <> "", Filtre, "") & " ORDER BY PONUMCLE, POARRET")
    
    dateCalcul = Format(m_dataHelper.GetParameter("SELECT PEDATEEXT FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")
    fraisGestionIncap = m_dataHelper.GetParameter("SELECT PEGESTIONINCAP FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode)
    fraisGestionInval = m_dataHelper.GetParameter("SELECT PEGESTIONINVAL FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode)
    
    If chkRecap.Value = 1 Then
      LanceEtat titre, IIf(typeTotaux, "type1.rpt", "type1bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
      LanceEtat titre, IIf(typeTotaux, "type2.rpt", "type2bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
      LanceEtat titre, IIf(typeTotaux, "type3.rpt", "type3bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
      LanceEtat titre, IIf(typeTotaux, "type4.rpt", "type4bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
      LanceEtat titre, IIf(typeTotaux, "type5.rpt", "type5bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
      LanceEtat titre, IIf(typeTotaux, "type6.rpt", "type6bis.rpt"), dateCalcul, fraisGestionIncap, fraisGestionInval
    End If
    
    If chkDetail.Value = 1 Then
      LanceEtat titre, "détail.rpt", dateCalcul, fraisGestionIncap, fraisGestionInval
    End If
  End If
End Sub

'##ModelId=5C8A67AF0119
Private Sub btnLancer_Click()
  Dim rq As String, Filtre As String
  Dim pos As Integer
    
  ' lancement des impressions
  Screen.MousePointer = vbHourglass
  
  ' charge les variables de parametrage : socx, garx
  Soc1 = m_dataHelper.GetParameter("SELECT SONOM FROM Societe WHERE SOCLE=" & cdSoc1)
  Soc2 = Soc1
  Soc3 = Soc1
  'Soc2 = m_dataHelper.GetParameter("SELECT SONOM FROM Societe WHERE SOCLE=" & cdSoc2)
  'Soc3 = m_dataHelper.GetParameter("SELECT SONOM FROM Societe WHERE SOCLE=" & cdSoc3)
  
  Gar11 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar11)
  Gar12 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar12)
  Gar13 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar13)
  Gar14 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar14)
  Gar15 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar15)
  Gar21 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar21)
  Gar22 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar22)
  Gar23 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar23)
  Gar31 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar31)
  Gar32 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar32)
  Gar33 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar33)
    
  Gar51 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar51)
  Gar53 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar53)
  Gar56 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar56)
  Gar57 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar57)
  Gar59 = m_dataHelper.GetParameter("SELECT GALIB FROM Garantie WHERE GAGARCLE=" & cdGar59)
  
  ' description de la période
  Dim dd As String, df As String
  
  dd = Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")
  df = Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode), "dd/mm/yyyy")

  DescriptionPourEtat = """Période " & numPeriode & " ( " & dd & " au " & df & " ) du Groupe " & NomGroupe & """"
  
  'pos = 1
  'Do
  '  pos = InStr(pos, DescriptionPourEtat, "/")
  '  If pos <> 0 Then
  '    DescriptionPourEtat = Left(DescriptionPourEtat, pos) & Mid(DescriptionPourEtat, pos)
  '    pos = pos + 2
  '  End If
  'Loop While pos <> 0
  'pos = InStr(1, DescriptionPourEtat, vbLf)
  'If pos <> 0 Then
  '  DescriptionPourEtat = Left(DescriptionPourEtat, pos - 1) & Mid(DescriptionPourEtat, pos + 1)
  'End If
  'DescriptionPourEtat = """" & DescriptionPourEtat & """"
    
  ' Montants payes
  'Filtre = "POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) : desactivé le 9/1/99
  Filtre = fmFilter.GetSelectionSQLString
  Call BuildAndLaunchReport(chkDetailMontantsPayes, chkRecapMontantsPayes, False, "POPRESTATION", Filtre, "Récapitulation des montants payés par année de survenance")
  
  ' Montants annualises
  'Filtre = "POTERME NOT " & m_datahelper.BuildDateLimit(dd, df) : desactivé le 9/1/99
  Filtre = fmFilter.GetSelectionSQLString
  Call BuildAndLaunchReport(chkDetailMontantsAnnualises, chkRecapMontantsAnnualises, False, "POPRESTATION_AN", Filtre, "Récapitulation des montants annualisés par année de survenance")
  
  ' provisions
  Filtre = fmFilter.GetSelectionSQLString & "AND NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df)
  Call BuildAndLaunchReport(chkDetailProvision, chkRecapProvision, True, "POPM", Filtre, "Provisions mathématiques par année de survenance")
  
  ' entrées
  Filtre = fmFilter.GetSelectionSQLString & "AND (NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) & " AND " & m_dataHelper.BuildDateLimit("POPREMIER_PAIEMENT", dd, df) & ")"
  Call BuildAndLaunchReport(chkDetailEntree, chkRecapEntree, True, "POPM", Filtre, "Récapitulatif des entrées par année de survenance")
  
  ' sorties
  Filtre = fmFilter.GetSelectionSQLString & "AND " & m_dataHelper.BuildDateLimit("POTERME", dd, df)
  Call BuildAndLaunchReport(chkDetailSortie, chkRecapSortie, True, "POPM", Filtre, "Récapitulatif des sorties par année de survenance")
  
  ' évolutions
  Filtre = fmFilter.GetSelectionSQLString & "AND (NOT " & m_dataHelper.BuildDateLimit("POTERME", dd, df) & " AND POPM_VAR <> 0)"
  Call BuildAndLaunchReport(chkDetailEvolution, chkRecapEvolution, True, "POPM_VAR", Filtre, "Récapitulatif des évolutions par année de survenance")
  
  ' Etat C7 - tableau A
  Dim rs As ADODB.Recordset
  
  If chkC7A.Value = 1 Then
    
    Set rs = m_dataSource.OpenRecordset("SELECT SOCLE, SONOM FROM Societe WHERE SOGROUPE = " & GroupeCle, Snapshot)
        
    While Not rs.EOF
      Call LanceEtatC7A(Year(df), " AND POSTECLE = " & rs.fields("SOCLE"), "Pour la société '" & rs.fields("SONOM") & "' du Groupe '" & NomGroupe & "'")
      
      rs.MoveNext
    Wend
        
    rs.Close
        
    Call LanceEtatC7A(Year(df), "", "Pour le groupe " & NomGroupe)
  End If
  
  ' Etat C7 - tableau B
  If chkC7B.Value = 1 Then
    Set rs = m_dataSource.OpenRecordset("SELECT SOCLE, SONOM FROM Societe WHERE SOGROUPE = " & GroupeCle, Snapshot)
        
    While Not rs.EOF
      Call LanceEtatC7B(Year(df), " AND POSTECLE = " & rs.fields("SOCLE"), "Pour la société '" & rs.fields("SONOM") & "' du Groupe '" & NomGroupe & "'")
      
      rs.MoveNext
    Wend
        
    rs.Close
    
    Call LanceEtatC7B(Year(df), "", "Pour le groupe " & NomGroupe)
  End If
  
  ' Etat C7 - tableau C
  If chkC7C.Value = 1 Then
    Set rs = m_dataSource.OpenRecordset("SELECT SOCLE, SONOM FROM Societe WHERE SOGROUPE = " & GroupeCle, Snapshot)
        
    While Not rs.EOF
      Call LanceEtatC7C(Year(df), " AND POSTECLE = " & rs.fields("SOCLE"), "Pour la société '" & rs.fields("SONOM") & "' du Groupe '" & NomGroupe & "'")
      
      rs.MoveNext
    Wend
        
    rs.Close
    
    Call LanceEtatC7C(Year(df), "", "Pour le groupe " & NomGroupe)
  End If
  
  ' etat de avec les données de revalorisation
  LanceEtatRevalo
  
  ' etat de reassurance
  LanceEtatReassurance
  
  ' etat liste des provisions par risques
  LanceEtatProvision
  
  Screen.MousePointer = vbDefault
End Sub

'##ModelId=5C8A67AF0129
Private Sub btnRecap_Click()
  If etatRecap = True Then
    chkRecapEvolution.Value = 1
    chkRecapSortie.Value = 1
    chkRecapEntree.Value = 1
    chkRecapProvision.Value = 1
    chkRecapMontantsAnnualises.Value = 1
    chkRecapMontantsPayes.Value = 1
  Else
    chkRecapEvolution.Value = 0
    chkRecapSortie.Value = 0
    chkRecapEntree.Value = 0
    chkRecapProvision.Value = 0
    chkRecapMontantsAnnualises.Value = 0
    chkRecapMontantsPayes.Value = 0
  End If
  
  etatRecap = Not etatRecap
End Sub

'##ModelId=5C8A67AF0138
Private Sub btnTousC7_Click()
  If btnTousC7.Caption = "Tous" Then
    chkC7A.Value = 1
    chkC7B.Value = 1
    chkC7C.Value = 1
    btnTousC7.Caption = "Aucun"
  Else
    chkC7A.Value = 0
    chkC7B.Value = 0
    chkC7C.Value = 0
    btnTousC7.Caption = "Tous"
  End If
End Sub

'##ModelId=5C8A67AF0148
Private Sub Command1_Click()
  Unload Me
End Sub

'##ModelId=5C8A67AF0157
Private Sub Form_Load()
  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  
  etatDetail = True
  etatRecap = True
  
  ' rempli le label de descrition de la periode
  txtDescription = DescriptionPeriode
  
  ' imprime a l'ecran par defaut
  rdoScreen.Value = True
  
  Unload frmWait
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Configuration initiale d'un rapport
'
'##ModelId=5C8A67AF0177
Private Function InitReport(filename As String) As CRAXDRT.Report
  Dim crxApp As CRAXDRT.Application

  Dim theReport As CRAXDRT.Report
  Dim crxDBTable As CRAXDRT.DatabaseTable

  Set crxApp = New CRAXDRT.Application
  Set theReport = crxApp.OpenReport(App.Path & "\" & filename)

  ' connexion DB
  Dim p As ADODB.Properties

  Set p = m_dataSource.Connection.Properties

  For Each crxDBTable In theReport.Database.Tables
    crxDBTable.SetLogOnInfo p("Data Source").Value, p("Initial Catalog").Value, p("User ID").Value, p("Password").Value
  Next

  Set InitReport = theReport
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Affecte la valeur à une formule dans un rapport
'
'##ModelId=5C8A67AF0196
Private Sub SetFormula(theReport As CRAXDRT.Report, Name As String, Value As String)
  Dim formula As CRAXDRT.FormulaFieldDefinition

  For Each formula In theReport.FormulaFields
    If formula.FormulaFieldName = Name Then
      formula.text = """" & Value & """"
      Exit For
    End If
  Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Affiche un rapport
'
'##ModelId=5C8A67AF01E4
Private Sub ShowReport(theReport As CRAXDRT.Report)
  Dim frm As frmReportViewer

  On Error GoTo err_ShowReport

  Set frm = New frmReportViewer

  Set frm.m_report = theReport

  frm.Caption = theReport.ReportTitle

  frm.Show vbModal, Me

  Exit Sub

err_ShowReport:
  DisplayError
  Resume Next

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Rapports revalo
'
'##ModelId=5C8A67AF0213
Private Sub LanceEtatRevalo()
  Dim Sq As String, cTableInca As String, cTablePass As String, cTableInva As String
  Dim rsParam As ADODB.Recordset
  Dim theReport As CRAXDRT.Report
  
  On Error GoTo err_LanceEtatRevalo
  
  ' Construction de la table "détail"
  m_dataSource.Execute "DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & frmNumPeriode
    
  Sq = "INSERT INTO EditionRevalo(CleGroupe, ClePeriode, PONUMCLE, GAGARCLE, GALIB, POPRESTATION, POPRESTATION_AN, " _
       & " PONAIS, PONOM, POEFFET, POTERME, POARRET, POREPRISE, POPM_INCAP_1R, POPM_PASS_1R, POPM_INVAL_1R, POPM_RENTE_1R, " _
       & " POPM_REVALO, POCOT_REVALO, POPM_RI, POPM, POSTECLE) " _
       & " SELECT " & GroupeCle & ", " & frmNumPeriode & ", Assure.PONUMCLE, Assure.POSIT, CP.Libelle, Assure.POPRESTATION, " _
       & " Assure.POPRESTATION_AN, Assure.PONAIS, Assure.PONOM, Assure.POEFFET, CASE WHEN Assure.POGARCLE=70 THEN Assure.PODEBUT ELSE Assure.POTERME END, CASE WHEN Assure.POGARCLE=70 THEN Assure.POFIN ELSE Assure.POARRET END, " _
       & " Assure.POREPRISE, Assure.POPM_INCAP_1R, Assure.POPM_PASS_1R, Assure.POPM_INVAL_1R, Assure.POPM_REDUC_1R, Assure.POPM_REVALO, " _
       & " Assure.POCOT_REVALO, Assure.POPM_RI, Assure.POPM, Assure.POSTECLE " _
       & " FROM Assure INNER JOIN CodePosition CP ON Assure.POSIT = CP.Position " _
       & " WHERE Assure.POPERCLE=" & frmNumPeriode & IIf(SoCle <> 0, " AND (Assure.POSTECLE=" & SoCle & ")", "") _
       & fmFilter.GetSelectionSQLString
  m_dataSource.Execute Sq
  
  
  ' Paramètres
  Set rsParam = m_dataSource.OpenRecordset("SELECT * FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode & " ORDER BY PENUMPARAMCALCUL", Snapshot)
  
  cTableInca = rsParam.fields("PELMINCAP")
  cTablePass = rsParam.fields("PELMPASSAGE")
  cTableInva = rsParam.fields("PELMINVAL")
  
  
  ' Etat Détail
  If chkDetailProvRevalo.Value = 1 Then
    Set theReport = InitReport("Detail_Revalo.rpt")

    ' Formules
    SetFormula theReport, "TableIncap", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInca & "'")
    SetFormula theReport, "TablePass", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTablePass & "'")
    SetFormula theReport, "TableInval", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInva & "'")
    SetFormula theReport, "TauxTechIncap", rsParam.fields("PETXINCAP")
    SetFormula theReport, "TauxTechInval", rsParam.fields("PETXINVAL")
    SetFormula theReport, "FraisGestIncap", rsParam.fields("PEGESTIONINCAP")
    SetFormula theReport, "FraisGestInval", rsParam.fields("PEGESTIONINVAL")
    SetFormula theReport, "TxRevalo", rsParam.fields("PETXREVALO")
    SetFormula theReport, "TMO", rsParam.fields("PETMO")
    SetFormula theReport, "DureeRevalo", rsParam.fields("PEDUREEREVALO") & " Ans"
    SetFormula theReport, "DebutPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "FinPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "DateImpression", Format(Now, "DD/MM/YYYY")
    SetFormula theReport, "Commentaire", m_dataHelper.GetParameter("SELECT PECOMMENTAIRE FROM PERIODE WHERE PENUMCLE = " & numPeriode)
    SetFormula theReport, "Lissage", rsParam.fields("PEDUREELISSAGE")
    
    SetFormula theReport, "NumPeriode", CStr(frmNumPeriode)
    
    If Trim(fmFilter.SelectionString) <> "" Then
      SetFormula theReport, "txtFiltre", fmFilter.SelectionString
    Else
      SetFormula theReport, "txtFiltre", ""
    End If
    
    theReport.RecordSelectionFormula = " {EditionRevalo.CleGroupe}=" & GroupeCle & " and {EditionRevalo.ClePeriode}=" & numPeriode
    
    theReport.ReportTitle = "Provisions pour sinistres avec revalorisation - Détail"
    theReport.LeftMargin = 400
    theReport.TopMargin = 400

    ShowReport theReport
  End If
  
  
  ' Etat Liste
  If chkListeRevalo.Value = 1 Then
    Set theReport = InitReport("Liste_Revalo.rpt")

    ' Formules
    SetFormula theReport, "TableIncap", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInca & "'")
    SetFormula theReport, "TablePass", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTablePass & "'")
    SetFormula theReport, "TableInval", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInva & "'")
    SetFormula theReport, "TauxTechIncap", rsParam.fields("PETXINCAP")
    SetFormula theReport, "TauxTechInval", rsParam.fields("PETXINVAL")
    SetFormula theReport, "FraisGestIncap", rsParam.fields("PEGESTIONINCAP")
    SetFormula theReport, "FraisGestInval", rsParam.fields("PEGESTIONINVAL")
    SetFormula theReport, "TxRevalo", rsParam.fields("PETXREVALO")
    SetFormula theReport, "TMO", rsParam.fields("PETMO")
    SetFormula theReport, "DureeRevalo", rsParam.fields("PEDUREEREVALO") & " Ans"
    SetFormula theReport, "DebutPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "FinPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "DateImpression", Format(Now, "DD/MM/YYYY")
    SetFormula theReport, "Commentaire", m_dataHelper.GetParameter("SELECT PECOMMENTAIRE FROM PERIODE WHERE PENUMCLE = " & numPeriode)
    SetFormula theReport, "Lissage", rsParam.fields("PEDUREELISSAGE")
        
    SetFormula theReport, "NumPeriode", CStr(frmNumPeriode)
    
    If Trim(fmFilter.SelectionString) <> "" Then
      SetFormula theReport, "txtFiltre", fmFilter.SelectionString
    Else
      SetFormula theReport, "txtFiltre", ""
    End If
    
    theReport.RecordSelectionFormula = " {EditionRevalo.CleGroupe}=" & GroupeCle & " and {EditionRevalo.ClePeriode}=" & numPeriode
    
    theReport.ReportTitle = "Provisions pour sinistres avec revalorisation - Liste"
    theReport.LeftMargin = 400
    theReport.TopMargin = 400

    ShowReport theReport
  End If
  
  
  ' Etat récapitulatif
  If chkRecapProvRevalo = 1 Then
   
    ' Construction de la table "détail"
    m_dataSource.Execute "DELETE FROM EditionRevalo WHERE CleGroupe=" & GroupeCle & " AND ClePeriode=" & frmNumPeriode
    
    Sq = "INSERT INTO EditionRevalo(CleGroupe, ClePeriode, GAGARCLE, GALIB, POSTECLE, POPRESTATION, " _
         & " POPRESTATION_AN, POPM, POPM_REVALO, POCOT_REVALO) " _
         & "SELECT " & GroupeCle & ", " & frmNumPeriode & ", Assure.POSIT, CP.Libelle, " _
         & " Assure.POSTECLE, Sum(Assure.POPRESTATION) , " _
         & " Sum(Assure.POPRESTATION_AN) , Sum(Assure.POPM) , " _
         & " Sum(Assure.POPM_REVALO), Sum(Assure.POCOT_REVALO) " _
         & " FROM Assure INNER JOIN CodePosition CP ON Assure.POSIT = CP.Position " _
         & " WHERE 1=1 " & fmFilter.GetSelectionSQLString _
         & " GROUP BY Assure.POPERCLE, Assure.POSIT, CP.Libelle, Assure.POSTECLE " _
         & " HAVING " & IIf(SoCle <> 0, "(Assure.POSTECLE=" & SoCle & ") AND ", "") & "(Assure.POPERCLE=" & frmNumPeriode & ")"
    m_dataSource.Execute Sq
    
    Set theReport = InitReport("Recap_Revalo.rpt")

    ' Formules
    SetFormula theReport, "TableIncap", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInca & "'")
    SetFormula theReport, "TablePass", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTablePass & "'")
    SetFormula theReport, "TableInval", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInva & "'")
    SetFormula theReport, "TauxTechIncap", rsParam.fields("PETXINCAP")
    SetFormula theReport, "TauxTechInval", rsParam.fields("PETXINVAL")
    SetFormula theReport, "FraisGestIncap", rsParam.fields("PEGESTIONINCAP")
    SetFormula theReport, "FraisGestInval", rsParam.fields("PEGESTIONINVAL")
    SetFormula theReport, "TxRevalo", rsParam.fields("PETXREVALO")
    SetFormula theReport, "TMO", rsParam.fields("PETMO")
    SetFormula theReport, "DureeRevalo", rsParam.fields("PEDUREEREVALO") & " Ans"
    SetFormula theReport, "DebutPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "FinPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "DateImpression", Format(Now, "DD/MM/YYYY")
    SetFormula theReport, "Commentaire", m_dataHelper.GetParameter("SELECT PECOMMENTAIRE FROM PERIODE WHERE PENUMCLE = " & numPeriode)
    SetFormula theReport, "Lissage", rsParam.fields("PEDUREELISSAGE")
        
    SetFormula theReport, "NumPeriode", CStr(frmNumPeriode)
    
    If Trim(fmFilter.SelectionString) <> "" Then
      SetFormula theReport, "txtFiltre", fmFilter.SelectionString
    Else
      SetFormula theReport, "txtFiltre", ""
    End If
    
    theReport.RecordSelectionFormula = " {EditionRevalo.CleGroupe}=" & GroupeCle & " and {EditionRevalo.ClePeriode}=" & numPeriode
    
    theReport.ReportTitle = "Provisions pour sinistres avec revalorisation - Récap"
    theReport.LeftMargin = 400
    theReport.TopMargin = 400

    ShowReport theReport
  End If
  
  If Not (rsParam Is Nothing) Then
    rsParam.Close
    Set rsParam = Nothing
  End If
  
  Exit Sub
  
err_LanceEtatRevalo:
  DisplayError
  Resume Next
End Sub

'##ModelId=5C8A67AF0232
Private Sub LanceEtatReassurance()
  Dim cTableInca As String, cTablePass As String, cTableInva As String
  Dim cTableEduc As String, cTableConjoint As String
  Dim theReport As CRAXDRT.Report
  
  Dim rsParam As ADODB.Recordset

  On Error GoTo err_LanceEtatReassurance


  ' Etat Liste
  If chkListReassurance.Value = vbChecked Then
    '
    ' creation de la vue
    '

    ' Create the command representing the view.
    cTableInca = "SELECT TOP 100 PERCENT Assure.POGPECLE, Assure.POPERCLE, Assure.POSIT as Regime, Year([POARRET]) AS An, " _
                & "Assure.POARRET, Assure.POCATEGORIE, Assure.POTRAITE_RASSUR, Assure.PONOM, Assure.PONUMCLE, " _
                & "Assure.POPM, Assure.POPSAP, Assure.POPM_RASSUR, Assure.POPSAP_RASSUR " _
                & "From Assure Assure "

    cTableInca = cTableInca & "WHERE (POGPECLE=" & GroupeCle & " AND POPERCLE=" & numPeriode & ") " & fmFilter.GetSelectionSQLString
    cTableInca = cTableInca & " AND (Assure.POPM<>0 OR Assure.POPSAP<>0)"
    cTableInca = cTableInca & " ORDER BY [POGARCLE], POCATEGORIE, Year([POARRET]), Assure.POTRAITE_RASSUR, Assure.PONOM"

    ' Create the new View
    m_dataSource.CreateView "EditReass", cTableInca

    '
    ' parametre de l'etat
    '
    Set rsParam = m_dataSource.OpenRecordset("SELECT * FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode & " ORDER BY PENUMPARAMCALCUL", Snapshot)

    cTableInca = rsParam.fields("PELMINCAP")
    cTablePass = rsParam.fields("PELMPASSAGE")
    cTableInva = rsParam.fields("PELMINVAL")

    cTableEduc = rsParam.fields("PETRENTEEDUC")
    cTableConjoint = rsParam.fields("PETRENTECONJOINT")

    Set theReport = InitReport("Liste_Reass.rpt")

    ' Formules
    SetFormula theReport, "TableIncap", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInca & "'")
    SetFormula theReport, "TablePass", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTablePass & "'")
    SetFormula theReport, "TableInval", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInva & "'")
    SetFormula theReport, "TauxTechIncap", rsParam.fields("PETXINCAP")
    SetFormula theReport, "TauxTechInval", rsParam.fields("PETXINVAL")
    SetFormula theReport, "FraisGestIncap", rsParam.fields("PEGESTIONINCAP")
    SetFormula theReport, "FraisGestInval", rsParam.fields("PEGESTIONINVAL")
    
    SetFormula theReport, "TableRenteEduc", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableEduc & "'")
    SetFormula theReport, "TauxTechRenteEduc", rsParam.fields("PETXRTEDUC")
    SetFormula theReport, "FraisGestRenteEduc", rsParam.fields("PEGESTIONEDUC")

    cTableEduc = "Paiement "
    Select Case rsParam.fields("PEFRACTIONEDUC")
      Case 1
        cTableEduc = cTableEduc & "Annuel"
      Case 2
        cTableEduc = cTableEduc & "Semestriel"
      Case 3
        cTableEduc = cTableEduc & "Trimestriel"
      Case 4
        cTableEduc = cTableEduc & "Mensuel"
    End Select
    Select Case rsParam.fields("PEPAIEMENTEDUC")
      Case 1
        cTableEduc = cTableEduc & " d'avance"
      Case 2
        cTableEduc = cTableEduc & " à terme échu"
    End Select
    SetFormula theReport, "PaiementRenteEduc", cTableEduc

    SetFormula theReport, "DebutPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEDEB FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "FinPeriode", Format(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM PERIODE WHERE PENUMCLE = " & numPeriode), "dd/mm/yyyy")
    SetFormula theReport, "Commentaire", m_dataHelper.GetParameter("SELECT PECOMMENTAIRE FROM PERIODE WHERE PENUMCLE = " & numPeriode)

    If Trim(fmFilter.SelectionString) <> "" Then
      SetFormula theReport, "txtFiltre", fmFilter.SelectionString
    Else
      SetFormula theReport, "txtFiltre", ""
    End If
    
    theReport.RecordSelectionFormula = "{EditReass.POGPECLE}=" & GroupeCle & " and {EditReass.POPERCLE}=" & numPeriode
    
    theReport.ReportTitle = "Provisions pour sinistres : Réassurance"
    theReport.LeftMargin = 400
    theReport.TopMargin = 400

    ShowReport theReport
  End If
  
  If Not (rsParam Is Nothing) Then
    rsParam.Close
    Set rsParam = Nothing
  End If
  
  Exit Sub
  
err_LanceEtatReassurance:
  DisplayError
  Resume Next
End Sub

'##ModelId=5C8A67AF0242
Private Sub LanceEtatProvision()
  Dim cTableInca As String, cTablePass As String, cTableInva As String
  Dim cTableEduc As String, cTableConjoint As String
  Dim theReport As CRAXDRT.Report
  
  Dim rsParam As ADODB.Recordset
  
  On Error GoTo err_LanceEtatProvision
  
  ' Etat Liste
  If chkListeProvision.Value = vbChecked Then

    '
    ' parametre de l'etat
    '
    Set rsParam = m_dataSource.OpenRecordset("SELECT * FROM ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode & " ORDER BY PENUMPARAMCALCUL", Snapshot)

    cTableInca = rsParam.fields("PELMINCAP")
    cTablePass = rsParam.fields("PELMPASSAGE")
    cTableInva = rsParam.fields("PELMINVAL")
    
    Set theReport = InitReport("Provisions.rpt")

    ' Formules
    SetFormula theReport, "TableIncap", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInca & "'")
    SetFormula theReport, "TablePass", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTablePass & "'")
    SetFormula theReport, "TableInval", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableInva & "'")
    SetFormula theReport, "TauxTechIncap", rsParam.fields("PETXINCAP")
    SetFormula theReport, "TauxTechInval", rsParam.fields("PETXINVAL")
    SetFormula theReport, "FraisGestIncap", rsParam.fields("PEGESTIONINCAP")
    SetFormula theReport, "FraisGestInval", rsParam.fields("PEGESTIONINVAL")
    
    SetFormula theReport, "TableRenteEduc", m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & cTableEduc & "'")
    SetFormula theReport, "TauxTechRenteEduc", rsParam.fields("PETXRTEDUC")
    SetFormula theReport, "FraisGestRenteEduc", rsParam.fields("PEGESTIONEDUC")
    
    If Trim(fmFilter.SelectionString) <> "" Then
      SetFormula theReport, "txtFiltre", fmFilter.SelectionString
    Else
      SetFormula theReport, "txtFiltre", ""
    End If

    SetFormula theReport, "Titre", "Période " & numPeriode & " - Provisions au " & m_dataHelper.GetParameter("SELECT " & m_dataHelper.BuildSQLDisplayDate("PEDATEEXT") & " FROM PERIODE WHERE PENUMCLE = " & numPeriode)

    theReport.RecordSelectionFormula = "{Assure.POGPECLE}=" & GroupeCle & " and {Assure.POPERCLE}=" & numPeriode & fmFilter.GetCRWFilterSQLString
    
    theReport.ReportTitle = "Provisions pour sinistres : Réassurance"
    theReport.LeftMargin = 400
    theReport.TopMargin = 400

    ShowReport theReport
  End If
  
  If Not (rsParam Is Nothing) Then
    rsParam.Close
    Set rsParam = Nothing
  End If
  
  Exit Sub
  
err_LanceEtatProvision:
  DisplayError
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' specifie les paramètres de connexion à la base de donnée
' il y a autant de LogonInfo(x) que de connexions
'
'##ModelId=5C8A67AF0251
Private Sub SetConnectionString()
'  'CrystalReport1.DataFiles(0) = DatabaseFileName
'  'CrystalReport1.Connect = DatabasePassword
'  'CrystalReport1.Password = Chr$(10) & "P3I32"
'
'  Dim NumberofTables As Integer, i As Integer, n As Integer
'  Dim str As String, DBName As String, UserName As String
'  Dim bCRWDebug As Boolean
'
'  bCRWDebug = GetSettingIni(SectionName, "DB", "CRWDebug", "0") = "1"
'
'  ' DataFiles
'  If bCRWDebug Then
'    NumberofTables = CrystalReport1.RetrieveDataFiles
'    str = ""
'    For i = 0 To NumberofTables - 1
'      str = str & "Table(" & i & ")=" & CrystalReport1.DataFiles(i) & vbLf
'    Next
'    MsgBox str, vbOKOnly + vbInformation, "Liste des tables du rapport " & CrystalReport1.ReportFileName
'  End If
'
'  ' changes DataFiles
'  DBName = GetSettingIni(SectionName, "DB", "CRWDBName", "#")
'  UserName = GetSettingIni(SectionName, "DB", "CRWUserName", "#")
'  If DBName <> "#" And UserName <> "#" Then
'    str = ""
'    For i = 0 To NumberofTables - 1
'      n = InStrRev(CrystalReport1.DataFiles(i), ".")
'      If n > 0 Then
'        CrystalReport1.DataFiles(i) = IIf(DBName <> "", DBName & ".", "") & IIf(UserName <> "", UserName & ".", "") & Mid(CrystalReport1.DataFiles(i), n + 1)
'        str = str & "Table(" & i & ")=" & CrystalReport1.DataFiles(i) & vbLf
'      End If
'    Next
'    If bCRWDebug Then
'      MsgBox str, vbOKOnly + vbInformation, "Liste  MODIFIEE des tables du rapport " & CrystalReport1.ReportFileName
'    End If
'  End If
'
'  ' LogonInfo
'  If bCRWDebug Then
'    NumberofTables = CrystalReport1.RetrieveLogonInfo
'    str = ""
'    For i = 0 To NumberofTables - 1
'      str = str & "LogonInfo(" & i & ")=" & CrystalReport1.LogonInfo(i) & vbLf
'    Next
'    MsgBox str, vbOKOnly + vbInformation, "Liste des LogonInfo du rapport " & CrystalReport1.ReportFileName
'  End If
'
'  ' LogonServer
'  If GetSettingIni(SectionName, "DB", "CRWDLLName", "#") <> "#" Then
'    CrystalReport1.LogOnServer GetSettingIni(SectionName, "DB", "CRWDLLName", "#"), GetSettingIni(SectionName, "DB", "CRWServerName", "#"), DBName, UserName, GetSettingIni(SectionName, "DB", "CRWPassword", "#")
'  End If
'
'  ' changes LogonInfo
'  NumberofTables = CrystalReport1.RetrieveLogonInfo
'  str = ""
'  For i = 0 To NumberofTables - 1
'    CrystalReport1.LogonInfo(i) = CRWDatabaseConnexion
'    str = str & "LogonInfo(" & i & ")=" & CrystalReport1.LogonInfo(i) & vbLf
'  Next
'  If bCRWDebug Then
'    MsgBox str, vbOKOnly + vbInformation, "Liste MODIFIEE des LogonInfo du rapport " & CrystalReport1.ReportFileName
'  End If
'
'  ' Query
'  If bCRWDebug Then
'    CrystalReport1.RetrieveSQLQuery
'    MsgBox CrystalReport1.SQLQuery, vbOKOnly + vbInformation, "SQL pour le rapport " & CrystalReport1.ReportFileName
'  End If
End Sub

'##ModelId=5C8A67AF0271
Private Sub DisplayError()
  Dim str As String
  
  str = "Erreur " & Err & vbLf & Err.Description
  
'  If CrystalReport1.LastErrorNumber <> 0 And CrystalReport1.LastErrorString <> Err.Description Then
'    str = str & vbLf & vbLf & "Erreur CrystalReport " & CrystalReport1.LastErrorNumber & vbLf & CrystalReport1.LastErrorString
'  End If
  
  MsgBox str, vbCritical
End Sub

