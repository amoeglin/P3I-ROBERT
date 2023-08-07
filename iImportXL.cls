VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iImportXLADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Feuille excel verrouillée par mot de passe : 'import'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements iP3IImportADO

Private m_Logger As clsLogger

' nb de rejet lors de l'import
Private NbRejet As Long
Private dateImport As Date
Private dateArrete As Date
Private avgPSAP_Regime6 As Double

Private rsCatR9 As ADODB.Recordset, rsCatInval As ADODB.Recordset, rsNCA As ADODB.Recordset
Private rsTaux As ADODB.Recordset

Private nbCatNotFound As Long
Private lstCatNotFound() As Long

Private m_dataHelper As DataHelper

Public Function Arrondi(Valeur As Double, Nbdecimale As Integer) As Double ' fonction arrondi
  Nbdecimale = Abs(Nbdecimale)
  Arrondi = Fix((Valeur * (10 ^ Nbdecimale) + 0.5)) / (10 ^ Nbdecimale)
End Function

Public Function BuildDateLimit(debut As String, Fin As String) As String
  Dim Limit As String
  
  Limit = "BETWEEN #" & Format(CDate(debut), "mm/dd/yyyy") & "# AND #" & Format(CDate(Fin), "mm/dd/yyyy") & "#"
  
  BuildDateLimit = Limit
End Function


Private Function FindTauxRenteEducation(rs As DAO.Recordset, dateFinPeriode As Date) As Double
  ' recherche du taux
  Dim age_b1 As Integer
  Dim taux_b1 As Double
  Dim age_b2 As Integer
  Dim taux_b2 As Double
  Dim age_b3 As Integer
  Dim taux_b3 As Double
  Dim age_b4 As Integer
  Dim taux_b4 As Double
  
  rsTaux.MoveFirst
  rsTaux.Find "Categorie=" & rs.Fields("Categorie")
  
  If Not rsTaux.EOF Then
    age_b1 = rsTaux.Fields("B1")
    taux_b1 = rsTaux.Fields("B1_PCT")
    age_b2 = rsTaux.Fields("B2")
    taux_b2 = rsTaux.Fields("B2_PCT")
    age_b3 = rsTaux.Fields("B3")
    taux_b3 = rsTaux.Fields("B3_PCT")
    age_b4 = rsTaux.Fields("B4")
    taux_b4 = rsTaux.Fields("B4_PCT")
  Else
    age_b1 = 0
    taux_b1 = 0
    age_b2 = 0
    taux_b2 = 0
    age_b3 = 0
    taux_b3 = 0
    age_b4 = 0
    taux_b4 = 0
    
    m_Logger.EcritTraceDansLog "Catégorie à ajouter dans la table 'TauxRenteEducation' : " & rs.Fields("Categorie")
  End If
  
  Dim ageEnfant As Long
  
  ' age de l'enfant
  If IsNull(rs.Fields("DateNaissance")) Then
    FindTauxRenteEducation = 0
    Exit Function
  End If
  
  ageEnfant = DateDiff("yyyy", rs.Fields("DateNaissance"), dateFinPeriode)
    
  If ageEnfant <= 0 Then
    FindTauxRenteEducation = 0
    Exit Function
  End If

  ' recherche du palier
  If ageEnfant < age_b1 Then
    FindTauxRenteEducation = taux_b1
  ElseIf ageEnfant < age_b2 Then
    FindTauxRenteEducation = taux_b2
  ElseIf ageEnfant < age_b3 Then
    FindTauxRenteEducation = taux_b3
  ElseIf ageEnfant < age_b4 Then
    FindTauxRenteEducation = taux_b4
  Else
    FindTauxRenteEducation = 0
  End If

  If FindTauxRenteEducation <= 0 Then
    FindTauxRenteEducation = 0
  End If
End Function

Private Function IsInPeriode(rs As DAO.Recordset, DateDebut As Date, DateFin As Date, nbJourDC As Long) As Boolean
  IsInPeriode = False
  
  If IsNull(rs.Fields("Regime")) Then
    Exit Function
  End If
  
  If IsNull(rs.Fields("DatePaiement")) Or Not IsDate(rs.Fields("DatePaiement")) Then
    If rs.Fields("Regime") = 6 Then
      If (rs.Fields("DateSinistre") + nbJourDC) >= DateDebut And (rs.Fields("DateSinistre") + nbJourDC) <= DateFin Then
        IsInPeriode = True
      End If
    End If
  Else
    If rs.Fields("DatePaiement") >= DateDebut And rs.Fields("DatePaiement") <= DateFin Then
      IsInPeriode = True
    End If
  End If
End Function

Private Function PassageCATR9(categorie As Long, NCA As String, ByRef bFoundCATR9 As Boolean) As Boolean
  PassageCATR9 = False
  bFoundCATR9 = False
  
  Dim nbAnnualForcee As Integer
      
  m_dataHelper.Multi_Find rsCatR9, "Categorie=" & categorie
  
  If Not rsCatR9.EOF Then
    bFoundCATR9 = True
    
    ' existe donc Incap+Passage ou Incap seul
    If CBool(rsCatR9.Fields("PassageSuivantNCA")) = True Then
      m_dataHelper.Multi_Find rsNCA, "NCA='" & NCA & "'"
      
      If rsNCA.EOF Then
        ' si pas d'entrée dans PassageNCA, c'est de l'incap avec passage
        PassageCATR9 = True ' incap avec passage
      Else
        ' decode l'entrée
        If CBool(rsNCA.Fields("Passage")) = True Then
          PassageCATR9 = True ' incap avec passage
        Else
          PassageCATR9 = False ' incap sans passage
        End If
      End If
    Else
      ' decode l'entree dans CATR9
      If CBool(rsCatR9.Fields("Passage")) = True Then
        PassageCATR9 = True ' incap avec passage
      Else
        PassageCATR9 = False ' incap sans passage
      End If
    End If
  End If
End Function

Private Sub AffecteSiNonNullEtNonZero(fd As ADODB.Field, fs As DAO.Field)
  fd = Null
  If Not IsNull(fs) Then
    If fs <> 0 Then
      fd = fs
    End If
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' on regroupe les paiements par sinistre. la date de debut est celle du premier paiement
' et la date de fin est celle du dernier paiement
'
Private Sub AddNewRecord(destDB As DataAccess, rs As DAO.Recordset, rsSalarie As ADODB.Recordset, _
                         CleGroupe As Long, NumPeriode As Long, fWait As Form, AgeRetraite As Byte, _
                         DateDebut As Date, DateFin As Date, nbJourMax As Long, nbJourDC As Long)
  Dim premierPaiement As Variant, dernierPaiement As Variant, poDebut As Variant, poFin As Variant
  Dim poDebutTotal As Variant, poFinTotal As Variant
  Dim dateNaissance As Date ', regroupement As String
  Dim tmpMontant As Double, montant As Double, montantTotal As Double, montantAn As Double, duree As Double, psap As Double
  Dim nbJPer As Long, nbJHors As Long
    
  Dim curSte As Long, n As Integer
  Dim curSS As String, curNCA As String, curRegime As Long, curCat As Long
  Dim curArret As Date, dateReprise As Variant
  Dim varBookmark As Variant
  Dim inPeriode As Boolean, bInit As Boolean
  
  Dim bPassageCATR9 As Boolean, bFoundCATR9 As Boolean
  Dim fieldName As String
  
  On Error GoTo GestionErreur
          
' ...pour debug...
'  fieldName = "Nom"
'  If Not IsNull(rs.Fields("Nom")) Then
'    If InStr(rs.Fields("Nom"), "SUC ") <> 0 Then
'      Stop
'    End If
'  End If
' ...pour debug...
  
  
  inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC)
  
  ' boucle sur un meme NumSociete+NumeroSS+Regime+Categorie+NCA+DateSinistre pour obtenir l'annualisation
  duree = 0
  montant = 0
  montantAn = 0
  montantTotal = 0
  bInit = False
  
  ' raz des dates
  premierPaiement = Null
  dernierPaiement = Null
  poDebut = Null
  poFin = Null
  poDebutTotal = Null
  poFinTotal = Null
  
  fieldName = "Regime"
  If IsNull(rs.Fields("Regime")) Then
    Call m_Logger.EcritTraceDansLog("Regime invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curRegime = rs.Fields("Regime")
  End If
  
  ' If inPeriode Then PHM 9/1/04 : on tient toujours compte des deces
  If inPeriode Or curRegime = 6 Then
    fieldName = "DatePaiement"
    If IsNull(rs.Fields("DatePaiement")) Or Not IsDate(rs.Fields("DatePaiement")) Then
      premierPaiement = Null
    Else
      premierPaiement = rs.Fields("DatePaiement")
    End If
    
    fieldName = "DateDebutPeriode"
    If IsNull(rs.Fields("DateDebutPeriode")) Or Not IsDate(rs.Fields("DateDebutPeriode")) Then
      poDebut = Null
    Else
      poDebut = rs.Fields("DateDebutPeriode")
    End If
    bInit = True
  Else
    inPeriode = False ' test debug
  End If
  
  If IsNull(rs.Fields("DateDebutPeriode")) Or Not IsDate(rs.Fields("DateDebutPeriode")) Then
    poDebutTotal = Null
  Else
    poDebutTotal = rs.Fields("DateDebutPeriode")
  End If
  
  ' init des variables
  fieldName = "NumSociete"
  If IsNull(rs.Fields("NumSociete")) Then
    Call m_Logger.EcritTraceDansLog("NumSociété invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curSte = rs.Fields("NumSociete")
  End If
  
  fieldName = "NumeroSS"
  If IsNull(rs.Fields("NumeroSS")) Then
    Call m_Logger.EcritTraceDansLog("NumeroSS invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curSS = rs.Fields("NumeroSS")
  End If
    
  fieldName = "Categorie"
  If IsNull(rs.Fields("Categorie")) Then
    Call m_Logger.EcritTraceDansLog("Categorie invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    If rs.Fields("Categorie") = 0 Then
      Call m_Logger.EcritTraceDansLog("Categorie invalide à la ligne: " & rs.Fields("RECNO"))
      NbRejet = NbRejet + 1
      Exit Sub
    End If
    curCat = rs.Fields("Categorie")
  End If
  
  fieldName = "NCA"
  If IsNull(rs.Fields("NCA")) Then
    Call m_Logger.EcritTraceDansLog("NCA invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curNCA = rs.Fields("NCA")
  End If
  
  fieldName = "DateSinistre"
  If IsNull(rs.Fields("DateSinistre")) Then
    Call m_Logger.EcritTraceDansLog("DateSinistre invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curArret = rs.Fields("DateSinistre")
  End If
  
  dateReprise = Null
  'regroupement = ""
  
  ' PHM 05/02/2004
  psap = 0
  ' PHM 05/02/2004
  
  ' boucle sur le meme sinistre
  Do
    ' test si fin de fichier
    If rs.EOF Then
      rs.Bookmark = varBookmark
      Exit Do
    End If
      
    inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC)
    
    ' test si toujour le meme gars
    If curSte <> rs.Fields("NumSociete") Or curSS <> rs.Fields("NumeroSS") _
       Or curRegime <> rs.Fields("Regime") Or curCat <> rs.Fields("Categorie") _
       Or curNCA <> rs.Fields("NCA") Or curArret <> rs.Fields("DateSinistre") Then
      rs.Bookmark = varBookmark

      Exit Do
    End If
    
    varBookmark = rs.Bookmark
    
    If inPeriode And Not bInit Then
      If IsNull(rs.Fields("DatePaiement")) Or Not IsDate(rs.Fields("DatePaiement")) Then
        premierPaiement = Null
      Else
        premierPaiement = rs.Fields("DatePaiement")
      End If
      
      If IsNull(rs.Fields("DateDebutPeriode")) Or Not IsDate(rs.Fields("DateDebutPeriode")) Then
        poDebut = Null
      Else
        poDebut = rs.Fields("DateDebutPeriode")
      End If
      bInit = True
    End If
    
    ' toujour le meme sinistre
    tmpMontant = 0
    fieldName = "MontantNet"
    If Not IsNull(rs.Fields("MontantNet")) Then
      tmpMontant = rs.Fields("MontantNet")
    End If
    
    ' cumul des prestations payées presente dans le fichier
    montantTotal = montantTotal + tmpMontant
    
    'If inPeriode Then PHM 9/1/04 : on tient toujours compte des deces
    If inPeriode Then
      nbJPer = 0
      fieldName = "DateFinPeriode"
      If Not IsNull(rs.Fields("DateDebutPeriode")) And Not IsNull(rs.Fields("DateFinPeriode")) Then
        If CDate(rs.Fields("DateDebutPeriode")) > CDate(rs.Fields("DateFinPeriode")) Then
          nbJPer = 0
          Call m_Logger.EcritTraceDansLog("!!!!!! DateDebutPeriode > DateFinPeriode à la ligne: " & rs.Fields("RECNO"))
        Else
          nbJPer = DateDiff("d", rs.Fields("DateDebutPeriode"), CDate(rs.Fields("DateFinPeriode"))) + 1
        End If
      End If
      
      If tmpMontant < 0 Then
        duree = duree - nbJPer
      Else
        duree = duree + nbJPer
      End If
      
      ' cumul prestation payées de la periode
      montant = montant + tmpMontant
      
      ' prestation prise en compte
      montantAn = montantAn + tmpMontant
    
      ' date de reprise
      fieldName = "DateReprise"
      If Not IsNull(rs.Fields("DateReprise")) Then
        dateReprise = rs.Fields("DateReprise")
      End If
      
      ' stocke les infos
      fieldName = "DatePaiement"
      If IsNull(rs.Fields("DatePaiement")) Or Not IsDate(rs.Fields("DatePaiement")) Then
        dernierPaiement = Null
      Else
        dernierPaiement = rs.Fields("DatePaiement")
      End If
      
      fieldName = "DateFinPeriode"
      If IsNull(rs.Fields("DateFinPeriode")) Or Not IsDate(rs.Fields("DateFinPeriode")) Then
        poFin = Null
      Else
        poFin = rs.Fields("DateFinPeriode")
      End If

' PHM 01/03/2004
    Else
      ' pour les régimes 6, les prestas payées apres la date de fin sont comptabilisés comme PSAP
      If curRegime = 6 Then
        Dim datePaiement As Date
        
        If IsNull(rs.Fields("DatePaiement")) Then
          datePaiement = rs.Fields("DateSinistre") + nbJourDC
        Else
          datePaiement = rs.Fields("DatePaiement")
        End If
        
        If datePaiement > DateFin Then
          ' ATTENTION : vérifier que tmpMontant contienne toujours la prestation
          psap = psap + tmpMontant
        End If
      End If
' PHM 01/03/2004
    End If

    If inPeriode Or curRegime = 6 Then
      If curRegime = 6 Then
        ' stocke les infos
        fieldName = "DatePaiement"
        If IsNull(rs.Fields("DatePaiement")) Or Not IsDate(rs.Fields("DatePaiement")) Then
          dernierPaiement = Null
        Else
          dernierPaiement = rs.Fields("DatePaiement")
        End If
        
        fieldName = "DateFinPeriode"
        If IsNull(rs.Fields("DateFinPeriode")) Or Not IsDate(rs.Fields("DateFinPeriode")) Then
          poFin = Null
        Else
          poFin = rs.Fields("DateFinPeriode")
        End If
      End If
      
      ' cumul des psap, tous les PSAP pour les décès
      tmpMontant = 0
      fieldName = "PSAP"
      If Not IsNull(rs.Fields("PSAP")) Then
        tmpMontant = rs.Fields("PSAP")
      End If
      psap = psap + tmpMontant
    End If

'*** OPTION : calcul juste de l'annualisation sur la periode en cours
'    ' avant la periode
'    If rs.Fields("DateDebutPeriode") < DateDebut Then
'      nbJHors = DateDiff("d", rs.Fields("DateDebutPeriode"), DateDebut) + 1
'
'      If nbJHors > nbJPer Then
'        tmpMontant = 0
'      Else
'        If nbJPer <> 0 And nbJHors <> 0 Then
'          tmpMontant = tmpMontant * (CDbl(nbJPer - nbJHors) / CDbl(nbJPer))
'        End If
'      End If
'    End If
'
'    ' apres la periode
'    If rs.Fields("DateFinPeriode") > DateFin Then
'      nbJHors = DateDiff("d", DateFin, rs.Fields("DatefinPeriode")) + 1
'
'      If nbJHors > nbJPer Then
'        tmpMontant = 0
'      Else
'        If nbJPer <> 0 And nbJHors <> 0 Then
'          tmpMontant = tmpMontant * (CDbl(nbJPer - nbJHors) / CDbl(nbJPer))
'        End If
'      End If
'    End If
    
    'If Not IsNull(rs.Fields("Regroupement")) Then
    '  regroupement = Trim(rs.Fields("Regroupement"))
    'End If
    
    ' affiche la position
    fWait.ProgressBar1.Value = rs.AbsolutePosition
    fWait.Label1(0).Caption = "Phase 2 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max & " - " & NbRejet & " rejets"
    fWait.Refresh
    DoEvents
    ' prochain paiement
    rs.MoveNext
  Loop
 
  ' elimine les ' ' du n° de sécu
  Do
    n = InStr(curSS, " ")
    If n <> 0 Then
      curSS = Left(curSS, n - 1) & Mid(curSS, n + 1)
    End If
  Loop Until n = 0
  
  ' elimine les ' ' du NCA
  Do
    n = InStr(curNCA, " ")
    If n <> 0 Then
      curNCA = Left(curNCA, n - 1) & Mid(curNCA, n + 1)
    End If
  Loop Until n = 0
  
  'If inPeriode then PHM 9/1/04 : on tient toujours compte des deces
  If inPeriode Or curRegime = 6 Then
    ' stocke les infos
    fieldName = "DatePaiement"
    If Not IsNull(rs.Fields("DatePaiement")) And IsDate(rs.Fields("DatePaiement")) Then
      dernierPaiement = rs.Fields("DatePaiement")
    End If
    
    fieldName = "DateFinPeriode"
    If IsNull(rs.Fields("DateFinPeriode")) Or Not IsDate(rs.Fields("DateFinPeriode")) Then
      poFin = rs.Fields("DateFinPeriode")
    End If
  End If
  
  fieldName = "DateFinPeriode"
  If IsNull(rs.Fields("DateFinPeriode")) Or Not IsDate(rs.Fields("DateFinPeriode")) Then
    poFinTotal = Null
  Else
    poFinTotal = rs.Fields("DateFinPeriode")
  End If

  ' store data
  rsSalarie.AddNew
  
  ' calcul de la date de paiement pour le regime 6 : il n'y a qu'UN seul paiement pour ces regimes
  'rsSalarie.Fields("PODATEPAIEMENTESTIMEE") = False
  If curRegime = 6 And (IsNull(premierPaiement) Or IsNull(dernierPaiement)) Then
    fieldName = "DateSinistre"
    premierPaiement = rs.Fields("DateSinistre") + nbJourDC
    dernierPaiement = premierPaiement
    rsSalarie.Fields("PODATEPAIEMENTESTIMEE") = True
  End If
  
  rsSalarie.Fields("POGPECLE") = CleGroupe
  rsSalarie.Fields("POPERCLE") = NumPeriode
  
  fieldName = "NumSociete"
  If Not IsNull(rs.Fields("NumSociete")) Then
    rsSalarie.Fields("POSTECLE") = rs.Fields("NumSociete")
  Else
    Call m_Logger.EcritTraceDansLog("Société invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  If curSS Then
    rsSalarie.Fields("PONUMCLE") = curSS
  Else
    Call m_Logger.EcritTraceDansLog("NumeroSS invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' garantie
  If curRegime <> 0 Then
    Select Case curRegime
      Case 1, 3, 6, 7, 9, 10, 11
        ' OK
      Case Else
        Call m_Logger.EcritTraceDansLog("Regime (=" & curRegime & ") invalide à la ligne: " & rs.Fields("RECNO"))
        rsSalarie.CancelUpdate
        NbRejet = NbRejet + 1
        Exit Sub
    End Select
    
    rsSalarie.Fields("POGARCLE") = curRegime + 50
  Else
    Call m_Logger.EcritTraceDansLog("Regime (=Null)invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  If curCat <> 0 Then
    rsSalarie.Fields("POCATEGORIE") = curCat
  Else
    rsSalarie.Fields("POCATEGORIE") = Null
  End If
  
  fieldName = "Nom"
  If Not IsNull(rs.Fields("Nom")) Then
    rsSalarie.Fields("PONOM") = rs.Fields("Nom")
  Else
    Call m_Logger.EcritTraceDansLog("Nom invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' NCA
  If curNCA <> "" Then
    rsSalarie.Fields("POCONVENTION") = curNCA
  Else
    Call m_Logger.EcritTraceDansLog("NCA invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' date du sinistre
  If Not IsNull(curArret) Then
    rsSalarie.Fields("POEFFET") = curArret
    rsSalarie.Fields("POARRET") = curArret
  Else
    Call m_Logger.EcritTraceDansLog("DateSinistre invalide à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
    
  ' date de naissance
  fieldName = "DateNaissance"
  If IsNull(rs.Fields("DateNaissance")) Then
    ' calcul en fct du n° de SS
    fieldName = "NumeroSS"
    dateNaissance = DateSerial(Mid(rs.Fields("NumeroSS"), 2, 2), Mid(rs.Fields("NumeroSS"), 4, 2), 1)
  Else
    dateNaissance = rs.Fields("DateNaissance")
  End If

  rsSalarie.Fields("PONAIS") = dateNaissance
  If curRegime = 9 And (DateDiff("yyyy", dateNaissance, DateFin) < 20 Or DateDiff("yyyy", dateNaissance, DateFin) > 65) Then
    Call m_Logger.EcritTraceDansLog("Attention : Age < 20 ans ou Age > 65 ans à la ligne: " & rs.Fields("RECNO"))
  End If
  
  If curRegime = 3 And DateDiff("yyyy", dateNaissance, DateFin) > 26 Then
    Call m_Logger.EcritTraceDansLog("Attention : Age > 26 ans à la ligne: " & rs.Fields("RECNO"))
  End If
  
  ' terme = age retraite
  If curRegime <> 1 And curRegime <> 10 And curRegime <> 11 Then ' sauf pour la rente de conjoint où terme=null => rente viagere
    rsSalarie.Fields("POTERME") = DateSerial(Year(dateNaissance) + AgeRetraite, Month(dateNaissance), Day(dateNaissance))
  End If
  
  ' debut garantie
  fieldName = "DebutGarantie"
  If Not IsNull(rs.Fields("DebutGarantie")) Then
    If IsDate(rs.Fields("DebutGarantie")) Then
      rsSalarie.Fields("POEFFET") = rs.Fields("DebutGarantie")
    End If
  End If
  
  ' Fin garantie
  fieldName = "FinGarantie"
  If Not IsNull(rs.Fields("FinGarantie")) Then
    If IsDate(rs.Fields("FinGarantie")) Then
      rsSalarie.Fields("POTERME") = rs.Fields("FinGarantie")
    End If
  End If
  
  ' la fin de garantie doit être spécifié 55, 60 ou 65 ans pour les rente de conjoint temporaire
  If curRegime = 10 And IsNull(rs.Fields("FinGarantie")) Then
    Call m_Logger.EcritTraceDansLog("!!!!!! FinDeGarantie invalide pour le regime 10 à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' la fin de garantie doit être spécifié 55, 60 ou 65 ans pour les rente de conjoint temporaire
  If curRegime = 11 And Not IsNull(rs.Fields("FinGarantie")) Then
    Call m_Logger.EcritTraceDansLog("FinDeGarantie présente (à confirmer) pour le regime 11 à la ligne: " & rs.Fields("RECNO"))
    Exit Sub
  End If
  
  ' date de reprise
  fieldName = "DateReprise"
  rsSalarie.Fields("POREPRISE") = rs.Fields("DateReprise")
  
  ' cause
  fieldName = "CauseArret"
  rsSalarie.Fields("POCAUSE") = rs.Fields("CauseArret")
  
  ' periode comptable
  rsSalarie.Fields("PODEBUT") = poDebut
  rsSalarie.Fields("POFIN") = poFin
  
  ' periode comptable totale
  rsSalarie.Fields("PODEBUTTOTAL") = poDebutTotal
  rsSalarie.Fields("POFINTOTAL") = poFinTotal
  
  ' delai de carence
  fieldName = "DelaiCarence"
  If Not IsNull(rs.Fields("DelaiCarence")) Then
    rsSalarie.Fields("PODELAI") = rs.Fields("DelaiCarence")
  Else
    rsSalarie.Fields("PODELAI") = 0
    'Call m_Logger.EcritTraceDansLog("DelaiCarence forcé à 0 à la ligne: " & rs.Fields("RECNO"))
  End If
  
  ' montant de la prestation
  rsSalarie.Fields("POPRESTATION") = montant ' non recalculé
  If duree <= 0 Then
    rsSalarie.Fields("POPRESTATION_AN") = 0
    If curRegime <> 6 Then
      Call m_Logger.EcritTraceDansLog("   Annual.= 0 car duree<=0 à la ligne: " & rs.Fields("RECNO") & " Régime " & curRegime)
    End If
  Else
    rsSalarie.Fields("POPRESTATION_AN") = Arrondi(montantAn * (365# / duree), 2) ' recalculé
  End If
    
  If rsSalarie.Fields("POPRESTATION_AN") < 0 Then
    Call m_Logger.EcritTraceDansLog("   Annual.= 0 car montant<0 à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.Fields("POPRESTATION_AN") = 0
  End If
  
  rsSalarie.Fields("POPRESTATIONTOTAL") = montantTotal
  
  ' rente education
  If curRegime = 3 Then
    ' annualisation pour une garantie de 100%
    rsSalarie.Fields("POPRESTATION_AN") = Arrondi(rsSalarie.Fields("POPRESTATION_AN") _
                                                   / (1 + FindTauxRenteEducation(rs, DateFin)) _
                                                   , 2 _
                                                 )
  End If
  ' premier et dernier paiement
  rsSalarie.Fields("POPREMIER_PAIEMENT") = premierPaiement
  rsSalarie.Fields("PODERNIERPAIEMENT") = dernierPaiement
  
  ' anciennete superieur à 180j
  rsSalarie.Fields("PODOSSIERCLOS") = False
  
  nbJPer = 0
  fieldName = "DateFinPeriode"
  If IsNull(rs.Fields("DateFinPeriode")) Then
    If rsSalarie.Fields("POGARCLE") <> cdGar56 Then
      Call m_Logger.EcritTraceDansLog("Fin de Période invalide à la ligne: " & rs.Fields("RECNO"))
    End If
  End If
  
  If Not IsNull(rs.Fields("DateFinPeriode")) Then
    nbJPer = DateDiff("d", rs.Fields("DateFinPeriode"), dateArrete) + 1
    If nbJPer > nbJourMax Then 'And (rsSalarie.Fields("POGARCLE") = cdGar57 Or rsSalarie.Fields("POGARCLE") = cdGar59) Then
      'Call m_Logger.EcritTraceDansLog("   Annual.= 0 (anc. paiement > " & nbJourMax & " j) à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.Fields("POPRESTATION_AN") = 0
      rsSalarie.Fields("POCOMMENTANNUL") = "Import: Annual.= 0 (anc. paiement > " & nbJourMax & " j)"
      If curRegime <> 6 Then
        rsSalarie.Fields("PODOSSIERCLOS") = True
      End If
    End If
  End If
  
  ' PSAP
  rsSalarie.Fields("POPSAPCAPMOYEN") = False
  rsSalarie.Fields("POPSAP") = 0
  If curRegime = 6 Then
    ' deces
'PHM 15/01/2004
'    If dernierPaiement > dateArrete Then
'      If rsSalarie.Fields("POPRESTATION") = 0 Then
'        If psap = 0 Then
'          rsSalarie.Fields("POPSAP") = avgPSAP_Regime6
'          rsSalarie.Fields("POPSAPCAPMOYEN") = True
'        End If
'      Else
'        rsSalarie.Fields("POPSAP") = rsSalarie.Fields("POPRESTATION")
'      End If
'    End If
    
    If rsSalarie.Fields("POPRESTATIONTOTAL") + rsSalarie.Fields("POPRESTATION") + rsSalarie.Fields("POPSAP") = 0 Then
      If psap = 0 Then
        rsSalarie.Fields("POPSAP") = avgPSAP_Regime6
        rsSalarie.Fields("POPSAPCAPMOYEN") = True
      End If
    End If
'PHM 15/01/2004
    
    ' on rajoute la PSAP du fichier
    rsSalarie.Fields("POPSAP") = rsSalarie.Fields("POPSAP") + psap
  Else
    ' autre regime
    If psap <> 0 Then
      rsSalarie.Fields("POPSAP") = psap
    Else
      If duree <> 0 And Not IsNull(rs.Fields("DateFinPeriode")) Then
        'Dim d1 As Double
        Dim d2 As Double
        'd1 = rsSalarie.Fields("POPRESTATION") * DateDiff("d", rsSalarie.Fields("POFIN"), DateFin) / (DateDiff("d", rsSalarie.Fields("PODEBUT"), rsSalarie.Fields("POFIN")) + 1)
        d2 = rsSalarie.Fields("POPRESTATION") * DateDiff("d", rs.Fields("DateFinPeriode"), dateArrete) / duree
        'If d1 <> d2 Then
        '  rsSalarie.Fields("POPSAP") = d1
        'End If
        rsSalarie.Fields("POPSAP") = Arrondi(d2, 2)
      Else
        rsSalarie.Fields("POPSAP") = 0
      End If
    End If
  End If
  
  If rsSalarie.Fields("POPSAP") < 0 Then
    rsSalarie.Fields("POPSAP") = 0
  End If
  
  
  ' position
'* rsSalarie.Fields("POSIT") = 1 ' incap avec passage
'* rsSalarie.Fields("POSIT") = 2 ' inval
'* rsSalarie.Fields("POSIT") = 3 ' incap sans passage
'* rsSalarie.Fields("POSIT") = 4 ' rente conjoint
'* rsSalarie.Fields("POSIT") = 5 ' rente education
'* rsSalarie.Fields("POSIT") = 6 ' deces
  
  Select Case rsSalarie.Fields("POGARCLE")
    Case cdGar51, cdGar510, cdGar511 ' rente de conjoint
      rsSalarie.Fields("POSIT") = 4 ' rente conjoint
    
    Case cdGar53 ' rente educ
      rsSalarie.Fields("POSIT") = 5 ' rente educ
    
    Case cdGar56 ' deces
      rsSalarie.Fields("POSIT") = 6 ' deces
    
    Case cdGar57 ' inaptitude à la conduite
      rsSalarie.Fields("POSIT") = 2
    
    Case cdGar59 ' arret de travail
      ' aller chercher dans la table R9 pour connaitre la position
      rsSalarie.Fields("POSIT") = 2 ' inval par defaut voir
      
      ' test NCA
      bPassageCATR9 = PassageCATR9(curCat, curNCA, bFoundCATR9)
      
      If bFoundCATR9 Then
        ' existe donc Incap+Passage ou Incap seul
        If bPassageCATR9 = True Then
          ' si passage dans l'entree dans CATR9
          ' ou si pas d'entrée dans PassageNCA, c'est de l'incap avec passage
          rsSalarie.Fields("POSIT") = 1 ' incap avec passage
        Else
          rsSalarie.Fields("POSIT") = 3 ' incap sans passage
        End If
      Else
        Dim iCat As Integer, bFound As Boolean
        
        ' recherche si la categorie a déjà été rencontré
        bFound = False
        For iCat = 0 To nbCatNotFound - 1
          If lstCatNotFound(iCat) = curCat Then
            bFound = True
            Exit For
          End If
        Next iCat
        
        ' recherche dans les catégories 9 étant de l'inval
        rsCatInval.MoveFirst
        rsCatInval.Find "Categorie=" & curCat
        
        If Not rsCatInval.EOF Then
          
          If Not bFound Then
            ReDim Preserve lstCatNotFound(nbCatNotFound + 1) As Long
            lstCatNotFound(nbCatNotFound) = curCat
            nbCatNotFound = nbCatNotFound + 1
            'Call m_Logger.EcritTraceDansLog("   Catégorie INVAL : " & curCat)
          End If
        
        Else

          ' ajoute si pas trouvé
          If Not bFound Then
            ReDim Preserve lstCatNotFound(nbCatNotFound + 1) As Long
            lstCatNotFound(nbCatNotFound) = curCat
            nbCatNotFound = nbCatNotFound + 1
            Call m_Logger.EcritTraceDansLog("   Catégorie à ajouter dans CATR9 ou CATR9INVAL : " & curCat)
          End If
        
        End If
      End If
  End Select
    
  ' valeurs par défaut ou calculées
  fieldName = "NumeroSS"
  rsSalarie.Fields("POSEXE") = Left(rs.Fields("NumeroSS"), 1)
  If rsSalarie.Fields("POSEXE") <> "1" And rsSalarie.Fields("POSEXE") <> "2" Then
    rsSalarie.Fields("POSEXE") = "1"
  End If
  
  rsSalarie.Fields("POPRESTATION_AN_PREC") = 0
  
  ' si le salarie a repris dans la periode comptable, on force l'annualisation à 0
  If Not IsNull(dateReprise) Then
    rsSalarie.Fields("POREPRISE") = dateReprise
    rsSalarie.Fields("POPRESTATION_AN") = 0
    Call m_Logger.EcritTraceDansLog("   Annual.= 0 (Date de reprise=" & Format(dateReprise, "dd/mm/yyyy") & ") à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.Fields("POCOMMENTANNUL") = "Import: Annual.= 0 (Date de reprise=" & Format(dateReprise, "dd/mm/yyyy") & ")"
  End If
  
  ' nouveau champs au 20/02/2003
  fieldName = "Nouv Regime"
  If Not IsNull(rs.Fields("Nouv Regime")) Then
    rsSalarie.Fields("POGARCLE_NEW") = rs.Fields("Nouv Regime")
  End If
  
  fieldName = "Nouv Categorie"
  If Not IsNull(rs.Fields("Nouv Categorie")) Then
    rsSalarie.Fields("POCATEGORIE_NEW") = rs.Fields("Nouv Categorie")
  End If
  
  fieldName = "Regroupement"
  If Not IsNull(rs.Fields("Regroupement")) Then
    rsSalarie.Fields("POREGROUPEMENT") = rs.Fields("Regroupement")
  End If
  
  fieldName = "Code Nature"
  If Not IsNull(rs.Fields("Code Nature")) Then
    rsSalarie.Fields("POCODENATURE") = rs.Fields("Code Nature")
  End If
  
  fieldName = "CCN"
  If Not IsNull(rs.Fields("CCN")) Then
    rsSalarie.Fields("POCCN") = rs.Fields("CCN")
  End If
  
  fieldName = "Code Risque"
  If Not IsNull(rs.Fields("Code Risque")) Then
    rsSalarie.Fields("POCODERISQUE") = rs.Fields("Code Risque")
  End If
  ' FIN nouveau champs au 20/02/2003
  
  
  ' DEBUT nouveau champs au 04/02/2004
  fieldName = "DateEntreeInval"
  If Not IsNull(rs.Fields("DateEntreeInval")) Then
    rsSalarie.Fields("PODATEENTREEINVAL") = rs.Fields("DateEntreeInval")
    
    ' test passage sinon on reste en 1
    If PassageCATR9(curCat, curNCA, bFoundCATR9) Then
      rsSalarie.Fields("POEFFET") = rsSalarie.Fields("PODATEENTREEINVAL")
      rsSalarie.Fields("POSIT") = 2 ' on force le passage en inval
    Else
      rsSalarie.Fields("PODATEENTREEINVAL") = rs.Fields("DateEntreeInval")
    End If
    
    If rsSalarie.Fields("PODATEENTREEINVAL") < rsSalarie.Fields("POARRET") Then
      Call m_Logger.EcritTraceDansLog("DateEntreeInval inferieur a date d'arret à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If
  
  If rsSalarie.Fields("POSIT") = 2 And IsNull(rsSalarie.Fields("PODATEENTREEINVAL")) Then
    ' la date d'entree en inval est obligatoire.
    Call m_Logger.EcritTraceDansLog("DateEntreeInval obligatoire et absente pour le régime " & rsSalarie.Fields("POGARCLE") - 50 & " à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  ' FIN nouveau champs au 04/02/2004
 
' les nouveaux champs DC au 08/10/2004
  fieldName = "CategorieRteEducDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteEduc"), rs.Fields("CategorieRteEducDC")
  
  fieldName = "RegimeRteEducDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteEduc"), rs.Fields("RegimeRteEducDC")
  
  fieldName = "CategorieRteConjointViagereDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteConjointViagere"), rs.Fields("CategorieRteConjointViagereDC")
  
  fieldName = "RegimeRteConjointViagereDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteConjointViagere"), rs.Fields("RegimeRteConjointViagereDC")
  
  fieldName = "CategorieRteConjointTempoDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteConjointTempo"), rs.Fields("CategorieRteConjointTempoDC")
  
  fieldName = "RegimeRteConjointTempoDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteConjointTempo"), rs.Fields("RegimeRteConjointTempoDC")
  
  fieldName = "MajoEnfant"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POMajoEnfant"), rs.Fields("MajoEnfant")
  If Not IsNull(rsSalarie.Fields("POMajoEnfant")) Then
    ' MajoEnfant >= 0
    If rs.Fields("MajoEnfant") < 0 Then
      Call m_Logger.EcritTraceDansLog("MajoEnfant (" & rs.Fields("MajoEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If

  
  
  'If rsSalarie.Fields("PONOM") = "ALEXANDRE Frédéric" Then Stop
  
  
  
  fieldName = "TauxGarantieDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POTauxGarantieDC"), rs.Fields("TauxGarantieDC")
  If Not IsNull(rsSalarie.Fields("POTauxGarantieDC")) Then
    ' TauxGarantieDC >= 0
    If rs.Fields("TauxGarantieDC") < 0 Then
      Call m_Logger.EcritTraceDansLog("TauxGarantieDC (" & rs.Fields("TauxGarantieDC") & ") invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If
  
  fieldName = "CategorieDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieDeces"), rs.Fields("CategorieDC")
  
  fieldName = "RegimeDC"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeDeces"), rs.Fields("RegimeDC")
  
  fieldName = "AgeEnfant"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POAgeMoyenEnfant"), rs.Fields("AgeEnfant")
  If Not IsNull(rsSalarie.Fields("POAgeMoyenEnfant")) Then
    ' age enfant >= 0
    If rsSalarie.Fields("POAgeMoyenEnfant") < 0 Then
      Call m_Logger.EcritTraceDansLog("AgeEnfant (" & rsSalarie.Fields("POAgeMoyenEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
    
    ' age enfant > 24
    If rsSalarie.Fields("POAgeMoyenEnfant") > 24 Then
      Call m_Logger.EcritTraceDansLog("AgeEnfant (" & rsSalarie.Fields("POAgeMoyenEnfant") & ") > 24 ans à la ligne: " & rs.Fields("RECNO"))
    End If
  End If
  
  fieldName = "NbEnfant"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("PONbEnfant"), rs.Fields("NbEnfant")
  If Not IsNull(rsSalarie.Fields("PONbEnfant")) Then
    ' nb enfant >= 0 et nb enfant entier
    If rsSalarie.Fields("PONbEnfant") < 0 Or CDbl(Fix(rsSalarie.Fields("PONbEnfant"))) <> CDbl(rsSalarie.Fields("PONbEnfant")) Then
      Call m_Logger.EcritTraceDansLog("NbEnfant (" & rsSalarie.Fields("PONbEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
    
  End If
  
  fieldName = "SalaireAnnuel"
  AffecteSiNonNullEtNonZero rsSalarie.Fields("POSalaireAnnuel"), rs.Fields("SalaireAnnuel")
  If Not IsNull(rsSalarie.Fields("POSalaireAnnuel")) Then
'    rsSalarie.Fields("POSalaireAnnuel") = rs.Fields("SalaireAnnuel")
    
    ' SalaireAnnuel >= 0
    If rsSalarie.Fields("POSalaireAnnuel") < 0 Then
      Call m_Logger.EcritTraceDansLog("SalaireAnnuel (" & rsSalarie.Fields("POSalaireAnnuel") & ") invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If
  
  fieldName = "Cadre"
  If Not IsNull(rs.Fields("Cadre")) Then
    rsSalarie.Fields("POIsCadre") = rs.Fields("Cadre")
  Else
    Call m_Logger.EcritTraceDansLog("Cadre doit être renseigné à la ligne: " & rs.Fields("RECNO"))
    rsSalarie.CancelUpdate
    NbRejet = NbRejet + 1
    Exit Sub
  End If

  fieldName = "SituationFamille"
  If Not IsNull(rs.Fields("SituationFamille")) Then
    n = m_dataHelper.GetParameterAsDouble("SELECT CleSituationFamille FROM P3IUser.SituationFamille WHERE Libelle='" & UCase(rs.Fields("SituationFamille")) & "'")
    If n <> 0 Then
      rsSalarie.Fields("POCleSituationFamille") = n
    Else
      Call m_Logger.EcritTraceDansLog("SituationFamille invalide à la ligne: " & rs.Fields("RECNO"))
      rsSalarie.CancelUpdate
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  Else
    rsSalarie.Fields("POCleSituationFamille") = Null
  End If
' les nouveaux champs DC au 08/10/2004

  rsSalarie.Fields("PODATEIMPORT") = dateArrete
  
  'If regroupement <> "" Then
  '  rsSalarie.Fields("POREGROUPEMENT") = regroupement
  'End If
  
  rsSalarie.Update
  
  Exit Sub
  
GestionErreur:
  Screen.MousePointer = vbDefault
  
  fWait.Visible = False
  
  If Err = 94 Then
  
  Else
    Select Case Err
      Case 3021
        MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
      
      Case 3265
        MsgBox "Erreur " & Err & " : Colonne '" & fieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & fieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO"))
      
      Case 3421
        MsgBox "Erreur " & Err & " : Colonne '" & fieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & fieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO"))
      
      Case -2147217873
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
      
      Case Else
        MsgBox "Erreur " & Err & " à la ligne " & rs.Fields("RECNO") & " : " & Err.Description, vbCritical
    End Select
  End If
  
  fWait.Visible = True
  
  Screen.MousePointer = vbHourglass
  
  Resume Next
End Sub

Private Sub AdjustDeces(destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, fWait As Form)
  On Error GoTo GestionErreur
      
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
  Dim rq As String, where As String
  
  ' fabrique la chaine de recherche des deces dont les enregistrements 7 et 9 relatifs sont à annualiser à 0
  Set rs = destDB.OpenRecordset("SELECT * FROM P3IUser.RisqueDeces WHERE AnnualisationZero<>0", Disconnected)
      
  where = ""
  Do Until rs.EOF
    If Not IsNull(rs.Fields("AnnualisationZero")) Then
      If CBool(rs.Fields("AnnualisationZero")) = True Then
        If where <> "" Then
          where = where & " OR "
        End If
        where = where & " POCODERISQUE = " & rs.Fields("Code")
      End If
    End If
    rs.MoveNext
  Loop
      
  rs.Close
      
  ' fabrique la requete
  rq = "SELECT * FROM P3IUser.ASSURE WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode & " AND POGARCLE=56 AND (" & where & ")"
      
  Dim mark As Variant
  Dim m_dataHelper As DataHelper, nbAnnualForcee As Integer
      
  Set m_dataHelper = destDB.CreateHelper
  
  nbAnnualForcee = 0
  
  Set rs = destDB.OpenRecordset(rq, Disconnected)
  Do Until rs.EOF
    ' affiche la position
    If (rs.AbsolutePosition Mod 10) = 0 Then
      fWait.ProgressBar1.Max = rs.RecordCount + 1
      fWait.ProgressBar1.Value = rs.AbsolutePosition
      fWait.Label1(0).Caption = "Phase 3 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max
      fWait.Refresh
      DoEvents
    End If
    
'    rq = "SELECT * FROM P3IUser.Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode _
'       & " AND (POGARCLE=57 OR POGARCLE=59) " _
'       & " AND PONUMCLE='" & rs.Fields("PONUMCLE") & "' AND POPRESTATION_AN <> 0"
'
'    Set rs2 = destDB.OpenRecordset(rq, Dynamic)
'
'    Do Until rs2.EOF
'      'rs2.Edit
'      Call m_Logger.EcritTraceDansLog("DECES : Annualisation de " & rs2.Fields("POPRESTATION_AN") & " forcée à 0 pour le n° SS " & rs2.Fields("PONUMCLE") & ", régime " & rs2.Fields("POGARCLE") - 50)
'      nbAnnualForcee = nbAnnualForcee + 1
'      rs2.Fields("POPRESTATION_AN") = 0
'      rs2.Fields("POCOMMENTANNUL") = "Import: DECES : Annualisation de " & rs2.Fields("POPRESTATION_AN") & " forcée à 0 pour le n° SS " & rs2.Fields("PONUMCLE") & ", régime " & rs2.Fields("POGARCLE") - 50
'      rs2.Update
'
'      rs2.MoveNext
'    Loop
'
'    rs2.Close
    
    rq = "UPDATE P3IUser.Assure SET " _
       & " POCOMMENTANNUL='Import: DECES : Annualisation de ' + convert(varchar, POPRESTATION_AN) + ' forcée à 0 pour le n° SS ' + PONUMCLE + ', régime ' + convert(varchar, POGARCLE-50), " _
       & " POPRESTATION_AN=0 " _
       & " WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode _
       & " AND (POGARCLE=57 OR POGARCLE=59) " _
       & " AND PONUMCLE='" & rs.Fields("PONUMCLE") & "' AND POPRESTATION_AN <> 0"
    
    destDB.Execute rq
    
    rs.MoveNext
  Loop
      
  rq = "SELECT POCOMMENTANNUL FROM P3IUser.Assure " _
      & " WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode _
      & "   AND POCOMMENTANNUL LIKE 'Import: DECES : Annualisation de %'"
  
  Set rs2 = destDB.OpenRecordset(rq, Disconnected)

  Do Until rs2.EOF
    Call m_Logger.EcritTraceDansLog(rs2.Fields("POCOMMENTANNUL"))
    nbAnnualForcee = nbAnnualForcee + 1

    rs2.MoveNext
  Loop

  rs2.Close
  
  If nbAnnualForcee <> 0 Then
    Call m_Logger.EcritTraceDansLog("DECES : Nb Annualisation forcée à 0 = " & nbAnnualForcee & " / " & rs.RecordCount)
  End If
  
  rs.Close
      
  Exit Sub
  
GestionErreur:
  Screen.MousePointer = vbDefault
  
  fWait.Visible = False
  
  MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
  
  fWait.Visible = True
  
  Screen.MousePointer = vbHourglass
  
  Resume Next
End Sub

Private Sub computeAvgRegime6(rs As DAO.Recordset, dateArrete As Date, nbJourDC As Long, fWait As Form)
  Dim tmpMontant As Double, tmpMontantPSAP As Double, montant As Double
  Dim nbJPer As Long, nbJHors As Long
    
  Dim datePaiement As Date
  Dim curSte As Long, maxRecord As Long, curArret As Date
  Dim curSS As String, curNCA As String, curRegime As Long, curCat As Long
  
  Dim bOk As Boolean
  
  Dim fieldName As String
  On Error GoTo GestionErreur
      
  avgPSAP_Regime6 = 0
  maxRecord = 0
  
  If Not rs.EOF Then
    ' parcours des enregistrements
    Do Until rs.EOF
      If (rs.AbsolutePosition Mod 10) = 0 Then
        ' affiche la position
        fWait.ProgressBar1.Max = rs.RecordCount + 1
        fWait.ProgressBar1.Value = rs.AbsolutePosition
        fWait.Label1(0).Caption = "Phase 1 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max
        fWait.Refresh
        DoEvents
      End If
      
      fieldName = "Regime"
      If rs.Fields("Regime") = 6 Then
        ' init des variables
        fieldName = "NumSociete"
        If IsNull(rs.Fields("NumSociete")) Then
          Call m_Logger.EcritTraceDansLog("NumSociété invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curSte = rs.Fields("NumSociete")
        End If
        
        fieldName = "NumeroSS"
        If IsNull(rs.Fields("NumeroSS")) Then
          Call m_Logger.EcritTraceDansLog("NumeroSS invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curSS = rs.Fields("NumeroSS")
        End If
        
        fieldName = "Regime"
        If IsNull(rs.Fields("Regime")) Then
          Call m_Logger.EcritTraceDansLog("Regime invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curRegime = rs.Fields("Regime")
        End If
        
        fieldName = "Categorie"
        If IsNull(rs.Fields("Categorie")) Then
          Call m_Logger.EcritTraceDansLog("Categorie invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          If rs.Fields("Categorie") = 0 Then
            Call m_Logger.EcritTraceDansLog("Categorie invalide à la ligne: " & rs.Fields("RECNO"))
            NbRejet = NbRejet + 1
            Exit Sub
          End If
          curCat = rs.Fields("Categorie")
        End If
        
        fieldName = "NCA"
        If IsNull(rs.Fields("NCA")) Then
          Call m_Logger.EcritTraceDansLog("NCA invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curNCA = rs.Fields("NCA")
        End If
        
        fieldName = "DateSinistre"
        If IsNull(rs.Fields("DateSinistre")) Then
          Call m_Logger.EcritTraceDansLog("DateSinistre invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curArret = rs.Fields("DateSinistre")
        End If
        
        ' boucle sur le meme sinistre
        bOk = False
        Do
          ' test si fin de fichier
          If rs.EOF Then
            Exit Do
          End If
            
          ' test si toujours le meme gars
          If curSte <> rs.Fields("NumSociete") Or curSS <> rs.Fields("NumeroSS") _
             Or curRegime <> rs.Fields("Regime") Or curCat <> rs.Fields("Categorie") _
             Or curNCA <> rs.Fields("NCA") Or curArret <> rs.Fields("DateSinistre") Then
            Exit Do
          End If
          
          ' toujours le meme sinistre
          tmpMontant = 0
          fieldName = "MontantNet"
          If Not IsNull(rs.Fields("MontantNet")) Then
            tmpMontant = rs.Fields("MontantNet")
          End If
          
          ' toujours le meme sinistre (on ajoute la psap de Excel)
          tmpMontantPSAP = 0
          fieldName = "PSAP"
          If Not IsNull(rs.Fields("PSAP")) Then
            tmpMontantPSAP = rs.Fields("PSAP")
          End If
          
          ' cumul prestation payées
          fieldName = "DatePaiement"
          If IsNull(rs.Fields("DatePaiement")) Then
            datePaiement = rs.Fields("DateSinistre") + nbJourDC
          Else
            datePaiement = rs.Fields("DatePaiement")
          End If
          If (tmpMontant + tmpMontantPSAP) <> 0 And datePaiement > dateArrete Then ' on ne tiens compte que de PSAP non nulle
            avgPSAP_Regime6 = avgPSAP_Regime6 + tmpMontant + tmpMontantPSAP
            bOk = True
          End If
          
          ' prochain paiement
          rs.MoveNext
        Loop
        
        ' compte les sinistres et non pas les mouvements
        If bOk Then
          maxRecord = maxRecord + 1
        End If
      Else
        rs.MoveNext
      End If ' risque 6
    Loop
    
    ' calcul de la moyenne
    If maxRecord <> 0 Then
      avgPSAP_Regime6 = Arrondi(avgPSAP_Regime6 / maxRecord, 2)
    Else
      avgPSAP_Regime6 = 0
    End If
  End If
  
  Exit Sub
  
GestionErreur:
  Screen.MousePointer = vbDefault
  
  fWait.Visible = False
  
  Select Case Err
    Case 3021
      MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
    
    Case 3265
      MsgBox "Erreur " & Err & " : Colonne '" & fieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
      Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & fieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO") + 1)
    
    Case 3421
      MsgBox "Erreur " & Err & " : Colonne '" & fieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
      Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & fieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO") + 1)
    
    Case Else
      MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
  End Select
  
  fWait.Visible = True
  
  Screen.MousePointer = vbHourglass
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période dans la
' table Assuré
'
'* ajouter un parametre dans les détails periode (et paramêtre) pour le
'* mode de calcul de l'annualisation : "dernier flux" ou "tous les flux de la periode"

Public Function iP3IImportADO_DoImport(CommonDialog1 As Object, destDB As DataAccess, CleGroupe As Long, _
                                    NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, _
                                    nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date) As Boolean
  Dim SrcDB As DAO.Database
  Dim rs As DAO.Recordset
  Dim rs3 As DAO.Recordset
  Dim rq As String
  Dim rs2 As ADODB.Recordset
  Dim rq2 As String
  Dim maxRecord As Long
  Dim codeRetour As Boolean
  Dim fWait As frmWait
  Dim datePaiement As Date
  
  On Error GoTo GestionErreur
  
  ' demande confirmation de la destruction des anciennes données
  If MsgBox("ATTENTION: le données vont être écrasées par celles importées." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) = vbNo Then
    iP3IImportADO_DoImport = False
    Exit Function
  End If
  
  ' demande le nom de la base (fichier xls)
  CommonDialog1.FileName = "*.xls"
  CommonDialog1.Filter = "Fichier Excel|*.xls"
  CommonDialog1.ShowOpen
  
  If CommonDialog1.FileName = "" Or CommonDialog1.FileName = "*.xls" Then
    iP3IImportADO_DoImport = False
    Exit Function
  End If
  
  ' date par defaut en cas d'absence d'information
  dateImport = Format(Now, "dd/mm/yyyy hh:nn")
  If IsNull(dateCloture) Then
    dateArrete = Format(Now, "dd/mm/yyyy")
  Else
    dateArrete = dateCloture
  End If
  
  Dim f As New frmSelectDate
  
  Load f
  
  f.gDateDebut = DateDebut
  f.gDateFin = DateFin
  f.DTPicker2 = dateArrete
  f.Show vbModal
  If ret_code = -1 Then
    iP3IImportADO_DoImport = False
    Exit Function
  End If
  
  dateArrete = f.DTPicker2
  
  Unload f
  Set f = Nothing
  
  NbRejet = 0
  
  Set fWait = New frmWait
  
  fWait.Caption = "Import en cours..."
  
  fWait.ProgressBar1.Min = 0
  fWait.ProgressBar1.Value = 0
  fWait.ProgressBar1.Max = 100
  
  Screen.MousePointer = vbHourglass
 
  fWait.Show vbModeless
  fWait.Refresh
  
  ' preparation du fichier log
  Dim n As Integer
  
  rq = ""
  n = InStrRev(CommonDialog1.FileName, "\")
  If n <> 0 Then
    rq = Left(CommonDialog1.FileName, n)
  End If
  
  Set m_Logger = New clsLogger
  
  m_Logger.FichierLog = rq & "ErreurImport.log"
  m_Logger.CreateLog "Import " & CommonDialog1.FileName & " Groupe " & CleGroupe & " Période " & NumPeriode
  
  ' ouvre la base de donnees maitre
  Set SrcDB = OpenDatabase(CommonDialog1.FileName, dbDriverNoPrompt, True, cdExcelExtendedPropertiesDAO)
  
  SrcDB.QueryTimeout = 120
  
  ' ATTENTION : la zone correspondant à la base de données dans la feuille Excel
  ' =========   doit être nommée 'Assure'
  
  '*** moyenne des PSAP des rente deces (regime 6) non null
  rq = "SELECT * FROM Assure WHERE Nom <>''" _
       & " ORDER BY NumSociete, NumeroSS, Regime, Categorie, NCA, DateSinistre, DateDebutPeriode, DateFinPeriode, DatePaiement"
  
  '*******  ATTENTION : ouverture de tous les recordset avant la transaction, sinon : crash
  
  Set m_dataHelper = destDB.CreateHelper
  
  ' ouvre la feuille XL en lecture
  Set rs = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
  Set rs3 = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
  
  ' ouvre la table assure en ecriture
  Set rs2 = destDB.OpenRecordset("P3IUser.Assure", Table)

  Set rsCatR9 = destDB.OpenRecordset("SELECT * From P3IUser.CATR9", Disconnected)
  
  Set rsCatInval = destDB.OpenRecordset("SELECT * From P3IUser.CATR9INVAL", Disconnected)
  
  Set rsNCA = destDB.OpenRecordset("SELECT * From P3IUser.PassageNCA", Disconnected)
  
  Set rsTaux = destDB.OpenRecordset("SELECT * FROM P3IUser.TauxRenteEducation", Disconnected)
  
  ' ouvre une transaction pour ne pas ecrire les changements en cas d'erreur
  codeRetour = True
  
  ' liste des catégorie non trouvée
  nbCatNotFound = 0
  Erase lstCatNotFound
  
  ' transaction...
  destDB.BeginTrans
  
  ' efface les anciennent valeurs
  destDB.Execute "DELETE FROM P3IUser.Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
  
  ' calcul la moyenne des PSAP pour le regime 6
  Call computeAvgRegime6(rs, dateArrete, nbJourDC, fWait)
  
  '*** import des données
  ' compte le nb d'article à traiter
  maxRecord = rs.RecordCount + 1
  
  If maxRecord <> 0 Then
    fWait.ProgressBar1.Max = maxRecord
    'Call m_Logger.EcritTraceDansLog(maxRecord & " articles dans le fichier")
  Else
    fWait.ProgressBar1.Max = 1
    fWait.Hide
    MsgBox "Aucun article trouvé pour la période spécifiée", vbInformation
  End If
  
  
'Dim fi As Field
'
'Debug.Print vbCrLf & "champ de rs"
'For Each fi In rs.Fields
'  Debug.Print vbTab & "'" & fi.Name & "'"
'Next fi
'
'Debug.Print vbCrLf & "champ de rs2"
'For Each fi In rs2.Fields
'  Debug.Print vbTab & "'" & fi.Name & "'"
'Next fi
  
  ' init de parametre de rupture
  Do Until rs3.EOF
    If (rs3.AbsolutePosition Mod 10) = 0 Then
      ' affiche la position
      fWait.ProgressBar1.Value = rs3.AbsolutePosition
      fWait.Label1(0).Caption = "Phase 2 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max & " - " & NbRejet & " rejets"
      fWait.Refresh
      DoEvents
    End If
    
    If fWait.fTravailAnnule = True Then
      Call m_Logger.EcritTraceDansLog("Import interrompu par l'utilisateur !")
      codeRetour = False
      Exit Do
    End If
        
    If rs3.Fields("Nom") <> "" Then
      ' !!!!!!!! AddNewRecord parcours le recordset sur le meme assure, sinistre, garantie...
      ' Ajoute et configure un enregistrement dans Assure
      Call AddNewRecord(destDB, rs3, rs2, CleGroupe, NumPeriode, fWait, AgeRetraite, CDate(DateDebut), CDate(DateFin), nbJourMax, nbJourDC)
    End If
    
    rs3.MoveNext
  Loop
  
  If codeRetour = False Then
    ' efface les valeurs incompletes
    'destDB.Execute "DELETE FROM P3IUser.Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    
    ' transaction...
    destDB.RollbackTrans
  Else
    ' phase 3 : annualisation à 0 des prestations pour les morts vraiment morts!
    Call AdjustDeces(destDB, CleGroupe, NumPeriode, fWait)
    
    ' transaction...
    destDB.CommitTrans
    
    ' mets à jour la date d'extraction
    rq = "UPDATE P3IUser.Periode SET PEDATEEXT = " & m_dataHelper.BuildDateString(dateArrete) & " WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE = " & NumPeriode
    destDB.Execute rq, dbFailOnError
  End If
  
  Call m_Logger.EcritTraceDansLog(rs3.RecordCount & " articles dans le fichier")
  
  Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
  
  rsCatR9.Close
  rsCatInval.Close
  
  rsNCA.Close
  
  rs2.Close
  rs3.Close
  rs.Close
  
  rsTaux.Close
  
  SrcDB.Close
  
  ' liste des catégorie non trouvée
  nbCatNotFound = 0
  Erase lstCatNotFound
  
  fWait.Hide
  Unload fWait
  
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  m_Logger.AfficheErreurLog
  
  iP3IImportADO_DoImport = codeRetour
  
  Set m_dataHelper = Nothing
  
  Exit Function
  
GestionErreur:
  If rs Is Nothing Then
    If Err = 3228 Then Resume
      
    m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
  Else
    fWait.Hide
    Unload fWait
    Set fWait = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
    
    ' efface les valeurs incompletes
    destDB.Execute "DELETE FROM P3IUser.Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    
    SrcDB.Close
    
    iP3IImportADO_DoImport = False
    
    Set m_dataHelper = Nothing
    
    Exit Function
    
    Resume Next
  End If
End Function


