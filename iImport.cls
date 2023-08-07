VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const typeGarCE As Byte = 1
Private Const typeGarCI As Byte = 2
Private Const typeGarCR As Byte = 3
 
Private Const cdDuvalletSoc1 As Byte = 1
Private Const cdDuvalletSoc2 As Byte = 2
Private Const cdDuvalletSoc3 As Byte = 3
 
Implements iP3IImport

' date par defaut en cas d'absence d'information
Private gDateNaissance30Ans As Date
Private gDateExtinction65Ans As Date

Private FichierLog As String
Private bErreurPendantImport As Boolean

' nb de rejet lors de l'import
Private NbRejet As Long

Private dateImport As Date


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ecrit un message dans le log
'
Private Sub EcritTraceDansLog(text As String)
  Dim f As Integer
  
  f = FreeFile(0)
  
  bErreurPendantImport = True
  
  Open FichierLog For Append As #f
    Print #f, Now() & " - " & text
  Close #f
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
Private Sub AddNewRecord(SrcDB As Database, rs As Recordset, rs2 As Recordset, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, fWait As Form)
  Dim rs3 As Recordset ' F2_CARTE
  Dim rs4 As Recordset ' F3_CARTE
  
  Dim NumSte As Long
  Dim Somme As Double
  Dim Garantie As Long
  Dim Delai As String, rq_Premier As String, rq_Prec As String, rq As String
  Dim Terme As Date
  Dim IdentA As String
  Dim fRejet As Boolean
  Dim fF1F2EnErreur As Boolean
  
  Dim nbJour As Long, NbJourDepuisDernierPaiement As Long
  Dim montant As Double
  Dim DateDebut_Prec As Date, DateFin_Prec As Date
  Dim DateDebut_Prec_String As String, DateFin_Prec_String As String
  
  On Error GoTo GestionErreur
  
  ' calcul de la période n-1
  DateDebut_Prec = CDate(DateDebut)
  DateDebut_Prec = DateSerial(Year(DateDebut_Prec) - 1, Month(DateDebut_Prec), Day(DateDebut_Prec))
  DateDebut_Prec_String = Format(DateDebut_Prec, "dd/mm/yyyy")
  
  DateFin_Prec = CDate(DateFin)
  DateFin_Prec = DateSerial(Year(DateFin_Prec) - 1, Month(DateFin_Prec), Day(DateFin_Prec))
  DateFin_Prec_String = Format(DateFin_Prec, "dd/mm/yyyy")
  
  
  ' clé identifiant l'assuré
  IdentA = rs.Fields("NUMERO_ADHERENT")
  
  'If IdentA = "09.19807.00" Then Stop
    
  ' source pour la prestation payée : à eclater en 3 articles si 3 paiements
  NumSte = cdSoc1
  Somme = rs.Fields("Montant")
  Garantie = 10 + rs.Fields("GarCle")
  
  '*************** upese
  If rs.Fields("CODE_GARANTIE") = "AT" Then
    Garantie = cdGar14 ' IJ 1 an
  End If
  
  If rs.Fields("CODE_GARANTIE") = "RP" Then
    Garantie = cdGar15 ' IJ 90 Jours
  End If
  '*************** upese
  
  If Somme < 0 Then
    ' rejet des remboursements de cotisation
    Call EcritTraceDansLog("***-> Rejet de la régularisation pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("DATE_ARRET_TRAVAIL") & ", Montant " & Somme & " F")
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' pas de rejet
  fRejet = False
  
  ' ajoute un enregistrement
  rs2.AddNew
  
  rs2.Fields("POGPECLE") = CleGroupe
  rs2.Fields("POPERCLE") = NumPeriode
  rs2.Fields("POSTECLE") = cdSoc1
  
  If IsNull(rs.Fields("Nom")) Then
    rs2.Fields("PONOM") = ">>> ASSURE INCONNU <<<"
    rs2.Fields("PONUMCLE") = IdentA
    rs2.Fields("POCONVENTION") = " "
    rs2.Fields("POSEXE") = 1
    rs2.Fields("POCSP") = " "
    rs2.Fields("PONAIS") = gDateNaissance30Ans
    rs2.Fields("POEFFET") = rs.Fields("DATE_ARRET_TRAVAIL")
    
    Call EcritTraceDansLog("   Valeurs fictives utilisées pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("DATE_ARRET_TRAVAIL"))
  Else
    rs2.Fields("PONOM") = rs.Fields("Nom")
    rs2.Fields("PONUMCLE") = IdentA
    rs2.Fields("POCONVENTION") = Left(IdentA, rs2.Fields("POCONVENTION").Size)
    rs2.Fields("POSEXE") = rs.Fields("Sexe")
    rs2.Fields("POCSP") = rs.Fields("CATEGORIE_PROFESSIONNELLE")
    rs2.Fields("PONAIS") = rs.Fields("DATE_NAISSANCE_ADHERENT")
    
    If IsNull(rs.Fields("DATE_DEBUT_EFFET_CONTRAT")) Then
      rs2.Fields("POEFFET") = rs.Fields("DATE_DEBUT_EFFET_CONTRAT")
    Else
      rs2.Fields("POEFFET") = rs.Fields("DATE_ARRET_TRAVAIL")
    End If
    
    rs2.Fields("POREPRISE") = rs.Fields("DATE_REPRISE_TRAVAIL")
  End If
      
  rs2.Fields("POCAUSE") = rs.Fields("CAUSE_ARRET_TRAVAIL")
  rs2.Fields("POSIT") = rs.Fields("Posit")
          
  If IsNull(rs.Fields("Delai")) Then
    Delai = "0"
  Else
    Delai = rs.Fields("Delai")
  End If
            
  If IsNull(rs.Fields("DATE_FIN_EFFET_CONTRAT")) Then
    Terme = gDateExtinction65Ans
  Else
    Terme = rs.Fields("DATE_FIN_EFFET_CONTRAT")
  End If
      
  '************upese
  Select Case Garantie
    Case cdGar14 ' incap 1 an
      If Terme > DateSerial(Year(rs2.Fields("POEFFET")), Month(rs2.Fields("POEFFET")), Day(rs2.Fields("POEFFET")) + 365 - 1) Then
        Terme = DateSerial(Year(rs2.Fields("POEFFET")), Month(rs2.Fields("POEFFET")), Day(rs2.Fields("POEFFET")) + 365 - 1)
      End If
      
    Case cdGar15 ' incap 90 j
      If Terme > DateSerial(Year(rs2.Fields("POEFFET")), Month(rs2.Fields("POEFFET")), Day(rs2.Fields("POEFFET")) + 90 - 1) Then
        Terme = DateSerial(Year(rs2.Fields("POEFFET")), Month(rs2.Fields("POEFFET")), Day(rs2.Fields("POEFFET")) + 90 - 1)
      End If
  End Select
  '************upese
      
  ' ajoute et configure un enregistrement dans Assure
  rs2.Fields("PODELAI") = Delai
  rs2.Fields("POTERME") = Terme
  rs2.Fields("POPRESTATION") = Somme
  rs2.Fields("POGARCLE") = Garantie
  
  rs2.Fields("POARRET") = rs.Fields("DATE_ARRET_TRAVAIL")
      
     
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' chargement de la requète pour rechercher le PREMIER PAIEMENT
  '                                          le DERNIER PAIEMENT >0 de la période précédente
  '                                          le DERNIER PAIEMENT >0 de la période en cours
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' champs necessaires
  rq_Premier = "SELECT DECOMPTES.DATE_REGLEMENT_DECOMPTE, PRESTATIONS.DATE_DEBUT_PERIODE_INDEMNISEE, " _
               & "PRESTATIONS.DATE_FIN_PERIODE_INDEMNISEE, SUM(PRESTATIONS.MONTANT_PRESTATION) as DernierPaiement " _
               & " FROM ADHERENTS INNER JOIN (CONTRATS INNER JOIN ((SINISTRES INNER JOIN DECOMPTES " _
               & " ON SINISTRES.NUMERO_SINISTRE = DECOMPTES.NUMERO_SINISTRE) INNER JOIN PRESTATIONS " _
               & " ON (DECOMPTES.NUMERO_DECOMPTE = PRESTATIONS.NUMERO_DECOMPTE) " _
               & " AND (SINISTRES.NUMERO_SINISTRE = PRESTATIONS.NUMERO_SINISTRE)) " _
               & " ON (CONTRATS.DATE_DEBUT_EFFET_CONTRAT = SINISTRES.DATE_DEBUT_EFFET_CONTRAT) " _
               & " AND (CONTRATS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT)) " _
               & " ON (ADHERENTS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT) " _
               & " AND (ADHERENTS.NUMERO_ADHERENT = CONTRATS.NUMERO_ADHERENT) "
  rq_Prec = rq_Premier
  rq = rq_Premier
  
  ' condition
  rq_Premier = rq_Premier & " WHERE ADHERENTS.NUMERO_ADHERENT = """ & IdentA & """ AND SINISTRES.DATE_ARRET_TRAVAIL = #" & Format(rs.Fields("DATE_ARRET_TRAVAIL"), "mm/dd/yyyy") & "#" _
               & " AND SINISTRES.ETAT_SINISTRE = ""O"" "

  rq_Prec = rq_Prec & " Where (DECOMPTES.DATE_REGLEMENT_DECOMPTE " & BuildDateLimit(DateDebut_Prec_String, DateFin_Prec_String) & ") " _
            & " AND ADHERENTS.NUMERO_ADHERENT = """ & IdentA & """ AND SINISTRES.DATE_ARRET_TRAVAIL = #" & Format(rs.Fields("DATE_ARRET_TRAVAIL"), "mm/dd/yyyy") & "#" _
            & " AND SINISTRES.ETAT_SINISTRE = ""O"" "

  rq = rq & " Where (DECOMPTES.DATE_REGLEMENT_DECOMPTE  " & BuildDateLimit(DateDebut, DateFin) & ") " _
       & " AND ADHERENTS.NUMERO_ADHERENT = """ & IdentA & """ AND SINISTRES.DATE_ARRET_TRAVAIL = #" & Format(rs.Fields("DATE_ARRET_TRAVAIL"), "mm/dd/yyyy") & "#" _
       & " AND SINISTRES.ETAT_SINISTRE = ""O"" "
  
  ' type de garantie
  rq_Premier = rq_Premier & " AND IIf(Left(Trim(PRESTATIONS.CODE_TYPE_PRESTATION), 1) = ""I"", 1, IIf(Trim(PRESTATIONS.CODE_TYPE_PRESTATION) = ""RI"", 3, 2))=" & rs.Fields("GarCle")
  rq_Prec = rq_Prec & " AND IIf(Left(Trim(PRESTATIONS.CODE_TYPE_PRESTATION), 1) = ""I"", 1, IIf(Trim(PRESTATIONS.CODE_TYPE_PRESTATION) = ""RI"", 3, 2))=" & rs.Fields("GarCle")
  rq = rq & " AND IIf(Left(Trim(PRESTATIONS.CODE_TYPE_PRESTATION), 1) = ""I"", 1, IIf(Trim(PRESTATIONS.CODE_TYPE_PRESTATION) = ""RI"", 3, 2))=" & rs.Fields("GarCle")
  
  '************upese
  Select Case Garantie
    Case cdGar14
      rq_Premier = rq_Premier & " AND PRESTATIONS.CODE_GARANTIE = ""AT"" "
      rq_Prec = rq_Prec & " AND PRESTATIONS.CODE_GARANTIE = ""AT"" "
      rq = rq & " AND PRESTATIONS.CODE_GARANTIE = ""AT"" "
      
    Case cdGar15
      rq_Premier = rq_Premier & " AND PRESTATIONS.CODE_GARANTIE = ""RP"" "
      rq_Prec = rq_Prec & " AND PRESTATIONS.CODE_GARANTIE = ""RP"" "
      rq = rq & " AND PRESTATIONS.CODE_GARANTIE = ""RP"" "
  
    Case Else
      rq_Premier = rq_Premier & " AND PRESTATIONS.CODE_GARANTIE = ""RR"" "
      rq_Prec = rq_Prec & " AND PRESTATIONS.CODE_GARANTIE = ""RR"" "
      rq = rq & " AND PRESTATIONS.CODE_GARANTIE = ""RR"" "
  End Select
  '************upese
  
  ' uniquement les sinistres ouverts
  rq_Premier = rq_Premier & " AND SINISTRES.ETAT_SINISTRE = ""O"" "
  rq_Prec = rq_Prec & " AND SINISTRES.ETAT_SINISTRE = ""O"" "
  rq = rq & " AND SINISTRES.ETAT_SINISTRE = ""O"" "

  ' group
  rq_Premier = rq_Premier & " GROUP BY DECOMPTES.DATE_REGLEMENT_DECOMPTE, PRESTATIONS.DATE_DEBUT_PERIODE_INDEMNISEE, PRESTATIONS.DATE_FIN_PERIODE_INDEMNISEE "
  rq_Prec = rq_Prec & " GROUP BY DECOMPTES.DATE_REGLEMENT_DECOMPTE, PRESTATIONS.DATE_DEBUT_PERIODE_INDEMNISEE, PRESTATIONS.DATE_FIN_PERIODE_INDEMNISEE "
  rq = rq & " GROUP BY DECOMPTES.DATE_REGLEMENT_DECOMPTE, PRESTATIONS.DATE_DEBUT_PERIODE_INDEMNISEE, PRESTATIONS.DATE_FIN_PERIODE_INDEMNISEE "
        
  
 ''''''''''''''''
 ' recherche la date comptabe du PREMIER PAIEMENT dans l'historique F3_CARTE
  
  Set rs4 = SrcDB.OpenRecordset(rq_Premier, dbOpenSnapshot)
  If Not rs4.EOF Then
    ' prends le PREMIER paiement
    rs4.MoveFirst
    rs2.Fields("POPREMIER_PAIEMENT") = rs4.Fields("DATE_REGLEMENT_DECOMPTE")
  End If
  
  rs4.Close
  '''''''''''''''
  
  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' calcul du montant annualisé pour la période PRECEDENTE
  ' cherche le dernier paiement pour une personne, un arret et une prestation
  ' chercher le précédent si la parestation payée est < 0 (=régul)
  
  Set rs4 = SrcDB.OpenRecordset(rq_Prec, dbOpenSnapshot)
  
  If Not rs4.EOF Then
    
    ' prends le dernier paiement
    rs4.MoveLast
    montant = rs4.Fields("DernierPaiement")

   ' recherche du précédent paiement positif
    Do While montant < 0
      rs4.MovePrevious  ' prends le paiement précédent
        If rs4.BOF Then
          montant = 0
          Exit Do
        End If
      montant = rs4.Fields("DernierPaiement")
    Loop
    
    If montant > 0 Then
        
      ' annualisation du montant
        nbJour = DateDiff("d", rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE")) + 1
        NbJourDepuisDernierPaiement = DateDiff("d", rs4.Fields("DATE_REGLEMENT_DECOMPTE"), CDate(DateFin_Prec)) + 1
           
      If nbJour > 0 Then
        montant = Fix((montant * (365 / nbJour) * 100) + 0.5) / 100
        rs2.Fields("POPRESTATION_AN_PREC") = montant
      Else
        Call EcritTraceDansLog("   Période de paiement PRECEDENT non valide pour " & IdentA & " Du " & Format(rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " AU " & Format(rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " : NbJour = " & nbJour)
        rs2.Fields("POPRESTATION_AN_PREC") = 0
      End If
    
    
      If NbJourDepuisDernierPaiement > nbJourMax Then
        Call EcritTraceDansLog("   Annual.= 0 PREC. (anc. paiement > " & nbJourMax & " j) " & IdentA & " du " & Format(rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " au " & Format(rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " : montant = " & montant)
        rs2.Fields("POPRESTATION_AN_PREC") = 0
      End If
    
    Else
      rs2.Fields("POPRESTATION_AN_PREC") = 0
    End If
  
  End If
  
  rs4.Close

  
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' calcul du montant annualisé pour la période EN_COURS
  ' cherche le dernier paiement pour une personne, un arret et une prestation
  ' chercher le précédent si la parestation payée est < 0 (=régul)
  
        
  Set rs4 = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
  
  If Not rs4.EOF Then
    ' prends le dernier paiement
    rs4.MoveLast
    montant = rs4.Fields("DernierPaiement")

   ' recherche du précédent paiement positif
    Do While montant < 0
      rs4.MovePrevious  ' prends le paiement précédent
        If rs4.BOF Then
          montant = 0
          Exit Do
        End If
      montant = rs4.Fields("DernierPaiement")
    Loop
    
    If montant > 0 Then
        
      ' annualisation du montant
        nbJour = DateDiff("d", rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE")) + 1
        NbJourDepuisDernierPaiement = DateDiff("d", rs4.Fields("DATE_REGLEMENT_DECOMPTE"), CDate(DateFin)) + 1
           
      If nbJour > 0 Then
        montant = Fix((montant * (365 / nbJour) * 100) + 0.5) / 100
        rs2.Fields("POPRESTATION_AN") = montant
      Else
        Call EcritTraceDansLog("   Période de paiement non valide pour " & IdentA & " Du " & Format(rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " AU " & Format(rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " : NbJour = " & nbJour)
        rs2.Fields("POPRESTATION_AN") = 0
      End If
    
    
      If NbJourDepuisDernierPaiement > nbJourMax Then
        Call EcritTraceDansLog("   Annual.= 0 (anc. paiement > " & nbJourMax & " j) " & IdentA & " du " & Format(rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " au " & Format(rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE"), "dd/mm/yyyy") & " : montant = " & montant)
        rs2.Fields("POPRESTATION_AN") = 0
      End If
    
      rs2.Fields("PODERNIERPAIEMENT") = 0
      rs2.Fields("PODEBUT") = rs4.Fields("DATE_DEBUT_PERIODE_INDEMNISEE")
      rs2.Fields("POFIN") = rs4.Fields("DATE_FIN_PERIODE_INDEMNISEE")
    Else
      rs2.Fields("POPRESTATION_AN") = 0
      rs2.Fields("PODERNIERPAIEMENT") = 0
    End If
  
  End If
  
  rs4.Close

  ''''''''''''''''''''''''''
  If rs.Fields("GarCle") = 2 Then
    ' rejet des remboursements de cotisation
    rs2.Fields("POPRESTATION_AN") = 0
    Call EcritTraceDansLog("   Annualisation forcée à 0 car Garantie = Exo pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("DATE_ARRET_TRAVAIL") & ", Montant " & Somme & " F")
  End If
  
  If Somme <= 0 Then
    rs2.Fields("POPRESTATION_AN") = 0
    Call EcritTraceDansLog("   Annualisation forcée à 0 car Total prestations payées = 0 " & IdentA & " Arret du " & Format(rs.Fields("DATE_ARRET_TRAVAIL"), "dd/mm/yyyy") & " Garantie " & rs.Fields("GarCle"))
  End If

  ''''''''''''''''''''''''''
  rs2.Fields("PODATEIMPORT") = dateImport
  
  ' ecrit l'enregistrement
  rs2.Update
  
  Exit Sub
  
GestionErreur:
  Screen.MousePointer = vbDefault
  
  fWait.Visible = False
  
  MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & _
         "IdentA = " & IdentA & " Arret du " & Format(rs.Fields("DATE_ARRET_TRAVAIL"), "dd/mm/yyyy"), vbCritical + vbMsgBoxSetForeground
  
  fWait.Visible = True
  
  Screen.MousePointer = vbHourglass
  
  Resume Next
End Sub

Private Function countRecord(SrcDB As Database, rq As String) As Long
  Dim rs As Recordset  ' F3_CARTE
  
  ' parcours les articles correspondant
  Set rs = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
  
  If Not rs.EOF Then
    rs.MoveLast
    
    countRecord = rs.RecordCount
  Else
    countRecord = 0
  End If
  
  rs.Close
End Function

Private Function DoIntermediate(SrcDB As Database, rs2 As Recordset, _
                                rq As String, fWait As frmWait, CleGroupe As Long, NumPeriode As Long, _
                                DateDebut As String, DateFin As String, nbJourMax As Long) As Boolean
  Dim rs As Recordset  ' F3_CARTE
  
  ' parcours les articles correspondant
  Set rs = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
  
  DoIntermediate = True
  
  ' init de parametre de rupture
  Do Until rs.EOF
    ' affiche la position
    fWait.ProgressBar1.Value = fWait.ProgressBar1.Value + 1
    fWait.Label1(0).Caption = "Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max & " - " & NbRejet & " rejets"
    fWait.Refresh
    DoEvents
        
    If fWait.fTravailAnnule = True Then
      Call EcritTraceDansLog("   Import interrompu par l'utilisateur !")
      DoIntermediate = False
      Exit Do
    End If
        
    If rs.Fields("Montant") <> "" Then
      ' ajoute et configure un enregistrement dans Assure pour la garantie Exo
      Call AddNewRecord(SrcDB, rs, rs2, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax, fWait)
    End If
    
    rs.MoveNext
  Loop
  
  rs.Close
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période dans la
' table Assuré
'
Public Function iP3IImport_DoImport(CommonDialog1 As Object, destDB As DAO.Database, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, AgeRetraite As Byte) As Boolean
  Dim SrcDB As Database
  Dim rs2 As Recordset ' Assurés
  Dim rq As String
  Dim maxRecord As Long
  
  Dim codeRetour As Boolean
  
  Dim fWait As frmWait
  
  CommonDialog1.filename = "*.mdb"
  CommonDialog1.ShowOpen
  
  If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.mdb" Then
    iP3IImport_DoImport = False
    Exit Function
  End If
  
  Set fWait = New frmWait
  
  fWait.Caption = "Import en cours..."
  
  fWait.ProgressBar1.Min = 0
  fWait.ProgressBar1.Value = 0
  fWait.ProgressBar1.Max = 100
  
  Screen.MousePointer = vbHourglass
 
  fWait.Show vbModeless
  fWait.Refresh
  
  ' date par defaut en cas d'absence d'information
  gDateNaissance30Ans = DateSerial(Year(DateFin) - 30, 1, 1)  ' age = 30 ans
  gDateExtinction65Ans = DateSerial(Year(DateFin) + 35, 1, 1) ' extinction garantie à 65 ans
  
  NbRejet = 0
  dateImport = Format(Now, "dd/mm/yyyy hh:nn")
  
  ' ouvre une transaction pour ne pas ecrire les changements en cas d'erreur
  DBEngine.BeginTrans
  codeRetour = True
  
  ' efface les anciennent valeurs
  destDB.Execute "DELETE FROM P3IUser.Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
  
  ' preparation du fichier log
  FichierLog = App.Path & "\" & "ErreurImport.log"
  bErreurPendantImport = False
  
  On Error Resume Next
  Kill FichierLog
  On Error GoTo 0
  
  Call EcritTraceDansLog("   Import " & CommonDialog1.filename & " Groupe " & CleGroupe & " Période " & NumPeriode & " du " & DateDebut & " au " & DateFin & " Anc. paiement " & nbJourMax & "j")
  
  ' ouvre la base de donnees maitre
  Set SrcDB = OpenDatabase(CommonDialog1.filename, , True)
  
  ' creer la requete : on parcours en testant DATE_COMPT comprise entre DateDebut et DateFin
  rq = "SELECT ADHERENTS.NUMERO_ADHERENT, [NOM_ADHERENT] & "" "" & [PRENOM_ADHERENT] AS Nom, " _
     & "IIf(ADHERENTS.DENOMINATION_ADHERENT=""Monsieur"",1,2) AS Sexe, CONTRATS.CATEGORIE_PROFESSIONNELLE, " _
     & "ADHERENTS.DATE_NAISSANCE_ADHERENT, CONTRATS.DATE_DEBUT_EFFET_CONTRAT, " _
     & "CONTRATS.DATE_FIN_EFFET_CONTRAT,  SINISTRES.DATE_ARRET_TRAVAIL, SINISTRES.DATE_REPRISE_TRAVAIL, " _
     & "IIf(SINISTRES.DATE_INVALIDITE Is Null,1,2) AS Posit, " _
     & "IIf(SINISTRES.CAUSE_ARRET_TRAVAIL=""ACCIDENT"",OPTIONSFRANCHISES.FRANCHISE_ACCIDENT, " _
     & "IIf(SINISTRES.CAUSE_ARRET_TRAVAIL=""MALADIE"",OPTIONSFRANCHISES.FRANCHISE_MALADIE, " _
     & "OPTIONSFRANCHISES.FRANCHISE_HOSPITALIERE)) AS Delai, sum(PRESTATIONS.MONTANT_PRESTATION) as Montant, " _
     & "IIf(Left(Trim(PRESTATIONS.CODE_TYPE_PRESTATION), 1) = ""I"", 1, IIf(Trim(PRESTATIONS.CODE_TYPE_PRESTATION) = ""RI"", 3, 2)) As GarCle, " _
     & " SINISTRES.CAUSE_ARRET_TRAVAIL, PRESTATIONS.CODE_GARANTIE "
'  rq = rq & "FROM (OPTIONSFRANCHISES INNER JOIN (ADHERENTS INNER JOIN GARANTIES " _
'     & "ON ADHERENTS.NUMERO_ADHERENT = GARANTIES.NUMERO_ADHERENT) " _
'     & "ON (OPTIONSFRANCHISES.OPTION_FRANCHISE = GARANTIES.OPTION_FRANCHISE) " _
'     & "AND (OPTIONSFRANCHISES.CODE_GARANTIE = GARANTIES.CODE_GARANTIE)) " _
'     & "INNER JOIN (CONTRATS INNER JOIN ((SINISTRES INNER JOIN DECOMPTES " _
'     & "ON SINISTRES.NUMERO_SINISTRE = DECOMPTES.NUMERO_SINISTRE) INNER JOIN PRESTATIONS " _
'     & "ON (DECOMPTES.NUMERO_DECOMPTE = PRESTATIONS.NUMERO_DECOMPTE) " _
'     & "AND (SINISTRES.NUMERO_SINISTRE = PRESTATIONS.NUMERO_SINISTRE)) " _
'     & "ON (CONTRATS.DATE_DEBUT_EFFET_CONTRAT = SINISTRES.DATE_DEBUT_EFFET_CONTRAT) " _
'     & "AND (CONTRATS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT)) " _
'     & "ON (ADHERENTS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT) " _
'     & "AND (CONTRATS.DATE_DEBUT_EFFET_CONTRAT = GARANTIES.DATE_DEBUT_EFFET_CONTRAT) " _
'     & "AND (CONTRATS.NUMERO_ADHERENT = GARANTIES.NUMERO_ADHERENT) " _
'     & "AND (ADHERENTS.NUMERO_ADHERENT = CONTRATS.NUMERO_ADHERENT)"
    rq = rq & "FROM (OPTIONSFRANCHISES INNER JOIN " _
      & "(ADHERENTS INNER JOIN GARANTIES ON ADHERENTS.NUMERO_ADHERENT = GARANTIES.NUMERO_ADHERENT) " _
      & "ON (OPTIONSFRANCHISES.CODE_GARANTIE = GARANTIES.CODE_GARANTIE) " _
      & "AND (OPTIONSFRANCHISES.OPTION_FRANCHISE = GARANTIES.OPTION_FRANCHISE)) " _
      & "INNER JOIN (CONTRATS INNER JOIN ((SINISTRES INNER JOIN DECOMPTES " _
      & "ON SINISTRES.NUMERO_SINISTRE = DECOMPTES.NUMERO_SINISTRE) INNER JOIN PRESTATIONS " _
      & "ON (DECOMPTES.NUMERO_DECOMPTE = PRESTATIONS.NUMERO_DECOMPTE) " _
      & "AND (SINISTRES.NUMERO_SINISTRE = PRESTATIONS.NUMERO_SINISTRE)) " _
      & "ON (CONTRATS.DATE_DEBUT_EFFET_CONTRAT = SINISTRES.DATE_DEBUT_EFFET_CONTRAT) " _
      & "AND (CONTRATS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT)) " _
      & "ON (PRESTATIONS.CODE_GARANTIE = GARANTIES.CODE_GARANTIE) " _
      & "AND (ADHERENTS.NUMERO_ADHERENT = SINISTRES.NUMERO_ADHERENT) " _
      & "AND (CONTRATS.DATE_DEBUT_EFFET_CONTRAT = GARANTIES.DATE_DEBUT_EFFET_CONTRAT) " _
      & "AND (CONTRATS.NUMERO_ADHERENT = GARANTIES.NUMERO_ADHERENT) " _
      & "AND (ADHERENTS.NUMERO_ADHERENT = CONTRATS.NUMERO_ADHERENT) "
  rq = rq & " WHERE DECOMPTES.DATE_REGLEMENT_DECOMPTE " & BuildDateLimit(DateDebut, DateFin) _
       & " AND SINISTRES.ETAT_SINISTRE = ""O"" "
  rq = rq & "GROUP BY ADHERENTS.NUMERO_ADHERENT, SINISTRES.DATE_ARRET_TRAVAIL, " _
     & "[NOM_ADHERENT] & "" "" & [PRENOM_ADHERENT], IIf(ADHERENTS.DENOMINATION_ADHERENT=""Monsieur"",1,2), " _
     & "CONTRATS.CATEGORIE_PROFESSIONNELLE, ADHERENTS.DATE_NAISSANCE_ADHERENT, " _
     & "CONTRATS.DATE_DEBUT_EFFET_CONTRAT, CONTRATS.DATE_FIN_EFFET_CONTRAT,SINISTRES.DATE_REPRISE_TRAVAIL, " _
     & "IIf(SINISTRES.DATE_INVALIDITE Is Null,1,2), " _
     & "IIf(SINISTRES.CAUSE_ARRET_TRAVAIL=""ACCIDENT"",OPTIONSFRANCHISES.FRANCHISE_ACCIDENT, " _
     & "IIf(SINISTRES.CAUSE_ARRET_TRAVAIL=""MALADIE"",OPTIONSFRANCHISES.FRANCHISE_MALADIE, " _
     & "OPTIONSFRANCHISES.FRANCHISE_HOSPITALIERE)), " _
     & "IIf(Left(Trim(PRESTATIONS.CODE_TYPE_PRESTATION),1)=""I"",1,IIf(Trim(PRESTATIONS.CODE_TYPE_PRESTATION)=""RI"",3,2)), " _
     & " SINISTRES.CAUSE_ARRET_TRAVAIL, PRESTATIONS.CODE_GARANTIE"

  ' compte le nb d'article à traiter
  maxRecord = countRecord(SrcDB, rq)
  
  If maxRecord <> 0 Then
    fWait.ProgressBar1.Max = maxRecord
  Else
    fWait.ProgressBar1.Max = 1
    
    MsgBox "Aucun article trouvé pour la période spécifiée", vbInformation
  End If
  
  ' ouvre la table assure en ecriture
  Set rs2 = destDB.OpenRecordset("P3IUser.Assure", dbOpenTable)
    
  ' importe les assurés
  codeRetour = DoIntermediate(SrcDB, rs2, rq, fWait, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax)
  
  ' valide si pas d'erreur
  If codeRetour = True Then
    ' valide les changements
    DBEngine.CommitTrans
  Else
    ' annule les changements
    DBEngine.Rollback
  End If
  
  rs2.Close
 
  SrcDB.Close
  
  Call EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
  
  fWait.Hide
  Unload fWait
  
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  If bErreurPendantImport = True Then
    Dim ret As Integer
    
    ret = MsgBox("Des erreurs se sont produites durant l'import," & vbLf & "Voulez-vous consultez le fichier " & FichierLog & " ?", vbInformation Or vbYesNo)
    If ret = vbYes Then
      Dim frm As New frmDisplayLog
      
      frm.FichierLog = FichierLog
      
      frm.Show vbModal
      
      Set frm = Nothing
    End If
  End If
  
  iP3IImport_DoImport = codeRetour
  
  Exit Function
  
GestionErreur:
  rs2.Close
 
  SrcDB.Close
  fWait.Hide
  Unload fWait
  Set fWait = Nothing
  
  Screen.MousePointer = vbDefault
  
  MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
  
  iP3IImport_DoImport = False
End Function



