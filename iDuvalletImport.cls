VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iDuvalletImport"
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
Private LastIdentA As String
Private dateImport As Date
'

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
Private Sub AddNewRecord(typeGar As Byte, SrcDB As Database, rs As Recordset, rs2 As Recordset, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, fWait As Form)
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
  IdentA = rs.Fields("Ident_A")
    
  ' source pour la prestation payée : à eclater en 3 articles si 3 paiements
  Select Case typeGar
    Case typeGarCE
      NumSte = rs.Fields("CE")
      Somme = rs.Fields("TotalCE")
      
      Select Case NumSte
        Case cdDuvalletSoc1
          Garantie = cdGar11
        Case cdDuvalletSoc2
          Garantie = cdGar21
        Case cdDuvalletSoc3
          Garantie = cdGar31
      End Select
  
    Case typeGarCI
      NumSte = rs.Fields("CI")
      Somme = rs.Fields("TotalCI")
      
      Select Case NumSte
        Case cdDuvalletSoc1
          Garantie = cdGar12
        Case cdDuvalletSoc2
          Garantie = cdGar22
        Case cdDuvalletSoc3
          Garantie = cdGar32
      End Select
    
    Case typeGarCR
      NumSte = rs.Fields("CR")
      Somme = rs.Fields("TotalCR")
      
      Select Case NumSte
        Case cdDuvalletSoc1
          Garantie = cdGar13
        Case cdDuvalletSoc2
          Garantie = cdGar23
        Case cdDuvalletSoc3
          Garantie = cdGar33
      End Select
  End Select
  
  ' pas de rejet
  fRejet = False
  
  ' test si num societe est 1, 2 ou 3
  If NumSte = 1 Or NumSte = 2 Or NumSte = 3 Then
    ' F3_CARTE est maitre mais ne contient que les prestations
    ' on charge les enregistrements correspondant de F2_CARTE (date de reprise)
    ' et de F1_CARTE (information sur l'assure)
  
    ' cherche les parametres dans les tables F2_CARTE liée à F1_CARTE
    Set rs3 = SrcDB.OpenRecordset("SELECT CONTRAT, CONVENTION, ASSURE, PRENOM, CODE_QUAL, CODE_PROF, DATEN, EFFET, " _
                                  & " REPRISE, CODE_SIN, INVALIDE, DELAIS_GEN, EXTINC_GEN, TERME, DELAIS_CON, EXTINC_CON, " _
                                  & " DELAIS_GRO, EXTINC_GRO" _
                                  & " FROM F2_CARTE LEFT JOIN F1_CARTE ON F1_CARTE.IDENT_A = F2_CARTE.IDENT_A " _
                                  & " WHERE F2_CARTE.IDENT_A = """ & IdentA & """" _
                                  & " AND F2_CARTE.ARRET = #" & Format(rs.Fields("ARRET"), "mm/dd/yyyy") & "#", dbOpenSnapshot)
    fF1F2EnErreur = False
    If rs3.EOF Then
      ' Pas d'enregistrement dans F2_CARTE
      
      ' Test si il y a un d'enregistrement dans F1_CARTE liée à F2_CARTE
      ' (RAPPEL : Il n'y a pas de record (champs à NULL) dans F2_CARTE, mais on a besoin des champs REPRISE et CODE_SIN)
      Set rs3 = SrcDB.OpenRecordset("SELECT CONTRAT, CONVENTION, ASSURE, PRENOM, CODE_QUAL, CODE_PROF, DATEN, EFFET, " _
                                    & " REPRISE, CODE_SIN, INVALIDE, DELAIS_GEN, EXTINC_GEN, TERME, DELAIS_CON, EXTINC_CON, " _
                                    & " DELAIS_GRO, EXTINC_GRO" _
                                    & " FROM F1_CARTE LEFT JOIN F2_CARTE ON F1_CARTE.IDENT_A = F2_CARTE.IDENT_A " _
                                    & " WHERE F1_CARTE.IDENT_A = """ & IdentA & """", dbOpenSnapshot)
      
      If rs3.EOF Then
        ' Pas d'enregistrement dans F1_CARTE
        'Call EcritTraceDansLog("Pas d'article dans F1_CARTE et F2_CARTE pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("ARRET") & " : Article non traité")
        'fRejet = True
        fF1F2EnErreur = True
      Else
        Call EcritTraceDansLog("Pas d'article dans F2_CARTE pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("ARRET") & " : Date de reprise inconnue")
      End If
    End If
    
    If fRejet = False Then
      ' ajoute un enregistrement
      rs2.AddNew
      
      rs2.Fields("POGPECLE") = CleGroupe
      rs2.Fields("POPERCLE") = NumPeriode
      
      Select Case NumSte
        Case cdDuvalletSoc1
          rs2.Fields("POSTECLE") = cdSoc1
      
        Case cdDuvalletSoc2
          rs2.Fields("POSTECLE") = cdSoc2
      
        Case cdDuvalletSoc3
          rs2.Fields("POSTECLE") = cdSoc3
      End Select
      
      If IsNull(rs3.Fields("ASSURE")) Or fF1F2EnErreur Then
        rs2.Fields("PONOM") = ">>> ASSURE INCONNU <<<"
        rs2.Fields("PONUMCLE") = Left(IdentA, rs2.Fields("PONUMCLE").Size)
        rs2.Fields("POCONVENTION") = " "
        rs2.Fields("POSEXE") = 1
        rs2.Fields("POCSP") = " "
        rs2.Fields("PONAIS") = gDateNaissance30Ans
        rs2.Fields("POEFFET") = rs.Fields("ARRET")
        
        Call EcritTraceDansLog("Valeurs fictives utilisées pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("ARRET") & " : Pb F1_CARTE")
      Else
        rs2.Fields("PONOM") = rs3.Fields("ASSURE") & " " & rs3.Fields("PRENOM")
        rs2.Fields("PONUMCLE") = rs3.Fields("CONTRAT")
        rs2.Fields("POCONVENTION") = rs3.Fields("CONVENTION")
        rs2.Fields("POSEXE") = IIf(InStr(1, rs3.Fields("CODE_QUAL"), "1") <> 0, 1, 2)
        rs2.Fields("POCSP") = rs3.Fields("CODE_PROF")
        rs2.Fields("PONAIS") = rs3.Fields("DATEN")
        If IsNull(rs3.Fields("EFFET")) Then
          rs2.Fields("POEFFET") = rs3.Fields("EFFET")
        Else
          rs2.Fields("POEFFET") = rs.Fields("ARRET")
        End If
        rs2.Fields("POREPRISE") = rs3.Fields("REPRISE")
      End If
      
      If Not fF1F2EnErreur Then
        rs2.Fields("POCAUSE") = rs3.Fields("CODE_SIN")
      End If
      'modif moeglin 24/10/98 : rs2.Fields("POSIT") = IIf(rs3.Fields("INVALIDE") = "N", 1, 2)
      rs2.Fields("POSIT") = IIf(typeGar = typeGarCR, 2, 1) ' on force l'invalidité si c'est une rente
          
      ' delai d'extinction
      If Not fF1F2EnErreur Then
        Select Case NumSte
          Case cdDuvalletSoc1
            If IsNull(rs3.Fields("DELAIS_GEN")) Then
              Delai = "0"
            Else
              Delai = rs3.Fields("DELAIS_GEN")
            End If
            
            If IsNull(rs3.Fields("EXTINC_GEN")) Then
              If IsNull(rs3.Fields("TERME")) Then
                Terme = gDateExtinction65Ans
              Else
                Terme = rs3.Fields("TERME")
              End If
            Else
              Terme = rs3.Fields("EXTINC_GEN")
            End If
            
          Case cdDuvalletSoc2
            If IsNull(rs3.Fields("DELAIS_CON")) Then
              Delai = "0"
            Else
              Delai = rs3.Fields("DELAIS_CON")
            End If
            
            If IsNull(rs3.Fields("EXTINC_CON")) Then
              If IsNull(rs3.Fields("TERME")) Then
                Terme = gDateExtinction65Ans
              Else
                Terme = rs3.Fields("TERME")
              End If
            Else
              Terme = rs3.Fields("EXTINC_CON")
            End If
            
          Case cdDuvalletSoc3
            If IsNull(rs3.Fields("DELAIS_GRO")) Then
              Delai = "0"
            Else
              Delai = rs3.Fields("DELAIS_GRO")
            End If
            
            If IsNull(rs3.Fields("EXTINC_GRO")) Then
              If IsNull(rs3.Fields("TERME")) Then
                Terme = gDateExtinction65Ans
              Else
                Terme = rs3.Fields("TERME")
              End If
            Else
              Terme = rs3.Fields("EXTINC_GRO")
            End If
        End Select
      Else
        Delai = "0"
        Terme = gDateExtinction65Ans
      End If
      
      ' ajoute et configure un enregistrement dans Assure
      rs2.Fields("PODELAI") = Delai
      rs2.Fields("POTERME") = Terme
      rs2.Fields("POPRESTATION") = Somme
      rs2.Fields("POGARCLE") = Garantie
      
      rs2.Fields("POARRET") = rs.Fields("ARRET")
      
     
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' chargement de la requète pour rechercher le PREMIER PAIEMENT
      '                                          le DERNIER PAIEMENT >0 de la période précédente
      '                                          le DERNIER PAIEMENT >0 de la période en cours
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      rq_Premier = "SELECT DU, AU, DATE_COMPT, "
      rq_Prec = "SELECT DU, AU, DATE_COMPT, "
      rq = "SELECT DU, AU, DATE_COMPT, "
      
      Select Case typeGar
        Case typeGarCE
          rq_Premier = rq_Premier & " EXO"
          rq_Prec = rq_Prec & " EXO"
          rq = rq & " EXO"
        
        Case typeGarCI
          rq_Premier = rq_Premier & " INDEMNITE"
          rq_Prec = rq_Prec & " INDEMNITE"
          rq = rq & " INDEMNITE"
        
        Case typeGarCR
          rq_Premier = rq_Premier & " RENTE"
          rq_Prec = rq_Prec & " RENTE"
          rq = rq & " RENTE"
      End Select
      
      rq_Premier = rq_Premier & " as DernierPaiement From F3_CARTE Where " _
            & "IDENT_A = """ & IdentA & """ AND ARRET = #" & Format(rs.Fields("ARRET"), "mm/dd/yyyy") & "#" _
  
      rq_Prec = rq_Prec & " as DernierPaiement From F3_CARTE Where (DATE_COMPT " & BuildDateLimit(DateDebut_Prec_String, DateFin_Prec_String) & ") " _
            & " AND IDENT_A = """ & IdentA & """ AND ARRET = #" & Format(rs.Fields("ARRET"), "mm/dd/yyyy") & "#" _
  
      rq = rq & " as DernierPaiement From F3_CARTE Where (DATE_COMPT " & BuildDateLimit(DateDebut, DateFin) & ") " _
            & " AND IDENT_A = """ & IdentA & """ AND ARRET = #" & Format(rs.Fields("ARRET"), "mm/dd/yyyy") & "#" _
  
  
      Select Case typeGar
        Case typeGarCE
          rq_Premier = rq_Premier & " AND CE = """ & NumSte & """"
          rq_Prec = rq_Prec & " AND CE = """ & NumSte & """"
          rq = rq & " AND CE = """ & NumSte & """"
        
        
        Case typeGarCI
          rq_Premier = rq_Premier & " AND CI = """ & NumSte & """"
          rq_Prec = rq_Prec & " AND CI = """ & NumSte & """"
          rq = rq & " AND CI = """ & NumSte & """"
        
        Case typeGarCR
          rq_Premier = rq_Premier & " AND CR = """ & NumSte & """"
          rq_Prec = rq_Prec & " AND CR = """ & NumSte & """"
          rq = rq & " AND CR = """ & NumSte & """"
      End Select
        
      rq_Premier = rq_Premier & " ORDER BY DATE_COMPT"
      rq_Prec = rq_Prec & " ORDER BY DATE_COMPT"
      rq = rq & " ORDER BY DATE_COMPT"
            
      
     ''''''''''''''''
     ' recherche la date comptabe du PREMIER PAIEMENT dans l'historique F3_CARTE
      
      Set rs4 = SrcDB.OpenRecordset(rq_Premier, dbOpenSnapshot)
      If Not rs4.EOF Then
        ' prends le PREMIER paiement
        rs4.MoveFirst
        rs2.Fields("POPREMIER_PAIEMENT") = rs4.Fields("DATE_COMPT")
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
            nbJour = DateDiff("d", rs4.Fields("DU"), rs4.Fields("AU"))
            NbJourDepuisDernierPaiement = DateDiff("d", rs4.Fields("DATE_COMPT"), CDate(DateFin_Prec))
               
          If nbJour > 0 Then
            montant = Fix((montant * (365 / nbJour) * 100) + 0.5) / 100
            rs2.Fields("POPRESTATION_AN_PREC") = montant
          Else
            Call EcritTraceDansLog("Période de paiement PRECEDENT non valide pour " & IdentA & " Du " & Format(rs4.Fields("DU"), "dd/mm/yyyy") & " AU " & Format(rs4.Fields("AU"), "dd/mm/yyyy") & " : NbJour = " & nbJour)
            rs2.Fields("POPRESTATION_AN_PREC") = 0
          End If
        
        
          If NbJourDepuisDernierPaiement > nbJourMax Then
            Call EcritTraceDansLog("Annual.= 0 PREC. (anc. paiement > " & nbJourMax & " j) " & IdentA & " du " & Format(rs4.Fields("DU"), "dd/mm/yyyy") & " au " & Format(rs4.Fields("AU"), "dd/mm/yyyy") & " : montant = " & montant)
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
            nbJour = DateDiff("d", rs4.Fields("DU"), rs4.Fields("AU"))
            NbJourDepuisDernierPaiement = DateDiff("d", rs4.Fields("DATE_COMPT"), CDate(DateFin))
               
          If nbJour > 0 Then
            montant = Fix((montant * (365 / nbJour) * 100) + 0.5) / 100
            rs2.Fields("POPRESTATION_AN") = montant
          Else
            Call EcritTraceDansLog("Période de paiement non valide pour " & IdentA & " Du " & Format(rs4.Fields("DU"), "dd/mm/yyyy") & " AU " & Format(rs4.Fields("AU"), "dd/mm/yyyy") & " : NbJour = " & nbJour)
            rs2.Fields("POPRESTATION_AN") = 0
          End If
        
        
          If NbJourDepuisDernierPaiement > nbJourMax Then
            Call EcritTraceDansLog("Annual.= 0 (anc. paiement > " & nbJourMax & " j) " & IdentA & " du " & Format(rs4.Fields("DU"), "dd/mm/yyyy") & " au " & Format(rs4.Fields("AU"), "dd/mm/yyyy") & " : montant = " & montant)
            rs2.Fields("POPRESTATION_AN") = 0
          End If
        
        Else
          rs2.Fields("POPRESTATION_AN") = 0
        End If
      
      End If
      
      rs4.Close
  
      ''''''''''''''''''''''''''
      rs2.Fields("PODATEIMPORT") = dateImport
      
      ' ecrit l'enregistrement
      rs2.Update
    Else
      fRejet = True
    End If
    
    rs3.Close
  End If
  
  If fRejet = True Then
    ' NumSte = 0
    If Somme <> 0 Then
      Call EcritTraceDansLog("Pas d'article dans F1_CARTE et F2_CARTE pour l'identifiant " & IdentA & ", Arret du " & rs.Fields("ARRET") & " : Article non traité")
      
      ' on comptabilise les rejet par personnes differentes
      If LastIdentA <> IdentA Then
        NbRejet = NbRejet + 1
      End If
    End If
    LastIdentA = IdentA
  End If
  
  Exit Sub
  
GestionErreur:
  Screen.MousePointer = vbDefault
  
  fWait.Visible = False
  
  MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & _
         "IdentA = " & IdentA & " Arret du " & Format(rs.Fields("ARRET"), "dd/mm/yyyy"), vbCritical + vbMsgBoxSetForeground
  
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

Private Function DoIntermediate(SrcDB As Database, rs2 As Recordset, typeGar As Byte, fieldName As String, _
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
      Call EcritTraceDansLog("Import interrompu par l'utilisateur !")
      DoIntermediate = False
      Exit Do
    End If
        
    If rs.Fields(fieldName) <> "" Then
      ' ajoute et configure un enregistrement dans Assure pour la garantie Exo
      Call AddNewRecord(typeGar, SrcDB, rs, rs2, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax, fWait)
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
  Dim rqCE As String, rqCI As String, rqCR As String
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
  LastIdentA = ""
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
  
  Call EcritTraceDansLog("Import " & CommonDialog1.filename & " Groupe " & CleGroupe & " Période " & NumPeriode & " du " & DateDebut & " au " & DateFin & " Anc. paiement " & nbJourMax & "j")
  
  ' ouvre la base de donnees maitre
  Set SrcDB = OpenDatabase(CommonDialog1.filename, , True)
  
  ' creer la requete : on parcours en testant DATE_COMPT comprise entre DateDebut et DateFin
'  rq = "SELECT IDENT_A, ARRET, CE, SUM(EXO) as TotalCE, CI, SUM(INDEMNITE) as TotalCI," _
'       & " CR, SUM(RENTE) as TotalCR" _
'       & " FROM F3_CARTE WHERE DATE_COMPT " & BuildDateLimit(DateDebut, DateFin) _
'       & " GROUP BY IDENT_A, ARRET, CE, CI, CR"
  
  rqCE = "SELECT IDENT_A, ARRET, CE, SUM(EXO) as TotalCE" _
       & " FROM F3_CARTE WHERE DATE_COMPT " & BuildDateLimit(DateDebut, DateFin) _
       & " GROUP BY IDENT_A, ARRET, CE"
  
  rqCI = "SELECT IDENT_A, ARRET, CI, SUM(INDEMNITE) as TotalCI" _
       & " FROM F3_CARTE WHERE DATE_COMPT " & BuildDateLimit(DateDebut, DateFin) _
       & " GROUP BY IDENT_A, ARRET, CI"
  
  rqCR = "SELECT IDENT_A, ARRET, CR, SUM(RENTE) as TotalCR" _
       & " FROM F3_CARTE WHERE DATE_COMPT " & BuildDateLimit(DateDebut, DateFin) _
       & " GROUP BY IDENT_A, ARRET, CR"
  
  ' compte le nb d'article à traiter
  maxRecord = countRecord(SrcDB, rqCE) + countRecord(SrcDB, rqCI) + countRecord(SrcDB, rqCR)
  
  If maxRecord <> 0 Then
    fWait.ProgressBar1.Max = maxRecord
  Else
    fWait.ProgressBar1.Max = 1
    
    MsgBox "Aucun article trouvé pour la période spécifiée", vbInformation
  End If
  
  ' ouvre la table assure en ecriture
  Set rs2 = destDB.OpenRecordset("P3IUser.Assure", dbOpenTable)
    
  ' importe les CE
  codeRetour = DoIntermediate(SrcDB, rs2, typeGarCE, "CE", rqCE, fWait, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax)
  
  ' ne valide que si il n'y a pas d'erreur
  If codeRetour = False Then
    ' annule les changements
    DBEngine.Rollback
  Else
    ' importe les CI
    codeRetour = DoIntermediate(SrcDB, rs2, typeGarCI, "CI", rqCI, fWait, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax)
    
    ' ne valide que si il n'y a pas d'erreur
    If codeRetour = False Then
      ' annule les changements
      DBEngine.Rollback
    Else
      ' importe les CR
      codeRetour = DoIntermediate(SrcDB, rs2, typeGarCR, "CR", rqCR, fWait, CleGroupe, NumPeriode, DateDebut, DateFin, nbJourMax)
      
      ' valide si pas d'erreur
      If codeRetour = True Then
        ' valide les changements
        DBEngine.CommitTrans
      Else
        ' annule les changements
        DBEngine.Rollback
      End If
    End If
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

