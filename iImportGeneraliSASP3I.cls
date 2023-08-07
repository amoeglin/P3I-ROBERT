VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iImportGeneraliSASP3I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C90E97302F3"
Option Explicit
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Feuille excel verrouillée par mot de passe : 'import'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements iP3IGeneraliImport


'Variables required for new Statutaire function
'##ModelId=5C90E97400D6
Private NumPeriodeStat As Long
'##ModelId=5C90E97400F6
Private NumPeriodeNonStat As Long
'##ModelId=5C90E9740118
Private PathSexFileExcel As String
'##ModelId=5C90E9740138
Private CategoryCodeSTAT As String
'##ModelId=5C90E9740172
Private SexAllMale As Boolean
'##ModelId=5C90E9740195
Private TwoLotImport As Boolean

'##ModelId=5C90E97401C5
Private autoMode As Boolean

'##ModelId=5C90E9740200
Private m_Logger As clsLogger

Private Enum eTypeImport
  eImportComplet = 1
  eImportDonneesSeules = 2
  eImportTablesParametresSeules = 3
End Enum

Private Enum eTypeDelaiInactivite
  eDatePaiement = 1
  eDateFinPeriodePaiement = 2
End Enum

Private Enum eTypeCalculAnnualisation
  eEnsemblePaiement = 1
  eDernierPaiement = 2
End Enum

' nb de rejet lors de l'import
'##ModelId=5C90E9740206
Private NbRejet As Long
'##ModelId=5C90E974023A
Private dateImport As Date
'##ModelId=5C90E9740272
Private dateArrete As Date
'##ModelId=5C90E97402A0
Private avgPSAP_Regime6 As Double
'##ModelId=5C90E97402E3
Private typeImport As eTypeImport
'##ModelId=5C90E97402F0
Private typeDelaiInactivite As eTypeDelaiInactivite
'##ModelId=5C90E97402FE
Private typeCalculAnnualisation As eTypeCalculAnnualisation

'##ModelId=5C90E9740316
Private rsCatR9 As ADODB.Recordset
'##ModelId=5C90E9740322
Private rsCatInval As ADODB.Recordset
'##ModelId=5C90E974032F
Private rsNCA As ADODB.Recordset
'Private rsTBQREGA As ADODB.Recordset, rsREGA01 As ADODB.Recordset, rsTaux As
'ADODB.Recordset
'##ModelId=5C90E9740340
Private rsCodesCat As ADODB.Recordset
'##ModelId=5C90E9740358
Private rsCodeCatInv As ADODB.Recordset

'Eclatement Statutaire 2 Lot
'##ModelId=5C90E9740366
Private rsSexe As DAO.Recordset

'##ModelId=5C90E974036D
Private nbCatNotFound As Long
'##ModelId=5C90E9740391
Private lstCatNotFound() As String

'##ModelId=5C90E97403B4
Private nbCodesCatNotFound As Long
'##ModelId=5C90E9750001
Private lstCodesCatNotFound() As String

'##ModelId=5C90E9750025
Private nbCodeCatInvNotFound As Long
'##ModelId=5C90E975004A
Private lstCodeCatInvNotFound() As String

'##ModelId=5C90E9750076
Private m_dataHelper As DataHelper

'##ModelId=5C90E975007E
Private m_TypePeriode As Integer

'##ModelId=5C90E97500A4
Private m_AssureP3IProvColl_SQLText As String
'


'##ModelId=5C90E97500D9
Public Function Arrondi(Valeur As Double, Nbdecimale As Integer) As Double ' fonction arrondi
  Nbdecimale = Abs(Nbdecimale)
  Arrondi = Fix((Valeur * (10 ^ Nbdecimale) + 0.5)) / (10 ^ Nbdecimale)
End Function


'##ModelId=5C90E975012D
Public Function BuildDateLimit(debut As String, Fin As String) As String
  Dim Limit As String
  
  Limit = "BETWEEN #" & Format(CDate(debut), "mm/dd/yyyy") & "# AND #" & Format(CDate(Fin), "mm/dd/yyyy") & "#"
  
  BuildDateLimit = Limit
End Function


'Private Function FindCodeProvision(Lib_Court_GE As String) As Integer
'  rsTBQREGA.MoveFirst
'  rsTBQREGA.Find "Lib_Court_GE='" & Lib_Court_GE & "'"
'
'  If Not rsTBQREGA.EOF Then
'
'    rsREGA01.MoveFirst
'    rsREGA01.Find "Code_GE='" & rsTBQREGA.Fields("Code_GE") & "'"
'
'    If Not rsREGA01.EOF Then
'
'      FindCodeProvision = rsREGA01.Fields("Code_PROV")
'
'    Else
'
'      m_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'REGA01' : " & rsTBQREGA.Fields("Code_GE")
'      FindCodeProvision = -1
'
'    End If
'
'  Else
'
'    m_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'TBQREGA' : " & Lib_Court_GE
'    FindCodeProvision = -1
'
'  End If
'End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'##ModelId=5C90E975016B
Private Function FindNumParamCalcul(Code_CIE As Long, Code_APP As Long, Code_Cat_Contrat As String) As Integer
  
  rsCodesCat.Filter = adFilterNone
  rsCodesCat.Filter = "Code_CIE=" & Code_CIE & " AND Code_APP=" & Code_APP & " AND Code_Cat_Contrat='" & Trim(Code_Cat_Contrat) & "'"

  If Not rsCodesCat.EOF Then

    FindNumParamCalcul = rsCodesCat.Fields("NumParamCalcul")

  Else
    Dim msg As String, iCat As Integer, bFound As Boolean
    
    msg = Code_CIE & " - " & Code_APP & " - " & Code_Cat_Contrat
    
    ' recherche si la categorie a déjà été rencontré
    bFound = False
    For iCat = 0 To nbCodesCatNotFound - 1
      If lstCodesCatNotFound(iCat) = msg Then
        bFound = True
        Exit For
      End If
    Next iCat
    
    If bFound = False Then
      ReDim Preserve lstCodesCatNotFound(nbCodesCatNotFound + 1) As String
      lstCodesCatNotFound(nbCodesCatNotFound) = Code_CIE & " - " & Code_APP & " - " & Code_Cat_Contrat
      nbCodesCatNotFound = nbCodesCatNotFound + 1
    End If
    
    FindNumParamCalcul = cdCategorieInvalParDefaut
    'm_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'CodesCat' : Code_CIE=" & Code_CIE & " AND Code_APP=" & Code_APP & " AND Code_Cat_Contrat='" & Code_Cat_Contrat & "'"

  End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recherche la catégorie inval
'
'##ModelId=5C90E97501DA
Private Function FindCategorieInval(CDCHOIXPREST As String) As Integer
  
  rsCodeCatInv.Filter = adFilterNone
'  rsCodeCatInv.Filter = "RTRIM(CDCHOIXPREST)='" & Trim(CDCHOIXPREST) & "'"
  rsCodeCatInv.Filter = "CDCHOIXPREST='" & Trim(CDCHOIXPREST) & "'"

  If Not rsCodeCatInv.EOF Then

    FindCategorieInval = rsCodeCatInv.Fields("CategorieInval")

  Else
    Dim iCat As Integer, bFound As Boolean
    
    ' recherche si la categorie a déjà été rencontré
    bFound = False
    For iCat = 0 To nbCodeCatInvNotFound - 1
      If lstCodeCatInvNotFound(iCat) = CDCHOIXPREST Then
        bFound = True
        Exit For
      End If
    Next iCat
    
    If bFound = False Then
      ReDim Preserve lstCodeCatInvNotFound(nbCodeCatInvNotFound + 1) As String
      lstCodeCatInvNotFound(nbCodeCatInvNotFound) = CDCHOIXPREST
      nbCodeCatInvNotFound = nbCodeCatInvNotFound + 1
    End If
    
    FindCategorieInval = cdCategorieInvalParDefaut

  End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'##ModelId=5C90E975021E
Private Function IsInPeriode(rs As ADODB.Recordset, DateDebut As Date, DateFin As Date, nbJourDC As Long) As Boolean
  Dim dDate As Variant, dSurv As Variant
  
  IsInPeriode = False
  
  If IsNull(rs.Fields("Regime")) Then
    Exit Function
  End If
  
  dDate = ReadDate(rs.Fields("DatePaiement"))
  dSurv = ReadDate(rs.Fields("Survenance"))
  
  If IsNull(dDate) Then
    
    ' Deces
    If rs.Fields("Regime") = cdGarDeces_Import Then
      If (dSurv + nbJourDC) >= DateDebut And (dSurv + nbJourDC) <= DateFin Then
        IsInPeriode = True
      End If
    End If
    
    ' Arret de travail
    If rs.Fields("Regime") = cdGarIncapInval_Import Then
      If dSurv >= DateDebut And dSurv <= DateFin Then
        IsInPeriode = True
      End If
    End If
    
  Else
    
    If dDate >= DateDebut And dDate <= DateFin Then
      IsInPeriode = True
    End If
  
  End If

  
  ' 13/05/2009 : les décès doivent être pris en compte
  If rs.Fields("Regime") = cdGarDeces_Import Then
    IsInPeriode = True
  End If

End Function


'##ModelId=5C90E97502C0
Private Function PassageCATR9(categorie As String, NCA As String, ByRef bFoundCATR9 As Boolean) As Boolean
  
  PassageCATR9 = False
  bFoundCATR9 = False
  
  Dim nbAnnualForcee As Integer
      
  m_dataHelper.Multi_Find rsCatR9, "Categorie='" & categorie & "'"
  
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

'##ModelId=5C90E9750338
Private Function IsStatutaire(catCode As String) As Boolean

  Dim arrCatCodeStat() As String
  Dim foundCodeStat As Boolean
  Dim i As Integer
  Dim codeStat As String
  
  arrCatCodeStat = Split(CategoryCodeSTAT, ",")
  
  foundCodeStat = False
  For i = 0 To UBound(arrCatCodeStat)
    codeStat = Trim$(arrCatCodeStat(i))
    codeStat = Replace(codeStat, ",", "")
    codeStat = Replace(codeStat, "'", "")
    
    If codeStat = Trim$(catCode) Then
      foundCodeStat = True
    End If
  Next
  
  IsStatutaire = foundCodeStat
    
End Function

'Modif 11/7/2018
'Function GetSettingIni(appname As String, Section As String, key As String, default As String, Optional fLocalMachine) As Variant
'  GetSettingIni = sReadIniFile(Section, key, default, 1024, sFichierIni)
'End Function

'Modif 11/7/2018
'##ModelId=5C90E9750387
Function sReadIniFile(sSection As String, sEntree As String, sDefaut, lLgMax As Long, sFichier As String) As String
  Dim lRes As Long
  Dim sValeur As String
  
  'Prépare la chaine de retour (longueur lLgMax +1, car caractère de fin de chaine)
  sValeur = Space(lLgMax + 1)
  
  'Lecture dans fichier en utilisant fonction API Win32 (lRes contient le nombre de caractères lus)
  lRes = GetPrivateProfileString(sSection, sEntree, sDefaut, sValeur, Len(sValeur), sFichier)
  
  'Retour de la valeur trouvée (ou de la valeur par défaut, sinon), uniquement les caractères "utiles"
  sReadIniFile = Left(sValeur, lRes)
End Function



' ATTENTION : ne fonctionne pas avec les membres d'une classe !!
'Private Sub AffecteSiNonNullEtNonZero(ByRef fd As Variant, fs As ADODB.Field)
'  fd = Null
'  If Not IsNull(fs) Then
'    If fs <> 0 Then
'      fd = fs
'    End If
'  End If
'End Sub
'
'Private Sub AffecteSiNonNull(ByRef fd As Variant, fs As ADODB.Field)
'  fd = Null
'  If Not IsNull(fs) Then
'    fd = fs
'  End If
'End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' on regroupe les paiements par sinistre. la date de debut est celle du
'premier paiement
' et la date de fin est celle du dernier paiement
'
'##ModelId=5C90E9760049
Private Sub AddNewRecord(destDB As DataAccess, rs As ADODB.Recordset, CleGroupe As Long, NumPeriode As Long, fWait As Form, AgeRetraite As Byte, DateDebut As Date, DateFin As Date, nbJourMax As Long, nbJourDC As Long, bFromInfocentre As Boolean, rsAssureP3IProvColl As ADODB.Recordset)
  
  Dim premierPaiement As Variant, dernierPaiement As Variant, poDebut As Variant, poFin As Variant
  Dim poDebutTotal As Variant, poFinTotal As Variant
  Dim dateNaissance As Date, dDate As Variant ', regroupement As String
  Dim tmpMontant As Double, montant As Double, montantTotal As Double, montantAn As Double, duree As Double, psap As Double
  Dim MontantBase As Double, MontantRevalo As Double, MontantMajo As Double
  Dim nbJPer As Long, nbJHors As Long, nbJIndemn As Integer, nbJAnnualisation As Integer, montantAnnualisation As Double
  Dim CodeProvision As Integer
    
  Dim curSte As Long, n As Integer, curBaseRevalo As Integer
  Dim curSS As String, curGE As String, curNCA As String, curRegime As Long, curCat As String
  Dim curArret As Date, dateReprise As Variant
  Dim varBookmark As Variant
  Dim inPeriode As Boolean, bInit As Boolean
  
  'Modif 11/7/2018
  Dim DateLimiteProvision As Date
  Dim CompanyName As String
  Dim SectionName As String
  Dim ageLimitViagere As Integer
  
  Dim bPassageCATR9 As Boolean, bFoundCATR9 As Boolean
  Dim FieldName As String
  
  Dim numParamCalcul As Long, bPortefeuilleSalarie As Boolean
  
  On Error GoTo GestionErreur
  
  CompanyName = "Moeglin"
  SectionName = "P3I"
          
' ...pour debug...
'  FieldName = "NomAssure"
'  If Not IsNull(rs.Fields("NomAssure")) Then
'    If InStr(rs.Fields("NomAssure"), "ASSURE4083") <> 0 Then
'      Stop
'    End If
'  End If
' ...pour debug...


  
  
  ' n° de jeu de parametres de calcul
  numParamCalcul = FindNumParamCalcul(rs.Fields("CDCOMPAGNIE"), rs.Fields("CDAPPLI"), Trim(rs.Fields("Code_Cat_Contrat")))
  
  ' Portefeuille Salaries ?
  bPortefeuilleSalarie = destDB.CreateHelper.GetParameterAsDouble("SELECT PEPortefeuilleSalaries FROM ParamCalcul WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode & " AND PENUMPARAMCALCUL=" & numParamCalcul)


  ' dans la période ?
  inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC)
 
  
  ' boucle sur un meme NumSociete+NumeroSS+Regime+Categorie+NCA+Survenance pour obtenir l'annualisation
  duree = 0
  montant = 0
  montantAn = 0
  montantTotal = 0
  nbJAnnualisation = 0
  montantAnnualisation = 0
  bInit = False
  
  ' raz des dates
  premierPaiement = Null
  dernierPaiement = Null
  poDebut = Null
  poFin = Null
  poDebutTotal = Null
  poFinTotal = Null
  curBaseRevalo = 0
  
  FieldName = "Regime"
  If IsNull(rs.Fields("Regime")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Regime invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curRegime = rs.Fields("Regime")
  End If
  
  
  
' DEBUG DEBUG
' DEBUG DEBUG
' DEBUG DEBUG
  
'  If rs.Fields("RECNO") = 7754 Then Stop

' DEBUG DEBUG
' DEBUG DEBUG
' DEBUG DEBUG




  'If curRegime = cdGarDeces_Import Then Stop
  
  
  
  If curRegime = cdGarRente_Import Then
    ' les dates correspondent à une définition de palier
    inPeriode = True
  End If
  
  
  ' If inPeriode Then PHM 9/1/04 : on tient toujours compte des deces
  If inPeriode Or curRegime = cdGarDeces_Import Then
    
    If curRegime = cdGarRente_Import Then
      premierPaiement = DateFin
    Else
      FieldName = "DatePaiement"
      dDate = ReadDate(rs.Fields("DatePaiement"))
      If IsNull(dDate) Then
        premierPaiement = Null
      Else
        premierPaiement = dDate
      End If
    End If
    
    FieldName = "PeriodeDu"
    dDate = ReadDate(rs.Fields("PeriodeDu"))
    If IsNull(dDate) Then
      poDebut = Null
    Else
      poDebut = dDate
    End If
    bInit = True
  
  Else
    inPeriode = False ' test debug
  End If
  
  dDate = ReadDate(rs.Fields("PeriodeDu"))
  If IsNull(dDate) Then
    poDebutTotal = Null
  Else
    poDebutTotal = dDate
  End If
  
  ' init des variables
  FieldName = "NumSociete"
  If IsNull(rs.Fields("NumSociete")) Then
    Call m_Logger.EcritTraceDansLog("REJET : NumSociété invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curSte = rs.Fields("NumSociete")
  End If
  
  FieldName = "Sinistre"
  If IsNull(rs.Fields("Sinistre")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Sinistre invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curSS = Trim(rs.Fields("Sinistre"))
  End If
    
  FieldName = "GE"
  If IsNull(rs.Fields("GE")) Then
    Call m_Logger.EcritTraceDansLog("REJET : GE invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curGE = rs.Fields("GE")
  End If
    
  FieldName = "Code_Cat_Contrat"
  If IsNull(rs.Fields("Code_Cat_Contrat")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Code_Cat_Contrat invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curCat = Trim(rs.Fields("Code_Cat_Contrat"))
    If curCat = vbNullString Then
      Call m_Logger.EcritTraceDansLog("REJET : Code_Cat_Contrat invalide à la ligne: " & rs.Fields("RECNO"))
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If
  
  FieldName = "Convention"
  If IsNull(rs.Fields("Convention")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Convention invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curNCA = Trim(rs.Fields("Convention"))
  End If
  
  FieldName = "Survenance"
  dDate = ReadDate(rs.Fields("Survenance"))
  If IsNull(dDate) Then
    Call m_Logger.EcritTraceDansLog("REJET : Survenance invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curArret = dDate
  End If
  
  FieldName = "BaseRevalo"
  If IsNull(rs.Fields("BaseRevalo")) Then
    'Call m_Logger.EcritTraceDansLog("REJET : BaseRevalo invalide à la ligne: " & rs.Fields("RECNO"))
    'NbRejet = NbRejet + 1
    'Exit Sub
    
    ' PhM 30/10/2008 : pour BANQUE, BaseRevalo est NULL
    curBaseRevalo = 0
    
  Else
    curBaseRevalo = IIf(rs.Fields("BaseRevalo") = "1", 1, 0)
  End If
  
  dateReprise = Null
  'regroupement = ""
  
  ' PHM 05/02/2004
  psap = 0
  ' PHM 05/02/2004
  
  
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' boucle sur le meme sinistre
  '
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Do
    ' test si fin de fichier
    If rs.EOF Then
      rs.bookmark = varBookmark
      Exit Do
    End If
      
    If rs.Fields("Regime") = cdGarRente_Import Then
      inPeriode = True
    Else
      inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC)
    End If
    
    
    ' test si toujours le meme assuré
    If curSte <> rs.Fields("NumSociete") Or curSS <> Trim(rs.Fields("Sinistre")) Or curGE <> rs.Fields("GE") _
       Or curRegime <> rs.Fields("Regime") Or curCat <> Trim(rs.Fields("Code_Cat_Contrat")) _
       Or curNCA <> Trim(rs.Fields("Convention")) Or curArret <> ReadDate(rs.Fields("Survenance")) _
       Or curBaseRevalo <> IIf(rs.Fields("BaseRevalo") = "1", 1, 0) Then
      rs.bookmark = varBookmark

      Exit Do
    End If
    
    
    ' sauvegarde de la ligne
    varBookmark = rs.bookmark
    
    If inPeriode And Not bInit Then
      
      If curRegime = cdGarRente_Import Then
        premierPaiement = DateFin
      Else
        FieldName = "DatePaiement"
        dDate = ReadDate(rs.Fields("DatePaiement"))
        If IsNull(dDate) Then
          premierPaiement = Null
        Else
          premierPaiement = dDate
        End If
      End If
      
      dDate = ReadDate(rs.Fields("PeriodeDu"))
      If IsNull(dDate) Then
        poDebut = Null
      Else
        poDebut = dDate
      End If
      
      bInit = True
    
    End If
    
    ' toujour le meme sinistre
    tmpMontant = 0
    FieldName = "MntRegle"
    If Not IsNull(rs.Fields("MntRegle")) Then
      tmpMontant = rs.Fields("MntRegle")
    End If
    
    ' cumul des prestations payées presente dans le fichier
    montantTotal = montantTotal + tmpMontant
    
    'If inPeriode Then PHM 9/1/04 : on tient toujours compte des deces
    If inPeriode Then
      nbJPer = 0
      
      If curRegime = cdGarRente_Import Then
        nbJPer = 365
      Else
        FieldName = "PeriodeAu"
'        If Not IsNull(rs.Fields("PeriodeDu")) And Not IsNull(rs.Fields("PeriodeAu")) _
'          And IsDate(rs.Fields("PeriodeDu")) And IsDate(rs.Fields("PeriodeAu")) Then
        If IsDate(ReadDate(rs.Fields("PeriodeDu"))) And IsDate(ReadDate(rs.Fields("PeriodeAu"))) Then
          If ReadDate(rs.Fields("PeriodeDu")) > ReadDate(rs.Fields("PeriodeAu")) Then
            nbJPer = 0
            Call m_Logger.EcritTraceDansLog("!!!!!! PeriodeDu > PeriodeAu à la ligne: " & rs.Fields("RECNO"))
          Else
            nbJPer = DateDiff("d", ReadDate(rs.Fields("PeriodeDu")), ReadDate(rs.Fields("PeriodeAu"))) + 1
          End If
        End If
      End If
            
            
      If nbJPer > 32767 Or nbJPer < -32767 Then
            Call m_Logger.EcritTraceDansLog("!!!!!! Période Incohérente PeriodeDu " & rs.Fields("PeriodeDu") & " et PeriodeAu " & rs.Fields("PeriodeAu") & " soit une durée= " & nbJPer & " jours ! à la ligne: " & rs.Fields("RECNO"))
            nbJPer = 0
      End If
            
            
      If tmpMontant < 0 Then
        duree = duree - nbJPer
      Else
        duree = duree + nbJPer
      End If
      
      ' prestation prise en compte
      montantAn = montantAn + tmpMontant
    
      ' cumul prestation payées de la periode
      montant = montant + tmpMontant
      
      ' annualisation
      If typeCalculAnnualisation = eEnsemblePaiement Or curRegime = cdGarRente_Import Then
        nbJAnnualisation = nbJAnnualisation + nbJPer
        montantAnnualisation = montantAnnualisation + tmpMontant
      Else
        nbJAnnualisation = nbJPer
        montantAnnualisation = tmpMontant
      End If
      
      ' Montant Base
      tmpMontant = 0
      FieldName = "MntBase"
      If Not IsNull(rs.Fields("MntBase")) Then
        tmpMontant = rs.Fields("MntBase")
      End If
      MontantBase = MontantBase + tmpMontant
      
      ' Montant Revalo
      tmpMontant = 0
      FieldName = "MntRevalo"
      If Not IsNull(rs.Fields("MntRevalo")) Then
        tmpMontant = rs.Fields("MntRevalo")
      End If
      MontantRevalo = MontantRevalo + tmpMontant
      
      ' Montant Majore
      tmpMontant = 0
      FieldName = "MntMajore"
      If Not IsNull(rs.Fields("MntMajore")) Then
        tmpMontant = rs.Fields("MntMajore")
      End If
      MontantMajo = MontantMajo + tmpMontant
      
      
      ' l'annualisation est deja calculée
      montantAnnualisation = MontantBase + MontantMajo + MontantRevalo
      
      
      
'      ' Nb Jours Indemnisés
'      FieldName = "NbJIndemnises"
'      If Not IsNull(rs.Fields("NbJIndemnises")) Then
'        If rs.Fields("NbJIndemnises") > 32767 Then
'          Call m_Logger.EcritTraceDansLog("Nb jours indemnisés invalide à la ligne: " & rs.Fields("RECNO"))
'        Else
'          nbJIndemn = nbJIndemn + rs.Fields("NbJIndemnises")
'        End If
'      End If
      
      ' Nb Jours Indemnisés = DATEDIFF(day, P.DTDEBPER, P.DTFINPER) AS NbJIndemnises
      FieldName = "NbJIndemnises"
      If Not IsNull(ReadDate(rs.Fields("PeriodeDu"))) And Not IsNull(ReadDate(rs.Fields("PeriodeAu"))) Then
        Dim nbJ As Long
        
        nbJ = DateDiff("d", ReadDate(rs.Fields("PeriodeDu")), ReadDate(rs.Fields("PeriodeAu")))
        If nbJ > 32767 Then
          Call m_Logger.EcritTraceDansLog("Nb jours indemnisés invalide à la ligne: " & rs.Fields("RECNO"))
        Else
          nbJIndemn = nbJIndemn + nbJ
        End If
      Else
        If curRegime <> cdGarDeces_Import And curRegime <> cdGarRente_Import Then
          Call m_Logger.EcritTraceDansLog("DTDEBPER ou DTFINPER invalide à la ligne: " & rs.Fields("RECNO"))
        Else
          nbJ = 1
          nbJIndemn = nbJIndemn + nbJ
        End If
      End If

      ' date de reprise
'      FieldName = "DateReprise"
'      If Not IsNull(rs.Fields("DateReprise")) Then
'        dateReprise = rs.Fields("DateReprise")
'      End If
      
      ' stocke les infos
      If curRegime = cdGarRente_Import Then
        dernierPaiement = DateFin
      Else
        FieldName = "DatePaiement"
        dDate = ReadDate(rs.Fields("DatePaiement"))
        If IsNull(dDate) Then
          dernierPaiement = Null
        Else
          dernierPaiement = dDate
        End If
      End If
      
      FieldName = "PeriodeAu"
      dDate = ReadDate(rs.Fields("PeriodeAu"))
      If IsNull(dDate) Then
        poFin = Null
      Else
        poFin = dDate
      End If

' PHM 01/03/2004
    Else
      ' pour les régimes 6, les prestas payées apres la date de fin sont comptabilisés comme PSAP
      If curRegime = cdGarDeces_Import Then
        Dim datePaiement As Date
        
        dDate = ReadDate(rs.Fields("DatePaiement"))
        If IsNull(dDate) Then
          datePaiement = ReadDate(rs.Fields("Survenance")) + nbJourDC
        Else
          datePaiement = dDate
        End If
        
        If datePaiement > DateFin Then
          ' ATTENTION : vérifier que tmpMontant contienne toujours la prestation
          psap = psap + tmpMontant
        End If
      End If
' PHM 01/03/2004
    End If

    If inPeriode Or curRegime = cdGarDeces_Import Then
      
      If curRegime = cdGarDeces_Import Then
        ' stocke les infos
        FieldName = "DatePaiement"
        dDate = ReadDate(rs.Fields("DatePaiement"))
        If IsNull(dDate) Then
          dernierPaiement = Null
        Else
          dernierPaiement = dDate
        End If
        
        FieldName = "PeriodeAu"
        dDate = ReadDate(rs.Fields("PeriodeAu"))
        If IsNull(dDate) Then
          poFin = Null
        Else
          poFin = dDate
        End If
      End If
      
      ' cumul des psap, tous les PSAP pour les décès
      tmpMontant = 0
      FieldName = "PSAP"
'      If Not IsNull(rs.Fields("PSAP")) Then
'        tmpMontant = rs.Fields("PSAP")
'      End If
      psap = psap + tmpMontant
    End If

'*** OPTION : calcul juste de l'annualisation sur la periode en cours
'    ' avant la periode
'    If rs.Fields("PeriodeDu") < DateDebut Then
'      nbJHors = DateDiff("d", rs.Fields("PeriodeDu"), DateDebut) + 1
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
'    If rs.Fields("PeriodeAu") > DateFin Then
'      nbJHors = DateDiff("d", DateFin, rs.Fields("PeriodeAu")) + 1
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
    If Not autoMode Then
      If (rs.AbsolutePosition Mod 9) = 0 Then
        fWait.ProgressBar1.Value = rs.AbsolutePosition
        fWait.Label1(0).Caption = "Phase 2 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max & " - " & NbRejet & " rejets"
        fWait.Refresh
        DoEvents
      End If
    End If
        
    '!!! PHM 22/02/2010 : L'import ne somme plus lignes pour en déduire le montant annualisé !!!
    Exit Do
    
    
    ' prochain paiement
    rs.MoveNext
  Loop
 
'  ' elimine les ' ' du n° de sécu
'  Do
'    n = InStr(curSS, " ")
'    If n <> 0 Then
'      curSS = Left(curSS, n - 1) & Mid(curSS, n + 1)
'    End If
'  Loop Until n = 0
  
  ' elimine les ' ' du NCA
  Do
    n = InStr(curNCA, " ")
    If n <> 0 Then
      curNCA = Left(curNCA, n - 1) & Mid(curNCA, n + 1)
    End If
  Loop Until n = 0
  
  'If inPeriode then PHM 9/1/04 : on tient toujours compte des deces
  If inPeriode Or curRegime = cdGarDeces_Import Then
    ' stocke les infos
    If curRegime = cdGarRente_Import Then
      dernierPaiement = DateFin
    Else
      FieldName = "DatePaiement"
      dDate = ReadDate(rs.Fields("DatePaiement"))
      If Not IsNull(dDate) Then
        dernierPaiement = dDate
      End If
    End If
    
    FieldName = "PeriodeAu"
    dDate = ReadDate(rs.Fields("PeriodeAu"))
    If IsNull(dDate) Then
      poFin = Null
    Else
      poFin = dDate
    End If
  End If
  
  FieldName = "PeriodeAu"
  dDate = ReadDate(rs.Fields("PeriodeAu"))
  If IsNull(dDate) Then
    poFinTotal = Null
  Else
    poFinTotal = dDate
  End If

  
  
  ' store data
  Dim theAssure As clsAssure
  Set theAssure = New clsAssure
  
  
  theAssure.m_NUENRP3I = rs.Fields("RECNO")
  
  
  
  ' calcul de la date de paiement pour le regime 6 : il n'y a qu'UN seul paiement pour ces regimes
  If curRegime = cdGarDeces_Import And (IsNull(premierPaiement) Or IsNull(dernierPaiement)) Then
    FieldName = "Survenance"
    premierPaiement = ReadDate(rs.Fields("Survenance")) + nbJourDC
    dernierPaiement = premierPaiement
    theAssure.m_PODATEPAIEMENTESTIMEE = True
  End If
  
  theAssure.m_POGPECLE = CleGroupe
  theAssure.m_POPERCLE = NumPeriode

  '### STATUTAIRE
  If TwoLotImport Then
    Dim isStat As Boolean
    Dim catCode As String
    'Dim arrCatCodeStat() As String
    'Dim i As Integer
    'Dim foundCodeStat As Boolean
    'Dim codeStat As String
    
    'get Category Code (CDPRODUIT) for this Assure
    catCode = rs.Fields("CDPRODUIT")
    
'    arrCatCodeStat = Split(CategoryCodeSTAT, ",")
'
'    foundCodeStat = False
'    For i = 0 To UBound(arrCatCodeStat)
'      codeStat = Trim$(arrCatCodeStat(i))
'      codeStat = Replace(codeStat, ",", "")
'      codeStat = Replace(codeStat, "'", "")
'      'If InStr(arrCatCodeStat(i), Trim$(catCode)) > 0 Then
'      If codeStat = Trim$(catCode) Then
'        foundCodeStat = True
'      End If
'    Next
    
    'check the IsStat field
    'If catCode = CategoryCodeSTAT Then
    'If InStr(CategoryCodeSTAT, Trim$(catCode)) > 0 Then
    'If foundCodeStat Then
    
    If IsStatutaire(catCode) Then
      theAssure.m_POPERCLE = NumPeriodeStat
    Else
      theAssure.m_POPERCLE = NumPeriodeNonStat
    End If
    ' modif am 17/10/2017
    ' Rechargement du Type de période en fonction de la période active chargée dans theAssure.m_POPERCLE
    m_TypePeriode = destDB.CreateHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & theAssure.m_POPERCLE)
    ' Rechargement du Portefeuille Salaries  en fonction de la période active chargée dans theAssure.m_POPERCLE
    bPortefeuilleSalarie = destDB.CreateHelper.GetParameterAsDouble("SELECT PEPortefeuilleSalaries FROM ParamCalcul WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & theAssure.m_POPERCLE & " AND PENUMPARAMCALCUL=" & numParamCalcul)

    End If
    

  
  FieldName = "NumSociete"
  If Not IsNull(rs.Fields("NumSociete")) Then
    theAssure.m_POSTECLE = rs.Fields("NumSociete")
  Else
    Call m_Logger.EcritTraceDansLog("REJET : Société invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  If curSS <> "" Then
    theAssure.m_PONUMCLE = curSS
  Else
    Call m_Logger.EcritTraceDansLog("REJET : Sinistre invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' garantie
  If curRegime <> 0 Then
    Select Case curRegime
      Case cdGarDeces_Import, cdGarInaptitudeConduite_Import, cdGarIncapInval_Import, cdGarRente_Import
        ' OK
      Case Else
        Call m_Logger.EcritTraceDansLog("REJET : Regime (=" & curRegime & ") invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
    End Select
    
    theAssure.m_POGARCLE = curRegime + 50
  Else
    Call m_Logger.EcritTraceDansLog("REJET : Regime (=Null)invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  If curCat <> vbNullString Then
    theAssure.m_POCATEGORIE = curCat
  Else
    theAssure.m_POCATEGORIE = Null
  End If
  
  ' Nom
  Select Case curRegime
    
    Case cdGarDeces_Import
      FieldName = "NomAssure"
      If Not IsNull(rs.Fields("NomAssure")) Then
        theAssure.m_PONOM = Trim(rs.Fields("NomAssure"))
      Else
        FieldName = "LBRENTIER"
        If Not IsNull(rs.Fields("LBRENTIER")) Then
          theAssure.m_PONOM = Trim(rs.Fields("LBRENTIER"))
        Else
          Call m_Logger.EcritTraceDansLog("REJET : Nom (LBRENTIER et LBASSURE) invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        End If
      End If
    
      FieldName = "DTNAISSREN"
      dDate = ReadDate(rs.Fields("DTNAISSREN"))
      If IsNull(dDate) Then
        FieldName = "DateNaiss"
        dDate = ReadDate(rs.Fields("DateNaiss"))
        If IsNull(dDate) Then
          Call m_Logger.EcritTraceDansLog("REJET : Date Naissance (DTNAISSASS et DTNAISSREN) invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          dateNaissance = dDate
        End If
      Else
        dateNaissance = dDate
      End If
    
    
    Case cdGarRente_Import
      FieldName = "LBRENTIER"
      If Not IsNull(rs.Fields("LBRENTIER")) Then
        theAssure.m_PONOM = Trim(rs.Fields("LBRENTIER"))
      Else
        Call m_Logger.EcritTraceDansLog("REJET : Nom (LBRENTIER) invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      End If
      
      FieldName = "DTNAISSREN"
      dDate = ReadDate(rs.Fields("DTNAISSREN"))
      If IsNull(dDate) Then
        Call m_Logger.EcritTraceDansLog("REJET : Date Naissance (DTNAISSREN) invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      Else
        dateNaissance = dDate
      End If

    
    Case Else
  
      FieldName = "NomAssure"
      If Not IsNull(rs.Fields("NomAssure")) Then
        theAssure.m_PONOM = Trim(rs.Fields("NomAssure"))
      Else
        Call m_Logger.EcritTraceDansLog("REJET : Nom (LBASSURE) invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      End If
      
      FieldName = "DateNaiss"
      dDate = ReadDate(rs.Fields("DateNaiss"))
      If IsNull(dDate) Then
        Call m_Logger.EcritTraceDansLog("REJET : Date Naissance (DTNAISSASS) invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      Else
        dateNaissance = dDate
      End If
  End Select
  
  ' NCA
  If curNCA <> "" Then
    theAssure.m_POCONVENTION = curNCA
  Else
    Call m_Logger.EcritTraceDansLog("REJET : NCA invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  ' date du sinistre
  If Not IsNull(curArret) Then
    theAssure.m_POEFFET = curArret
    theAssure.m_POARRET = curArret
  Else
    Call m_Logger.EcritTraceDansLog("REJET : Survenance invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
    
  
  FieldName = "BaseRevalo"
  theAssure.m_POBaseRevalo = curBaseRevalo
     
    
  If curArret <= dateNaissance Then
    Call m_Logger.EcritTraceDansLog("Attention : Survenance <= Date Naissance à la ligne: " & rs.Fields("RECNO"))
  End If

  theAssure.m_PONAIS = dateNaissance
  If curRegime = cdGarIncapInval_Import And (DateDiff("yyyy", dateNaissance, DateFin) < 18 Or DateDiff("yyyy", dateNaissance, DateFin) > 65) Then
    Call m_Logger.EcritTraceDansLog("Attention : Age < 18 ans ou Age > 65 ans à la ligne: " & rs.Fields("RECNO"))
  End If
  
  
  'Modif 11/7/2018 - age limit force a 120 - recalculé correctement dans calculer
  ageLimitViagere = 120  'm_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeLimitViagere", "120"))
  CodeProvision = rs.Fields("CodeProv")
  
  ' terme = age retraite
  ' sauf pour la rente de conjoint où terme=null => rente viagere
  If curRegime <> cdGarRente_Import Then
    theAssure.m_POTERME = DateSerial(Year(dateNaissance) + AgeRetraite, Month(dateNaissance), Day(dateNaissance))
  End If
  
  
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    DateLimiteProvision = DateSerial(Year(dateNaissance) + ageLimitViagere, Month(dateNaissance), Day(dateNaissance))
    theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
  
  
  
  
  ' debut garantie
  FieldName = "EffetConv"
  dDate = ReadDate(rs.Fields("EffetConv"))
  If Not IsNull(dDate) Then
    theAssure.m_POEFFET = dDate
  End If
  
  ' Fin garantie
  FieldName = "FinGarantie"
  dDate = ReadDate(rs.Fields("FinGarantie"))
  If Not IsNull(dDate) Then
    theAssure.m_POTERME = dDate
  End If
  
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
      theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
  
  ' date de reprise
  FieldName = "Reprise=(null)"
  theAssure.m_POREPRISE = Null
  
  ' cause
  FieldName = "Cause"
  theAssure.m_POCAUSE = Trim(rs.Fields("LBRISQUE"))
  
  ' Franchise
  FieldName = "Franchise"
  If Not IsNull(rs.Fields("Franchise")) Then
    If IsNumeric(rs.Fields("Franchise")) Then
      theAssure.m_POTYPEF = rs.Fields("Franchise")
    End If
  End If
  
  ' periode comptable
  theAssure.m_PODEBUT = poDebut
  theAssure.m_POFIN = poFin
  
  ' periode comptable totale
  theAssure.m_PODEBUTTOTAL = poDebutTotal
  theAssure.m_POFINTOTAL = poFinTotal
  
  ' delai de carence
  FieldName = "DelaiCarence=0"
  theAssure.m_PODELAI = 0
  
  ' CSP
  'theAssure.m_POCSP = ""
  
  ' montant de la prestation
  theAssure.m_POPRESTATION = montant ' non recalculé
  
  If curRegime = cdGarRente_Import Then
    
    ' Rentes
    theAssure.m_POPRESTATION_AN = montantAnnualisation ' pour les rentes le montant est directement le montant annualisé
  
  ElseIf bFromInfocentre = True And curRegime = cdGarIncapInval_Import Then
      
      If montantAnnualisation = 0 Then
        ' Montant Base
        tmpMontant = 0
        FieldName = "MntBase"
        If Not IsNull(rs.Fields("MntBase")) Then
          tmpMontant = rs.Fields("MntBase")
        End If
        MontantBase = MontantBase + tmpMontant
        
        ' Montant Revalo
        tmpMontant = 0
        FieldName = "MntRevalo"
        If Not IsNull(rs.Fields("MntRevalo")) Then
          tmpMontant = rs.Fields("MntRevalo")
        End If
        MontantRevalo = MontantRevalo + tmpMontant
        
        ' Montant Majore
        tmpMontant = 0
        FieldName = "MntMajore"
        If Not IsNull(rs.Fields("MntMajore")) Then
          tmpMontant = rs.Fields("MntMajore")
        End If
        MontantMajo = MontantMajo + tmpMontant
      
        ' l'annualisation est deja calculée
        montantAnnualisation = MontantBase + MontantMajo + MontantRevalo
        
        If Not IsDate(poFin) Then
          ' Dernière période réglée = survenance si elle est nulle
          theAssure.m_POFIN = theAssure.m_POARRET
        End If
      
      End If
      
      If montantTotal = 0 Then
        ' montant regle
        tmpMontant = 0
        FieldName = "MntRegle"
        If Not IsNull(rs.Fields("MntRegle")) Then
          tmpMontant = rs.Fields("MntRegle")
        End If
        
        ' cumul des prestations payées presente dans le fichier
        montantTotal = tmpMontant
      End If
    
      theAssure.m_POPRESTATION_AN = montantAnnualisation ' le montant est directement le montant annualisé
  
  Else
    
    ' autre
    If nbJAnnualisation <= 0 And curRegime <> cdGarDeces_Import Then
      
      theAssure.m_POPRESTATION_AN = 0
      
      If curRegime <> cdGarDeces_Import Then
        If IsNull(poDebut) And IsNull(poFin) And IsNull(premierPaiement) And IsNull(dernierPaiement) And IsNull(ReadDate(rs.Fields("DebIndemn"))) Then
          ' sera traité plus tard...
          theAssure.m_POPRESTATION_AN = 0
        Else
          theAssure.m_POPRESTATION_AN = 0
          If nbJAnnualisation < 0 Then
            Call m_Logger.EcritTraceDansLog("   Annual.= 0 car nbJAnnualisation<=0 à la ligne: " & rs.Fields("RECNO") & " Régime " & curRegime)
          End If
        End If
      End If
    
    Else
      
      theAssure.m_POPRESTATION_AN = montantAnnualisation ' le montant est directement le montant annualisé
    
    End If
  
  End If
    
  If theAssure.m_POPRESTATION_AN < 0 Then
    Call m_Logger.EcritTraceDansLog("   Annual.= 0 car montant<0 à la ligne: " & rs.Fields("RECNO"))
    theAssure.m_POPRESTATION_AN = 0
  End If
  
  theAssure.m_POPRESTATIONTOTAL = montantTotal
  
'  ' rente education
'  If curRegime = 3 Then
'    ' annualisation pour une garantie de 100%
'    rsSalarie.Fields("POPRESTATION_AN") = Arrondi(rsSalarie.Fields("POPRESTATION_AN") _
'                                                   / (1 + FindTauxRenteEducation(rs, DateFin)) _
'                                                   , 2 _
'                                                 )
'  End If
  
  ' premier et dernier paiement
  theAssure.m_POPREMIER_PAIEMENT = premierPaiement
  theAssure.m_PODERNIERPAIEMENT = dernierPaiement
  
  ' anciennete superieur à 180j
  theAssure.m_PODOSSIERCLOS = False
  
  nbJPer = 0
  
  If bFromInfocentre = True Then
    theAssure.m_POIJ = Null
  Else
    FieldName = "IndemnitesJournalieres"
    'AffecteSiNonNull theAssure.m_POIJ, rs.Fields("IndemnitesJournalieres")
    theAssure.m_POIJ = rs.Fields("IndemnitesJournalieres")
  End If
  
  If IsNull(poDebut) And IsNull(poFin) And IsNull(premierPaiement) And IsNull(dernierPaiement) And IsNull(ReadDate(rs.Fields("DebIndemn"))) Then
    
    If bFromInfocentre = False Then
      
      If IsNull(theAssure.m_POIJ) Then
        theAssure.m_POPRESTATION_AN = 0
        Call m_Logger.EcritTraceDansLog("   Annual.= 0 car IndemnitesJournalieres non renseignées à la ligne: " & rs.Fields("RECNO"))
      Else
        theAssure.m_POPRESTATION_AN = 365 * theAssure.m_POIJ
        Call m_Logger.EcritTraceDansLog("   Annual.= 365 * IndemnitesJournalieres = " & theAssure.m_POPRESTATION_AN & " à la ligne: " & rs.Fields("RECNO"))
      End If
    
    End If
    
  Else
  
    FieldName = "PeriodeAu"
    dDate = ReadDate(rs.Fields("PeriodeAu"))
    If IsNull(dDate) Then
      If theAssure.m_POGARCLE <> cdGar56 And theAssure.m_POGARCLE <> cdGarRente Then
        Call m_Logger.EcritTraceDansLog("Fin de Période invalide à la ligne: " & rs.Fields("RECNO"))
      End If
    End If
  
  End If
  
  FieldName = "IndicCC"
  Dim bContentieux As Boolean, sIndicCC As String
  
  sIndicCC = vbNullString
  If Not IsNull(rs.Fields("IndicCC")) Then
    sIndicCC = Trim(rs.Fields("IndicCC"))
  End If
  
  If sIndicCC <> vbNullString Then
    bContentieux = IIf(UCase(sIndicCC) = "O" Or sIndicCC = "1", True, False)
  Else
    bContentieux = False
  End If
  
  Dim bKeep As Boolean
  
  If typeDelaiInactivite = eDateFinPeriodePaiement Then
    
    dDate = ReadDate(rs.Fields("PeriodeAu"))
    If Not IsNull(dDate) Then
      nbJPer = DateDiff("d", dDate, dateArrete) + 1
      If nbJPer > nbJourMax And bContentieux = False Then 'And (rsSalarie.Fields("POGARCLE") = cdGar57 Or rsSalarie.Fields("POGARCLE") = cdGar59) Then
                
        bKeep = True
        
        If curRegime = cdGarIncapInval_Import And (rs.Fields("CodeProv") = cdPositImport_IncapProf Or rs.Fields("CodeProv") = cdPositImport_IncapNonProf) And bPortefeuilleSalarie Then
          bKeep = False
' PHM 28/04/2010 : la règle ne s'applique que au code provision 1 & 2
'        If curRegime = cdGarIncapInval_Import And (rs.Fields("CodeProv") = 3 Or rs.Fields("CodeProv") = 4) And bPortefeuilleSalarie Then
'          bKeep = True
'        ElseIf curRegime = cdGarDeces_Import And bPortefeuilleSalarie Then
'          bKeep = True
'        Else
'          bKeep = False
' PHM 28/04/2010 : la règle ne s'applique que au code provision 1 & 2
        End If
        
        If bKeep = False Then
          'Call m_Logger.EcritTraceDansLog("   Annual.= 0 (anc. paiement > " & nbJourMax & " j) à la ligne: " & rs.Fields("RECNO"))
          theAssure.m_POPRESTATION_AN = 0
          theAssure.m_POCOMMENTANNUL = "Import: Annual.= 0 (anc. PeriodeAu > " & nbJourMax & " j)"
          If curRegime <> 6 Then
            theAssure.m_PODOSSIERCLOS = True
          End If
        End If
      
      End If
    End If
  
  Else
    
    nbJPer = DateDiff("d", dernierPaiement, dateArrete) + 1
    If nbJPer > nbJourMax And bContentieux = False Then 'And (rsSalarie.Fields("POGARCLE") = cdGar57 Or rsSalarie.Fields("POGARCLE") = cdGar59) Then
      
      bKeep = True
      
      If curRegime = cdGarIncapInval_Import And (rs.Fields("CodeProv") = cdPositImport_IncapProf Or rs.Fields("CodeProv") = cdPositImport_IncapNonProf) And bPortefeuilleSalarie Then
        bKeep = False
' PHM 28/04/2010 : la règle ne s'applique que au code provision 1 & 2
'        If curRegime = cdGarIncapInval_Import And (rs.Fields("CodeProv") = 3 Or rs.Fields("CodeProv") = 4) And bPortefeuilleSalarie Then
'          bKeep = True
'        ElseIf curRegime = cdGarDeces_Import And bPortefeuilleSalarie Then
'          bKeep = True
'        Else
'          bKeep = False
' PHM 28/04/2010 : la règle ne s'applique que au code provision 1 & 2
      End If
      
      If bKeep = False Then
        'Call m_Logger.EcritTraceDansLog("   Annual.= 0 (anc. paiement > " & nbJourMax & " j) à la ligne: " & rs.Fields("RECNO"))
        theAssure.m_POPRESTATION_AN = 0
        theAssure.m_POCOMMENTANNUL = "Import: Annual.= 0 (anc. paiement > " & nbJourMax & " j)"
        If curRegime <> 6 Then
          theAssure.m_PODOSSIERCLOS = True
        End If
      End If
    
    End If
  
  End If
  
  ' PSAP
  theAssure.m_POPSAPCAPMOYEN = False
  theAssure.m_POPSAP = 0
  
  If curRegime = cdGarDeces_Import Then
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
    
'    If rsSalarie.Fields("POPRESTATIONTOTAL") + rsSalarie.Fields("POPRESTATION") + rsSalarie.Fields("POPSAP") = 0 Then
'      If psap = 0 Then
'        rsSalarie.Fields("POPSAP") = avgPSAP_Regime6
'        rsSalarie.Fields("POPSAPCAPMOYEN") = True
'      End If
'    End If

'PHM 15/01/2004
    
    'rsSalarie.Fields("POPRESTATIONTOTAL") = montantAnnualisation
    'rsSalarie.Fields("POPRESTATION") = montantAnnualisation
    theAssure.m_POPSAP = montantAnnualisation
    
    ' on rajoute la PSAP du fichier
    theAssure.m_POPSAP = theAssure.m_POPSAP + psap
  Else
    ' autre regime
    If psap <> 0 Then
      theAssure.m_POPSAP = psap
    Else
      dDate = ReadDate(rs.Fields("PeriodeAu"))
      If duree <> 0 And Not IsNull(dDate) Then
        Dim d2 As Double
        
        d2 = theAssure.m_POPRESTATION * DateDiff("d", dDate, dateArrete) / duree
        
        theAssure.m_POPSAP = Arrondi(d2, 2)
      Else
        theAssure.m_POPSAP = 0
      End If
    End If
  End If
  
  If theAssure.m_POPSAP < 0 Then
    theAssure.m_POPSAP = 0
  End If
  
  
  
  ' date de fin de provisionnement
  FieldName = "FinGarantie"
  dDate = ReadDate(rs.Fields("FinGarantie"))
  If Not IsNull(dDate) Then
    theAssure.m_POTERME = dDate
  End If
  
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
    
  
  
'  If rs.Fields("RECNO") = 2 Then Stop
'  If rs.Fields("RECNO") = 5154 Then Stop
  
  
  
  ' position
'* rsSalarie.Fields("POSIT") = 1 ' incap avec passage
'* rsSalarie.Fields("POSIT") = 2 ' inval
'* rsSalarie.Fields("POSIT") = 3 ' incap sans passage
'* rsSalarie.Fields("POSIT") = 6 ' deces
'* rsSalarie.Fields("POSIT") = 7 ' maternite
'* rsSalarie.Fields("POSIT") = 20 ' rente conjoint
  Dim bTemporaire As Boolean, dDateLimite As Variant
  
  dDateLimite = ReadDate(rs.Fields("FinGarantie"))
  
  Select Case theAssure.m_POGARCLE
    ' Rentes
    Case cdGarRente
      theAssure.m_POSIT = cdPosit_RenteAutreViagere ' par defaut : autre rente
  
      ' recherche le type de rente
      'CodeProvision = FindCodeProvision(rs.Fields("GE"))
      CodeProvision = 0
      CodeProvision = rs.Fields("CodeProv")
      If CodeProvision < 1 Then
        Call m_Logger.EcritTraceDansLog("REJET : Code_Prov < 0 invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      End If
      
      ' rente temporaire ou viagere ?
      FieldName = "FinGarantie"
      bTemporaire = False
      If Not IsNull(dDateLimite) Then
        bTemporaire = True
      End If
        
      ' adaptation aux differents types de rente
      Select Case CodeProvision
        ' rente certaine
        Case cdPositImport_RenteCertaine
          theAssure.m_POSIT = cdPosit_RenteCertaine
        
        ' rente education
        Case cdPositImport_RenteEducation, cdPositImport_RenteEducation_Handicape
          theAssure.m_POSIT = IIf(bTemporaire, cdPosit_RenteEducationTemporaire, cdPosit_RenteEducationViagere)
          
          If Not IsNull(dDateLimite) Then
            If (DateDiff("yyyy", dateNaissance, dDateLimite) > 26) Then
              Call m_Logger.EcritTraceDansLog("Attention : Age limite rente éducation > 26 ans à la ligne: " & rs.Fields("RECNO"))
            End If
          Else
            If (DateDiff("yyyy", dateNaissance, dDateLimite) > 26) Then
              theAssure.m_POSIT = cdPosit_RenteEducationViagere
            Else
              If CodeProvision = cdPositImport_RenteEducation_Handicape Then
                theAssure.m_POSIT = cdPosit_RenteEducationViagere
              Else
                CodeProvision = cdPosit_RenteEducationTemporaire
                theAssure.m_POSIT = CodeProvision
                dDateLimite = DateAdd("yyyy", 26, dateNaissance)
                theAssure.m_POTERME = dDateLimite
              End If
            End If
          End If
          
          
    
        ' rente de conjoint
        Case cdPositImport_RenteConjoint
          theAssure.m_POSIT = IIf(bTemporaire, cdPosit_RenteConjointTemporaire, cdPosit_RenteConjointViagere)
    
        ' rente retraite
        Case cdPositImport_RenteRetraite
          theAssure.m_POSIT = IIf(bTemporaire, cdPosit_RenteRetraiteTemporaire, cdPosit_RenteRetraiteViagere)
    
        ' autre rente
        Case cdPositImport_RenteAutre
          theAssure.m_POSIT = IIf(bTemporaire, cdPosit_RenteAutreTemporaire, cdPosit_RenteAutreViagere)
    
        ' erreur
        Case Else
          Call m_Logger.EcritTraceDansLog("REJET : Code_PROV " & CodeProvision & " (GE='" & rs.Fields("GE") & "') non supporté à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
      End Select
    
    ' deces
    Case cdGar56
      theAssure.m_POSIT = cdPosit_Deces ' deces
    
    ' inaptitude à la conduite
    Case cdGar57
      theAssure.m_POSIT = cdPosit_Inval
    
    ' arret de travail
    Case cdGar59
      theAssure.m_POSIT = cdPosit_Inval ' inval par defaut puis décodage
      
      ' recherche du type de garantie
      'CodeProvision = FindCodeProvision(rs.Fields("GE"))
      CodeProvision = 0
      CodeProvision = rs.Fields("CodeProv")
      If CodeProvision < 1 Then
        Call m_Logger.EcritTraceDansLog("REJET : Code_Prov < 0 invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      End If
      
'      If CodeProvision = 6001 Or CodeProvision = 6003 Then
'        NbRejet = NbRejet
'      End If
        
      
      Select Case CodeProvision
        ' Incap
        Case cdPositImport_IncapProf, cdPositImport_IncapNonProf, cdProvision_Incap_AvecPassage_Viager '6001=incap avec passage viager
          ' test NCA
          bPassageCATR9 = PassageCATR9(curCat, curNCA, bFoundCATR9)
          
          If bFoundCATR9 Then
            ' existe donc Incap+Passage ou Incap seul
            If bPassageCATR9 = True Then
              ' si passage dans l'entree dans CATR9
              ' ou si pas d'entrée dans PassageNCA, c'est de l'incap avec passage
              theAssure.m_POSIT = cdPosit_IncapAvecPassage ' incap avec passage
            Else
              theAssure.m_POSIT = cdPosit_IncapSansPassage ' incap sans passage
            End If
            
            If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
                theAssure.m_POSIT = cdProvision_Incap_AvecPassage_Viager
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
            If Not rsCatInval.EOF Then
              rsCatInval.MoveFirst
            End If
            rsCatInval.Find "Categorie='" & curCat & "'"
            
            If Not rsCatInval.EOF Then
              
              If Not bFound Then
                ReDim Preserve lstCatNotFound(nbCatNotFound + 1) As String
                lstCatNotFound(nbCatNotFound) = curCat
                nbCatNotFound = nbCatNotFound + 1
                'Call m_Logger.EcritTraceDansLog("   Catégorie INVAL : " & curCat)
              End If
            
            Else
    
              ' ajoute si pas trouvé
              If Not bFound Then
                ReDim Preserve lstCatNotFound(nbCatNotFound + 1) As String
                lstCatNotFound(nbCatNotFound) = curCat
                nbCatNotFound = nbCatNotFound + 1
                'Call m_Logger.EcritTraceDansLog("   Catégorie à ajouter dans CATR9 ou CATR9INVAL : " & curCat & " - ATTENTION : Assuré considéré en INVAL")
              End If
            
            End If
          End If
          
          If CodeProvision = cdPositImport_IncapProf Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
            theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 65, Month(dateNaissance), Day(dateNaissance))  ' 65 ans
          ElseIf Year(curArret) - Year(dateNaissance) < 57 Then
            theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 60, Month(dateNaissance), Day(dateNaissance)) ' 60 ans
          Else
            theAssure.m_POTERME = DateSerial(Year(curArret) + 3, Month(curArret), Day(curArret)) ' Arret + 3 ans
          End If
          
          'Modif 11/7/2018
          If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
              theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
          End If
                
        
        
        ' Inval
        Case cdPositImport_InvalProf, cdPositImport_InvalNonProf, cdProvision_Inval_Viager '6003=Inval viagère
          theAssure.m_POSIT = cdPosit_Inval
          
          If CodeProvision = cdProvision_Inval_Viager Then
              theAssure.m_POSIT = cdProvision_Inval_Viager
          End If
          
          
          ' Décodage de CDCHOIXPREST pour Catégori Inval
          If Not IsNull(rs.Fields("CDCHOIXPREST")) Then
            theAssure.m_POCategorieInval = FindCategorieInval(rs.Fields("CDCHOIXPREST"))
          Else
            Call m_Logger.EcritTraceDansLog("CDCHOIXPREST non spécifié (Assuré forcé en Inval Catégorie " & cdCategorieInvalParDefaut & ") à la ligne: " & rs.Fields("RECNO"))
            theAssure.m_POCategorieInval = cdCategorieInvalParDefaut ' Catégorie par défaut (majorité des invalides de Generali)
          End If
          
          If CodeProvision = cdPositImport_InvalProf Or CodeProvision = cdProvision_Inval_Viager Then
            theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 65, Month(dateNaissance), Day(dateNaissance)) ' 65 ans
          Else
            If theAssure.m_POCategorieInval = 1 Then
              theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 65, Month(dateNaissance), Day(dateNaissance)) ' 65 ans
            Else
              theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 60, Month(dateNaissance), Day(dateNaissance)) ' 60 ans
            End If
            
            ' 14/05/2009 : prolongé jusqu'à 65 ans pour éviter coeff négatif.
            If bPortefeuilleSalarie = True And dateArrete > theAssure.m_POTERME Then
              theAssure.m_POTERME = DateSerial(Year(dateNaissance) + 65, Month(dateNaissance), Day(dateNaissance)) ' 65 ans
            End If
          
          End If
          
          'Modif 11/7/2018
          If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
              theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
          End If
          
          If CodeProvision = cdProvision_Inval_Viager Then
            'on force le code provision inval viager
            theAssure.m_POSIT = cdProvision_Inval_Viager
          End If
        
        
        
        ' Maternite
        Case 5
          theAssure.m_POSIT = cdPosit_Maternite
          theAssure.m_POTERME = DateSerial(Year(curArret), Month(curArret) + 10, Day(curArret))  ' 10 Mois max
        
        ' Mensualisation
        Case 6
          theAssure.m_POSIT = cdPosit_Mensualisation
          theAssure.m_POTERME = DateSerial(Year(curArret), Month(curArret) + 12, Day(curArret)) ' 12 Mois max
        
        ' Chomage
        Case 8
          theAssure.m_POSIT = cdPosit_Chomage
          theAssure.m_POTERME = DateSerial(Year(curArret), Month(curArret) + 24, Day(curArret)) ' 24 Mois max
          
         
        Case Else
          Call m_Logger.EcritTraceDansLog("REJET : Code_PROV " & CodeProvision & " (GE='" & rs.Fields("GE") & "') non supporté à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
      End Select
  End Select
  
  
  ' PSAP pour les codes provision 1 à 4
  ' fonction -> DESACTIVEE LE 04/02/2018 AM
  
'  Select Case CodeProvision
'    Case cdPositImport_IncapProf, cdPositImport_IncapNonProf, cdPositImport_InvalProf, cdPositImport_InvalNonProf
'      dDate = ReadDate(rs.Fields("PeriodeAu"))
'      If Not IsNull(dDate) Then
'
'        d2 = theAssure.m_POPRESTATION_AN * DateDiff("d", dDate, dateArrete) / 365.25
'
'        theAssure.m_POPSAP = Arrondi(d2, 2)
'      Else
'        dDate = ReadDate(rs.Fields("Survenance"))
'        If Not IsNull(dDate) Then
'
'          d2 = theAssure.m_POPRESTATION_AN * DateDiff("d", dDate, dateArrete) / 365.25
'
'          theAssure.m_POPSAP = Arrondi(d2, 2)
'        Else
'          theAssure.m_POPSAP = 0
'        End If
'      End If
'  End Select
  
  If theAssure.m_POPSAP < 0 Then
    theAssure.m_POPSAP = 0
  End If
  
  
  ' date de fin de provisionnement
  ' PHM 01/07/2009 : si DTLIMPRO/FinGarantie précisée, on l'utilise
  FieldName = "FinGarantie"
  dDate = ReadDate(rs.Fields("FinGarantie"))
  If Not IsNull(dDate) Then
    theAssure.m_POTERME = dDate
  End If
  
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
  
  ' PHM 01/07/2009
  
  ' valeurs par défaut = Male puis importée
  DoEvents
  theAssure.m_POSEXE = "1"
  FieldName = "Sexe"
  
  ' modification ali am 15/10/2017 IsStatutaire(catCode)
  If TwoLotImport And IsStatutaire(catCode) Then
  '### Modify this after we get the final structure of the Excel File
    If Not SexAllMale Then
      
      'get data from Excel File -- 1==Homme ; 2==Femme
      Dim sexe As String
      Dim sexeExcel As String
      Dim dataVersion As String
      
      dataVersion = rs.Fields("DataVersion")
      If dataVersion = "0" Then
        dataVersion = "I"
      Else
        dataVersion = "M"
      End If
      
      sexe = "1"
      
      'rsSexe.FindFirst ("AssID=" & 1)  '"AssID='" & 1 & "'"
      rsSexe.MoveFirst
      rsSexe.FindFirst ("NUTRAITP3I=" & rs.Fields("NUTRAITP3I") & " AND NUENRP3I=" & rs.Fields("NUENRP3I") & " AND DataVersion='" & dataVersion & "'")
            
      If rsSexe.NoMatch Then
        '###
        'MsgBox "Le sexe de l'assuré n'était pas trouvé dans le fichier Excel fourni"
        Call m_Logger.EcritTraceDansLog("Sexe non renseignée à la ligne: " & rs.Fields("RECNO") & _
          "pour les paramètres suivantes : NUTRAITP3I=" & rs.Fields("NUTRAITP3I") & " AND NUENRP3I=" & rs.Fields("NUENRP3I") & _
          " AND DataVersion='" & dataVersion & "' - le sexe sera force à Homme !")
        'NbRejet = NbRejet + 1
        'Exit Sub
      Else
        If Not IsNull(rsSexe.Fields("CDSEXASSURE")) Then
          sexeExcel = rsSexe!CDSEXASSURE
        End If
        
'        Do While Not sexe.NoMatch
'            sexe = rsSexe!sexe
'            rsSexe.FindNext "AssID='" & 1 & "'"
'        Loop
      End If
      
      If sexeExcel = "F" Then
        sexe = "2"
      End If
    
      theAssure.m_POSEXE = sexe
  
    
    End If
  Else
    'Standard Import
    If Not IsNull(rs.Fields("Sexe")) Then
      theAssure.m_POSEXE = rs.Fields("Sexe")
    End If
    
    ' PHM 26/04/2010
    If theAssure.m_POGARCLE = cdGarRente Then
      theAssure.m_POSEXE = "2" ' pour les rentes, on force à Femme
    End If
    ' PHM 26/04/2010
  End If
  
  
  
  
  theAssure.m_POPRESTATION_AN_PREC = 0
  
  ' si le salarie a repris dans la periode comptable, on force l'annualisation à 0
  If Not IsNull(dateReprise) Then
    theAssure.m_POREPRISE = dateReprise
    theAssure.m_POPRESTATION_AN = 0
    Call m_Logger.EcritTraceDansLog("   Annual.= 0 (Date de reprise=" & Format(dateReprise, "dd/mm/yyyy") & ") à la ligne: " & rs.Fields("RECNO"))
    theAssure.m_POCOMMENTANNUL = "Import: Annual.= 0 (Date de reprise=" & Format(dateReprise, "dd/mm/yyyy") & ")"
  End If
  
  ' nouveau champs au 20/02/2003
  FieldName = "GE->Nouv Regime"
  If Not IsNull(rs.Fields("GE")) Then
    theAssure.m_POGARCLE_NEW = rs.Fields("GE")
  End If
  
  FieldName = "CodeProvision"
  theAssure.m_POCATEGORIE_NEW = CodeProvision
  
  If bFromInfocentre = False Then
    FieldName = "RegrAnnexe"
    If Not IsNull(rs.Fields("RegrAnnexe")) Then
      theAssure.m_POREGROUPEMENT = rs.Fields("RegrAnnexe")
    End If
    
    FieldName = "RegrStat"
    If Not IsNull(rs.Fields("RegrStat")) Then
      theAssure.m_POCODENATURE = rs.Fields("RegrStat")
    End If
  End If
  
'  FieldName = "CCN"
'  If Not IsNull(rs.Fields("CCN")) Then
'    rsSalarie.Fields("POCCN") = rs.Fields("CCN")
'  End If
  
'  FieldName = "Code Risque"
'  If Not IsNull(rs.Fields("Code Risque")) Then
'    rsSalarie.Fields("POCODERISQUE") = rs.Fields("Code Risque")
'  End If
  ' FIN nouveau champs au 20/02/2003
  
  
  ' DEBUT nouveau champs au 04/02/2004
  FieldName = "DateEntreeInvalidite"
  dDate = ReadDate(rs.Fields("DateEntreeInvalidite"))
  If Not IsNull(dDate) Then
    theAssure.m_PODATEENTREEINVAL = dDate
    
    ' test passage sinon on reste en 1
    If PassageCATR9(curCat, curNCA, bFoundCATR9) Then
      theAssure.m_POEFFET = theAssure.m_PODATEENTREEINVAL
      theAssure.m_POSIT = cdPosit_Inval ' on force le passage en inval
      
      'ajout 14/4/2018 RSAM
      If CodeProvision = cdProvision_Inval_Viager Then
        'on force le code provision inval viager
        theAssure.m_POSIT = cdProvision_Inval_Viager
      End If
      
      If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
          theAssure.m_POSIT = cdProvision_Incap_AvecPassage_Viager
      End If
      'fin ajout
    Else
      theAssure.m_PODATEENTREEINVAL = dDate
    End If
    
    If theAssure.m_PODATEENTREEINVAL < theAssure.m_POARRET Then
      Call m_Logger.EcritTraceDansLog("REJET : DateEntreeInval inferieur a date d'arret à la ligne: " & rs.Fields("RECNO"))
      NbRejet = NbRejet + 1
      Exit Sub
    End If
  End If
  
' PHM 14/06/2006
'    ' la date d'entree en inval est obligatoire.
'    Call m_Logger.EcritTraceDansLog("DateEntreeInval obligatoire et absente pour le régime " & rsSalarie.Fields("POGARCLE") - 50 & " à la ligne: " & rs.Fields("RECNO"))
'    rsSalarie.CancelUpdate
'    NbRejet = NbRejet + 1
'    Exit Sub
'  End If

  'ajout 14/4/2018 RSAM : cdProvision_Inval_Viager Or cdProvision_Incap_AvecPassage_Viager
  'desactive le 06 07 2018 AM
  ' If (theAssure.m_POSIT = cdPosit_Inval Or theAssure.m_POSIT = cdProvision_Inval_Viager Or theAssure.m_POSIT = cdProvision_Incap_AvecPassage_Viager) And IsNull(dDate) Then
  'active le 06 07 2018 AM
  If (theAssure.m_POSIT = cdPosit_Inval Or theAssure.m_POSIT = cdProvision_Inval_Viager) And IsNull(dDate) Then
  
    ' date d'entree en inval = POARRET + 3 Ans.
' PHM 27/11/2008
'    rsSalarie.Fields("PODATEENTREEINVAL") = DateAdd("yyyy", 3, rsSalarie.Fields("POARRET"))
'    Call m_Logger.EcritTraceDansLog("DateEntreeInval absente forcée à " & rsSalarie.Fields("PODATEENTREEINVAL") & " à la ligne: " & rs.Fields("RECNO"))
    
    theAssure.m_PODATEENTREEINVAL = theAssure.m_POARRET
  
  End If
' PHM 14/06/2006
  ' FIN nouveau champs au 04/02/2004
 
' les nouveaux champs DC au 08/10/2004
''  FieldName = "CategorieRteEducDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteEduc"), rs.Fields("CategorieRteEducDC")
''
''  FieldName = "RegimeRteEducDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteEduc"), rs.Fields("RegimeRteEducDC")
''
''  FieldName = "CategorieRteConjointViagereDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteConjointViagere"), rs.Fields("CategorieRteConjointViagereDC")
''
''  FieldName = "RegimeRteConjointViagereDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteConjointViagere"), rs.Fields("RegimeRteConjointViagereDC")
''
''  FieldName = "CategorieRteConjointTempoDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieRenteConjointTempo"), rs.Fields("CategorieRteConjointTempoDC")
''
''  FieldName = "RegimeRteConjointTempoDC"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeRenteConjointTempo"), rs.Fields("RegimeRteConjointTempoDC")
''
''  FieldName = "MajoEnfant"
''  AffecteSiNonNullEtNonZero rsSalarie.Fields("POMajoEnfant"), rs.Fields("MajoEnfant")
''  If Not IsNull(rsSalarie.Fields("POMajoEnfant")) Then
''    ' MajoEnfant >= 0
''    If rs.Fields("MajoEnfant") < 0 Then
''      Call m_Logger.EcritTraceDansLog("MajoEnfant (" & rs.Fields("MajoEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
''      rsSalarie.CancelUpdate
''      NbRejet = NbRejet + 1
''      Exit Sub
''    End If
''  End If
'
'
'
'  'If rsSalarie.Fields("PONOM") = "ALEXANDRE Frédéric" Then Stop
'
'
'
'  FieldName = "TauxGarantieDC"
'  AffecteSiNonNullEtNonZero rsSalarie.Fields("POTauxGarantieDC"), rs.Fields("TauxGarantieDC")
'  If Not IsNull(rsSalarie.Fields("POTauxGarantieDC")) Then
'    ' TauxGarantieDC >= 0
'    If rs.Fields("TauxGarantieDC") < 0 Then
'      Call m_Logger.EcritTraceDansLog("TauxGarantieDC (" & rs.Fields("TauxGarantieDC") & ") invalide à la ligne: " & rs.Fields("RECNO"))
'      rsSalarie.CancelUpdate
'      NbRejet = NbRejet + 1
'      Exit Sub
'    End If
'  End If
'
'  FieldName = "CategorieDC"
'  AffecteSiNonNullEtNonZero rsSalarie.Fields("POCategorieDeces"), rs.Fields("CategorieDC")
'
'  FieldName = "RegimeDC"
'  AffecteSiNonNullEtNonZero rsSalarie.Fields("PORegimeDeces"), rs.Fields("RegimeDC")
'
'  FieldName = "AgeEnfant"
'  AffecteSiNonNullEtNonZero rsSalarie.Fields("POAgeMoyenEnfant"), rs.Fields("AgeEnfant")
'  If Not IsNull(rsSalarie.Fields("POAgeMoyenEnfant")) Then
'    ' age enfant >= 0
'    If rsSalarie.Fields("POAgeMoyenEnfant") < 0 Then
'      Call m_Logger.EcritTraceDansLog("AgeEnfant (" & rsSalarie.Fields("POAgeMoyenEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
'      rsSalarie.CancelUpdate
'      NbRejet = NbRejet + 1
'      Exit Sub
'    End If
'
'    ' age enfant > 24
'    If rsSalarie.Fields("POAgeMoyenEnfant") > 24 Then
'      Call m_Logger.EcritTraceDansLog("AgeEnfant (" & rsSalarie.Fields("POAgeMoyenEnfant") & ") > 24 ans à la ligne: " & rs.Fields("RECNO"))
'    End If
'  End If
'
'  FieldName = "NbEnfant"
'  AffecteSiNonNullEtNonZero rsSalarie.Fields("PONbEnfant"), rs.Fields("NbEnfant")
'  If Not IsNull(rsSalarie.Fields("PONbEnfant")) Then
'    ' nb enfant >= 0 et nb enfant entier
'    If rsSalarie.Fields("PONbEnfant") < 0 Or CDbl(Fix(rsSalarie.Fields("PONbEnfant"))) <> CDbl(rsSalarie.Fields("PONbEnfant")) Then
'      Call m_Logger.EcritTraceDansLog("NbEnfant (" & rsSalarie.Fields("PONbEnfant") & ") invalide à la ligne: " & rs.Fields("RECNO"))
'      rsSalarie.CancelUpdate
'      NbRejet = NbRejet + 1
'      Exit Sub
'    End If
'
'  End If
  
  FieldName = "SalBrut"
  If Not IsNull(rs.Fields("SalBrut")) Then
    theAssure.m_POSalaireAnnuel = rs.Fields("SalBrut")
  End If
  
  FieldName = "Cadre"
  theAssure.m_POIsCadre = False

'  If Not IsNull(rs.Fields("Cadre")) Then
'    rsSalarie.Fields("POIsCadre") = rs.Fields("Cadre")
'  Else
'    Call m_Logger.EcritTraceDansLog("Cadre doit être renseigné à la ligne: " & rs.Fields("RECNO"))
'    rsSalarie.CancelUpdate
'    NbRejet = NbRejet + 1
'    Exit Sub
'  End If
'
'  FieldName = "SituationFamille"
'  If Not IsNull(rs.Fields("SituationFamille")) Then
'    n = m_dataHelper.GetParameterAsDouble("SELECT CleSituationFamille FROM SituationFamille WHERE Libelle='" & UCase(rs.Fields("SituationFamille")) & "'")
'    If n <> 0 Then
'      rsSalarie.Fields("POCleSituationFamille") = n
'    Else
'      Call m_Logger.EcritTraceDansLog("SituationFamille invalide à la ligne: " & rs.Fields("RECNO"))
'      rsSalarie.CancelUpdate
'      NbRejet = NbRejet + 1
'      Exit Sub
'    End If
'  Else
'    rsSalarie.Fields("POCleSituationFamille") = Null
'  End If
' les nouveaux champs DC au 08/10/2004

  
' PHM 30/05/2006 : les nouveaux champs générali
  FieldName = "Contractant"
  theAssure.m_POContractant = rs.Fields("Contractant")
  
  ' CDSITUASIN stocké dans POSituConv
  FieldName = "CDSITUATSIN"
  theAssure.m_POSituConv = rs.Fields("CDSITUATSIN")
  
  
  If bFromInfocentre = False Then
  
    FieldName = "EffetSitu"
    theAssure.m_POEffetSitu = rs.Fields("EffetSitu")
    
    FieldName = "Etablissement"
    theAssure.m_POEtablissement = rs.Fields("Etablissement")
    
    FieldName = "TypeRglt"
    theAssure.m_POTypeReglement = rs.Fields("TypeRglt")
    
    FieldName = "CreationDossier"
    theAssure.m_POCreationDossier = rs.Fields("CreationDossier")
    
    FieldName = "DebDossier"
    theAssure.m_PODebutDossier = rs.Fields("DebDossier")
    
    FieldName = "FinDossier"
    theAssure.m_POFinDossier = rs.Fields("FinDossier")
    
    FieldName = "MotifCloture"
    theAssure.m_POMotifCloture = rs.Fields("MotifCloture")
    
    FieldName = "DebRefSal"
    theAssure.m_PODebutRefSalaire = rs.Fields("DebRefSal")
    
    FieldName = "FinRefSal"
    theAssure.m_POFinRefSalaire = rs.Fields("FinRefSal")
  End If
  
  FieldName = "CodeOption"
  theAssure.m_POCode_Option_Contrat = rs.Fields("CodeOption")
  If IsNull(rs.Fields("CodeOption")) Then
    Call m_Logger.EcritTraceDansLog("REJET : CodeOption non renseignée à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  FieldName = "Dossier"
  theAssure.m_PONumSinistre = rs.Fields("Dossier")
  
  FieldName = "TypeMvt"
  theAssure.m_POTypeMvt = rs.Fields("TypeMvt")
  
  theAssure.m_POMontantBase = MontantBase
  theAssure.m_POMontantRevalo = MontantRevalo
  theAssure.m_POMontantMajore = MontantMajo
  
  FieldName = "DebIndemn"
  dDate = ReadDate(rs.Fields("DebIndemn"))
  If Not IsNull(dDate) Then
    theAssure.m_PODebutIndemnisation = dDate
  End If
  
  theAssure.m_PONbJourIndemn = nbJIndemn
  
  If bFromInfocentre = False Then
    FieldName = "FamComptable"
    
    theAssure.m_POFamilleComptable = rs.Fields("FamComptable")
    
    If IsNull(rs.Fields("FamComptable")) Then
      Call m_Logger.EcritTraceDansLog("REJET : FamComptable non renseignée à la ligne: " & rs.Fields("RECNO"))
      NbRejet = NbRejet + 1
      Exit Sub
    End If
    
    FieldName = "Inspecteur"
    theAssure.m_POInspecteur = rs.Fields("Inspecteur")
    
    FieldName = "Insp"
    theAssure.m_POInsp = rs.Fields("Insp")
    
    FieldName = "Apport"
    theAssure.m_POApport = rs.Fields("Apport")
    
    FieldName = "Apport2"
    theAssure.m_POApport2 = rs.Fields("Apport2")
    
    FieldName = "Gestionnaire"
    theAssure.m_POGestionnaire = rs.Fields("Gestionnaire")
    
    FieldName = "NbIntervenant"
    theAssure.m_PONbIntervenant = rs.Fields("NbIntervenant")
    
    FieldName = "Compagnie"
    theAssure.m_POCompagnie = rs.Fields("Compagnie")
  Else
    ' famille comptable pour compatibilité (=30)
    FieldName = "FamComptable"
    theAssure.m_POFamilleComptable = rs.Fields("FamComptable")
    
    FieldName = "CDCOMPAGNIE"
    theAssure.m_POCompagnie = rs.Fields("CDCOMPAGNIE")
    
    FieldName = "CDAPPLI"
    theAssure.m_POAppli = rs.Fields("CDAPPLI")
  End If
  
  FieldName = "IndicCC"
  theAssure.m_POIndicCC = rs.Fields("IndicCC")
' PHM 30/05/2006 : les nouveaux champs générali
  
  
  FieldName = "Fractionnement"
  If Not IsNull(rs.Fields("Fractionnement")) Then
    theAssure.m_POFRACT = rs.Fields("Fractionnement")
  End If
  
  FieldName = "TermeEchu"
  If Not IsNull(rs.Fields("TermeEchu")) Then
    theAssure.m_POECHU = rs.Fields("TermeEchu")
    Select Case theAssure.m_POECHU
      Case EcheancePaiement.ePaiementAvance, EcheancePaiement.ePaiementEchu
      
      Case Else
        Call m_Logger.EcritTraceDansLog("REJET : TermeEchu invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
    End Select
  End If
  
  FieldName = "ParamCalcul"
  
  theAssure.m_PONumParamCalcul = FindNumParamCalcul(rs.Fields("CDCOMPAGNIE"), rs.Fields("CDAPPLI"), Trim(rs.Fields("Code_Cat_Contrat")))
  If theAssure.m_PONumParamCalcul = -1 Then
  
    Call m_Logger.EcritTraceDansLog("REJET : ParamCalcul non renseigné à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If

  
  
  ' Evol 2010 - Lot 2
  FieldName = "CDCONTENTIEUX"
  If Not IsNull(rs.Fields("CDCONTENTIEUX")) Then
    theAssure.m_POCDCONTENTIEUX = rs.Fields("CDCONTENTIEUX")
  End If
  
  FieldName = "NUSINISTRE"
  If Not IsNull(rs.Fields("NUSINISTRE")) Then
    theAssure.m_PONUSINISTRE = rs.Fields("NUSINISTRE")
  End If
  
  FieldName = "MTCAPCON"
  theAssure.m_POMontantCapConstit = 0
  'If Not IsNull(rs.Fields("MTCAPCON")) Then
  '  theAssure.m_POMontantCapConstit = rs.Fields("MTCAPCON")
  'End If
  
  FieldName = "MTCAPSSRISQ"
  theAssure.m_POMontantCapSousRisque = 0
  If Not IsNull(rs.Fields("MTCAPSSRISQ")) Then
    theAssure.m_POMontantCapSousRisque = rs.Fields("MTCAPSSRISQ")
  End If
  
  
  theAssure.m_PODATEIMPORT = dateArrete
  
  
  FieldName = "Commentaire"
  If Not IsNull(rs.Fields("Commentaire")) Then
    theAssure.m_Commentaire = rs.Fields("Commentaire")
  End If
  
  
  '### STATUTAIRE : Set the POSIT field in the Assure Table if this is an assure of type statutaire => verify the CDPRODUIT column:
  If IsStatutaire(catCode) Then
    theAssure.m_POSIT = 5000
    theAssure.m_POCATEGORIE_NEW = 5000
  End If
  
  
  
'  theAssure.m_Commentaire = Null
'  If Not IsNull(rs.Fields("Commentaire")) Then
'    theAssure.m_Commentaire = rs.Fields("Commentaire")
'  End If

  
  
  
  'If regroupement <> "" Then
  '  rsSalarie.Fields("POREGROUPEMENT") = regroupement
  'End If
  
  '
  ' Sauvegarde de la ligne AVANT la réforme des retraite
  '
  theAssure.Save destDB, False
  
  
  '
  ' Sauvegarde de la ligne APRES la réforme des retraite
  '
  'modif am rechargement du TYPE période en fonction du no de période
  
  '
  If (eProvisionRetraite = m_TypePeriode Or eProvisionRetraiteRevalo = m_TypePeriode) And theAssure.m_POGARCLE = cdGar59 And bPortefeuilleSalarie Then
    theAssure.m_IdAssure = theAssure.m_RECNO
    theAssure.m_RECNO = 0
    
    theAssure.m_POTERME = ReadDate(rs.Fields("FinGarantie")) '''''''''''''' A recalculer en fonction de la garantie
    
    'Modif 11/7/2018
    If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
        theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
    End If
          
    theAssure.m_CoeffAmortissement = 1
    theAssure.m_TopAmortissable = rs.Fields("FLAMORTISSABLE")
    
    theAssure.Save destDB, True
  End If
  
  
  ' sauvegarde des données de P3IPROVCOLL pour Symphonie
  'CopieVersAssureP3IProvColl CleGroupe, NumPeriode, rsAssureP3IProvColl, rs, m_Logger
  CopieVersAssureP3IProvColl_ADO CleGroupe, NumPeriode, rsAssureP3IProvColl, rs, m_Logger
  
  
  Exit Sub
  
GestionErreur:
  
  If Err = 94 Then
  
    Resume Next
  
  Else
  
    If autoMode Then
      Select Case Err
        Case 3021
          'MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
          m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
        
        Case 3265
          MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
          Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO"))
        
        Case 3421
          'MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
          Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO"))
        
        Case -2147217873
          m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
        
        Case Else
          'MsgBox "Erreur " & Err & " à la ligne " & rs.Fields("RECNO") & " : " & Err.Description, vbCritical
          m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
      End Select
    Else
      Screen.MousePointer = vbDefault
      fWait.Visible = False
    
      Select Case Err
        Case 3021
          MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
        
        Case 3265
          MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
          Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO"))
        
        Case 3421
          MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
          Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO"))
        
        Case -2147217873
          m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
        
        Case Else
          MsgBox "Erreur " & Err & " à la ligne " & rs.Fields("RECNO") & " : " & Err.Description, vbCritical
      End Select
    
      fWait.Visible = True
      Screen.MousePointer = vbHourglass
    End If
  
    Resume Next
  End If

End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'##ModelId=5C90E9760250
Private Sub CopieVersAssureP3IProvColl(CleGroupe As Long, NumPeriode As Long, rsAssureP3IProvColl As ADODB.Recordset, rs As ADODB.Recordset, cLogger As clsLogger)
  On Error GoTo err_CopyLot
  
  Dim i As Integer, f As ADODB.Field, nb As Long, bOk As Boolean
  
  
  rsAssureP3IProvColl.AddNew
  
  
  ' on ne recopie que les champs present dans AssureP3IProvColl les autres sont dans Assure
  For i = 0 To rsAssureP3IProvColl.Fields.Count - 1
    
    Set f = rsAssureP3IProvColl.Fields(i)
      
    If f.Name = "CleGroupe" Then
      f.Value = CleGroupe
    ElseIf f.Name = "NumPeriode" Then
      f.Value = NumPeriode
    Else
      f.Value = rs.Fields(f.Name).Value
    End If
    
  Next
    
  
  rsAssureP3IProvColl.Update
    
    
  Exit Sub
  
err_CopyLot:
  
  If Not autoMode Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Sub




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'##ModelId=5C90E976030A
Private Sub CopieVersAssureP3IProvColl_ADO(CleGroupe As Long, NumPeriode As Long, rsAssureP3IProvColl As ADODB.Recordset, rs As ADODB.Recordset, cLogger As clsLogger)
  On Error GoTo err_CopieVersAssureP3IProvColl_ADO
  
  Dim f As ADODB.Field
  
  ' creation de la requete d'insert
  If m_AssureP3IProvColl_SQLText = "" Then
    Dim cmdText As String, cmdValues As String
    
    cmdText = "INSERT INTO Assure_P3IPROVCOLL("
    cmdValues = " VALUES("
    
    For Each f In rsAssureP3IProvColl.Fields
      cmdText = cmdText & f.Name & ", "
      cmdValues = cmdValues & "?, "
    Next
    
    cmdText = Left(cmdText, Len(cmdText) - 2) & ")"
    cmdValues = Left(cmdValues, Len(cmdValues) - 2) & ")"
    
    m_AssureP3IProvColl_SQLText = cmdText & cmdValues
  End If
  
  ' creation de la commande
  Dim cmd As ADODB.Command, prm As ADODB.Parameter, Fin As ADODB.Field
  
  Set cmd = New ADODB.Command
  
  Set cmd.ActiveConnection = m_dataHelper.DataSource.Connection
  cmd.CommandType = adCmdText
  cmd.CommandText = m_AssureP3IProvColl_SQLText
  
  Set prm = cmd.CreateParameter("@CleGroupe", adInteger, adParamInput, 0, CleGroupe)
  cmd.Parameters.Append prm
  Set prm = cmd.CreateParameter("@NumPeriode", adInteger, adParamInput, 0, NumPeriode)
  cmd.Parameters.Append prm

  ' ajjout des valeurs
  For Each f In rsAssureP3IProvColl.Fields
    If f.Name <> "CleGroupe" And f.Name <> "NumPeriode" Then
      Dim l As Long, t As ADODB.DataTypeEnum
      
      Set Fin = rs.Fields(f.Name)
      
      Select Case f.Type
        Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
          If IsNull(Fin.Value) Then
            l = 1
          Else
            l = Len(Fin.Value)
            If l = 0 Then
              l = 1
              Fin.Value = Null
            End If
          End If
          t = adVarChar
        
        Case adNumeric, adDecimal, adDouble
          l = 0
          t = adDouble
       
        Case Else
          l = 0
          t = f.Type
      End Select
      
      Set prm = cmd.CreateParameter("@" & f.Name, t, adParamInput, l, Fin.Value)
      
      cmd.Parameters.Append prm
    End If
  Next
  
  cmd.Execute
  
  Set cmd = Nothing
  
  Exit Sub
  
err_CopieVersAssureP3IProvColl_ADO:

  If autoMode Then
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Else
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  Resume Next
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'##ModelId=5C90E9770005
Private Sub CopieVersAssureP3ILogTrait(CleGroupe As Long, NumPeriode As Long, destDB As DataAccess, NumeroLot As Long, cLogger As clsLogger)
  On Error GoTo err_CopyLot
  
  Dim i As Integer, f As ADODB.Field, nb As Long, bOk As Boolean
  
  Dim rsAssureP3ILogTrait As ADODB.Recordset, rsIn As ADODB.Recordset
  
  
  destDB.Execute "DELETE FROM Assure_P3ILOGTRAIT WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriode & " AND NUTRAITP3I=" & NumeroLot
  
  
  Set rsIn = destDB.OpenRecordset("SELECT * FROM P3ILOGTRAIT WHERE NUTRAITP3I=" & NumeroLot, Snapshot)
  
  Set rsAssureP3ILogTrait = destDB.OpenRecordset("Assure_P3ILOGTRAIT", Table)
  
  rsAssureP3ILogTrait.AddNew
  
  rsAssureP3ILogTrait.Fields("CleGroupe") = CleGroupe
  rsAssureP3ILogTrait.Fields("NumPeriode") = NumPeriode
  
  rsAssureP3ILogTrait.Fields("NUTRAITP3I") = rsIn.Fields("NUTRAITP3I")
  rsAssureP3ILogTrait.Fields("NUTRAIT") = rsIn.Fields("NUTRAIT")
  rsAssureP3ILogTrait.Fields("DTTRAIT") = rsIn.Fields("DTTRAIT")
  rsAssureP3ILogTrait.Fields("HHTRAIT") = rsIn.Fields("HHTRAIT")
  rsAssureP3ILogTrait.Fields("DTDEBPER") = rsIn.Fields("DTDEBPER")
  rsAssureP3ILogTrait.Fields("DTFINPER") = rsIn.Fields("DTFINPER")
  rsAssureP3ILogTrait.Fields("IDTABLESAS") = rsIn.Fields("IDTABLESAS")
  rsAssureP3ILogTrait.Fields("NBLIGTRAIT") = rsIn.Fields("NBLIGTRAIT")
  rsAssureP3ILogTrait.Fields("MTTRAIT") = rsIn.Fields("MTTRAIT")
  rsAssureP3ILogTrait.Fields("Commentaire") = rsIn.Fields("Commentaire")
    
  rsAssureP3ILogTrait.Update
  
  
  rsAssureP3ILogTrait.Close
  rsIn.Close
    
  Exit Sub
  
err_CopyLot:

  If autoMode Then
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Else
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  End If
  
  Resume Next
End Sub


'##ModelId=5C90E97700A3
Private Sub AdjustDeces(destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, fWait As Form)
  On Error GoTo GestionErreur
      
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
  Dim rq As String, where As String
  
  ' fabrique la chaine de recherche des deces dont les enregistrements 7 et 9 relatifs sont à annualiser à 0
  Set rs = destDB.OpenRecordset("SELECT * FROM RisqueDeces WHERE AnnualisationZero<>0", Disconnected)
      
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
  rq = "SELECT * FROM ASSURE WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode & " AND POGARCLE=56 AND (" & where & ")"
      
  Dim mark As Variant
  Dim nbAnnualForcee As Integer
      
  Set m_dataHelper = destDB.CreateHelper
  
  nbAnnualForcee = 0
  
  Set rs = destDB.OpenRecordset(rq, Disconnected)
  Do Until rs.EOF
    ' affiche la position
    If Not autoMode Then
      If (rs.AbsolutePosition Mod 9) = 0 Then
        fWait.ProgressBar1.Max = rs.RecordCount + 1
        fWait.ProgressBar1.Value = rs.AbsolutePosition
        fWait.Label1(0).Caption = "Phase 3 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max
        fWait.Refresh
        DoEvents
      End If
    End If
    
'    rq = "SELECT * FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode _
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
    
    rq = "UPDATE Assure SET " _
       & " POCOMMENTANNUL='Import: DECES : Annualisation de ' + convert(varchar, POPRESTATION_AN) + ' forcée à 0 pour le n° SS ' + PONUMCLE + ', régime ' + convert(varchar, POGARCLE-50), " _
       & " POPRESTATION_AN=0 " _
       & " WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode _
       & " AND (POGARCLE=57 OR POGARCLE=59) " _
       & " AND PONUMCLE='" & rs.Fields("PONUMCLE") & "' AND POPRESTATION_AN <> 0"
    
    destDB.Execute rq
    
    rs.MoveNext
  Loop
      
  rq = "SELECT POCOMMENTANNUL FROM Assure " _
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

  If autoMode Then
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  Else
    Screen.MousePointer = vbDefault
    fWait.Visible = False
    MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
    fWait.Visible = True
    Screen.MousePointer = vbHourglass
  End If
  
  Resume Next
End Sub

'##ModelId=5C90E9770148
Private Sub computeAvgRegime6(rs As ADODB.Recordset, dateArrete As Date, nbJourDC As Long, fWait As Form)
  Dim tmpMontant As Double, tmpMontantPSAP As Double, montant As Double
  Dim nbJPer As Long, nbJHors As Long
    
  Dim datePaiement As Date
  Dim curSte As Long, maxRecord As Long, curArret As Date
  Dim curSS As String, curNCA As String, curRegime As Long, curCat As Long
  
  Dim bOk As Boolean
  
  Dim FieldName As String
  On Error GoTo GestionErreur
      
  avgPSAP_Regime6 = 0
  maxRecord = 0
  
  If Not rs.EOF Then
    ' parcours des enregistrements
    Do Until rs.EOF
      If Not autoMode Then
        If (rs.AbsolutePosition Mod 10) = 0 Then
          ' affiche la position
          fWait.ProgressBar1.Max = rs.RecordCount + 1
          fWait.ProgressBar1.Value = rs.AbsolutePosition
          fWait.Label1(0).Caption = "Phase 1 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max
          fWait.Refresh
          DoEvents
        End If
      End If
      
      FieldName = "Regime"
      If rs.Fields("Regime") = 6 Then
        ' init des variables
        FieldName = "NumSociete"
        If IsNull(rs.Fields("NumSociete")) Then
          Call m_Logger.EcritTraceDansLog("REJET : NumSociété invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curSte = rs.Fields("NumSociete")
        End If
        
        FieldName = "NumeroSS"
        If IsNull(rs.Fields("NumeroSS")) Then
          Call m_Logger.EcritTraceDansLog("REJET : NumeroSS invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curSS = Trim(rs.Fields("NumeroSS"))
        End If
        
        FieldName = "Regime"
        If IsNull(rs.Fields("Regime")) Then
          Call m_Logger.EcritTraceDansLog("REJET : Regime invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curRegime = rs.Fields("Regime")
        End If
        
        FieldName = "Categorie"
        If IsNull(rs.Fields("Categorie")) Then
          Call m_Logger.EcritTraceDansLog("REJET : Categorie invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          If Trim(rs.Fields("Categorie")) = vbNullString Then
            Call m_Logger.EcritTraceDansLog("REJET : Categorie invalide à la ligne: " & rs.Fields("RECNO"))
            NbRejet = NbRejet + 1
            Exit Sub
          End If
          curCat = rs.Fields("Categorie")
        End If
        
        FieldName = "NCA"
        If IsNull(rs.Fields("NCA")) Then
          Call m_Logger.EcritTraceDansLog("REJET : NCA invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curNCA = Trim(rs.Fields("NCA"))
        End If
        
        FieldName = "Survenance"
        If IsNull(rs.Fields("Survenance")) Then
          Call m_Logger.EcritTraceDansLog("REJET : Survenance invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curArret = ReadDate(rs.Fields("Survenance"))
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
             Or curNCA <> rs.Fields("NCA") Or curArret <> rs.Fields("Survenance") Then
            Exit Do
          End If
          
          ' toujours le meme sinistre
          tmpMontant = 0
          FieldName = "MontantNet"
          If Not IsNull(rs.Fields("MontantNet")) Then
            tmpMontant = rs.Fields("MontantNet")
          End If
          
          ' toujours le meme sinistre (on ajoute la psap de Excel)
          tmpMontantPSAP = 0
          FieldName = "PSAP"
          If Not IsNull(rs.Fields("PSAP")) Then
            tmpMontantPSAP = rs.Fields("PSAP")
          End If
          
          ' cumul prestation payées
          FieldName = "DatePaiement"
          If IsNull(rs.Fields("DatePaiement")) Then
            datePaiement = rs.Fields("Survenance") + nbJourDC
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

  If autoMode Then
    Select Case Err
      Case 3021
        'MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
      
      Case 3265
        'MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO") + 1)
      
      Case 3421
        'MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO") + 1)
      
      Case Else
        'MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rs.Fields("RECNO")
    End Select
  Else
    Screen.MousePointer = vbDefault
    fWait.Visible = False
    
    Select Case Err
      Case 3021
        MsgBox "Erreur " & Err & " : " & Err.Description & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
      
      Case 3265
        MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import - Ligne " & rs.Fields("RECNO") + 1)
      
      Case 3421
        MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte" & vbLf & "Ligne " & rs.Fields("RECNO") + 1, vbCritical
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rs.Fields("RECNO") + 1)
      
      Case Else
        MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
    End Select
    
    fWait.Visible = True
    Screen.MousePointer = vbHourglass
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période dans la
' table Assuré
'
'* ajouter un parametre dans les détails periode (et paramêtre) pour le
'* mode de calcul de l'annualisation : "dernier flux" ou "tous les flux de la periode"

'##ModelId=5C90E97701D6
Public Function iP3IGeneraliImport_DoImport(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, Optional autoMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactiviteA As Long, Optional typeCalculAnnualisationA As Long, Optional excelFileA As String, Optional logFileA As String) As Boolean
  
  MsgBox "DoImport() not implemented in object iP3IGeneraliImport.iImportGeneraliSASP3I !", vbCritical
       
  iP3IGeneraliImport_DoImport = False

End Function


'********************************************************************
' Controle des données avant import
'
'##ModelId=5C90E9780030
Public Function iP3IGeneraliImport_DoControle(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Integer, NumeroLot As Long, NumPeriode As Long, sFichierIni As String) As Boolean
  
  iP3IGeneraliImport_DoControle = DoControle(CommonDialog1, destDB, CleGroupe, NumeroLot, NumPeriode, sFichierIni)
  
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période
' depuis un lot présent dans P3ILOGTRAIT et P3IPROVCOLL
'
'##ModelId=5C90E978010F
Public Function iP3IGeneraliImport_DoImportSASP3I(NumeroLot As Long, logPath As String, ByRef destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, Optional automatMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactiviteA As Long, Optional typeCalculAnnualisationA As Long, Optional logFileA As String) As Boolean

  Dim rs3 As ADODB.Recordset, rsAssureP3IProvColl As ADODB.Recordset
  Dim rq As String
  Dim maxRecord As Long
  Dim codeRetour As Boolean
  Dim fWait As frmWait
  Dim datePaiement As Date
  
  Dim res As Long
  Dim sqlStr As String
  Dim rsCount As ADODB.Recordset
  
'destDB.BeginTrans
'GoTo starttest

  On Error GoTo GestionErreur
  
  If Not automatMode Then
    autoMode = False
  Else
    autoMode = True
  End If
  
  NbRejet = 0
  
  If autoMode Then
    typeDelaiInactivite = typeDelaiInactiviteA
    typeImport = eImportDonneesSeules
    typeCalculAnnualisation = typeCalculAnnualisationA
    dateArrete = dateArreteA
    Set fWait = New frmWait
  Else
  
   ' demande confirmation de la destruction des anciennes données
   If MsgBox("ATTENTION: le données vont être écrasées par celles importées." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) = vbNo Then
     iP3IGeneraliImport_DoImportSASP3I = False
     Exit Function
   End If
   
   ' date par defaut en cas d'absence d'information
   dateImport = Format(Now, "dd/mm/yyyy hh:nn")
   If IsNull(dateCloture) Then
     dateArrete = Format(Now, "dd/mm/yyyy")
   Else
     dateArrete = dateCloture
   End If
   
   ' choix de la date d'arreté des comptes et des autres parametres d'import
   Dim f As New frmSelectDate
   
   Load f
   
   f.gDateDebut = DateDebut
   f.gDateFin = DateFin
   f.DTPicker2 = dateArrete
   
   f.rdoImportComplet.Enabled = False
   f.rdoImportDonneesSeules.Value = True
   f.rdoImportTableParametre.Enabled = False
   
   f.lblDate2.Caption = Replace(f.lblDate2.Caption, "888", nbJourMax)
   
   f.Show vbModal
   If ret_code = -1 Then
     iP3IGeneraliImport_DoImportSASP3I = False
     Exit Function
   End If
   
   ' date d'arreté des comptes
   dateArrete = f.DTPicker2
   
   ' type d'import
   If f.rdoImportComplet.Value = True Then
     typeImport = eImportComplet
   ElseIf f.rdoImportDonneesSeules.Value = True Then
     typeImport = eImportDonneesSeules
   ElseIf f.rdoImportTableParametre.Value = True Then
     typeImport = eImportTablesParametresSeules
   End If
   
   ' type de calcul de delai
   If f.rdoDatePaiement.Value = True Then
     typeDelaiInactivite = eDatePaiement
   ElseIf f.rdoDateFinPeriode.Value = True Then
     typeDelaiInactivite = eDateFinPeriodePaiement
   End If
   
   ' type de calcul de l'annualisation
   If f.rdoEnsemblePaiement.Value = True Then
     typeCalculAnnualisation = eEnsemblePaiement
   ElseIf f.rdoDernierPaiement.Value = True Then
     typeCalculAnnualisation = eDernierPaiement
   End If
   ' choix de la date d'arreté des comptes et des autres parametres d'import
   
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
  
  End If 'auto mode
  
  ' preparation du fichier log
  Set m_Logger = New clsLogger
  
  
  If autoMode Then
    m_Logger.modeAuto = True
    m_Logger.FichierLog = logFileA
  Else
    m_Logger.FichierLog = logPath & GetWinUser & "_ErreurImport.log"
    m_Logger.CreateLog "Import depuis le lot " & NumeroLot & " vers Groupe " & CleGroupe & " Période " & NumPeriode
  End If
  
  
  '*******  ATTENTION : ouverture de tous les recordset avant la transaction, sinon : crash
  
  Set m_dataHelper = destDB.CreateHelper
  
  ' transaction...
  Dim transOpened As Boolean
  
  
  destDB.BeginTrans
  transOpened = True
  
  '***
  '*** Import des tables de paramètres
  '***
  
  If typeImport = eImportComplet Or typeImport = eImportTablesParametresSeules Then
     
    If Not autoMode Then
      MsgBox "Type d'import incorrect !", vbCritical
      Stop
    End If
          
  End If
  
  '***
  '*** Import des données
  '***
  
  If typeImport = eImportComplet Or typeImport = eImportDonneesSeules Then
  
    Call m_Logger.EcritTraceDansLog("   ")
    Call m_Logger.EcritTraceDansLog(">>>>> Import des données")
  
    If Not autoMode Then
      fWait.Caption = "Import des données en cours..."
    End If
    
    ' ATTENTION : la zone correspondant à la base de données dans la feuille Excel
    ' =========   doit être nommée 'Assure'
    
      
    '
    ' Data
    '
    
    rq = "SELECT P.NUENRP3I AS RECNO, P.CDCOMPAGNIE AS NumSociete, P.CDPRODUIT AS Code_Cat_Contrat, P.NUCONTRA AS Convention, " _
           & "    P.NUDOSSIERPREST AS Sinistre, P.CDGARAN AS GE, P.CDTYPMVT AS TypeMvt, P.CDOPTION AS CodeOption, P.LBSOUSCR AS Contractant, " _
           & "    P.LBASSURE AS NomAssure, P.DTNAISSASS AS DateNaiss, P.DTSURVSIN AS Survenance, P.DTSURVSIN AS EffetConv, P.CDPERIODICITE AS Fractionnement, " _
           & "    F.m AS Fractionnement, P.CDTYPTERME AS TermeEchu, P.DTEFFREN AS DebIndemn, P.DTLIMPRO AS FinGarantie, P.DTDERREG AS DatePaiement, " _
           & "    P.DTDEBPER AS PeriodeDu, P.DTFINPER AS PeriodeAu, P.CDSINCON AS IndicCC, P.DTMISINV AS DateEntreeInvalidite, " _
           & "    P.MTDERPERRGLTADA AS MntRegle, P.INDBASREV AS BaseRevalo, P.MTPREANN AS MntBase, P.MTPREREV AS MntRevalo, "
    '     rq = rq & "    P.MTPREMAJ AS MntMajore, DATEDIFF(day, P.DTDEBPER, P.DTFINPER) AS NbJIndemnises, P.NUTRAITP3I, P.DataVersion, "
    rq = rq & "    P.MTPREMAJ AS MntMajore, P.NUTRAITP3I, P.DataVersion, P.NUSOUSDOSSIERPREST as Dossier, P.Commentaire as Commentaire, " _
           & "    S.Code AS Sexe, C.CodeProv, C.Regime, 0 as Franchise, NULL as SalBrut, 0 as Cadre, 30 as FamComptable, " _
           & "    P.* "
    
    rq = rq & " FROM P3IUser.TypeSexe AS S  " _
          & "    RIGHT OUTER JOIN P3IUser.P3IPROVCOLL AS P ON S.CDSEXE = P.CDSEXASSURE " _
          & "    LEFT OUTER JOIN P3IUser.Fractionnement AS F ON P.CDPERIODICITE = F.CDPERIODICITE  " _
          & "    LEFT OUTER JOIN P3IUser.CodeProvision AS C  " _
          & "    INNER JOIN P3IUser.TBQREGA AS R ON C.CodeProv = R.Code_PROV " _
          & "         ON P.CDGARAN = cast(R.Code_GE as int) AND P.CDCOMPAGNIE = R.Code_CIE AND P.CDAPPLI = R.Code_APP And R.GroupeCle = " & CleGroupe & " And R.NumPeriode = " & NumPeriode

    rq = rq & " WHERE P.NUTRAITP3I = " & NumeroLot _
         & " AND ( " _
         & "      P.DataVersion = 1 " _
         & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
         & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
         & "     ) "
    
    'rq = rq & " ORDER BY NumSociete, C.Regime, Dossier, GE, Code_Cat_Contrat, Survenance, BaseRevalo, PeriodeDu, PeriodeAu, DatePaiement "
    rq = rq & " ORDER BY RECNO " ' plus de boucle pour les sommations
 
   
    ' Type de période
    m_TypePeriode = destDB.CreateHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode)
    
    
    Set rs3 = destDB.OpenRecordset(rq, Snapshot)
    
    
    Set rsCatR9 = destDB.OpenRecordset("SELECT * From CATR9 WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCodesCat = destDB.OpenRecordset("SELECT * From CodesCat WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCodeCatInv = destDB.OpenRecordset("SELECT * From CodeCatInv WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCatInval = destDB.OpenRecordset("SELECT * From CATR9INVAL WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    
    Set rsNCA = destDB.OpenRecordset("SELECT * From PassageNCA WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    
    'Set rsTBQREGA = destDB.OpenRecordset("SELECT * From TBQREGA WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, snapshot)
    
    'Set rsREGA01 = destDB.OpenRecordset("SELECT * From REGA01 WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    
    'Set rsTaux = destDB.OpenRecordset("SELECT * FROM TauxRenteEducation", snapshot)
    
    
    ' efface les anciennent valeurs
    'destDB.Execute "DELETE FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    
    'Load the Sexe Recordset for the function
    If PathSexFileExcel <> "" Then
      Dim SrcDB As DAO.Database
      
      Set SrcDB = OpenDatabase(PathSexFileExcel, dbDriverNoPrompt, True, cdExcelExtendedPropertiesDAO)
      SrcDB.QueryTimeout = 600
      
      Set rsSexe = SrcDB.OpenRecordset("Select * From DONNEES_LOT", dbOpenSnapshot)  'Select * From DONNEES_LOT Where AssID = 1
      'rsSexe.MoveLast
      'rsSexe.MoveFirst
    End If
    
        
       
'Dim start As Single
'start = Timer

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modif ali am du 16/10/2017
' si un traite l'import d'un SEUL lot,  on supprime ASSURE et Assure_P3IPROVCOLL uniquement pour la période courante NumPeriode
' si un traite l'import dans DEUX lots, on supprime ASSURE et Assure_P3IPROVCOLL pour la période NumPeriodeNonStat ET la période NumPeriodeStat
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''
    'Delete for NumPeriode - 1 SEUL LOT
    ''''''''''''''''''''''''''''

  If Not TwoLotImport Then
  
Delete_Assure_1lot:
    sqlStr = "DELETE TOP (1000) FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    destDB.Execute sqlStr
    
    Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode, Snapshot)
    res = rsCount(0).Value
    
    DoEvents
    'frmPeriode.lblStatus.Caption = NumPeriode & ":" + " -- Assurées à supprimer : " & res
    'DoEvents
    
    If res > 0 Then GoTo Delete_Assure_1lot
    

'    Debug.Print "Duration: " & Format(Timer - start, "0.000")
'    MsgBox "Duration " & NumPeriode & " : " & Format(Timer - start, "0.000")
'    destDB.RollbackTrans
'    Screen.MousePointer = vbDefault
'    Exit Function
  

    ' efface les anciennent valeurs
    'destDB.Execute "DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriode
  
    
Delete_P3IPROVCOLL_1lot:
    sqlStr = "DELETE TOP (1000) FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriode
    destDB.Execute sqlStr
    
    Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure_P3IPROVCOLL where CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    res = rsCount(0).Value
    
    DoEvents
    'frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
    'DoEvents
    
    If res > 0 Then GoTo Delete_P3IPROVCOLL_1lot
    
    End If
    
    ''''''''''''''''''''''''''''
    'Delete for NONstatutaire et Statutaire - Eclatement en 2 LOTS
    ''''''''''''''''''''''''''''
    
    If TwoLotImport Then
    
Delete_Assure:
    sqlStr = "DELETE TOP (1000) FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriodeNonStat
    destDB.Execute sqlStr
    
    Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriodeNonStat, Snapshot)
    res = rsCount(0).Value
    
    DoEvents
    'frmPeriode.lblStatus.Caption = NumPeriode & ":" + " -- Assurées à supprimer : " & res
    'DoEvents
    
    If res > 0 Then GoTo Delete_Assure
Delete_P3IPROVCOLL:
    sqlStr = "DELETE TOP (1000) FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriodeNonStat
    destDB.Execute sqlStr
    
    Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure_P3IPROVCOLL where CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriodeNonStat, Snapshot)
    res = rsCount(0).Value
    
    DoEvents
    'frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
    'DoEvents
    
    If res > 0 Then GoTo Delete_P3IPROVCOLL
    
   
  
Delete_Assure2:
      sqlStr = "DELETE TOP (1000) FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriodeStat
      destDB.Execute sqlStr
      
      Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriodeStat, Snapshot)
      res = rsCount(0).Value
      
      DoEvents
      
      If res > 0 Then GoTo Delete_Assure2
    
    
Delete_P3IPROVCOLL2:
      sqlStr = "DELETE TOP (1000) FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriodeStat
      destDB.Execute sqlStr
      
      Set rsCount = destDB.OpenRecordset("SELECT COUNT(*)FROM Assure_P3IPROVCOLL where CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriodeStat, Snapshot)
      res = rsCount(0).Value
      
      DoEvents
      
      If res > 0 Then GoTo Delete_P3IPROVCOLL2
        
    End If
    
    
    ' ouvre la Assure_P3IProvColl (en lecture pour la liste des champs)
    
    m_AssureP3IProvColl_SQLText = ""
    Set rsAssureP3IProvColl = destDB.OpenRecordset("SELECT * FROM Assure_P3IPROVCOLL WHERE CleGroupe=-1 AND NumPeriode=-1", Snapshot)
   
      
    ' ouvre une transaction pour ne pas ecrire les changements en cas d'erreur
    codeRetour = True
    
    ' liste des catégorie non trouvée
    nbCatNotFound = 0
    Erase lstCatNotFound
    
    nbCodesCatNotFound = 0
    Erase lstCodesCatNotFound
    
    nbCodeCatInvNotFound = 0
    Erase lstCodeCatInvNotFound
    
    ' calcul la moyenne des PSAP pour le regime 6
  'PHM 30/05/2006 : pas utilisé par Generali
    'Call computeAvgRegime6(rs, dateArrete, nbJourDC, fWait)
  'PHM 30/05/2006 : pas utilisé par Generali
    
    '*** import des données
    ' compte le nb d'article à traiter
    rs3.MoveLast
    rs3.MoveFirst
    maxRecord = rs3.RecordCount + 1
    
    If maxRecord <> 0 Then
      If Not autoMode Then
        fWait.ProgressBar1.Max = maxRecord
        'Call m_Logger.EcritTraceDansLog(maxRecord & " articles dans le fichier")
      End If
    Else
      If autoMode Then
        m_Logger.EcritTraceDansLog "Aucun article trouvé pour la période spécifiée"
      Else
        fWait.ProgressBar1.Max = 1
        fWait.Hide
        MsgBox "Aucun article trouvé pour la période spécifiée", vbInformation
      End If
    End If
    
    
  'Dim fi As Field
  '
  'Debug.Print vbCrLf & "champ de rs"
  'For Each fi In rs.Fields
  '  Debug.Print vbTab & "'" & fi.Name & "'"
  'Next fi
  
    
    ' init de parametre de rupture
    Do Until rs3.EOF
    
      If (rs3.AbsolutePosition Mod 10) = 0 Then
        If Not autoMode Then
          ' affiche la position
          fWait.ProgressBar1.Value = rs3.AbsolutePosition
          fWait.Label1(0).Caption = "Phase 2 - Article n°" & fWait.ProgressBar1.Value & " / " & fWait.ProgressBar1.Max & " - " & NbRejet & " rejets"
          fWait.Refresh
          DoEvents
        End If
      End If
      
      If Not autoMode Then
        If fWait.fTravailAnnule = True Then
          Call m_Logger.EcritTraceDansLog("Import interrompu par l'utilisateur !")
          codeRetour = False
          Exit Do
        End If
      End If
          
      If Not IsNull(rs3.Fields("Regime")) Then
      
        ' !!!!!!!! AddNewRecord parcours le recordset sur le meme assure, sinistre, garantie...
        ' Ajoute et configure un enregistrement dans Assure
        Call AddNewRecord(destDB, rs3, CleGroupe, NumPeriode, fWait, AgeRetraite, CDate(DateDebut), CDate(DateFin), nbJourMax, nbJourDC, True, rsAssureP3IProvColl)
        
      Else
        
        Call m_Logger.EcritTraceDansLog("REJET : Table TBQREGA mal configurée pour CDGARAN=Code_GE=" & rs3.Fields("GE") & ", Code_CIE=" & rs3.Fields("NumSociete") & ", Code_APP=" & rs3.Fields("CDAPPLI") & " : CodeProvision " & rs3.Fields("CodeProv") & " introuvable pour la ligne: " & rs3.Fields("RECNO"))
        NbRejet = NbRejet + 1
      
      End If
      
      rs3.MoveNext
           
    Loop
    
    If codeRetour = False Then
      ' transaction...
      destDB.RollbackTrans
    Else
      ' phase 3 : annualisation à 0 des prestations pour les morts vraiment morts!
      Call AdjustDeces(destDB, CleGroupe, NumPeriode, fWait)
      
      ' transaction...
      destDB.CommitTrans
      
      ' mets à jour la date d'extraction
      'rq = "UPDATE Periode SET PEDATEEXT = " & m_dataHelper.BuildDateString(dateArrete) & ", NUTRAITP3I=" & NumeroLot _
           & " WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE = " & NumPeriode
      rq = "UPDATE Periode SET PEDATEEXT = " & m_dataHelper.BuildDateString(dateArrete) & ", NUTRAITP3I=" & NumeroLot & " , PECALCULER = 'False' " _
           & " WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE = " & NumPeriode
      destDB.Execute rq, dbFailOnError
      
      ' Import de P3ILOGTRAIT
      CopieVersAssureP3ILogTrait CleGroupe, NumPeriode, destDB, NumeroLot, m_Logger
    End If
    
    m_Logger.EcritTraceDansLog " "
    
    Call m_Logger.EcritTraceDansLog(rs3.RecordCount & " articles dans le fichier")
    
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
    
    rsCatR9.Close
    rsCodesCat.Close
    rsCodeCatInv.Close
    rsCatInval.Close
    rsNCA.Close
    'rsTBQREGA.Close
    'rsREGA01.Close
    'rsTaux.Close
    
    rsAssureP3IProvColl.Close
    
    rs3.Close
    
    If Not rsSexe Is Nothing Then
      rsSexe.Close
    End If
    
    'PHM 30/05/2006 : pas utilisé par Generali
    '  rs.Close
    'PHM 30/05/2006 : pas utilisé par Generali
  
  Else ' Import Données ?
      
    If transOpened = True Then
      destDB.CommitTrans
      transOpened = False
    End If
    
    codeRetour = True
  
  End If
  
  ' liste des catégorie non trouvée
  ' CATR9 / CATR9INVAL
  If nbCatNotFound > 0 Then
    m_Logger.EcritTraceDansLog " "
    m_Logger.EcritTraceDansLog "Catégories à ajouter dans CATR9 ou CATR9INVAL : (ATTENTION Assuré concidéré en INVALIDITE !!!)"
    For maxRecord = 0 To nbCatNotFound - 1
      m_Logger.EcritTraceDansLog "  " & lstCatNotFound(maxRecord)
    Next
  End If
  
  nbCatNotFound = 0
  Erase lstCatNotFound
  
  ' CODESCAT
  If nbCodesCatNotFound > 0 Then
    m_Logger.EcritTraceDansLog " "
    m_Logger.EcritTraceDansLog "Catégories à ajouter dans la table 'CODESCAT' :"
    m_Logger.EcritTraceDansLog "Code_CIE - Code_APP - Code_Cat_Contrat"
    For maxRecord = 0 To nbCodesCatNotFound - 1
      m_Logger.EcritTraceDansLog "  " & lstCodesCatNotFound(maxRecord)
    Next
  End If
  
  nbCodesCatNotFound = 0
  Erase lstCodesCatNotFound
    

  ' CODECATINV
  If nbCodeCatInvNotFound > 0 Then
    m_Logger.EcritTraceDansLog " "
    m_Logger.EcritTraceDansLog "CDCHOIXPREST à ajouter dans la table 'CODECATINV' :"
    For maxRecord = 0 To nbCodeCatInvNotFound - 1
      m_Logger.EcritTraceDansLog "  " & lstCodeCatInvNotFound(maxRecord)
    Next
  End If
  
  nbCodeCatInvNotFound = 0
  Erase lstCodeCatInvNotFound
  
  
  ' Date d'import
  destDB.Execute ("UPDATE Periode SET DateDernierImport=GetDate(), DateDernierControle=NULL WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode)
  
  
  If Not autoMode Then
    fWait.Hide
    Unload fWait
  End If
  
  Set fWait = Nothing
  
  If Not autoMode Then
    Screen.MousePointer = vbDefault
    m_Logger.AfficheErreurLog
  End If
  
  iP3IGeneraliImport_DoImportSASP3I = codeRetour
  
  Set m_dataHelper = Nothing
  
  Exit Function
  
GestionErreur:
  If rs3 Is Nothing Then
    If Err = 3228 Then Resume
    
    If Not autoMode Then
      If Not fWait Is Nothing Then
        fWait.Hide
        Unload fWait
        Set fWait = Nothing
      End If
    End If
      
    m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
    
    If transOpened = True Then
      destDB.RollbackTrans
      transOpened = False
    End If
    
  Else
    If Not autoMode Then
      fWait.Hide
      Unload fWait
      Set fWait = Nothing
    End If
    
    If autoMode Then
      m_Logger.EcritTraceDansLog "Erreur pendant l'import : " & Err.Number & " : " & Err.Description
    Else
      Screen.MousePointer = vbDefault
      MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
    End If
    
    'Resume Next
    ' efface les valeurs incompletes
'    destDB.Execute "DELETE FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    If transOpened = True Then
      destDB.RollbackTrans
      transOpened = False
    End If
    
   
    iP3IGeneraliImport_DoImportSASP3I = False
    
    Set m_dataHelper = Nothing
  End If

  Exit Function
  
  Resume Next

End Function

'##ModelId=5C90E978034B
Public Function iP3IGeneraliImport_SetStatutaireVariables(mNumPeriodeStat As Long, mNumPeriodeNonStat As Long, mPathSexFileExcel As String, mCategoryCodeSTAT As String, mSexAllMale As Boolean, mTwoLotImport As Boolean) As Boolean
       
    NumPeriodeStat = mNumPeriodeStat
    NumPeriodeNonStat = mNumPeriodeNonStat
    PathSexFileExcel = mPathSexFileExcel
    CategoryCodeSTAT = mCategoryCodeSTAT
    SexAllMale = mSexAllMale
    TwoLotImport = mTwoLotImport
       
    iP3IGeneraliImport_SetStatutaireVariables = True

End Function

