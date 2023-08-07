VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iImportGeneraliRevalo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C90E97A0041"
Option Explicit
 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Feuille excel verrouillée par mot de passe : 'import'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements iP3IGeneraliImport

'##ModelId=5C90E97A0204
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

'##ModelId=5C90E97A020B
Private autoMode As Boolean

' nb de rejet lors de l'import
'##ModelId=5C90E97A022B
Private NbRejet As Long
'##ModelId=5C90E97A0256
Private dateImport As Date
'##ModelId=5C90E97A028F
Private dateArrete As Date
'##ModelId=5C90E97A02B0
Private avgPSAP_Regime6 As Double
'##ModelId=5C90E97A02E8
Private typeImport As eTypeImport
'##ModelId=5C90E97A02F5
Private typeDelaiInactivite As eTypeDelaiInactivite
'##ModelId=5C90E97A0302
Private typeCalculAnnualisation As eTypeCalculAnnualisation

'##ModelId=5C90E97A030F
Private rsCatR9 As ADODB.Recordset
'##ModelId=5C90E97A031B
Private rsCatInval As ADODB.Recordset
'##ModelId=5C90E97A032C
Private rsNCA As ADODB.Recordset
'##ModelId=5C90E97A0339
Private rsCodesCat As ADODB.Recordset
'##ModelId=5C90E97A0346
Private rsCodeCatInv As ADODB.Recordset
'##ModelId=5C90E97A0353
Private rsTBQREGA As ADODB.Recordset ', rsREGA01 As ADODB.Recordset
'##ModelId=5C90E97A035F
Private rsTaux As ADODB.Recordset
'##ModelId=5C90E97A036C
Private rsCodeProvision As ADODB.Recordset

'##ModelId=5C90E97A0372
Private nbCatNotFound As Long
'##ModelId=5C90E97A0394
Private lstCatNotFound() As String

'##ModelId=5C90E97A03B5
Private nbCodesCatNotFound As Long
'##ModelId=5C90E97A03D6
Private lstCodesCatNotFound() As String

'##ModelId=5C90E97B001D
Private nbCodeCatInvNotFound As Long
'##ModelId=5C90E97B0052
Private lstCodeCatInvNotFound() As String

'##ModelId=5C90E97B008A
Private m_dataHelper As DataHelper

'##ModelId=5C90E97B0090
Private m_TypePeriode As Integer


' Valeurs par défaut pour remplacer celles de l'import qui ne sont pas
'spécifiées (import format compact)
'##ModelId=5C90E97B00BD
Private Const SECTION_IMPORTXLSREVALO As String = "ImportXLSRevalo"

'##ModelId=5C90E97B00FB
Private def_Code_CIE As Variant
'##ModelId=5C90E97B011E
Private def_Code_APP As Variant
'##ModelId=5C90E97B0153
Private def_Contractant  As Variant
'##ModelId=5C90E97B0190
Private def_Sexe  As Variant
'##ModelId=5C90E97B01C3
Private def_SituConv As Variant
'##ModelId=5C90E97B01F8
Private def_EffetSitu As Variant
'##ModelId=5C90E97B022F
Private def_EffetConv As Variant
'##ModelId=5C90E97B0255
Private def_CodeOption  As Variant
'##ModelId=5C90E97B0284
Private def_Etablissement As Variant
'##ModelId=5C90E97B02B9
Private def_TypeMvt As Variant
'##ModelId=5C90E97B02DE
Private def_IndemnitesJournalieres As Variant
'##ModelId=5C90E97B0303
Private def_BaseRevalo  As Variant
'##ModelId=5C90E97B033C
Private def_MntRevalo  As Variant
'##ModelId=5C90E97B0377
Private def_MntMajore  As Variant
'##ModelId=5C90E97B039F
Private def_MntRegle  As Variant
'##ModelId=5C90E97B03C5
Private def_TypeRglt  As Variant
'##ModelId=5C90E97C0001
Private def_NbJIndemnises As Variant
'##ModelId=5C90E97C0027
Private def_datePaiement As Variant
'##ModelId=5C90E97C004D
Private def_CreationDossier As Variant
'##ModelId=5C90E97C008A
Private def_FinDossier As Variant
'##ModelId=5C90E97C00B3
Private def_MotifCloture As Variant
'##ModelId=5C90E97C00E9
Private def_SalBrut As Variant
'##ModelId=5C90E97C0113
Private def_DebRefSal As Variant
'##ModelId=5C90E97C013C
Private def_FinRefSal As Variant
'##ModelId=5C90E97C0170
Private def_Code_Cat_Contrat  As Variant
'##ModelId=5C90E97C0197
Private def_RegrStat As Variant
'##ModelId=5C90E97C01D9
Private def_RegrAnnexe  As Variant
'##ModelId=5C90E97C0203
Private def_FamComptable  As Variant
'##ModelId=5C90E97C022A
Private def_Inspecteur As Variant
'##ModelId=5C90E97C0251
Private def_Insp As Variant
'##ModelId=5C90E97C0286
Private def_Apport As Variant
'##ModelId=5C90E97C02BC
Private def_Apport2 As Variant
'##ModelId=5C90E97C02F6
Private def_Gestionnaire As Variant
'##ModelId=5C90E97C032B
Private def_IndicCC As Variant
'##ModelId=5C90E97C036C
Private def_Cause As Variant
'##ModelId=5C90E97C0392
Private def_Franchise As Variant
'##ModelId=5C90E97C03BB
Private def_Compagnie As Variant
'##ModelId=5C90E97C03E3
Private def_NbIntervenant As Variant
'##ModelId=5C90E97D0031
Private def_CDCONTENTIEUX As Variant
'
' Valeurs déduites d'autres champs
' Convention=<Sinistre>
' Dossier=<Sinistre>
' DateEntreeInvalidite=<Survenance>
' MntBase=<Annualisation>
' DebIndemn=<Survenance>
' DebDossier=<Survenance>
'
'


'
' lecture des valeurs par défaut depuis P3I.INI
'
'##ModelId=5C90E97D005F
Private Function ReadDefValue(entry As String, sFichierIni As String) As Variant
  
  ReadDefValue = sReadIniFile(SECTION_IMPORTXLSREVALO, entry, "###", 255, sFichierIni)
  If ReadDefValue = "###" Or UCase(ReadDefValue) = "NULL" Then
    ReadDefValue = Null
  End If
   
End Function


'##ModelId=5C90E97D0108
Private Sub InitDefaultValues(sFichierIni As String)
  
  def_Code_CIE = ReadDefValue("Code_CIE", sFichierIni)
  def_Code_APP = ReadDefValue("Code_APP", sFichierIni)
  def_Contractant = ReadDefValue("Contractant", sFichierIni)
  def_Sexe = ReadDefValue("Sexe", sFichierIni)
  def_SituConv = ReadDefValue("SituConv", sFichierIni)
  def_EffetSitu = ReadDefValue("EffetSitu", sFichierIni)
  def_EffetConv = ReadDefValue("EffetConv", sFichierIni)
  def_CodeOption = ReadDefValue("CodeOption", sFichierIni)
  def_Etablissement = ReadDefValue("Etablissement", sFichierIni)
  def_TypeMvt = ReadDefValue("TypeMvt", sFichierIni)
  def_IndemnitesJournalieres = ReadDefValue("IndemnitesJournalieres", sFichierIni)
  def_BaseRevalo = ReadDefValue("BaseRevalo", sFichierIni)
  def_MntRevalo = ReadDefValue("MntRevalo", sFichierIni)
  def_MntMajore = ReadDefValue("MntMajore", sFichierIni)
  def_MntRegle = ReadDefValue("MntRegle", sFichierIni)
  def_TypeRglt = ReadDefValue("TypeRglt", sFichierIni)
  def_NbJIndemnises = ReadDefValue("NbJIndemnises", sFichierIni)
  def_datePaiement = ReadDefValue("datePaiement", sFichierIni)
  def_CreationDossier = ReadDefValue("CreationDossier", sFichierIni)
  def_FinDossier = ReadDefValue("FinDossier", sFichierIni)
  def_MotifCloture = ReadDefValue("MotifCloture", sFichierIni)
  def_SalBrut = ReadDefValue("SalBrut", sFichierIni)
  def_DebRefSal = ReadDefValue("DebRefSal", sFichierIni)
  def_FinRefSal = ReadDefValue("FinRefSal", sFichierIni)
  def_Code_Cat_Contrat = ReadDefValue("Code_Cat_Contrat", sFichierIni)
  def_RegrStat = ReadDefValue("RegrStat", sFichierIni)
  def_RegrAnnexe = ReadDefValue("RegrAnnexe", sFichierIni)
  def_FamComptable = ReadDefValue("FamComptable", sFichierIni)
  def_Inspecteur = ReadDefValue("Inspecteur", sFichierIni)
  def_Insp = ReadDefValue("Insp", sFichierIni)
  def_Apport = ReadDefValue("Apport", sFichierIni)
  def_Apport2 = ReadDefValue("Apport2", sFichierIni)
  def_Gestionnaire = ReadDefValue("Gestionnaire", sFichierIni)
  def_IndicCC = ReadDefValue("IndicCC", sFichierIni)
  def_Cause = ReadDefValue("Cause", sFichierIni)
  def_Franchise = ReadDefValue("Franchise", sFichierIni)
  def_Compagnie = ReadDefValue("Compagnie", sFichierIni)
  def_NbIntervenant = ReadDefValue("NbIntervenant", sFichierIni)
  def_CDCONTENTIEUX = ReadDefValue("CDCONTENTIEUX", sFichierIni)
  
End Sub


'##ModelId=5C90E97D0148
Public Function Arrondi(Valeur As Double, Nbdecimale As Integer) As Double ' fonction arrondi
  Nbdecimale = Abs(Nbdecimale)
  Arrondi = Fix((Valeur * (10 ^ Nbdecimale) + 0.5)) / (10 ^ Nbdecimale)
End Function

'##ModelId=5C90E97D01B1
Public Function BuildDateLimit(debut As String, Fin As String) As String
  Dim Limit As String
  
  Limit = "BETWEEN #" & Format(CDate(debut), "mm/dd/yyyy") & "# AND #" & Format(CDate(Fin), "mm/dd/yyyy") & "#"
  
  BuildDateLimit = Limit
End Function


'
' recherche du code provision correspondant à une garantie.
' on peut préciser le Lib_Court_GE (texte), le Code_GE (n° interne à
'Generali), le Code_PROV (code provision de P3I)
'
'##ModelId=5C90E97D020D
Private Function FindCodeProvision(Garantie As String) As Integer
  
  m_dataHelper.Multi_Find rsTBQREGA, "Lib_Court_GE='" & Garantie & "'"
  
  If Not rsTBQREGA.EOF Then
    
    FindCodeProvision = rsTBQREGA.Fields("Code_PROV")

  Else
    
    m_dataHelper.Multi_Find rsTBQREGA, "Code_GE='" & Garantie & "'"
    
    If rsTBQREGA.EOF Then
      
      m_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'TBQREGA' : " & Garantie
      FindCodeProvision = -1
    
    Else
      
      FindCodeProvision = rsTBQREGA.Fields("Code_PROV")
    
    End If
  
  End If

End Function


'
' recherche du regime correspondant à une garantie.
' on peut préciser le Lib_Court_GE (texte), le Code_GE (n° interne à
'Generali), le Code_PROV (code provision de P3I)
'
'##ModelId=5C90E97D025B
Private Function FindRegime(Garantie As String) As Integer
  
  Dim code_prov As Integer
  
  
  code_prov = FindCodeProvision(Garantie)
  
  If code_prov = -1 Then
    FindRegime = -1
    Exit Function
  End If
  
  
  m_dataHelper.Multi_Find rsCodeProvision, "CodeProv=" & code_prov
  
  If Not rsCodeProvision.EOF Then
    
    FindRegime = rsCodeProvision.Fields("Regime")

  Else
    
    m_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'TBQREGA' : " & Garantie
    FindRegime = -1
  
  End If

End Function


'
' recherche du numero de parametre de calcul en fonction du Code_Cat_Contrat
'
'##ModelId=5C90E97D02AC
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
    
    FindNumParamCalcul = -1
    'm_Logger.EcritTraceDansLog "Garantie Elémentaire à ajouter dans la table 'CodesCat' : Code_CIE=" & Code_CIE & " AND Code_APP=" & Code_APP & " AND Code_Cat_Contrat='" & Code_Cat_Contrat & "'"

  End If

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recherche la catégorie inval
'
'##ModelId=5C90E97D034A
Private Function FindCategorieInval(CDCHOIXPREST As String) As Integer
  
  rsCodeCatInv.Filter = adFilterNone
  rsCodeCatInv.Filter = "CDCHOIXPREST='" & Trim(CDCHOIXPREST) & "'"

  If rsCodeCatInv.EOF = False Then

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
' la ligne doit elle etre prise en compte dans la période
'
'##ModelId=5C90E97D0390
Private Function IsInPeriode(rs As DAO.Recordset, DateDebut As Date, DateFin As Date, nbJourDC As Long, curRegime As Long) As Boolean
  Dim dDate As Variant, dSurv As Variant
  
  IsInPeriode = False
  
  'dDate = ReadDateDAO(rs.Fields("DatePaiement"))
  dDate = ReadDateDAO(rs.Fields("Survenance"))
  dSurv = ReadDateDAO(rs.Fields("Survenance"))
  
  If IsNull(dDate) Then
    
    ' Deces
    If curRegime = cdGarDeces_Import Then
      If (dSurv + nbJourDC) >= DateDebut And (dSurv + nbJourDC) <= DateFin Then
        IsInPeriode = True
      End If
    End If
    
    ' Arret de travail
    If curRegime = cdGarIncapInval_Import Then
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
  If curRegime = cdGarDeces_Import Then
    IsInPeriode = True
  End If
End Function


'
' Recherche dans la table CATR9 si cette catégorie d'incap contient une
'garantie passage en inval
'
'##ModelId=5C90E97E006B
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


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' on regroupe les paiements par sinistre. la date de debut est celle du
'premier paiement
' et la date de fin est celle du dernier paiement
'
'##ModelId=5C90E97E00F4
Private Sub AddNewRecord(destDB As DataAccess, rs As DAO.Recordset, CleGroupe As Long, NumPeriode As Long, fWait As Form, AgeRetraite As Byte, DateDebut As Date, DateFin As Date, nbJourMax As Long, nbJourDC As Long, bFromInfocentre As Boolean, bImportForRevalo As Boolean)
  
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
          

' ...pour debug...
'  FieldName = "NomAssure"
'  If Not IsNull(rs.Fields("NomAssure")) Then
'    If InStr(rs.Fields("NomAssure"), "ASSURE4083") <> 0 Then
'      Stop
'    End If
'  End If
' ...pour debug...
  
  
  ' n° de jeu de parametres de calcul
  numParamCalcul = 0
  If Not IsNull(rs.Fields("ParamCalcul")) Then
    numParamCalcul = rs.Fields("ParamCalcul")
  Else
    If bImportForRevalo Then
      numParamCalcul = FindNumParamCalcul(rs.Fields("COMPAGNIE"), rs.Fields("APPLI"), Trim(rs.Fields("Code_Cat_Contrat")))
    Else
      numParamCalcul = FindNumParamCalcul(rs.Fields("COMPAGNIE"), rs.Fields("APPLI"), Trim(rs.Fields("Code_Cat_Contrat")))
    End If
  End If
  
  
  ' Portefeuille Salaries ?
  bPortefeuilleSalarie = destDB.CreateHelper.GetParameterAsDouble("SELECT PEPortefeuilleSalaries FROM ParamCalcul WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode & " AND PENUMPARAMCALCUL=" & numParamCalcul)
  
  
  ' Garantie = Code_GE
  FieldName = "Garantie"
  If IsNull(rs.Fields("Garantie")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Garantie invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curGE = Trim(rs.Fields("Garantie"))
  End If
  
  
  ' Regime deduit de Lib_Court_GE
  FieldName = "Regime"
  curRegime = FindRegime(curGE)
  If curRegime = -1 Then
    Call m_Logger.EcritTraceDansLog("REJET : Garantie invalide (impossible d'en déduire le regime) à la ligne : " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
  
  inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC, curRegime)
  
  
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
  
  
 
  
' DEBUG DEBUG
' DEBUG DEBUG
' DEBUG DEBUG
  
  'If rs.Fields("RECNO") = 162 Then Stop

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
      ' premierPaiement = ReadDateXL(rs.Fields("DatePaiement"))
      premierPaiement = Null
    End If
    
    FieldName = "PeriodeDu"
    poDebut = ReadDateXL(rs.Fields("PeriodeDu"))
    bInit = True
  Else
    inPeriode = False ' test debug
  End If
  
  FieldName = "PeriodeDu"
  poDebutTotal = ReadDateXL(rs.Fields("PeriodeDu"))
  
  ' init des variables
  FieldName = "NumSociete"
  curSte = def_Code_CIE
  
  FieldName = "Sinistre"
  If IsNull(rs.Fields("Sinistre")) Then
    Call m_Logger.EcritTraceDansLog("REJET : Sinistre invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curSS = Trim(rs.Fields("Sinistre"))
  End If
    
  FieldName = "Code_Cat_Contrat"
  curCat = def_Code_Cat_Contrat
  
  FieldName = "Convention"
  curNCA = curSS
  
  FieldName = "Survenance"
  dDate = ReadDateXL(rs.Fields("Survenance"))
  If IsNull(dDate) Then
    Call m_Logger.EcritTraceDansLog("REJET : Survenance invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    curArret = rs.Fields("Survenance")
  End If
  
  FieldName = "BaseRevalo"
  curBaseRevalo = def_BaseRevalo
  
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
      
    If curRegime = cdGarRente_Import Then
      inPeriode = True
    Else
      inPeriode = IsInPeriode(rs, DateDebut, DateFin, nbJourDC, curRegime)
    End If
    
    
    ' test si toujours le meme assuré
    If curSS <> Trim(rs.Fields("Sinistre")) Or curGE <> rs.Fields("Garantie") _
       Or curNCA <> Trim(rs.Fields("Sinistre")) Or curArret <> rs.Fields("Survenance") Then
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
        ' premierPaiement = ReadDateXL(rs.Fields("DatePaiement"))
        premierPaiement = Null
      End If
      
      FieldName = "PeriodeDu"
      poDebut = ReadDateXL(rs.Fields("PeriodeDu"))
      
      bInit = True
    End If
    
    ' toujour le meme sinistre
    FieldName = "MntRegle"
    tmpMontant = def_MntRegle
    
    ' cumul des prestations payées presente dans le fichier
    montantTotal = montantTotal + tmpMontant
    
    'If inPeriode Then PHM 9/1/04 : on tient toujours compte des deces
    If inPeriode Then
      nbJPer = 0
      
      If curRegime = cdGarRente_Import Then
        nbJPer = 365
      Else
        FieldName = "PeriodeAu"
        If IsDate(ReadDateXL(rs.Fields("PeriodeDu"))) And IsDate(ReadDateXL(rs.Fields("PeriodeAu"))) Then
          If CDate(rs.Fields("PeriodeDu")) > CDate(rs.Fields("PeriodeAu")) Then
            nbJPer = 0
            Call m_Logger.EcritTraceDansLog("!!!!!! PeriodeDu > PeriodeAu à la ligne: " & rs.Fields("RECNO"))
          Else
            nbJPer = DateDiff("d", rs.Fields("PeriodeDu"), CDate(rs.Fields("PeriodeAu"))) + 1
          End If
        End If
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
      FieldName = "MntRevalo"
      tmpMontant = def_MntRevalo
      MontantRevalo = MontantRevalo + tmpMontant
      
      ' Montant Majore
      FieldName = "MntMajore"
      tmpMontant = def_MntMajore
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
      If IsDate(ReadDateXL(rs.Fields("PeriodeDu"))) And IsDate(ReadDateXL(rs.Fields("PeriodeAu"))) Then
        Dim nbJ As Long
        
        nbJ = DateDiff("d", (rs.Fields("PeriodeDu")), (rs.Fields("PeriodeAu")))
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
        ' dernierPaiement = ReadDateXL(rs.Fields("DatePaiement"))
        dernierPaiement = Null
      End If
      
      FieldName = "PeriodeAu"
      poFin = ReadDateXL(rs.Fields("PeriodeAu"))

' PHM 01/03/2004
    Else
      ' pour les régimes 6, les prestas payées apres la date de fin sont comptabilisés comme PSAP
      If curRegime = cdGarDeces_Import Then
        Dim datePaiement As Date
        
        FieldName = "DatePaiement"
        'dDate = ReadDateXL(rs.Fields("DatePaiement"))
        dDate = Null
        If IsNull(dDate) Then
          datePaiement = rs.Fields("Survenance") + nbJourDC
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
        ' dernierPaiement = ReadDateXL(rs.Fields("DatePaiement"))
        dernierPaiement = Null
        
        FieldName = "PeriodeAu"
        poFin = ReadDateXL(rs.Fields("PeriodeAu"))
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
    If (rs.AbsolutePosition Mod 9) = 0 Then
      If Not autoMode Then
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
      dDate = def_datePaiement
      If Not IsNull(dDate) Then
        dernierPaiement = dDate
      End If
    End If
    
    FieldName = "PeriodeAu"
    poFin = ReadDateXL(rs.Fields("PeriodeAu"))
  End If
  
  FieldName = "PeriodeAu"
  poFinTotal = ReadDateXL(rs.Fields("PeriodeAu"))



  ' store data
  Dim theAssure As clsAssure
  Set theAssure = New clsAssure
  
  
  theAssure.m_NUENRP3I = rs.Fields("RECNO")
  
    
  
  ' calcul de la date de paiement pour le regime 6 : il n'y a qu'UN seul paiement pour ces regimes
  If curRegime = cdGarDeces_Import And (IsNull(premierPaiement) Or IsNull(dernierPaiement)) Then
    FieldName = "Survenance"
    premierPaiement = rs.Fields("Survenance") + nbJourDC
    dernierPaiement = premierPaiement
    theAssure.m_PODATEPAIEMENTESTIMEE = True
  End If
  
  theAssure.m_POGPECLE = CleGroupe
  theAssure.m_POPERCLE = NumPeriode
  
  FieldName = "NumSociete"
  theAssure.m_POSTECLE = def_Code_CIE
  theAssure.m_POCompagnie = def_Code_CIE
  
  If curSS Then
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
  
  If Trim(curCat) <> "" Then
    theAssure.m_POCATEGORIE = curCat
  Else
    theAssure.m_POCATEGORIE = Null
 End If
  
  FieldName = "NomAssure"
  If Not IsNull(rs.Fields("NomAssure")) Then
    theAssure.m_PONOM = Trim(rs.Fields("NomAssure"))
    If Not IsNull(rs.Fields("PrenomAssure")) Then
      theAssure.m_PONOM = theAssure.m_PONOM & " " & Trim(rs.Fields("PrenomAssure"))
    End If
  Else
    Call m_Logger.EcritTraceDansLog("REJET : Nom invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If
  
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
     
    
  ' date de naissance
  FieldName = "DateNaiss"
  dDate = ReadDateXL(rs.Fields("DateNaiss"))
  If IsNull(dDate) Then
'    ' calcul en fct du n° de SS
'    FieldName = "NumeroSS"
'    dateNaissance = DateSerial(Mid(rs.Fields("NumeroSS"), 2, 2), Mid(rs.Fields("NumeroSS"), 4, 2), 1)
    Call m_Logger.EcritTraceDansLog("REJET : DateNaiss invalide à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  Else
    dateNaissance = rs.Fields("DateNaiss")
  End If

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
  theAssure.m_POEFFET = def_EffetConv
  
  ' Fin garantie
  FieldName = "FinGarantie"
  If theAssure.m_POTERME = Null Then
    theAssure.m_POTERME = ReadDateXL(rs.Fields("FinGarantie"))
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
  theAssure.m_POCAUSE = def_Cause
  
  ' Franchise
  FieldName = "Franchise"
  theAssure.m_POTYPEF = def_Franchise
  
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
  theAssure.m_POCSP = ""
  
  ' montant de la prestation
  theAssure.m_POPRESTATION = montant ' non recalculé
  
  If curRegime = cdGarRente_Import Then
    ' Rentes
    
    theAssure.m_POPRESTATION_AN = montantAnnualisation ' pour les rentes le montant est directement le montant annualisé
  
'  ElseIf bFromInfocentre = True And curRegime = cdGarIncapInval_Import Then
  ElseIf curRegime = cdGarIncapInval_Import Then
      
      If montantAnnualisation = 0 Then
        ' Montant Base
        tmpMontant = 0
        FieldName = "MntBase"
        If Not IsNull(rs.Fields("MntBase")) Then
          tmpMontant = rs.Fields("MntBase")
        End If
        MontantBase = MontantBase + tmpMontant
        
        ' Montant Revalo
        tmpMontant = def_MntRevalo
        FieldName = "MntRevalo"
        MontantRevalo = MontantRevalo + tmpMontant
        
        ' Montant Majore
        tmpMontant = def_MntMajore
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
        tmpMontant = def_MntRegle
        FieldName = "MntRegle"
        
        ' cumul des prestations payées presente dans le fichier
        montantTotal = tmpMontant
      End If
    
      theAssure.m_POPRESTATION_AN = montantAnnualisation ' le montant est directement le montant annualisé
  
  Else
    
    ' autre
    If nbJAnnualisation <= 0 And curRegime <> cdGarDeces_Import Then
      
      theAssure.m_POPRESTATION_AN = 0
      
      If curRegime <> cdGarDeces_Import Then
        If IsNull(poDebut) And IsNull(poFin) And IsNull(premierPaiement) And IsNull(dernierPaiement) And IsNull(rs.Fields("DebIndemn")) Then
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
    theAssure.m_POIJ = def_IndemnitesJournalieres
  End If
  
  If IsNull(poDebut) And IsNull(poFin) And IsNull(premierPaiement) And IsNull(dernierPaiement) And IsNull((rs.Fields("Survenance"))) Then
    
    If bFromInfocentre = False Then

''      'If IsNull(rsSalarie.Fields("POIJ")) Then
''      If IsNull(theAssure.m_POIJ) Then
''        'rsSalarie.Fields("POPRESTATION_AN") = 0
''        theAssure.m_POPRESTATION_AN = 0
''        Call m_Logger.EcritTraceDansLog("   Annual.= 0 car IndemnitesJournalieres non renseignées à la ligne: " & rs.Fields("RECNO"))
''      Else
''        'rsSalarie.Fields("POPRESTATION_AN") = 365 * rsSalarie.Fields("POIJ")
''        theAssure.m_POPRESTATION_AN = 365 * theAssure.m_POIJ
''        'Call m_Logger.EcritTraceDansLog("   Annual.= 365 * IndemnitesJournalieres = " & rsSalarie.Fields("POPRESTATION_AN") & " à la ligne: " & rs.Fields("RECNO"))
''        Call m_Logger.EcritTraceDansLog("   Annual.= 365 * IndemnitesJournalieres = " & theAssure.m_POPRESTATION_AN & " à la ligne: " & rs.Fields("RECNO"))
''      End If

    End If
    
  Else
  
    FieldName = "PeriodeAu"
    dDate = ReadDateXL(rs.Fields("PeriodeAu"))
    If IsNull(dDate) Then
      If theAssure.m_POGARCLE <> cdGar56 And theAssure.m_POGARCLE <> cdGarRente Then
        Call m_Logger.EcritTraceDansLog("Fin de Période invalide à la ligne: " & rs.Fields("RECNO"))
      End If
    End If
  
  End If
  
  FieldName = "IndicCC"
  Dim bContentieux As Boolean, sIndicCC As Variant
  
  sIndicCC = def_IndicCC
  If sIndicCC <> Null Then
    bContentieux = IIf(sIndicCC = "O" Or sIndicCC = "Y" Or sIndicCC = "0", True, False)
  Else
    bContentieux = False
  End If
  
  Dim bKeep As Boolean
  
  If typeDelaiInactivite = eDateFinPeriodePaiement Then
    
    dDate = (rs.Fields("PeriodeAu"))
    If Not IsNull(dDate) Then
      nbJPer = DateDiff("d", dDate, dateArrete) + 1
      If nbJPer > nbJourMax And bContentieux = False Then 'And (rsSalarie.Fields("POGARCLE") = cdGar57 Or rsSalarie.Fields("POGARCLE") = cdGar59) Then
        
        bKeep = True
        
        CodeProvision = FindCodeProvision(rs.Fields("Garantie"))
        
        If curRegime = cdGarIncapInval_Import And (CodeProvision = cdPositImport_IncapProf Or CodeProvision = cdPositImport_IncapNonProf) And bPortefeuilleSalarie Then
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
      
      CodeProvision = FindCodeProvision(rs.Fields("Garantie"))
      
      If curRegime = cdGarIncapInval_Import And (CodeProvision = cdPositImport_IncapProf Or CodeProvision = cdPositImport_IncapNonProf) And bPortefeuilleSalarie Then
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
  
  'If rs.Fields("RECNO") = 100 Then Stop
  
    ' autre regime
    If psap <> 0 Then
      theAssure.m_POPSAP = psap
    Else
      dDate = ReadDateXL(rs.Fields("PeriodeAu"))
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
  dDate = ReadDateXL(rs.Fields("FinGarantie"))
  If Not IsNull(dDate) Then
    theAssure.m_POTERME = dDate
  End If
    
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
  
'  If rs.Fields("RECNO") = 5153 Then Stop
'  If rs.Fields("RECNO") = 5154 Then Stop
  
  
  
  ' position
'* rsSalarie.Fields("POSIT") = 1 ' incap avec passage
'* rsSalarie.Fields("POSIT") = 2 ' inval
'* rsSalarie.Fields("POSIT") = 3 ' incap sans passage
'* rsSalarie.Fields("POSIT") = 6 ' deces
'* rsSalarie.Fields("POSIT") = 7 ' maternite
'* rsSalarie.Fields("POSIT") = 20 ' rente conjoint
  Dim bTemporaire As Boolean, dDateLimite As Variant
  
  dDateLimite = ReadDateXL(rs.Fields("FinGarantie"))
  
  Select Case theAssure.m_POGARCLE
    ' Rentes
    Case cdGarRente
      theAssure.m_POSIT = cdPosit_RenteAutreViagere ' par defaut : autre rente
  
      ' recherche le type de rente
      CodeProvision = FindCodeProvision(rs.Fields("Garantie"))
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
          Call m_Logger.EcritTraceDansLog("REJET : Code_PROV " & CodeProvision & " (GE='" & rs.Fields("Garantie") & "') non supporté à la ligne: " & rs.Fields("RECNO"))
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
      theAssure.m_POSIT = cdPosit_Inval ' inval par defaut voir
      
      ' recherche du type de garantie
      CodeProvision = FindCodeProvision(rs.Fields("Garantie"))
      If CodeProvision < 1 Then
        Call m_Logger.EcritTraceDansLog("REJET : Code_Prov < 0 invalide à la ligne: " & rs.Fields("RECNO"))
        NbRejet = NbRejet + 1
        Exit Sub
      End If
      
      Select Case CodeProvision
        ' Incap
        Case cdPositImport_IncapProf, cdPositImport_IncapNonProf, cdProvision_Incap_AvecPassage_Viager
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
        Case cdPositImport_InvalProf, cdPositImport_InvalNonProf, cdProvision_Inval_Viager
          theAssure.m_POSIT = cdPosit_Inval
          
          
          ' Décodage de CDCHOIXPREST pour Catégori Inval
          FieldName = "CDCHOIXPREST"
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
          Call m_Logger.EcritTraceDansLog("REJET : Code_PROV " & CodeProvision & " (GE='" & rs.Fields("Garantie") & "') non supporté à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
      End Select
  End Select
  
  
  ' PSAP pour les codes provision 1 à 4
  ' fonction -> DESACTIVEE LE 04/02/2018 AM
  
'  Select Case CodeProvision
'    Case cdPositImport_IncapProf, cdPositImport_IncapNonProf, cdPositImport_InvalProf, cdPositImport_InvalNonProf
'      dDate = ReadDateXL(rs.Fields("PeriodeAu"))
'      If Not IsNull(dDate) Then
'
'        d2 = theAssure.m_POPRESTATION_AN * DateDiff("d", dDate, dateArrete) / 365.25
'
'        theAssure.m_POPSAP = Arrondi(d2, 2)
'      Else
'        dDate = ReadDateXL(rs.Fields("Survenance"))
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
  dDate = ReadDateXL(rs.Fields("FinGarantie"))
  If Not IsNull(dDate) Then
    theAssure.m_POTERME = dDate
  End If
  
  'Modif 11/7/2018
  If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
  End If
  
  ' PHM 01/07/2009
  
    
  ' valeurs par défaut = Male puis importées
  FieldName = "Sexe"
  theAssure.m_POSEXE = def_Sexe
  
  
  ' PHM 26/04/2010
  If theAssure.m_POGARCLE = cdGarRente Then
    theAssure.m_POSEXE = "2" ' pour les rentes, on force à Femme
  End If
  ' PHM 26/04/2010
  
  
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
  If Not IsNull(rs.Fields("Garantie")) Then
    theAssure.m_POGARCLE_NEW = rs.Fields("Garantie")
  End If
  
  FieldName = "CodeProvision"
  theAssure.m_POCATEGORIE_NEW = CodeProvision
  
  If bFromInfocentre = False Then
    FieldName = "RegrAnnexe"
    theAssure.m_POREGROUPEMENT = def_RegrAnnexe
    
    FieldName = "RegrStat"
    theAssure.m_POCODENATURE = def_RegrStat
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
  If theAssure.m_POSIT = cdPosit_Inval Then
    
    FieldName = "DateEntreeInvalidite"
    dDate = ReadDateXL(rs.Fields("DateEntreeInvalidite"))
    If IsDate(dDate) Then
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
  
  Else
    theAssure.m_POEFFET = ReadDateXL(rs.Fields("Survenance"))
  End If
  
' PHM 14/06/2006
'  If rsSalarie.Fields("POSIT") = cdPosit_Inval And IsNull(rsSalarie.Fields("PODATEENTREEINVAL")) Then
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
 
 
  FieldName = "SalBrut"
  theAssure.m_POSalaireAnnuel = def_SalBrut
  
  FieldName = "Cadre"
  theAssure.m_POIsCadre = False

  
' PHM 30/05/2006 : les nouveaux champs générali
  FieldName = "Contractant"
  theAssure.m_POContractant = def_Contractant
  
  ' CDSITUASIN stocké dans POSituConv
  FieldName = "SituConv"
  theAssure.m_POSituConv = def_SituConv
  
  
  If bFromInfocentre = False Then
    
    FieldName = "EffetSitu"
    theAssure.m_POEffetSitu = def_EffetSitu
    
    FieldName = "Etablissement"
    theAssure.m_POEtablissement = def_Etablissement
    
    FieldName = "TypeRglt"
    theAssure.m_POTypeReglement = def_TypeRglt
    
    FieldName = "CreationDossier"
    theAssure.m_POCreationDossier = def_CreationDossier
    
    FieldName = "DebDossier"
    theAssure.m_PODebutDossier = rs.Fields("Survenance")
    
    FieldName = "FinDossier"
    theAssure.m_POFinDossier = def_FinDossier
    
    FieldName = "MotifCloture"
    theAssure.m_POMotifCloture = def_MotifCloture
    
    FieldName = "DebRefSal"
    theAssure.m_PODebutRefSalaire = def_DebRefSal
    
    FieldName = "FinRefSal"
    theAssure.m_POFinRefSalaire = def_FinRefSal
  End If
  
  FieldName = "CodeOption"
  theAssure.m_POCode_Option_Contrat = rs.Fields("CodeOption")
  If IsNull(theAssure.m_POCode_Option_Contrat) Then
    Call m_Logger.EcritTraceDansLog("CodeOption non renseignée : valeur utilisée = " & def_CodeOption & " à la ligne: " & rs.Fields("RECNO"))
    theAssure.m_POCode_Option_Contrat = def_CodeOption
'    NbRejet = NbRejet + 1
'    Exit Sub
  End If
  
  If bFromInfocentre = False Then
    FieldName = "Etablissement"
    theAssure.m_POEtablissement = def_Etablissement
  End If
  
  FieldName = "Dossier"
  theAssure.m_PONumSinistre = curSS
  
  FieldName = "TypeMvt"
  theAssure.m_POTypeMvt = def_TypeMvt
  
  theAssure.m_POMontantBase = MontantBase
  theAssure.m_POMontantRevalo = MontantRevalo
  theAssure.m_POMontantMajore = MontantMajo
  
  FieldName = "DebIndemn"
  dDate = ReadDateXL(rs.Fields("Survenance"))
  If Not IsNull(dDate) Then
    theAssure.m_PODebutIndemnisation = dDate
  End If
  
  theAssure.m_PONbJourIndemn = nbJIndemn
  
  If bFromInfocentre = False Then
    FieldName = "FamComptable"
    theAssure.m_POFamilleComptable = def_FamComptable
    
    FieldName = "Inspecteur"
    theAssure.m_POInspecteur = def_Inspecteur
    
    FieldName = "Insp"
    theAssure.m_POInsp = def_Insp
    
    FieldName = "Apport"
    theAssure.m_POApport = def_Apport
    
    FieldName = "Apport2"
    theAssure.m_POApport2 = def_Apport2
    
    FieldName = "Gestionnaire"
    theAssure.m_POGestionnaire = def_Gestionnaire
    
    FieldName = "NbIntervenant"
    theAssure.m_PONbIntervenant = def_NbIntervenant
    
    FieldName = "Compagnie"
    theAssure.m_POCompagnie = def_Code_CIE
    
    FieldName = "APPLI"
    theAssure.m_POAppli = def_Code_APP
  Else
    ' famille comptable pour compatibilité (=30)
    FieldName = "FamComptable"
    theAssure.m_POFamilleComptable = def_FamComptable
    
    FieldName = "Compagnie"
    theAssure.m_POCompagnie = def_Code_CIE
    
    FieldName = "APPLI"
    theAssure.m_POAppli = def_Code_APP
  End If
  
  FieldName = "IndicCC"
  theAssure.m_POIndicCC = def_IndicCC
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
  
  theAssure.m_PONumParamCalcul = rs.Fields("ParamCalcul")
  If IsNull(theAssure.m_PONumParamCalcul) Then
  
    Call m_Logger.EcritTraceDansLog("REJET : ParamCalcul non renseigné à la ligne: " & rs.Fields("RECNO"))
    NbRejet = NbRejet + 1
    Exit Sub
  End If

  
  
  ' Evol 2010 - Lot 2
  FieldName = "CDCONTENTIEUX"
  theAssure.m_POCDCONTENTIEUX = def_CDCONTENTIEUX
  
  FieldName = "NUSINISTRE"
  theAssure.m_PONUSINISTRE = curSS
  
  FieldName = "MTCAPCON"
  If Not IsNull(rs.Fields("MTCAPCON")) Then
    theAssure.m_POMontantCapConstit = rs.Fields("MTCAPCON")
  End If
  
  
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
  If (eProvisionRetraite = m_TypePeriode Or eProvisionRetraiteRevalo = m_TypePeriode) And theAssure.m_POGARCLE = cdGar59 And bPortefeuilleSalarie Then
    theAssure.m_IdAssure = theAssure.m_RECNO
    theAssure.m_RECNO = 0
    
    FieldName = "FinGarantie"
    theAssure.m_POTERME = rs.Fields("FinGarantie") '''''''''''''' A recalculer en fonction de la garantie
    
    'Modif 11/7/2018
    If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
        theAssure.m_POTERME = DateLimiteProvision  ' âge terme est forcé à 120 ans
    End If
    
    theAssure.m_CoeffAmortissement = 1
  
    FieldName = "FLAMORTISSABLE"
    theAssure.m_TopAmortissable = rs.Fields("FLAMORTISSABLE")
    
    theAssure.Save destDB, True
  End If
  
  ' Rien n'est écrit dans P3IPROVCOLL (Pas d'info pour Symphonie)
  
  
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
          'MsgBox "Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import" & vbLf & "Ligne " & rs.Fields("RECNO"), vbCritical
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

'##ModelId=5C90E97E02D8
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
  Dim m_dataHelper As DataHelper, nbAnnualForcee As Integer
      
  Set m_dataHelper = destDB.CreateHelper
  
  nbAnnualForcee = 0
  
  Set rs = destDB.OpenRecordset(rq, Disconnected)
  Do Until rs.EOF
    ' affiche la position
    If Not autoMode Then
      If (rs.AbsolutePosition Mod 10) = 0 Then
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

  If Not autoMode Then
    Screen.MousePointer = vbDefault
    fWait.Visible = False
    
    MsgBox "Erreur " & Err & " : " & Err.Description, vbCritical
    
    fWait.Visible = True
    Screen.MousePointer = vbHourglass
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Sub

'##ModelId=5C90E97E0396
Private Sub computeAvgRegime6(rs As DAO.Recordset, dateArrete As Date, nbJourDC As Long, fWait As Form)
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
          curSS = rs.Fields("NumeroSS")
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
          If rs.Fields("Categorie") = 0 Then
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
          curNCA = rs.Fields("NCA")
        End If
        
        FieldName = "Survenance"
        If IsNull(rs.Fields("Survenance")) Then
          Call m_Logger.EcritTraceDansLog("REJET : Survenance invalide à la ligne: " & rs.Fields("RECNO"))
          NbRejet = NbRejet + 1
          Exit Sub
        Else
          curArret = rs.Fields("Survenance")
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

'##ModelId=5C90E97F006C
Public Function iP3IGeneraliImport_DoImport(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, Optional autoMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactiviteA As Long, Optional typeCalculAnnualisationA As Long, Optional excelFileA As String, Optional logFileA As String) As Boolean
  
  ' Type de période
  m_TypePeriode = destDB.CreateHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode)

  ' lecture des valeurs par défaut depuis P3I.INI
  InitDefaultValues sFichierIni
  
  ' import depuis feuille Excel au format simplifié
  iP3IGeneraliImport_DoImport = DoImportXLS(CommonDialog1, destDB, CleGroupe, NumPeriode, DateDebut, DateFin, _
                     nbJourMax, nbJourDC, AgeRetraite, dateCloture, sFichierIni, m_bP3I_Individuel, _
                     autoMode, dateArreteA, typeDelaiInactiviteA, typeCalculAnnualisationA, excelFileA, logFileA)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' import depuis feuille Excel
'
'##ModelId=5C90E97F0329
Public Function DoImportXLS(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, Optional automatMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactiviteA As Long, Optional typeCalculAnnualisationA As Long, Optional excelFileA As String, Optional logFileA As String) As Boolean
                      
  Dim SrcDB As DAO.Database
  Dim rs3 As DAO.Recordset
  Dim rq As String
  Dim maxRecord As Long
  Dim codeRetour As Boolean
  Dim fWait As frmWait
  Dim datePaiement As Date
  Dim filename As String
  
  On Error GoTo GestionErreur
  
  If Not automatMode Then
    filename = CommonDialog1.filename
    autoMode = False
  Else
    filename = excelFileA
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
      DoImportXLS = False
      Exit Function
    End If
    
    
    ' Type de période
    m_TypePeriode = destDB.CreateHelper.GetParameterAsDouble("SELECT PETypePeriode FROM Periode WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE=" & NumPeriode)
  
  
    ' demande le nom de la base (fichier xls)
    CommonDialog1.filename = "*.xls"
    CommonDialog1.Filter = "Fichier Excel|*.xls|Base de données MS Access|*.mdb|"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Or CommonDialog1.filename = "*.xlsx" Then
      DoImportXLS = False
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
    
    f.lblDate2.Caption = Replace(f.lblDate2.Caption, "888", nbJourMax)
    
    f.rdoImportComplet.Value = True
    
    f.Show vbModal
    If ret_code = -1 Then
      DoImportXLS = False
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
  Dim n As Integer
  
  rq = ""
  n = InStrRev(CommonDialog1.filename, "\")
  If n <> 0 Then
    rq = Left(CommonDialog1.filename, n)
  End If
  
  Set m_Logger = New clsLogger
  
  If autoMode Then
    m_Logger.modeAuto = True
    m_Logger.FichierLog = logFileA
  Else
    m_Logger.FichierLog = rq & GetWinUser & "_ErreurImport.log"
    m_Logger.CreateLog "Import " & filename & " Groupe " & CleGroupe & " Période " & NumPeriode
  End If
  
 
  If autoMode Then
    Set SrcDB = OpenDatabase(filename, dbDriverNoPrompt, True, cdExcelExtendedPropertiesDAO)
  Else
    ' ouvre la base de donnees maitre
    If Right(UCase(CommonDialog1.filename), 5) = ".XLSX" Then
      ' Fichier Excel 2007 .XLSX
      fWait.Hide
      Unload fWait
      
      Set fWait = Nothing
      
      Screen.MousePointer = vbDefault
      
      MsgBox "Le format de fichier Excel 2007 n'est pas compatible avec ce logiciel." & vbLf & "Veuillez enregistrer votre fichier au format Excel 2003 !", vbExclamation + vbOKOnly, "Import des données"
      
      DoImportXLS = False
      Exit Function
    ElseIf Right(UCase(CommonDialog1.filename), 4) = ".XLS" Then
      ' Fichier Excel .XLS
      Set SrcDB = OpenDatabase(CommonDialog1.filename, dbDriverNoPrompt, True, cdExcelExtendedPropertiesDAO)
    Else
      ' Base Access .MDB
      Set SrcDB = OpenDatabase(CommonDialog1.filename)
    End If
  End If
  
  SrcDB.QueryTimeout = 120
  
  '*******  ATTENTION : ouverture de tous les recordset avant la transaction, sinon : crash
  
  Set m_dataHelper = destDB.CreateHelper
  
  
  ' Type d'import
  Dim bImportForRevalo As Boolean
  
  bImportForRevalo = False
  If m_dataHelper.GetParameterAsLong("SELECT PETYPEPERIODE FROM Periode WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE = " & NumPeriode) = eRevalo Then
    bImportForRevalo = True
  End If
  
  
  ' transaction...
  Dim transOpened As Boolean
  
  destDB.BeginTrans
  transOpened = True
  
  '***
  '*** Import des tables de paramètres
  '***
  
  If typeImport = eImportComplet Or typeImport = eImportTablesParametresSeules Then
      
    Call m_Logger.EcritTraceDansLog("   ")
    Call m_Logger.EcritTraceDansLog(">>>>> Import des tables de paramètres")
    
    ImportGenerique SrcDB, destDB, "TBQREGA", CleGroupe, NumPeriode, m_Logger, fWait
    
    'ImportGenerique SrcDB, destDB, "REGA01", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "CDSITUAT", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "CodesCat", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "CodeCatInv", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "Correspondance_CatOption", CleGroupe, NumPeriode, m_Logger, fWait
    
    'ImportGenerique SrcDB, destDB, "Produits", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "Capitaux_Moyens", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "CATR9", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "CATR9INVAL", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "PassageNCA", CleGroupe, NumPeriode, m_Logger, fWait
    
    ImportGenerique SrcDB, destDB, "PM_Retenue", CleGroupe, NumPeriode, m_Logger, fWait
  
    ImportGenerique SrcDB, destDB, "Reassurance", CleGroupe, NumPeriode, m_Logger, fWait
  
    ImportGenerique SrcDB, destDB, "ParamRentes", CleGroupe, NumPeriode, m_Logger, fWait
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
    
    '*** moyenne des PSAP des rente deces (regime 6) non null
'    rq = "SELECT * FROM Assure WHERE NomAssure <>""""" _
'         & " ORDER BY NumSociete, Regime, Sinistre, GE, Code_Cat_Contrat, Survenance, BaseRevalo, PeriodeDu, PeriodeAu, DatePaiement"
    rq = "SELECT * FROM Assure WHERE NomAssure <>""""" _
       & " ORDER BY RECNO " ' plus de boucle pour les sommations
    
    ' ouvre la feuille XL Assuré en lecture
    Set rs3 = SrcDB.OpenRecordset(rq, dbOpenSnapshot)
    
  
    Set rsCatR9 = destDB.OpenRecordset("SELECT * From CATR9 WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCodesCat = destDB.OpenRecordset("SELECT * From CodesCat WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCodeCatInv = destDB.OpenRecordset("SELECT * From CodeCatInv WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    Set rsCatInval = destDB.OpenRecordset("SELECT * From CATR9INVAL WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    
    Set rsNCA = destDB.OpenRecordset("SELECT * From PassageNCA WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Snapshot)
    
    Set rsTBQREGA = destDB.OpenRecordset("SELECT * From TBQREGA WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, Disconnected)
    
    'Set rsREGA01 = destDB.OpenRecordset("SELECT * From REGA01 WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode, snapshot)
    
    Set rsTaux = destDB.OpenRecordset("SELECT * FROM TauxRenteEducation", Snapshot)
    
    Set rsCodeProvision = destDB.OpenRecordset("SELECT * FROM CodeProvision", Snapshot)
    
    
    ' efface les anciennent valeurs
    destDB.Execute "DELETE FROM Assure WHERE POGPECLE=" & CleGroupe & " AND POPERCLE=" & NumPeriode
    
    ' efface les anciennent valeurs
    destDB.Execute "DELETE FROM Assure_P3IPROVCOLL WHERE CleGroupe=" & CleGroupe & " AND NumPeriode=" & NumPeriode
    
    
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
    'maxRecord = rs.RecordCount + 1
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
          
      If rs3.Fields("NomAssure") <> "" Then
      
        ' !!!!!!!! AddNewRecord parcours le recordset sur le meme assure, sinistre, garantie...
        ' Ajoute et configure un enregistrement dans Assure
        Call AddNewRecord(destDB, rs3, CleGroupe, NumPeriode, fWait, AgeRetraite, CDate(DateDebut), CDate(DateFin), nbJourMax, nbJourDC, False, bImportForRevalo)
        
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
      
      ' mets à jour la date d'extraction et PECALCULER = FALSE
      rq = "UPDATE Periode SET PEDATEEXT = " & m_dataHelper.BuildDateString(dateArrete) & " , PECALCULER = 'False' WHERE PEGPECLE = " & CleGroupe & " AND PENUMCLE = " & NumPeriode
      destDB.Execute rq, dbFailOnError
    End If
    
    m_Logger.EcritTraceDansLog " "
    
    Call m_Logger.EcritTraceDansLog(rs3.RecordCount & " articles dans le fichier")
    
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
    
    rsCatR9.Close
    rsCodesCat.Close
    rsCodeCatInv.Close
    rsCatInval.Close
    rsNCA.Close
    rsTBQREGA.Close
    'rsREGA01.Close
    rsTaux.Close
    rsCodeProvision.Close
    
    rs3.Close
  
  Else ' Import Données ?
      
    If transOpened = True Then
      destDB.CommitTrans
      transOpened = False
    End If
    
    codeRetour = True
  
  End If
  
  SrcDB.Close
  
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
  
  DoImportXLS = codeRetour
  
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
    
    SrcDB.Close
    
    DoImportXLS = False
    
    Set m_dataHelper = Nothing
  End If
  
  Exit Function
  
  Resume Next ' pour debug
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Import des tables de paramètre
'
'##ModelId=5C90E98001E1
Private Function ImportGenerique(SrcDB As DAO.Database, ByRef destDB As DataAccess, nomTable As String, CleGroupe As Long, NumPeriode As Long, m_Logger As clsLogger, fWait As frmWait) As Boolean
  Dim rsXL As DAO.Recordset
  Dim rsCR As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean
  Dim f As ADODB.Field, FieldName As String, cs As String
  Dim n As Integer, bookmark As Variant
    
  ImportGenerique = False
    
  If NumPeriode <= 0 Then Exit Function
  
  On Error GoTo GestionErreur
  
  If Not autoMode Then
    fWait.Caption = "Import de '" & nomTable & "' en cours..."
    
    fWait.ProgressBar1.Min = 0
    fWait.ProgressBar1.Value = 0
    fWait.ProgressBar1.Max = 100
    fWait.ProgressBar1.Refresh
    
    Screen.MousePointer = vbHourglass
  End If
 
  m_Logger.EcritTraceDansLog "------------------------------------------------------------"
  m_Logger.EcritTraceDansLog "Import de la table '" & nomTable & "'"

  ' cree une transaction
  fFoundError = False
  destDB.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode
  
  ' ouvre les tables de destianation
  Set rsCR = destDB.OpenRecordset("" & nomTable, Table)
  
  Dim idField As Integer
  
  If NumPeriode > 0 Then
    idField = IIf(nomTable = "ParamRentes", 4, 2)
  Else
    idField = 0
  End If
  Set rsXL = SrcDB.OpenRecordset("SELECT * FROM " & nomTable & " WHERE " & rsCR.Fields(idField).Name & " IS NOT NULL", dbOpenSnapshot)
  
  idField = IIf(idField >= 2, idField - 2, idField)
  
  If Not rsXL Is Nothing Then
    If rsXL.EOF Then
      If Not autoMode Then
        fWait.ProgressBar1.Max = 1
      End If
      m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
    Else
      rsXL.MoveLast
      rsXL.MoveFirst
    
      If Not autoMode Then
        fWait.ProgressBar1.Max = rsXL.RecordCount + 1
      End If
      
      ' lit les enregistrements
      Do Until rsXL.EOF
        bRejet = False
        If (rsXL.AbsolutePosition Mod 5) = 0 Then
          ' affiche la position
          If Not autoMode Then
            fWait.ProgressBar1.Value = rsXL.AbsolutePosition
            fWait.ProgressBar1.Refresh
          End If
        End If
        
        If Not IsNull(rsXL.Fields(idField)) Then
          rsCR.AddNew
          
          For Each f In rsCR.Fields

          
            FieldName = f.Name
            If f.Name = "NumPeriode" Then
              f.Value = NumPeriode
            ElseIf f.Name = "GroupeCle" Then
              f.Value = CleGroupe
            Else
              Select Case rsXL.Fields(f.Name).Type
                Case dbChar, dbText, dbMemo
                  If IsNull(rsXL.Fields(f.Name)) Then
                    cs = ""
                  Else
                    cs = Trim(rsXL.Fields(f.Name))
                  End If
                  If Len(cs) > f.DefinedSize Then
                    f.Value = Left(cs, f.DefinedSize)
                    Call m_Logger.EcritTraceDansLog("   Erreur : Colonne '" & FieldName & "' chaine trop longue (len>" & f.DefinedSize & ") valeur tronquée - Ligne " & rsXL.AbsolutePosition)
                  Else
                    f.Value = cs
                  End If
                
                Case Else
                  f.Value = rsXL.Fields(f.Name)
              End Select
            End If
          Next
          
          If fFoundError = True Then
            rsCR.CancelUpdate
            Exit Do
          Else
            rsCR.Update
          End If
        End If
        
        rsXL.MoveNext
      Loop
    End If
  End If
  
  Call m_Logger.EcritTraceDansLog("   " & rsXL.RecordCount & " lignes dans le fichier Excel pour la table '" & nomTable & "'")
  
  If fFoundError Then
    Call m_Logger.EcritTraceDansLog(">>>>> Fichier rejetté à cause des erreurs durant l'import !")
    
    destDB.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & CleGroupe & " AND NumPeriode=" & NumPeriode
  Else
    Call m_Logger.EcritTraceDansLog("   " & NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
  End If
  
  m_Logger.EcritTraceDansLog "Fin Import de la table '" & nomTable & "'"
  
  If Not rsXL Is Nothing Then
    rsXL.Close
    Set rsXL = Nothing
  End If
  
  rsCR.Close
  Set rsCR = Nothing
  
  ImportGenerique = fFoundError
  
  If Not autoMode Then
    Screen.MousePointer = vbDefault
  End If
  
  Exit Function
  
GestionErreur:
  If rsXL Is Nothing Then
    If Err = -2147217865 Then
      m_Logger.EcritTraceDansLog "   Format Incorrect : le fichier ne correspond pas au format de la table '" & nomTable & "'. " & Err.Description
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
      
      Case -2147352571
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée incompatible - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  fFoundError = True
  Resume Next
End Function


'********************************************************************
' Controle des données avant import
'
'##ModelId=5C90E980031A
Public Function iP3IGeneraliImport_DoControle(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Integer, NumeroLot As Long, NumPeriode As Long, sFichierIni As String) As Boolean
  
  MsgBox "DoControle() not implemented in object iP3IGeneraliImport.iImportGeneraliRevalo !", vbCritical
  
  iP3IGeneraliImport_DoControle = False
  
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période
' depuis un lot présent dans P3ILOGTRAIT et P3IPROVCOLL
'
'##ModelId=5C90E981002B
Public Function iP3IGeneraliImport_DoImportSASP3I(NumeroLot As Long, logPath As String, ByRef destDB As DataAccess, CleGroupe As Long, NumPeriode As Long, DateDebut As String, DateFin As String, nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, Optional autoMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactivitA As Long, Optional typeCalculAnnualisationA As Long, Optional logFileA As String) As Boolean

  MsgBox "DoImportSASP3I() not implemented in object iP3IGeneraliImport.iImportGeneraliRevalo !", vbCritical
       
  iP3IGeneraliImport_DoImportSASP3I = False
  
End Function

'##ModelId=5C90E98102E7
Public Function iP3IGeneraliImport_SetStatutaireVariables(mNumPeriodeStat As Long, mNumPeriodeNonStat As Long, mPathSexFileExcel As String, mCategoryCodeSTAT As String, mSexAllMale As Boolean, mTwoLotImport As Boolean) As Boolean
    
    MsgBox "SetStatutaireVariables() not implemented in object iP3IGeneraliImport.iImportGenerali !", vbCritical
       
    iP3IGeneraliImport_SetStatutaireVariables = False

End Function



