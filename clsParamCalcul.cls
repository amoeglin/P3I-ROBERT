VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamCalcul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75F601DE"
Option Explicit

'##ModelId=5C8A75F602DE
Private loadedRECNO As Long
'##ModelId=5C8A75F602F4
Private loadedCodeParamCalcul As Long

'##ModelId=5C8A75F60313
Public CodeParamCalcul As Long
'##ModelId=5C8A75F60323
Public NomParamCalcul As String
'##ModelId=5C8A75F60333
Public Commentaire As String

'risque statutaire
'##ModelId=5C8A75F60352
Public statYear As String
'##ModelId=5C8A75F60362
Public statAgeRet As Integer
'##ModelId=5C8A75F60381
Public statAgeMax As Integer
'##ModelId=5C8A75F60390
Public statAgeMin As Integer
'##ModelId=5C8A75F603B0
Public statMO As Integer
'##ModelId=5C8A75F603BF
Public statCLM As Integer
'##ModelId=5C8A75F603CF
Public statCLD As Integer
'##ModelId=5C8A75F70006
Public statMAT As Integer
'##ModelId=5C8A75F70016
Public statAT As Integer

'##ModelId=5C8A75F70035
Public TableStatutaire As String

' Incap
'##ModelId=5C8A75F70045
Public TauxIncapacite As Double
'##ModelId=5C8A75F70064
Public FraisGestionIncapacite As Double

Public TauxIncapPassageSpecifiqueInval1 As Double
'##ModelId=5C8A75F70083
Public LoiIncapacite As String
'##ModelId=5C8A75F700A2
Public LoiPassage As String
'##ModelId=5C8A75F700B2
Public LoiIncapacite_Retraite As String
'##ModelId=5C8A75F700D1
Public LoiPassage_Retraite As String
'##ModelId=5C8A75F700F0
Public DureeChomage As Integer

' Inval
'##ModelId=5C8A75F70100
Public TauxInvalidite As Double
'##ModelId=5C8A75F7011F
Public FraisGestionInvalidite As Double
'##ModelId=5C8A75F7013F
Public LoiInvalidite As String
'##ModelId=5C8A75F7015E
Public LoiInvalidite_Retraite As String
'##ModelId=5C8A75F7016D
Public LoiInvalidite_Viagere As String
'##ModelId=5C8A75F7018F
Public InterpolationInvalidite As InterpolationInval
'##ModelId=5C8A75F7019C
Public AgeLimiteInvalCat1 As Integer
'##ModelId=5C8A75F701AC
Public AgeLimiteInvalCat1_Retraite As Integer
'##ModelId=5C8A75F701CB
Public Inval1Cat As Double


'Dépendance
'##ModelId=5C8A75F701EA
Public LoiDependance As String


' Etat C7
'##ModelId=5C8A75F7020A
Public TauxRevenuC7 As Double
'##ModelId=5C8A75F7021C
Public FraisGestionC7 As Double

' Revalo
'##ModelId=5C8A75F7023B
Public TauxIndexation As Double
'##ModelId=5C8A75F7025A
Public DureeIndexation As Integer
'##ModelId=5C8A75F70279
Public TMO As Double
'##ModelId=5C8A75F70299
Public DureeLissage As Integer

' Rente Conjoint
'##ModelId=5C8A75F702B8
Public TauxTechniqueRenteConjoint As Double
'##ModelId=5C8A75F702E2
Public FraisGestionRenteConjoint As Double
'##ModelId=5C8A75F70301
Public FractionnementRenteConjoint As Fractionnement
'##ModelId=5C8A75F70304
Public PaiementRenteConjoint As EcheancePaiement
'##ModelId=5C8A75F70311
Public TableRenteConjoint As String

' Rente Education
'##ModelId=5C8A75F70330
Public TauxTechniqueRenteEducation As Double
'##ModelId=5C8A75F7033F
Public FraisGestionRenteEducation As Double
'##ModelId=5C8A75F70370
Public FractionnementRenteEducation As Fractionnement
'##ModelId=5C8A75F70380
Public PaiementRenteEducation As EcheancePaiement
'##ModelId=5C8A75F70381
Public TableRenteEducation As String

' maintien deces
'##ModelId=5C8A75F7039D
Public FraisGestionCapitauxDecesDC As Double
'##ModelId=5C8A75F703BC
Public FraisGestionRenteEducationDC As Double
  
'##ModelId=5C8A75F703DC
Public PourcentageAccident As Double
  
'##ModelId=5C8A75F80013
Public FraisGestionRenteConjointDC As Double
'##ModelId=5C8A75F80032
Public CapitalMoyenRenteConjointTempoDC As Double
'##ModelId=5C8A75F80051
Public CapitalMoyenRenteConjointViagereDC As Double
'##ModelId=5C8A75F80071
Public ForcerCapitalMoyenRenteConjointDC As Boolean
'##ModelId=5C8A75F80080
Public AgeConjointRenteConjointDC As Double
  
'##ModelId=5C8A75F800A0
Public UtiliserTableLissageProvision As Boolean

'##ModelId=5C8A75F800BF
Public RecalculCoeffBCAC As Boolean
'##ModelId=5C8A75F800E0
Public MethodeCalculDC As MethodeCalculMaintienDC
'##ModelId=5C8A75F800EE
Public PctPMCalculeeDC As Double

'##ModelId=5C8A75F8010D
Public LoiIncapaciteDC As String
'##ModelId=5C8A75F8012C
Public LoiInvaliditeDC As String
'##ModelId=5C8A75F8014B
Public LoiIncapaciteDC_Retraite As String
'##ModelId=5C8A75F8016B
Public LoiInvaliditeDC_Retraite As String
'##ModelId=5C8A75F8018A
Public TauxTechnicCalculDC As Double
'##ModelId=5C8A75F801B9
Public FraisGestionCalculDC As Double
'##ModelId=5C8A75F801C8
Public AgeLimiteCalulDC As Integer
'##ModelId=5C8A75F801E8
Public AgeLimiteCalulDC_Retraite As Integer

Public AgeLimiteCalulDC_Inval1 As Integer

Public AgeLimiteCalulDC_Retraite_Inval1 As Integer

'##ModelId=5C8A75F80213
Public TableIncapacitePrecalculDC As String
'##ModelId=5C8A75F8022A
Public TableInvaliditePrecalculDC As String
'##ModelId=5C8A75F80258
Public TableIncapacitePrecalculDC_Retraite As String
'##ModelId=5C8A75F80278
Public TableInvaliditePrecalculDC_Retraite As String

' Calcul de l'Age
'##ModelId=5C8A75F80297
Public CalculAge_Anniversaire As Boolean

' Passage
'##ModelId=5C8A75F802B6
Public LissageCoeffPassage As Boolean
'##ModelId=5C8A75F802D0
Public BridageAgeLimiteTable As Boolean
'##ModelId=5C8A75F802EF
Public AnnualisationPassage As Boolean

' Portefeuille Salaries
'##ModelId=5C8A75F8030E
Public PortefeuilleSalaries As Boolean

'##ModelId=5C8A75F8032D
Public autoMode As Boolean



'

'##ModelId=5C8A75F8034D
Private Sub Class_Initialize()
  loadedCodeParamCalcul = -1
  CodeParamCalcul = -1
  loadedRECNO = -1
End Sub


'##ModelId=5C8A75F8036C
Public Property Get RECNO() As Long
  RECNO = loadedRECNO
End Property


'##ModelId=5C8A75F8038B
Public Function GetCleTable(nomTable As String) As Long
  GetCleTable = m_dataHelper.GetParameterAsDouble("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
  If GetCleTable = 0 Then
    GetCleTable = -1
  End If
End Function


'##ModelId=5C8A75F803BA
Public Function GetLibelleTable(nomTable As String) As String
  GetLibelleTable = m_dataHelper.GetParameterAsStringCRW("SELECT LIBTABLE FROM ListeTableLoi WHERE NOMTABLE = '" & nomTable & "'")
End Function


'##ModelId=5C8A75F90001
Public Function GetNomTable(ByVal CleTable As Long) As String
  Dim ret As Variant
  
  ret = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE = " & CleTable)
  If IsNull(ret) Then
    GetNomTable = ""
  Else
    GetNomTable = ret
  End If
End Function



'##ModelId=5C8A75F90030
Public Function LoadFromIni(NumParam As Long, typeParam As String) As Boolean
  Dim theParamSectionName As String
  
  LoadFromIni = False
  
  loadedCodeParamCalcul = NumParam
  CodeParamCalcul = NumParam
  loadedRECNO = -1
  
  
  If NumParam = -1 Then
  
    theParamSectionName = DEFAULT_SECTION
  
  Else
    
    If typeParam <> "BILAN" And typeParam <> "CLIENT" And typeParam <> "SIMULATION" Then
      typeParam = "BILAN"
      MsgBox "clsParamCalcul::LoadFromIni(" & NumParam & ", " & typeParam & ") : typeParam inconnu, BILAN utilisé !", vbCritical
    End If
    
    theParamSectionName = DEFAULT_PARAM_SECTION & typeParam & "_" & CodeParamCalcul
  
  End If
  
  
  NomParamCalcul = GetSettingIni(CompanyName, theParamSectionName, "Nom", "???")
  If NomParamCalcul = "???" Then
    Exit Function
  End If
  
  Commentaire = GetSettingIni(CompanyName, theParamSectionName, "Commentaire", "")
   
  'risque statutaire
  statAgeRet = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatAgeRetraite", "62"))
  statAgeMax = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "statAgeMax", "64"))
  statAgeMin = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "statAgeMin", "20"))
  statMO = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatMOMaxSemaine", "52"))
  statCLM = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatCLMMaxSemaine", "156"))
  statCLD = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatCLDMaxSemaine", "252"))
  statMAT = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatMATMaxSemaine", "23"))
  statAT = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatATMaxSemaine", "366"))
  
  statYear = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "StatAnneeBareme", "2015"))
  TableStatutaire = GetSettingIni(CompanyName, theParamSectionName, "TableStatutaireBaremes", "Statutaire_Generali_2004")
  
  
  ' Incap
  TauxIncapacite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxIncapacite", "2.75"))
  FraisGestionIncapacite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionIncapacite", "3"))
  TauxIncapPassageSpecifiqueInval1 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxIncapPassageSpecifiqueInval1", "27"))
  LoiIncapacite = GetSettingIni(CompanyName, theParamSectionName, "LoiIncapacite", "LINCAP1996")
  LoiPassage = GetSettingIni(CompanyName, theParamSectionName, "LoiPassage", "LPASS1996")
  LoiIncapacite_Retraite = GetSettingIni(CompanyName, theParamSectionName, "LoiIncapacite", "LINCAP1996")
  LoiPassage_Retraite = GetSettingIni(CompanyName, theParamSectionName, "LoiPassage", "LPASS1996")
  DureeChomage = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "DureeChomage", "36"))

  ' Inval
  TauxInvalidite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxInvalidité", "2.75"))
  FraisGestionInvalidite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionInvalidite", "3"))
  LoiInvalidite = GetSettingIni(CompanyName, theParamSectionName, "LoiInvalidite", "LINVAL1996")
  LoiInvalidite_Retraite = GetSettingIni(CompanyName, theParamSectionName, "LoiInvalidite", "LINVAL1996")
  InterpolationInvalidite = GetSettingIni(CompanyName, theParamSectionName, "InterpolationInvalidite", eInterpolationInval_NON)
  AgeLimiteInvalCat1 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeLimiteInvalCat1", "65"))
  Inval1Cat = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "Inval1Cat", "30"))
  
  ' Inval Viagere
  LoiInvalidite_Viagere = GetSettingIni(CompanyName, theParamSectionName, "LoiInvaliditeViagere", "LINVAL_VIAGERE_BCAC")
  
  ' Dépendance
  LoiDependance = GetSettingIni(CompanyName, theParamSectionName, "LoiDependance", "DependanceBCAC")
  
  ' Etats C7
  TauxRevenuC7 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxRevenuC7", "3.81"))
  FraisGestionC7 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionC7", "3"))
  
  TauxIndexation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxIndexation", "0.94"))
  DureeIndexation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "DureeIndexation", "7"))
  TMO = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TMO", "3.81"))
  DureeLissage = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "DureeLissage", "7"))
    
  ' rente conjoint
  TauxTechniqueRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxTechniqueRenteConjoint", 2.25))
  FraisGestionRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteConjoint", "3"))
  FractionnementRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FractionnementRenteConjoint", eFractionnementTrimestriel))
  PaiementRenteConjoint = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "PaiementRenteConjoint", ePaiementEchu))
  TableRenteConjoint = GetSettingIni(CompanyName, theParamSectionName, "TableRenteConjoint", "TF0002")

  ' rente education
  TauxTechniqueRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxTechniqueRenteEducation", 2.25))
  FraisGestionRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteEducation", "3"))
  FractionnementRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FractionnementRenteEducation", eFractionnementTrimestriel))
  PaiementRenteEducation = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "PaiementRenteEducation", ePaiementEchu))
  TableRenteEducation = GetSettingIni(CompanyName, theParamSectionName, "TableRenteEducation", "TF0002")

  ' maintien deces
  FraisGestionCapitauxDecesDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionCapitauxDecesDC", "3"))
  FraisGestionRenteEducationDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteEducationDC", "3"))
  
  PourcentageAccident = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "PourcentageAccident", "20"))
  
  FraisGestionRenteConjointDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteConjointDC", "3"))
  CapitalMoyenRenteConjointTempoDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "CapitalMoyenRenteConjointTempoDC", "1"))
  CapitalMoyenRenteConjointViagereDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "CapitalMoyenRenteConjointViagereDC", "0.5"))
  ForcerCapitalMoyenRenteConjointDC = IIf(GetSettingIni(CompanyName, theParamSectionName, "ForcerCapitalMoyenRenteConjointDC", "0") <> "0", True, False)
  AgeConjointRenteConjointDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeConjointRenteConjointDC", "2.5"))
  
  UtiliserTableLissageProvision = IIf(GetSettingIni(CompanyName, theParamSectionName, "UtiliserTableLissageProvision", "0") <> "0", True, False)
  
  '
  ' recalcul
  '
  
  '=GetSettingIni(CompanyName, theParamSectionName, "PMGDForcerInval", "0")
  RecalculCoeffBCAC = IIf(GetSettingIni(CompanyName, theParamSectionName, "RecalculCoeffBCAC", "0") <> "0", True, False)
  
  MethodeCalculDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "MethodeCalculDC", eCapitauxConstitutifs))
  PctPMCalculeeDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "PctPMCalculeeDC", 0))
  
  ' combo incap
  LoiIncapaciteDC = GetSettingIni(CompanyName, theParamSectionName, "LoiIncapaciteDC", "MortIncap")
  LoiIncapaciteDC_Retraite = GetSettingIni(CompanyName, theParamSectionName, "LoiIncapaciteDC", "MortIncap")
  

  ' combo invalidite
  LoiInvaliditeDC = GetSettingIni(CompanyName, theParamSectionName, "LoiInvaliditeDC", "MortInval")
  LoiInvaliditeDC_Retraite = GetSettingIni(CompanyName, theParamSectionName, "LoiInvaliditeDC", "MortInval")
  
  TauxTechnicCalculDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "TauxTechniqueDC", "2.75"))
  FraisGestionCalculDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "FraisGestionDC", "3"))
  
  AgeLimiteCalulDC = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC", "60"))
  AgeLimiteCalulDC_Retraite = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC", "60"))
  
  AgeLimiteCalulDC_Inval1 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC_Inval1", "60"))
  AgeLimiteCalulDC_Retraite_Inval1 = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC_Inval1", "60"))
  
  
  ' combo incap
  TableIncapacitePrecalculDC = GetSettingIni(CompanyName, theParamSectionName, "TableIncapacitePrecalculDC", "PMIJ60-2,5%")
  TableIncapacitePrecalculDC_Retraite = GetSettingIni(CompanyName, theParamSectionName, "TableIncapacitePrecalculDC", "PMIJ60-2,5%")
  
  ' combo invalidite
  TableInvaliditePrecalculDC = GetSettingIni(CompanyName, theParamSectionName, "TableInvaliditePrecalculDC", "PMInv60-2,5%")
  TableInvaliditePrecalculDC_Retraite = GetSettingIni(CompanyName, theParamSectionName, "TableInvaliditePrecalculDC", "PMInv60-2,5%")

  ' calcul de l'age
  CalculAge_Anniversaire = IIf(GetSettingIni(CompanyName, theParamSectionName, "CalculAge_Anniversaire", "0") <> "0", True, False)


  LissageCoeffPassage = GetSettingIni(CompanyName, theParamSectionName, "LissageCoeffPassage", "0") <> "0"
  AnnualisationPassage = GetSettingIni(CompanyName, theParamSectionName, "AnnualisationPassage", "0") <> "0"
  BridageAgeLimiteTable = GetSettingIni(CompanyName, theParamSectionName, "BridageAgeLimiteTable", "0") <> "0"
  
  PortefeuilleSalaries = GetSettingIni(CompanyName, theParamSectionName, "PortefeuilleSalaries", "0") <> "0"
  
  LoadFromIni = True
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sauvegarde des paramètres dans le .INI
'
'##ModelId=5C8A75F9007E
Public Function SaveToIni(bToDefaultSection As Boolean, typeParam As String) As Boolean
  Dim theParamSectionName As String
  
  SaveToIni = False
  
  ' Ou sauvegarder ?
  If bToDefaultSection = True Then
    
    '
    ' paramètres par défaut
    '
    theParamSectionName = DEFAULT_SECTION
  
  Else
    
    '
    ' paramètres spécifique
    '
    
    If typeParam <> "BILAN" And typeParam <> "CLIENT" And typeParam <> "SIMULATION" Then
      typeParam = "BILAN"
      MsgBox "clsParamCalcul::SaveFromIni(" & bToDefaultSection & ", " & typeParam & ") : typeParam inconnu, BILAN utilisé !", vbCritical
    End If
    
    ' efface la section courante
    theParamSectionName = DEFAULT_PARAM_SECTION & typeParam & "_" & CodeParamCalcul
    Call DeleteSection(CompanyName, theParamSectionName)
    
    ' efface l'ancienne section si le numero a changé
    If loadedCodeParamCalcul <> CodeParamCalcul Then
      Call DeleteSection(CompanyName, DEFAULT_PARAM_SECTION & typeParam & "_" & loadedCodeParamCalcul)
    End If
  
  End If
    
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "NumParamCalcul", CodeParamCalcul)
  Call SaveSettingIni(CompanyName, theParamSectionName, "Nom", NomParamCalcul)
  Call SaveSettingIni(CompanyName, theParamSectionName, "Commentaire", Commentaire)
  
  ' statutaire
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatAgeRetraite", statAgeRet)
  Call SaveSettingIni(CompanyName, theParamSectionName, "statAgeMax", statAgeMax)
  Call SaveSettingIni(CompanyName, theParamSectionName, "statAgeMin", statAgeMin)
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatMOMaxSemaine", statMO)
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatCLMMaxSemaine", statCLM)
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatCLDMaxSemaine", statCLD)
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatMATMaxSemaine", statMAT)
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatATMaxSemaine", statAT)
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "StatAnneeBareme", statYear)
  Call SaveSettingIni(CompanyName, theParamSectionName, "TableStatutaireBaremes", TableStatutaire)
 
   
  ' Incap
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxIncapacite", TauxIncapacite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionIncapacite", FraisGestionIncapacite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxIncapPassageSpecifiqueInval1", TauxIncapPassageSpecifiqueInval1)
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiIncapacite", LoiIncapacite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiPassage", LoiPassage)
  Call SaveSettingIni(CompanyName, theParamSectionName, "DureeChomage", DureeChomage)
  
  ' Inval
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxInvalidité", TauxInvalidite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionInvalidite", FraisGestionInvalidite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiInvalidite", LoiInvalidite)
  Call SaveSettingIni(CompanyName, theParamSectionName, "AgeLimiteInvalCat1", AgeLimiteInvalCat1)
  Call SaveSettingIni(CompanyName, theParamSectionName, "Inval1cat", Inval1Cat)
 
 
  ' Dépendance
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiDependance", LoiDependance)
 
  ' etats C7
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxRevenuC7", TauxRevenuC7)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionC7", FraisGestionC7)
  
  ' Revalo
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxIndexation", TauxIndexation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "DureeIndexation", DureeIndexation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "TMO", TMO)
  Call SaveSettingIni(CompanyName, theParamSectionName, "DureeLissage", DureeLissage)
    
  ' rente conjoint
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxTechniqueRenteConjoint", TauxTechniqueRenteConjoint)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteConjoint", FraisGestionRenteConjoint)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FractionnementRenteConjoint", FractionnementRenteConjoint)
  Call SaveSettingIni(CompanyName, theParamSectionName, "PaiementRenteConjoint", PaiementRenteConjoint)
  Call SaveSettingIni(CompanyName, theParamSectionName, "TableRenteConjoint", TableRenteConjoint)

  ' rente education
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxTechniqueRenteEducation", TauxTechniqueRenteEducation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteEducation", FraisGestionRenteEducation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FractionnementRenteEducation", FractionnementRenteEducation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "PaiementRenteEducation", PaiementRenteEducation)
  Call SaveSettingIni(CompanyName, theParamSectionName, "TableRenteEducation", TableRenteEducation)

  ' maintien deces
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionCapitauxDecesDC", FraisGestionCapitauxDecesDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteEducationDC", FraisGestionRenteEducationDC)
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "PourcentageAccident", PourcentageAccident)
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionRenteConjointDC", FraisGestionRenteConjointDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "CapitalMoyenRenteConjointTempoDC", CapitalMoyenRenteConjointTempoDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "CapitalMoyenRenteConjointViagereDC", CapitalMoyenRenteConjointViagereDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "ForcerCapitalMoyenRenteConjointDC", IIf(ForcerCapitalMoyenRenteConjointDC = vbChecked, "1", "0"))
  Call SaveSettingIni(CompanyName, theParamSectionName, "AgeConjointRenteConjointDC", AgeConjointRenteConjointDC)
  
  ' Methode lissage provision
  Call SaveSettingIni(CompanyName, theParamSectionName, "UtiliserTableLissageProvision", IIf(UtiliserTableLissageProvision = True, "1", "0"))
  
  '
  ' recalcul
  '
  
  'Call SaveSettingIni(CompanyName, theParamSectionName, "PMGDForcerInval", "0")
  Call SaveSettingIni(CompanyName, theParamSectionName, "RecalculCoeffBCAC", IIf(RecalculCoeffBCAC = True, "1", "0"))
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "MethodeCalculDC", MethodeCalculDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "PctPMCalculeeDC", PctPMCalculeeDC)

  ' combo incap
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiIncapaciteDC", LoiIncapaciteDC)
  
  ' combo invalidite
  Call SaveSettingIni(CompanyName, theParamSectionName, "LoiInvaliditeDC", LoiInvaliditeDC)
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "TauxTechniqueDC", TauxTechnicCalculDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "FraisGestionDC", FraisGestionCalculDC)
  
  Call SaveSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC", AgeLimiteCalulDC)
  Call SaveSettingIni(CompanyName, theParamSectionName, "AgeLimiteCalulDC_Inval1", AgeLimiteCalulDC_Inval1)
  
  ' coeff incap precalculé dc
  Call SaveSettingIni(CompanyName, theParamSectionName, "TableIncapacitePrecalculDC", TableIncapacitePrecalculDC)
  
  ' coeff inval precalculé dc
  Call SaveSettingIni(CompanyName, theParamSectionName, "TableInvaliditePrecalculDC", TableInvaliditePrecalculDC)

  ' calcul de l'age
  Call SaveSettingIni(CompanyName, theParamSectionName, "CalculAge_Anniversaire", IIf(CalculAge_Anniversaire = True, "1", "0"))


  Call SaveSettingIni(CompanyName, theParamSectionName, "LissageCoeffPassage", IIf(LissageCoeffPassage = True, "1", "0"))
  Call SaveSettingIni(CompanyName, theParamSectionName, "BridageAgeLimiteTable", IIf(BridageAgeLimiteTable = True, "1", "0"))
  Call SaveSettingIni(CompanyName, theParamSectionName, "AnnualisationPassage", IIf(AnnualisationPassage = True, "1", "0"))
  
  ' interpolation de l'inval
  Call SaveSettingIni(CompanyName, theParamSectionName, "InterpolationInvalidite", InterpolationInvalidite)

  Call SaveSettingIni(CompanyName, theParamSectionName, "PortefeuilleSalaries", IIf(PortefeuilleSalaries = True, "1", "0"))

  SaveToIni = True
End Function


'##ModelId=5C8A75F900BC
Private Function ReadDouble(f As ADODB.field, ByVal default As Double) As Double
  
  If IsNull(f) Then
    ReadDouble = default
  Else
    ReadDouble = f.Value
  End If
    
End Function


'##ModelId=5C8A75F9011A
Private Function ReadString(f As ADODB.field, ByVal default As String) As String
  
  If IsNull(f) Then
    ReadString = default
  Else
    ReadString = f.Value
  End If
    
End Function


'##ModelId=5C8A75F90168
Public Function LoadFromDB(NumGroupe As Integer, numPeriode As Long, NumParamCalcul As Long) As Boolean
  Dim rs As ADODB.Recordset
  Dim theParamSectionName As String

  
  theParamSectionName = DEFAULT_SECTION
  
  On Error GoTo err_LoadFromDB
  
  LoadFromDB = False
  
  
  ' code
  loadedCodeParamCalcul = NumParamCalcul
  CodeParamCalcul = NumParamCalcul
  
  ' ouverture de la table ParamCalcul
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM ParamCalcul WHERE PEGPECLE = " & NumGroupe & " AND PENUMCLE=" & numPeriode & " AND PENUMPARAMCALCUL=" & CodeParamCalcul, Snapshot)
  If rs.EOF = True Then
    ' creation
    Exit Function
  End If
  
  loadedRECNO = rs.fields("RECNO")
  CodeParamCalcul = rs.fields("PENUMPARAMCALCUL").Value
  
  NomParamCalcul = rs.fields("PENOMPARAM").Value
  Commentaire = rs.fields("PECOMMENTAIREPARAM").Value
  
  ' statutaire
  statAgeRet = IIf(IsNull(rs.fields("PESTAT_AGERET").Value), 62, rs.fields("PESTAT_AGERET").Value)
  statAgeMax = IIf(IsNull(rs.fields("PESTAT_AGEMAX").Value), 64, rs.fields("PESTAT_AGEMAX").Value)
  statAgeMin = IIf(IsNull(rs.fields("PESTAT_AGEMIN").Value), 20, rs.fields("PESTAT_AGEMIN").Value)
  statMO = IIf(IsNull(rs.fields("PESTAT_MO").Value), 52, rs.fields("PESTAT_MO").Value)
  statCLM = IIf(IsNull(rs.fields("PESTAT_CLM").Value), 156, rs.fields("PESTAT_CLM").Value)
  statCLD = IIf(IsNull(rs.fields("PESTAT_CLD").Value), 252, rs.fields("PESTAT_CLD").Value)
  statMAT = IIf(IsNull(rs.fields("PESTAT_MAT").Value), 23, rs.fields("PESTAT_MAT").Value)
  statAT = IIf(IsNull(rs.fields("PESTAT_AT").Value), 366, rs.fields("PESTAT_AT").Value)
  
  statYear = IIf(IsNull(rs.fields("PESTAT_ANNBAR").Value), 2015, rs.fields("PESTAT_ANNBAR").Value)
  TableStatutaire = GetNomTable(IIf(IsNull(rs.fields("PECleTableStatutaire").Value), 0, rs.fields("PECleTableStatutaire").Value))
     
     
  ' Incap
  LoiIncapacite = rs.fields("PELMINCAP").Value
  LoiIncapacite_Retraite = ReadString(rs.fields("PELMINCAP_Retraite"), LoiIncapacite)
  TauxIncapacite = rs.fields("PETXINCAP").Value
  FraisGestionIncapacite = rs.fields("PEGESTIONINCAP").Value
  TauxIncapPassageSpecifiqueInval1 = IIf(IsNull(rs.fields("PETXINCAPINVAL1").Value), 0, rs.fields("PETXINCAPINVAL1").Value)
  LoiPassage = rs.fields("PELMPASSAGE").Value
  LoiPassage_Retraite = ReadString(rs.fields("PELMPASSAGE_Retraite"), LoiPassage_Retraite)
  
' duree maxi de la prestation chômage 02/09/2013 AM
  If IsNull(rs.fields("PEDUREECHOMAGE")) Then
    DureeChomage = GetSettingIni(CompanyName, theParamSectionName, "DureeChomage", 24) ' ancienne valeur dans P3I collective
  Else
    DureeChomage = rs.fields("PEDUREECHOMAGE").Value
  End If
  
  ' Inval
  LoiInvalidite = rs.fields("PELMINVAL").Value
  LoiInvalidite_Retraite = ReadString(rs.fields("PELMINVAL_Retraite"), LoiInvalidite)
  TauxInvalidite = rs.fields("PETXINVAL").Value
  FraisGestionInvalidite = rs.fields("PEGESTIONINVAL").Value
  
  'Inval Viagere
  If IsNull(rs.fields("PELMINVAL_Viagere")) Then
    LoiInvalidite_Viagere = GetSettingIni(CompanyName, theParamSectionName, "LoiInvaliditeViagere", "LINVAL_VIAGERE_BCAC")
  Else
    LoiInvalidite_Viagere = rs.fields("PELMINVAL_Viagere").Value
  End If
    
  AgeLimiteInvalCat1 = ReadDouble(rs.fields("PEAgeLimiteInvalCat1"), 65)
  AgeLimiteInvalCat1_Retraite = ReadDouble(rs.fields("PEAgeLimiteInvalCat1_Retraite"), 65)
  
  ' proportion des invalides 1 ère cat 02/09/2013 AM
  If IsNull(rs.fields("PEINVAL1CAT")) Then
    Inval1Cat = GetSettingIni(CompanyName, theParamSectionName, "Inval1Cat", 30)
  Else
    Inval1Cat = rs.fields("PEINVAL1CAT").Value
  End If
 
  ' Dépendance
  If IsNull(rs.fields("PELMDEPENDANCE")) Then
    LoiDependance = GetSettingIni(CompanyName, theParamSectionName, "LoiDependance", "DependanceBCAC")
  Else
    LoiDependance = rs.fields("PELMDEPENDANCE").Value
  End If
  TauxRevenuC7 = rs.fields("PETXREVENUC7").Value
  FraisGestionC7 = rs.fields("PEGESTIONC7").Value
      
  TauxIndexation = rs.fields("PETXREVALO").Value
  DureeIndexation = rs.fields("PEDUREEREVALO").Value
  DureeLissage = rs.fields("PEDUREELISSAGE").Value
  TMO = rs.fields("PETMO").Value
  
  'rs.fields("PEAGERETRAITE").Value = 65
  
  ' rente education
  TableRenteEducation = rs.fields("PETRENTEEDUC").Value
  TauxTechniqueRenteEducation = rs.fields("PETXRTEDUC").Value
  FraisGestionRenteEducation = rs.fields("PEGESTIONEDUC").Value
  PaiementRenteEducation = rs.fields("PEPAIEMENTEDUC").Value
  FractionnementRenteEducation = rs.fields("PEFRACTIONEDUC").Value
  
  ' rente conjoint
  TableRenteConjoint = rs.fields("PETRENTECONJOINT").Value
  TauxTechniqueRenteConjoint = rs.fields("PETXRTCONJOINT").Value
  FraisGestionRenteConjoint = rs.fields("PEGESTIONCONJOINT").Value
  PaiementRenteConjoint = rs.fields("PEPAIEMENTCONJOINT").Value
  FractionnementRenteConjoint = rs.fields("PEFRACTIONCONJOINT").Value
  
  ' maintien deces
  FraisGestionCapitauxDecesDC = rs.fields("PEGestionDeces_MDC").Value
  FraisGestionRenteEducationDC = rs.fields("PEGestionEduc_MDC").Value
  
  PourcentageAccident = rs.fields("PEPCTACCIDENT").Value
  
  FraisGestionRenteConjointDC = rs.fields("PEGestionConjoint_MDC").Value
  CapitalMoyenRenteConjointTempoDC = rs.fields("PECapConstitMoyenTempo_MDC").Value
  CapitalMoyenRenteConjointViagereDC = rs.fields("PECapConstitMoyenViagere_MDC").Value
  ForcerCapitalMoyenRenteConjointDC = rs.fields("PEForcerCapConstitMoyen_MDC").Value
  AgeConjointRenteConjointDC = rs.fields("PEAgeConjoint_MDC").Value
  
  UtiliserTableLissageProvision = rs.fields("PELissageProvision").Value
  
  
  '
  ' Maintien DC
  '
  
  'datPrimaryRS.Recordset.fields("PEPMGDForcerInval") = IIf(chkPMGDForcerInval = vbChecked, 1, 0)
  RecalculCoeffBCAC = rs.fields("PERecalculBCAC").Value
  
  AgeLimiteCalulDC = rs.fields("PEAgeLimiteCalul_MDC").Value
  AgeLimiteCalulDC_Retraite = ReadDouble(rs.fields("PEAgeLimiteCalul_MDC_Retraite"), AgeLimiteCalulDC)
  
  'Ajout le 06/11/2023 AG
  If IsNull(rs.fields("PEAgeLimiteCalul_MDC_Inval1")) Then
    AgeLimiteCalulDC_Inval1 = AgeLimiteCalulDC
  Else
    AgeLimiteCalulDC_Inval1 = rs.fields("PEAgeLimiteCalul_MDC_Inval1").Value
  End If
  
  If IsNull(rs.fields("PEAgeLimiteCalul_MDC_Retraite_Inval1")) Then
   AgeLimiteCalulDC_Retraite_Inval1 = AgeLimiteCalulDC_Retraite
  Else
   AgeLimiteCalulDC_Retraite_Inval1 = ReadDouble(rs.fields("PEAgeLimiteCalul_MDC_Retraite_Inval1"), AgeLimiteCalulDC_Inval1)
  End If
  MethodeCalculDC = rs.fields("PEMethodeCalcul").Value
  PctPMCalculeeDC = ReadDouble(rs.fields("PEPctPMCalculeeDC"), 0)


  ' Incap DC : ID et non NOM !!!
  LoiIncapaciteDC = GetNomTable(rs.fields("PECleTableMortBCACIncap_MDC").Value)
  LoiIncapaciteDC_Retraite = GetNomTable(ReadDouble(rs.fields("PECleTableMortBCACIncap_MDC_Retraite"), rs.fields("PECleTableMortBCACIncap_MDC").Value))
  
  ' Invalidite DC : ID et non NOM !!!
  LoiInvaliditeDC = GetNomTable(rs.fields("PECleTableMortBCACInval_MDC").Value)
  LoiInvaliditeDC_Retraite = GetNomTable(ReadDouble(rs.fields("PECleTableMortBCACInval_MDC_Retraite"), rs.fields("PECleTableMortBCACInval_MDC").Value))
  
  TauxTechnicCalculDC = rs.fields("PETauxTechRecalculBCAC_MDC").Value
  FraisGestionCalculDC = rs.fields("PEGestionRecalculBCAC_MDC").Value
  
  
  ' table precalculée Incap DC : ID et non NOM !!!
  TableIncapacitePrecalculDC = GetNomTable(rs.fields("PECleTableBCACIncap_MDC").Value)
  TableIncapacitePrecalculDC_Retraite = GetNomTable(ReadDouble(rs.fields("PECleTableBCACIncap_MDC_Retraite"), rs.fields("PECleTableBCACIncap_MDC").Value))
  
  ' table precalculée Inval DC : ID et non NOM !!!
  TableInvaliditePrecalculDC = GetNomTable(rs.fields("PECleTableBCACInval_MDC").Value)
  TableInvaliditePrecalculDC_Retraite = GetNomTable(ReadDouble(rs.fields("PECleTableBCACInval_MDC_Retraite"), rs.fields("PECleTableBCACInval_MDC").Value))
  
  
  ' calcul de l'age
  CalculAge_Anniversaire = rs.fields("PEAge_Anniversaire").Value

  
  ' méthode d'interpollation inval
  InterpolationInvalidite = rs.fields("PEINTERPOLATION_INVAL").Value
  
  ' Calcul du passage sans lissage : INVERSER DANS LA BDD !!!
  LissageCoeffPassage = Not rs.fields("PESANSLISSAGE_PASSAGE")
  AnnualisationPassage = rs.fields("PEANNUALISATION_PASSAGE")
  
  BridageAgeLimiteTable = rs.fields("PEBRIDAGE_AGE")
  
  
    
  PortefeuilleSalaries = rs.fields("PEPortefeuilleSalaries")
    
  rs.Close
  
  LoadFromDB = True
  
  Exit Function
  
err_LoadFromDB:
  MsgBox "erreur " & Err & vbLf & Err.Description & vbLf & vbLf & " Chargement interrompu "
  LoadFromDB = False
  
  Exit Function
  
  
  Resume Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sauvegarde dans la base de données
'
'##ModelId=5C8A75F901D6
Public Function SaveToDB(NumGroupe As Integer, numPeriode As Long) As Boolean
  Dim rs As ADODB.Recordset
  
  
  On Error GoTo err_SaveToDB
  
  SaveToDB = False
  
  
  ' changement de code ?
  If loadedCodeParamCalcul <> -1 And loadedCodeParamCalcul <> CodeParamCalcul Then
    m_dataSource.Execute "DELETE FROM ParamCalcul WHERE RECNO=" & loadedRECNO
    
    loadedCodeParamCalcul = CodeParamCalcul
  End If
  
  
  ' ouverture de la table ParamCalcul
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM ParamCalcul WHERE PEGPECLE = " & NumGroupe & " AND PENUMCLE=" & numPeriode & " AND PENUMPARAMCALCUL=" & CodeParamCalcul, Dynamic)
  If rs.EOF = True Then
    ' creation
    rs.AddNew
  End If
  
  
  rs.fields("PEGPECLE").Value = NumGroupe
  rs.fields("PENUMCLE").Value = numPeriode
  rs.fields("PENUMPARAMCALCUL").Value = CodeParamCalcul
  
  rs.fields("PENOMPARAM").Value = NomParamCalcul
  rs.fields("PECOMMENTAIREPARAM").Value = Commentaire
  
  ' statutaire
  rs.fields("PESTAT_AGERET").Value = statAgeRet
  rs.fields("PESTAT_AGEMAX").Value = statAgeMax
  rs.fields("PESTAT_AGEMIN").Value = statAgeMin
  rs.fields("PESTAT_MO").Value = statMO
  rs.fields("PESTAT_CLM").Value = statCLM
  rs.fields("PESTAT_CLD").Value = statCLD
  rs.fields("PESTAT_MAT").Value = statMAT
  rs.fields("PESTAT_AT").Value = statAT
  
  rs.fields("PESTAT_ANNBAR").Value = statYear
  rs.fields("PECleTableStatutaire").Value = GetCleTable(TableStatutaire)
  
    
  ' Incap
  rs.fields("PELMINCAP").Value = LoiIncapacite
  rs.fields("PELMINCAP_Retraite").Value = LoiIncapacite_Retraite
  rs.fields("PETXINCAP").Value = TauxIncapacite
  rs.fields("PEGESTIONINCAP").Value = FraisGestionIncapacite
  rs.fields("PETXINCAPINVAL1").Value = TauxIncapPassageSpecifiqueInval1
  rs.fields("PELMPASSAGE").Value = LoiPassage
  rs.fields("PELMPASSAGE_Retraite").Value = LoiPassage_Retraite
  rs.fields("PEDUREECHOMAGE").Value = DureeChomage
  
  
  ' Inval
  rs.fields("PELMINVAL").Value = LoiInvalidite
  rs.fields("PETXINVAL").Value = TauxInvalidite
  rs.fields("PELMINVAL_Retraite").Value = LoiInvalidite_Retraite
  rs.fields("PEGESTIONINVAL").Value = FraisGestionInvalidite
  rs.fields("PEAgeLimiteInvalCat1").Value = AgeLimiteInvalCat1
  rs.fields("PEAgeLimiteInvalCat1_Retraite").Value = AgeLimiteInvalCat1_Retraite
  rs.fields("PEINVAL1CAT").Value = Inval1Cat
  
  ' Inval Viagere
  rs.fields("PELMINVAL_Viagere").Value = LoiInvalidite_Viagere
  
  ' Dépendance
  rs.fields("PELMDEPENDANCE").Value = LoiDependance
  
  rs.fields("PETXREVENUC7").Value = TauxRevenuC7
  rs.fields("PEGESTIONC7").Value = FraisGestionC7
      
  rs.fields("PETXREVALO").Value = TauxIndexation
  rs.fields("PEDUREEREVALO").Value = DureeIndexation
  rs.fields("PEDUREELISSAGE").Value = DureeLissage
  rs.fields("PETMO").Value = TMO
  
  rs.fields("PEAGERETRAITE").Value = 65
  
  ' rente education
  rs.fields("PETRENTEEDUC").Value = TableRenteEducation
  rs.fields("PETXRTEDUC").Value = TauxTechniqueRenteEducation
  rs.fields("PEGESTIONEDUC").Value = FraisGestionRenteEducation
  rs.fields("PEPAIEMENTEDUC").Value = PaiementRenteEducation
  rs.fields("PEFRACTIONEDUC").Value = FractionnementRenteEducation
  
  ' rente conjoint
  rs.fields("PETRENTECONJOINT").Value = TableRenteConjoint
  rs.fields("PETXRTCONJOINT").Value = TauxTechniqueRenteConjoint
  rs.fields("PEGESTIONCONJOINT").Value = FraisGestionRenteConjoint
  rs.fields("PEPAIEMENTCONJOINT").Value = PaiementRenteConjoint
  rs.fields("PEFRACTIONCONJOINT").Value = FractionnementRenteConjoint
  
  ' maintien deces
  rs.fields("PEGestionDeces_MDC").Value = FraisGestionCapitauxDecesDC
  rs.fields("PEGestionEduc_MDC").Value = FraisGestionRenteEducationDC
  
  rs.fields("PEPCTACCIDENT").Value = PourcentageAccident
  
  rs.fields("PEGestionConjoint_MDC").Value = FraisGestionRenteConjointDC
  rs.fields("PECapConstitMoyenTempo_MDC").Value = CapitalMoyenRenteConjointTempoDC
  rs.fields("PECapConstitMoyenViagere_MDC").Value = CapitalMoyenRenteConjointViagereDC
  rs.fields("PEForcerCapConstitMoyen_MDC").Value = ForcerCapitalMoyenRenteConjointDC
  rs.fields("PEAgeConjoint_MDC").Value = AgeConjointRenteConjointDC
  
  rs.fields("PELissageProvision").Value = UtiliserTableLissageProvision
  
  
  '
  ' Maintien DC
  '
  
  'datPrimaryRS.Recordset.fields("PEPMGDForcerInval") = IIf(chkPMGDForcerInval = vbChecked= 1= 0)
  rs.fields("PERecalculBCAC").Value = RecalculCoeffBCAC
  
  rs.fields("PEAgeLimiteCalul_MDC").Value = AgeLimiteCalulDC
  rs.fields("PEAgeLimiteCalul_MDC_Retraite").Value = AgeLimiteCalulDC_Retraite
  
  rs.fields("PEAgeLimiteCalul_MDC_Inval1").Value = AgeLimiteCalulDC_Inval1
  rs.fields("PEAgeLimiteCalul_MDC_Retraite_Inval1").Value = AgeLimiteCalulDC_Retraite_Inval1
  
  rs.fields("PEMethodeCalcul").Value = MethodeCalculDC
  rs.fields("PEPctPMCalculeeDC").Value = PctPMCalculeeDC

  ' Incap DC : ID et non NOM !!!
  rs.fields("PECleTableMortBCACIncap_MDC").Value = GetCleTable(LoiIncapaciteDC)
  rs.fields("PECleTableMortBCACIncap_MDC_Retraite").Value = GetCleTable(LoiIncapaciteDC_Retraite)
  
  ' Invalidite DC : ID et non NOM !!!
  rs.fields("PECleTableMortBCACInval_MDC").Value = GetCleTable(LoiInvaliditeDC)
  rs.fields("PECleTableMortBCACInval_MDC_Retraite").Value = GetCleTable(LoiInvaliditeDC_Retraite)
  
  rs.fields("PETauxTechRecalculBCAC_MDC").Value = TauxTechnicCalculDC
  rs.fields("PEGestionRecalculBCAC_MDC").Value = FraisGestionCalculDC
  
  ' table precalculée Incap DC : ID et non NOM !!!
  rs.fields("PECleTableBCACIncap_MDC").Value = GetCleTable(TableIncapacitePrecalculDC)
  rs.fields("PECleTableBCACIncap_MDC_Retraite").Value = GetCleTable(TableIncapacitePrecalculDC_Retraite)
  
  ' table precalculée Inval DC : ID et non NOM !!!
  rs.fields("PECleTableBCACInval_MDC").Value = GetCleTable(TableInvaliditePrecalculDC)
  rs.fields("PECleTableBCACInval_MDC_Retraite").Value = GetCleTable(TableInvaliditePrecalculDC_Retraite)
  
  
  ' calcul de l'age
  rs.fields("PEAge_Anniversaire").Value = CalculAge_Anniversaire

  
  ' méthode d'interpollation inval
  rs.fields("PEINTERPOLATION_INVAL").Value = InterpolationInvalidite
  
  ' Calcul du passage sans lissage : INVERSER DANS LA BDD !!!
  rs.fields("PESANSLISSAGE_PASSAGE") = Not LissageCoeffPassage
  
  rs.fields("PEANNUALISATION_PASSAGE") = AnnualisationPassage
  
  rs.fields("PEBRIDAGE_AGE") = BridageAgeLimiteTable
  
  
  rs.fields("PEPortefeuilleSalaries") = PortefeuilleSalaries
  
  
  ' sauvegarde
  rs.Update
  
  loadedRECNO = rs.fields("RECNO").Value
  
  rs.Close
  
  SaveToDB = True
  
  Exit Function
  
err_SaveToDB:
  MsgBox "erreur " & Err & vbLf & Err.Description & vbLf & vbLf & " Sauvegarde annulée !" = vbCritical
  rs.CancelUpdate
  SaveToDB = False
  
  Exit Function
  
  Resume Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fabrique une chaine de description des paramètres
'
'##ModelId=5C8A75F90218
Public Function DumpParam() As String
  Dim sParam As String
  
  On Error GoTo err_DumpParam
  
  sParam = "======================================================================" & vbLf
  
  sParam = sParam & "Paramètre de calcul n° = " & CodeParamCalcul & vbLf & vbLf
  sParam = sParam & "Nom = " & NomParamCalcul & vbLf & vbLf
  sParam = sParam & "Commentaire = " & Commentaire & vbLf & vbLf
   
  sParam = sParam & "PortefeuilleSalaries = " & IIf(PortefeuilleSalaries = True, "Oui", "Non") & vbLf & vbLf
  
  ' calcul de l'age
  sParam = sParam & "Calcul des âges : " & vbLf
  sParam = sParam & "    Méthode = " & IIf(CalculAge_Anniversaire = True, "Age anniversaire", "Différence de millesimes") & vbLf
  sParam = sParam & "    Bridage âge limite de la table = " & IIf(BridageAgeLimiteTable = True, "Oui", "Non") & vbLf & vbLf
    
  ' Incap
  sParam = sParam & "Provisions Incapacité : " & vbLf
  sParam = sParam & "    Loi de maintien avant réforme = " & GetLibelleTable(LoiIncapacite) & vbLf
  sParam = sParam & "    Loi de passage avant réforme = " & GetLibelleTable(LoiPassage) & vbLf
  sParam = sParam & "    Taux Technique = " & TauxIncapacite & "%" & vbLf
  sParam = sParam & "    Frais de gestion = " & FraisGestionIncapacite & "%" & vbLf
  sParam = sParam & "    Taux Incap avec Passage Inval Cat 1 nouvelle table age depart retraite specifique inval 1 = " & TauxIncapPassageSpecifiqueInval1 & "%" & vbLf
  sParam = sParam & "    Loi de maintien en incapacite après réforme = " & GetLibelleTable(LoiIncapacite_Retraite) & vbLf
  sParam = sParam & "    Loi de passage après réforme = " & GetLibelleTable(LoiPassage_Retraite) & vbLf
  sParam = sParam & "    Interpoler les coeffs de provision Incapacité et Passage = Oui" & vbLf
  sParam = sParam & "    Lissage du coefficient de passage = " & IIf(LissageCoeffPassage = True, "Oui", "Non") & vbLf & vbLf
  sParam = sParam & "    Annualisation réduite de la provision de passage = " & IIf(AnnualisationPassage = True, "Oui", "Non") & vbLf & vbLf
  sParam = sParam & "    Durée Chômage = " & DureeChomage & " mois " & vbLf

  ' Inval
  sParam = sParam & "Provisions Invalidité : " & vbLf
  sParam = sParam & "    Loi de maintien en invalidité avant réforme = " & GetLibelleTable(LoiInvalidite) & vbLf
  sParam = sParam & "    Taux technique = " & TauxInvalidite & "%" & vbLf
  sParam = sParam & "    Frais de gestion = " & FraisGestionInvalidite & "%" & vbLf
  sParam = sParam & "    Loi de maintien en invalidité après réforme = " & GetLibelleTable(LoiInvalidite_Retraite) & vbLf
  sParam = sParam & "    Age limite pour les invalides en 1ere catégorie avant réforme = " & AgeLimiteInvalCat1 & " ans" & vbLf
  sParam = sParam & "    Age limite pour les invalides en 1ere catégorie après réforme = " & AgeLimiteInvalCat1_Retraite & " ans" & vbLf
  sParam = sParam & "    Interpolation = " & InterpollationInvaliditeToString(InterpolationInvalidite) & vbLf & vbLf
  sParam = sParam & "    % des invalides en 1ere catégorie = " & Inval1Cat & "% " & vbLf
  
  ' Inval Viagere
  sParam = sParam & "    Loi de maintien en invalidité viagère = " & GetLibelleTable(LoiInvalidite_Viagere) & vbLf
  
  ' Dépendance
  sParam = sParam & "Provisions Dépendance : " & vbLf
  sParam = sParam & "    Loi de maintien  = " & GetLibelleTable(LoiDependance) & vbLf
  
  ' etats C7
  sParam = sParam & "Etats C7 : " & vbLf
  sParam = sParam & "    Taux des revenus financiers = " & TauxRevenuC7 & "%" & vbLf
  sParam = sParam & "    Frais de gestion = " & FraisGestionC7 & "%" & vbLf & vbLf
  
  ' Revalo
  sParam = sParam & "Reprise de revalorisation : " & vbLf
  sParam = sParam & "    Taux d'ndexation = " & TauxIndexation & "%" & vbLf
  sParam = sParam & "    Duree d'indexation = " & DureeIndexation & " ans" & vbLf
  sParam = sParam & "    Taux financier TME  = " & TMO & "%" & vbLf
  sParam = sParam & "    Durée de lissage de la Prime Unique = " & DureeLissage & " ans" & vbLf & vbLf
    
  ' rentes
  sParam = sParam & "Rentes = Paramètres par défaut si Catégorie/Survenance absents : " & vbLf
  sParam = sParam & "    Table = " & GetLibelleTable(TableRenteEducation) & vbLf
  sParam = sParam & "    Taux technique = " & TauxTechniqueRenteEducation & "%" & vbLf
  sParam = sParam & "    Frais de gestion = " & FraisGestionRenteEducation & "%" & vbLf
  sParam = sParam & "    Fractionnement = " & FractionnementToString(FractionnementRenteEducation) & vbLf
  sParam = sParam & "    Paiement = " & TermeToString(PaiementRenteEducation) & vbLf & vbLf

  ' maintien deces
  sParam = sParam & "Provisions Maintien en garantie décès : " & vbLf
  sParam = sParam & "    Capitaux décès : " & vbLf
  sParam = sParam & "        Frais de gestion = " & FraisGestionCapitauxDecesDC & "%" & vbLf
  sParam = sParam & "        Pourcentage de décès par accident = " & PourcentageAccident & "%" & vbLf & vbLf
  
  sParam = sParam & "    Rente éducation : " & vbLf
  sParam = sParam & "        Frais de gestion = " & FraisGestionRenteEducationDC & "%" & vbLf & vbLf
  
  sParam = sParam & "    Rente de conjoint : " & vbLf
  sParam = sParam & "        Frais de gestion = " & FraisGestionRenteConjointDC & "%" & vbLf
  sParam = sParam & "        Age du conjoint = " & IIf(AgeConjointRenteConjointDC >= 0, "+", "-") & AgeConjointRenteConjointDC & " ans" & vbLf
  sParam = sParam & "        Forcer capital constitutif moyen = " & IIf(ForcerCapitalMoyenRenteConjointDC = vbChecked, "Oui", "Non") & vbLf
  sParam = sParam & "        Capital moyen rente temporaire = " & CapitalMoyenRenteConjointTempoDC & vbLf
  sParam = sParam & "        Capital moyen rente viagère = " & CapitalMoyenRenteConjointViagereDC & vbLf & vbLf
  
  ' Methode lissage provision
  sParam = sParam & "    Pourcentage de provisions à constituer = " & IIf(UtiliserTableLissageProvision = True, "Utiliser la table 'LissageProvision'", "100%") & vbLf & vbLf
  
  '
  ' recalcul
  '
  
  sParam = sParam & "    Méthode de calcul des Provisions : " & vbLf
  sParam = sParam & "        Méthode  = " & MethodeCalculDCToString(MethodeCalculDC) & vbLf
  sParam = sParam & "        % de la Provision calculée = " & PctPMCalculeeDC & "%" & vbLf & vbLf
  
  sParam = sParam & "    Coefficient de provision : " & vbLf
  sParam = sParam & "        Méthode = " & IIf(RecalculCoeffBCAC = True, "Calcul du coefficient", "Utiliser les tables de coefficients précalculés du BCAC") & vbLf & vbLf
  
  sParam = sParam & "        Calcul du coefficient :" & vbLf
  sParam = sParam & "            Mortalité Incap avant réforme = " & GetLibelleTable(LoiIncapaciteDC) & vbLf
  sParam = sParam & "            Mortalité Incap après réforme = " & GetLibelleTable(LoiIncapaciteDC_Retraite) & vbLf
  sParam = sParam & "            Mortalité Inval avant réforme = " & GetLibelleTable(LoiInvaliditeDC) & vbLf
  sParam = sParam & "            Mortalité Inval après réforme = " & GetLibelleTable(LoiInvaliditeDC_Retraite) & vbLf
  sParam = sParam & "            Taux technique = " & TauxTechnicCalculDC & "%" & vbLf
  sParam = sParam & "            Frais de gestion = " & FraisGestionCalculDC & "%" & vbLf
  sParam = sParam & "            Age limite avant réforme = " & AgeLimiteCalulDC & " ans" & vbLf
  sParam = sParam & "            Age limite après réforme = " & AgeLimiteCalulDC_Retraite & " ans" & vbLf & vbLf
  sParam = sParam & "            Age limite catégorie invadité 1 = " & AgeLimiteCalulDC_Inval1 & " ans" & vbLf
  sParam = sParam & "            Age limite catégorie invadité 1 après réforme = " & AgeLimiteCalulDC_Retraite_Inval1 & " ans" & vbLf & vbLf
  
  
  ' coeff precalculé dc
  sParam = sParam & "        Coefficient précalculé :" & vbLf
  sParam = sParam & "            Table Incap avant réforme = " & GetLibelleTable(TableIncapacitePrecalculDC) & vbLf
  sParam = sParam & "            Table Incap après réforme = " & GetLibelleTable(TableIncapacitePrecalculDC_Retraite) & vbLf
  sParam = sParam & "            Table Inval avant réforme = " & GetLibelleTable(TableInvaliditePrecalculDC) & vbLf
  sParam = sParam & "            Table Inval après réforme = " & GetLibelleTable(TableInvaliditePrecalculDC_Retraite) & vbLf & vbLf

  sParam = sParam & "======================================================================" & vbLf

  DumpParam = sParam
  
  Exit Function
  
err_DumpParam:
  DumpParam = Err.Description
  
  Exit Function
  
  Resume Next
End Function


'##ModelId=5C8A75F90237
Public Function TermeToString(tp As EcheancePaiement) As String

 Select Case tp
    Case ePaiementAvance
      TermeToString = "d'avance"
    Case ePaiementEchu
      TermeToString = "à terme échu"
      
    Case Else
      TermeToString = tp
  End Select

End Function


'##ModelId=5C8A75F90264
Public Function FractionnementToString(tp As Fractionnement) As String
  
  Select Case tp
    Case eFractionnementAnnuel
      FractionnementToString = "Annuel"
    Case eFractionnementSemestriel
      FractionnementToString = "Semestriel"
    Case eFractionnementTrimestriel
      FractionnementToString = "Trimestriel"
    Case eFractionnementMensuel
      FractionnementToString = "Mensuel"
    
    Case Else
      FractionnementToString = tp
  End Select
  
End Function


'##ModelId=5C8A75F90284
Public Function InterpollationInvaliditeToString(tp As InterpolationInval) As String

  Select Case tp
    Case eInterpolationInval_CorrectionDuree
      InterpollationInvaliditeToString = "Correction si durée restante <12 mois"
    Case eInterpolationInval_Age
      InterpollationInvaliditeToString = "Interpolation sur l'âge"
    Case eInterpolationInval_AgeDuree
      InterpollationInvaliditeToString = "Interpolation sur l'âge et la durée"
    Case Else
      InterpollationInvaliditeToString = "Aucune interpolation"
  End Select

End Function


'##ModelId=5C8A75F902B2
Public Function MethodeCalculDCToString(tp As MethodeCalculMaintienDC) As String

  Select Case tp
    Case eCapitauxConstitutifs
      MethodeCalculDCToString = "Capitaux constitutifs sous risque"
    Case ePctProvisionCalculee
      MethodeCalculDCToString = "% de la Provision calculée"
    Case ePasDeCalcul
      MethodeCalculDCToString = "Aucun calcul"
  End Select

End Function

