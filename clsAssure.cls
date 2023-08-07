VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C90E96B0126"

Option Explicit

'
' Members
'

'##ModelId=5C90E96B02B9
Public m_RECNO As Variant
'##ModelId=5C90E96B02DA
Public m_POGPECLE As Variant
'##ModelId=5C90E96B0302
Public m_POPERCLE As Variant
'##ModelId=5C90E96B033B
Public m_POSTECLE As Variant
'##ModelId=5C90E96B035A
Public m_PONumParamCalcul As Variant
'##ModelId=5C90E96B037B
Public m_PONUMCLE As Variant
'##ModelId=5C90E96B0399
Public m_POGARCLE As Variant
'##ModelId=5C90E96B03B7
Public m_POCONVENTION As Variant
'##ModelId=5C90E96B03D5
Public m_PONOM As Variant
'##ModelId=5C90E96C000E
Public m_POSEXE As Variant
'##ModelId=5C90E96C002A
Public m_POCSP As Variant
'##ModelId=5C90E96C0044
Public m_PONAIS As Variant
'##ModelId=5C90E96C0061
Public m_POEFFET As Variant
'##ModelId=5C90E96C0093
Public m_POTERME As Variant
'##ModelId=5C90E96C00AF
Public m_POARRET As Variant
'##ModelId=5C90E96C00D0
Public m_POREPRISE As Variant
'##ModelId=5C90E96C00EC
Public m_POCAUSE As Variant
'##ModelId=5C90E96C0113
Public m_POTYPEF As Variant
'##ModelId=5C90E96C0141
Public m_PODELAI As Variant
'##ModelId=5C90E96C015B
Public m_POPREMIER_PAIEMENT As Variant
'##ModelId=5C90E96C0178
Public m_POPRESTATION As Variant
'##ModelId=5C90E96C01AF
Public m_POPRESTATIONTOTAL As Variant
'##ModelId=5C90E96C01CB
Public m_POPRESTATION_AN_PREC As Variant
'##ModelId=5C90E96C01E6
Public m_POPRESTATION_AN As Variant
'##ModelId=5C90E96C0211
Public m_POSIT As Variant
'##ModelId=5C90E96C024D
Public m_POPM_INCAP_1F As Variant
'##ModelId=5C90E96C027A
Public m_POPM_PASS_1F As Variant
'##ModelId=5C90E96C029C
Public m_POPM_INVAL_1F As Variant
'##ModelId=5C90E96C02CD
Public m_POPM As Variant
'##ModelId=5C90E96C02EC
Public m_POPM_SORTIE As Variant
'##ModelId=5C90E96C032A
Public m_POPM_VAR As Variant
'##ModelId=5C90E96C034E
Public m_PODERNIERPAIEMENT As Variant
'##ModelId=5C90E96C0378
Public m_PODEBUT As Variant
'##ModelId=5C90E96C03A5
Public m_POFIN As Variant
'##ModelId=5C90E96C03DD
Public m_PODEBUTTOTAL As Variant
'##ModelId=5C90E96D001D
Public m_POFINTOTAL As Variant
'##ModelId=5C90E96D0049
Public m_POPSAP As Variant
'##ModelId=5C90E96D006F
Public m_POPM_RI As Variant
'##ModelId=5C90E96D0092
Public m_POPM_REVALO As Variant
'##ModelId=5C90E96D00CF
Public m_POCOT_REVALO As Variant
'##ModelId=5C90E96D00F2
Public m_PODATEIMPORT As Variant
'##ModelId=5C90E96D0114
Public m_POPM_INCAP_1R As Variant
'##ModelId=5C90E96D0153
Public m_POPM_PASS_1R As Variant
'##ModelId=5C90E96D0175
Public m_POPM_INVAL_1R As Variant
'##ModelId=5C90E96D01A4
Public m_POCATEGORIE As Variant
'##ModelId=5C90E96D01C8
Public m_POPM_RCJT_1F As Variant
'##ModelId=5C90E96D01FF
Public m_POPM_REDUC_1F As Variant
'##ModelId=5C90E96D0220
Public m_POPRESTA_RASSUR As Variant
'##ModelId=5C90E96D0248
Public m_POPM_RASSUR As Variant
'##ModelId=5C90E96D0286
Public m_POPSAP_RASSUR As Variant
'##ModelId=5C90E96D02AB
Public m_POTRAITE_RASSUR As Variant
'##ModelId=5C90E96D02DA
Public m_POTRAITE_RASSUR2 As Variant
'##ModelId=5C90E96D02FE
Public m_POTRAITE_RASSUR_TAUX As Variant
'##ModelId=5C90E96D0334
Public m_PODATEPAIEMENTESTIMEE As Variant
'##ModelId=5C90E96D035A
Public m_POPSAPCAPMOYEN As Variant
'##ModelId=5C90E96D0394
Public m_PODOSSIERCLOS As Variant
'##ModelId=5C90E96D03B7
Public m_POGARCLE_NEW As Variant
'##ModelId=5C90E96E0008
Public m_POCATEGORIE_NEW As Variant
'##ModelId=5C90E96E002D
Public m_POREGROUPEMENT As Variant
'##ModelId=5C90E96E0065
Public m_POCODENATURE As Variant
'##ModelId=5C90E96E008A
Public m_POCCN As Variant
'##ModelId=5C90E96E00C3
Public m_POCODERISQUE As Variant
'##ModelId=5C90E96E00E7
Public m_POPMANNULEE As Variant
'##ModelId=5C90E96E0120
Public m_POPSAPANNULEE As Variant
'##ModelId=5C90E96E0144
Public m_POCOMMENTANNUL As Variant
'##ModelId=5C90E96E0168
Public m_PODATEENTREEINVAL As Variant
'##ModelId=5C90E96E018E
Public m_POPM_RCJT_1R As Variant
'##ModelId=5C90E96E01CD
Public m_POPM_REDUC_1R As Variant
'##ModelId=5C90E96E0209
Public m_POIsCadre As Variant
'##ModelId=5C90E96E0239
Public m_POSalaireAnnuel As Variant
'##ModelId=5C90E96E0261
Public m_POCleSituationFamille As Variant
'##ModelId=5C90E96E0288
Public m_POTauxGarantieDC As Variant
'##ModelId=5C90E96E02BE
Public m_PONbEnfant As Variant
'##ModelId=5C90E96E02E6
Public m_POAgeMoyenEnfant As Variant
'##ModelId=5C90E96E0316
Public m_POMajoEnfant As Variant
'##ModelId=5C90E96E0346
Public m_PORegimeDeces As Variant
'##ModelId=5C90E96E0374
Public m_POCategorieDeces As Variant
'##ModelId=5C90E96E039A
Public m_PORegimeRenteConjointTempo As Variant
'##ModelId=5C90E96E03CE
Public m_POCategorieRenteConjointTempo As Variant
'##ModelId=5C90E96F0018
Public m_PORegimeRenteConjointViagere As Variant
'##ModelId=5C90E96F0049
Public m_POCategorieRenteConjointViagere As Variant
'##ModelId=5C90E96F0070
Public m_PORegimeRenteEduc As Variant
'##ModelId=5C90E96F0097
Public m_POCategorieRenteEduc As Variant
'##ModelId=5C90E96F00BE
Public m_POPourcentLissage As Variant
'##ModelId=5C90E96F00E7
Public m_POCoeffBCAC As Variant
'##ModelId=5C90E96F0126
Public m_POContractant As Variant
'##ModelId=5C90E96F015C
Public m_POSituConv As Variant
'##ModelId=5C90E96F0186
Public m_POEffetSitu As Variant
'##ModelId=5C90E96F01B9
Public m_POCode_Option_Contrat As Variant
'##ModelId=5C90E96F01E3
Public m_POEtablissement As Variant
'##ModelId=5C90E96F0225
Public m_PONumSinistre As Variant
'##ModelId=5C90E96F025B
Public m_POTypeMvt As Variant
'##ModelId=5C90E96F0287
Public m_POIJ As Variant
'##ModelId=5C90E96F02BC
Public m_POMontantBase As Variant
'##ModelId=5C90E96F02E8
Public m_POMontantRevalo As Variant
'##ModelId=5C90E96F0310
Public m_POTypeReglement As Variant
'##ModelId=5C90E96F0354
Public m_PODebutIndemnisation As Variant
'##ModelId=5C90E96F038D
Public m_PONbJourIndemn As Variant
'##ModelId=5C90E96F03B7
Public m_POCreationDossier As Variant
'##ModelId=5C90E9700006
Public m_PODebutDossier As Variant
'##ModelId=5C90E970003E
Public m_POFinDossier As Variant
'##ModelId=5C90E970007C
Public m_POMotifCloture As Variant
'##ModelId=5C90E97000B3
Public m_PODebutRefSalaire As Variant
'##ModelId=5C90E97000DC
Public m_POFinRefSalaire As Variant
'##ModelId=5C90E9700122
Public m_POFamilleComptable As Variant
'##ModelId=5C90E970015A
Public m_POInspecteur As Variant
'##ModelId=5C90E9700185
Public m_POInsp As Variant
'##ModelId=5C90E97001BB
Public m_POApport As Variant
'##ModelId=5C90E97001F3
Public m_POApport2 As Variant
'##ModelId=5C90E9700239
Public m_POGestionnaire As Variant
'##ModelId=5C90E9700270
Public m_POIndicCC As Variant
'##ModelId=5C90E97002B7
Public m_POCompagnie As Variant
'##ModelId=5C90E9700302
Public m_POAppli As Variant
'##ModelId=5C90E970032D
Public m_PONbIntervenant As Variant
'##ModelId=5C90E970035D
Public m_POPM_AvecCorrectif As Variant
'##ModelId=5C90E9700389
Public m_POPM_SansCorrectif As Variant
'##ModelId=5C90E97003B2
Public m_POCorrectif As Variant
'##ModelId=5C90E97003DD
Public m_POPMReassAvecCorrectif As Variant
'##ModelId=5C90E9710022
Public m_POPMAvecCorrectif As Variant
'##ModelId=5C90E971005B
Public m_POPM_X As Variant
'##ModelId=5C90E9710086
Public m_POPM_XTERME As Variant
'##ModelId=5C90E97100C0
Public m_POPM_ANC As Variant
'##ModelId=5C90E97100FE
Public m_POPM_DUREE As Variant
'##ModelId=5C90E971013A
Public m_POCaptive As Variant
'##ModelId=5C90E971017A
Public m_POBaseRevalo As Variant
'##ModelId=5C90E97101BE
Public m_POFRACT As Variant
'##ModelId=5C90E9710204
Public m_POECHU As Variant
'##ModelId=5C90E971022F
Public m_POMontantMajore As Variant
'##ModelId=5C90E9710266
Public m_NUENRP3I As Variant
'##ModelId=5C90E9710293
Public m_TXTECHN As Variant
'##ModelId=5C90E97102C0
Public m_TXFRAIS As Variant
  
'Evol 2010 - Lot 2
'##ModelId=5C90E97102F0
Public m_POCDCONTENTIEUX  As Variant
'##ModelId=5C90E9710339
Public m_PONUSINISTRE  As Variant
'##ModelId=5C90E9710366
Public m_POCategorieInval  As Variant
'##ModelId=5C90E971039F
Public m_POMontantCapConstit  As Variant
'##ModelId=5C90E97103CE
Public m_POMontantCapSousRisque  As Variant

' Evolution Reforme des retraites, table Assure_Retraite
'##ModelId=5C90E9720017
Public m_IdAssure As Variant
'##ModelId=5C90E9720052
Public m_TopAmortissable As Variant
'##ModelId=5C90E9720087
Public m_CoeffAmortissement As Variant
'##ModelId=5C90E97200BA
Public m_AgeMiniDepartRetraite As Variant
'##ModelId=5C90E9720102
Public m_AgeDepartRetraiteTauxPlein As Variant

'Genarali : Add column comment
'##ModelId=5C90E972013B
Public m_Commentaire As Variant



'
' Constructors
'


'##ModelId=5C90E9720169
Private Sub Class_Initialize()
    
  m_RECNO = Null
  m_POGPECLE = Null
  m_POPERCLE = Null
  m_POSTECLE = Null
  m_PONumParamCalcul = Null
  m_PONUMCLE = Null
  m_POGARCLE = Null
  m_POCONVENTION = Null
  m_PONOM = Null
  m_POSEXE = Null
  m_POCSP = Null
  m_PONAIS = Null
  m_POEFFET = Null
  m_POTERME = Null
  m_POARRET = Null
  m_POREPRISE = Null
  m_POCAUSE = Null
  m_POTYPEF = Null
  m_PODELAI = Null
  m_POPREMIER_PAIEMENT = Null
  m_POPRESTATION = Null
  m_POPRESTATIONTOTAL = Null
  m_POPRESTATION_AN_PREC = Null
  m_POPRESTATION_AN = Null
  m_POSIT = Null
  m_POPM_INCAP_1F = Null
  m_POPM_PASS_1F = Null
  m_POPM_INVAL_1F = Null
  m_POPM = Null
  m_POPM_SORTIE = Null
  m_POPM_VAR = Null
  m_PODERNIERPAIEMENT = Null
  m_PODEBUT = Null
  m_POFIN = Null
  m_PODEBUTTOTAL = Null
  m_POFINTOTAL = Null
  m_POPSAP = Null
  m_POPM_RI = Null
  m_POPM_REVALO = Null
  m_POCOT_REVALO = Null
  m_PODATEIMPORT = Null
  m_POPM_INCAP_1R = Null
  m_POPM_PASS_1R = Null
  m_POPM_INVAL_1R = Null
  m_POCATEGORIE = Null
  m_POPM_RCJT_1F = Null
  m_POPM_REDUC_1F = Null
  m_POPRESTA_RASSUR = Null
  m_POPM_RASSUR = Null
  m_POPSAP_RASSUR = Null
  m_POTRAITE_RASSUR = Null
  m_POTRAITE_RASSUR2 = Null
  m_POTRAITE_RASSUR_TAUX = Null
  m_PODATEPAIEMENTESTIMEE = Null
  m_POPSAPCAPMOYEN = Null
  m_PODOSSIERCLOS = Null
  m_POGARCLE_NEW = Null
  m_POCATEGORIE_NEW = Null
  m_POREGROUPEMENT = Null
  m_POCODENATURE = Null
  m_POCCN = Null
  m_POCODERISQUE = Null
  m_POPMANNULEE = Null
  m_POPSAPANNULEE = Null
  m_POCOMMENTANNUL = Null
  m_PODATEENTREEINVAL = Null
  m_POPM_RCJT_1R = Null
  m_POPM_REDUC_1R = Null
  m_POIsCadre = Null
  m_POSalaireAnnuel = Null
  m_POCleSituationFamille = Null
  m_POTauxGarantieDC = Null
  m_PONbEnfant = Null
  m_POAgeMoyenEnfant = Null
  m_POMajoEnfant = Null
  m_PORegimeDeces = Null
  m_POCategorieDeces = Null
  m_PORegimeRenteConjointTempo = Null
  m_POCategorieRenteConjointTempo = Null
  m_PORegimeRenteConjointViagere = Null
  m_POCategorieRenteConjointViagere = Null
  m_PORegimeRenteEduc = Null
  m_POCategorieRenteEduc = Null
  m_POPourcentLissage = Null
  m_POCoeffBCAC = Null
  m_POContractant = Null
  m_POSituConv = Null
  m_POEffetSitu = Null
  m_POCode_Option_Contrat = Null
  m_POEtablissement = Null
  m_PONumSinistre = Null
  m_POTypeMvt = Null
  m_POIJ = Null
  m_POMontantBase = Null
  m_POMontantRevalo = Null
  m_POTypeReglement = Null
  m_PODebutIndemnisation = Null
  m_PONbJourIndemn = Null
  m_POCreationDossier = Null
  m_PODebutDossier = Null
  m_POFinDossier = Null
  m_POMotifCloture = Null
  m_PODebutRefSalaire = Null
  m_POFinRefSalaire = Null
  m_POFamilleComptable = Null
  m_POInspecteur = Null
  m_POInsp = Null
  m_POApport = Null
  m_POApport2 = Null
  m_POGestionnaire = Null
  m_POIndicCC = Null
  m_POCompagnie = Null
  m_POAppli = Null
  m_PONbIntervenant = Null
  m_POPM_AvecCorrectif = Null
  m_POPM_SansCorrectif = Null
  m_POCorrectif = Null
  m_POPMReassAvecCorrectif = Null
  m_POPMAvecCorrectif = Null
  m_POPM_X = Null
  m_POPM_XTERME = Null
  m_POPM_ANC = Null
  m_POPM_DUREE = Null
  m_POCaptive = Null
  m_POBaseRevalo = Null
  m_POFRACT = Null
  m_POECHU = Null
  m_POMontantMajore = Null
  m_NUENRP3I = Null
  m_TXTECHN = Null
  m_TXFRAIS = Null
  
  m_POCDCONTENTIEUX = Null
  m_PONUSINISTRE = Null
  m_POCategorieInval = Null
  m_POMontantCapConstit = Null
  m_POMontantCapSousRisque = Null

  m_IdAssure = Null
  m_TopAmortissable = Null
  m_CoeffAmortissement = Null
  m_AgeMiniDepartRetraite = Null
  m_AgeDepartRetraiteTauxPlein = Null
  
  m_Commentaire = Null
  
End Sub


'##ModelId=5C90E97201A5
Public Sub Copy(ByRef src As clsAssure)

  m_RECNO = src.m_RECNO
  m_POGPECLE = src.m_POGPECLE
  m_POPERCLE = src.m_POPERCLE
  m_POSTECLE = src.m_POSTECLE
  m_PONumParamCalcul = src.m_PONumParamCalcul
  m_PONUMCLE = src.m_PONUMCLE
  m_POGARCLE = src.m_POGARCLE
  m_POCONVENTION = src.m_POCONVENTION
  m_PONOM = src.m_PONOM
  m_POSEXE = src.m_POSEXE
  m_POCSP = src.m_POCSP
  m_PONAIS = src.m_PONAIS
  m_POEFFET = src.m_POEFFET
  m_POTERME = src.m_POTERME
  m_POARRET = src.m_POARRET
  m_POREPRISE = src.m_POREPRISE
  m_POCAUSE = src.m_POCAUSE
  m_POTYPEF = src.m_POTYPEF
  m_PODELAI = src.m_PODELAI
  m_POPREMIER_PAIEMENT = src.m_POPREMIER_PAIEMENT
  m_POPRESTATION = src.m_POPRESTATION
  m_POPRESTATIONTOTAL = src.m_POPRESTATIONTOTAL
  m_POPRESTATION_AN_PREC = src.m_POPRESTATION_AN_PREC
  m_POPRESTATION_AN = src.m_POPRESTATION_AN
  m_POSIT = src.m_POSIT
  m_POPM_INCAP_1F = src.m_POPM_INCAP_1F
  m_POPM_PASS_1F = src.m_POPM_PASS_1F
  m_POPM_INVAL_1F = src.m_POPM_INVAL_1F
  m_POPM = src.m_POPM
  m_POPM_SORTIE = src.m_POPM_SORTIE
  m_POPM_VAR = src.m_POPM_VAR
  m_PODERNIERPAIEMENT = src.m_PODERNIERPAIEMENT
  m_PODEBUT = src.m_PODEBUT
  m_POFIN = src.m_POFIN
  m_PODEBUTTOTAL = src.m_PODEBUTTOTAL
  m_POFINTOTAL = src.m_POFINTOTAL
  m_POPSAP = src.m_POPSAP
  m_POPM_RI = src.m_POPM_RI
  m_POPM_REVALO = src.m_POPM_REVALO
  m_POCOT_REVALO = src.m_POCOT_REVALO
  m_PODATEIMPORT = src.m_PODATEIMPORT
  m_POPM_INCAP_1R = src.m_POPM_INCAP_1R
  m_POPM_PASS_1R = src.m_POPM_PASS_1R
  m_POPM_INVAL_1R = src.m_POPM_INVAL_1R
  m_POCATEGORIE = src.m_POCATEGORIE
  m_POPM_RCJT_1F = src.m_POPM_RCJT_1F
  m_POPM_REDUC_1F = src.m_POPM_REDUC_1F
  m_POPRESTA_RASSUR = src.m_POPRESTA_RASSUR
  m_POPM_RASSUR = src.m_POPM_RASSUR
  m_POPSAP_RASSUR = src.m_POPSAP_RASSUR
  m_POTRAITE_RASSUR = src.m_POTRAITE_RASSUR
  m_POTRAITE_RASSUR2 = src.m_POTRAITE_RASSUR2
  m_POTRAITE_RASSUR_TAUX = src.m_POTRAITE_RASSUR_TAUX
  m_PODATEPAIEMENTESTIMEE = src.m_PODATEPAIEMENTESTIMEE
  m_POPSAPCAPMOYEN = src.m_POPSAPCAPMOYEN
  m_PODOSSIERCLOS = src.m_PODOSSIERCLOS
  m_POGARCLE_NEW = src.m_POGARCLE_NEW
  m_POCATEGORIE_NEW = src.m_POCATEGORIE_NEW
  m_POREGROUPEMENT = src.m_POREGROUPEMENT
  m_POCODENATURE = src.m_POCODENATURE
  m_POCCN = src.m_POCCN
  m_POCODERISQUE = src.m_POCODERISQUE
  m_POPMANNULEE = src.m_POPMANNULEE
  m_POPSAPANNULEE = src.m_POPSAPANNULEE
  m_POCOMMENTANNUL = src.m_POCOMMENTANNUL
  m_PODATEENTREEINVAL = src.m_PODATEENTREEINVAL
  m_POPM_RCJT_1R = src.m_POPM_RCJT_1R
  m_POPM_REDUC_1R = src.m_POPM_REDUC_1R
  m_POIsCadre = src.m_POIsCadre
  m_POSalaireAnnuel = src.m_POSalaireAnnuel
  m_POCleSituationFamille = src.m_POCleSituationFamille
  m_POTauxGarantieDC = src.m_POTauxGarantieDC
  m_PONbEnfant = src.m_PONbEnfant
  m_POAgeMoyenEnfant = src.m_POAgeMoyenEnfant
  m_POMajoEnfant = src.m_POMajoEnfant
  m_PORegimeDeces = src.m_PORegimeDeces
  m_POCategorieDeces = src.m_POCategorieDeces
  m_PORegimeRenteConjointTempo = src.m_PORegimeRenteConjointTempo
  m_POCategorieRenteConjointTempo = src.m_POCategorieRenteConjointTempo
  m_PORegimeRenteConjointViagere = src.m_PORegimeRenteConjointViagere
  m_POCategorieRenteConjointViagere = src.m_POCategorieRenteConjointViagere
  m_PORegimeRenteEduc = src.m_PORegimeRenteEduc
  m_POCategorieRenteEduc = src.m_POCategorieRenteEduc
  m_POPourcentLissage = src.m_POPourcentLissage
  m_POCoeffBCAC = src.m_POCoeffBCAC
  m_POContractant = src.m_POContractant
  m_POSituConv = src.m_POSituConv
  m_POEffetSitu = src.m_POEffetSitu
  m_POCode_Option_Contrat = src.m_POCode_Option_Contrat
  m_POEtablissement = src.m_POEtablissement
  m_PONumSinistre = src.m_PONumSinistre
  m_POTypeMvt = src.m_POTypeMvt
  m_POIJ = src.m_POIJ
  m_POMontantBase = src.m_POMontantBase
  m_POMontantRevalo = src.m_POMontantRevalo
  m_POTypeReglement = src.m_POTypeReglement
  m_PODebutIndemnisation = src.m_PODebutIndemnisation
  m_PONbJourIndemn = src.m_PONbJourIndemn
  m_POCreationDossier = src.m_POCreationDossier
  m_PODebutDossier = src.m_PODebutDossier
  m_POFinDossier = src.m_POFinDossier
  m_POMotifCloture = src.m_POMotifCloture
  m_PODebutRefSalaire = src.m_PODebutRefSalaire
  m_POFinRefSalaire = src.m_POFinRefSalaire
  m_POFamilleComptable = src.m_POFamilleComptable
  m_POInspecteur = src.m_POInspecteur
  m_POInsp = src.m_POInsp
  m_POApport = src.m_POApport
  m_POApport2 = src.m_POApport2
  m_POGestionnaire = src.m_POGestionnaire
  m_POIndicCC = src.m_POIndicCC
  m_POCompagnie = src.m_POCompagnie
  m_POAppli = src.m_POAppli
  m_PONbIntervenant = src.m_PONbIntervenant
  m_POPM_AvecCorrectif = src.m_POPM_AvecCorrectif
  m_POPM_SansCorrectif = src.m_POPM_SansCorrectif
  m_POCorrectif = src.m_POCorrectif
  m_POPMReassAvecCorrectif = src.m_POPMReassAvecCorrectif
  m_POPMAvecCorrectif = src.m_POPMAvecCorrectif
  m_POPM_X = src.m_POPM_X
  m_POPM_XTERME = src.m_POPM_XTERME
  m_POPM_ANC = src.m_POPM_ANC
  m_POPM_DUREE = src.m_POPM_DUREE
  m_POCaptive = src.m_POCaptive
  m_POBaseRevalo = src.m_POBaseRevalo
  m_POFRACT = src.m_POFRACT
  m_POECHU = src.m_POECHU
  m_POMontantMajore = src.m_POMontantMajore
  m_NUENRP3I = src.m_NUENRP3I
  m_TXTECHN = src.m_TXTECHN
  m_TXFRAIS = src.m_TXFRAIS
  
  m_POCDCONTENTIEUX = src.m_POCDCONTENTIEUX
  m_PONUSINISTRE = src.m_PONUSINISTRE
  m_POCategorieInval = src.m_POCategorieInval
  m_POMontantCapConstit = src.m_POMontantCapConstit
  m_POMontantCapSousRisque = src.m_POMontantCapSousRisque

  m_IdAssure = src.m_IdAssure
  m_TopAmortissable = src.m_TopAmortissable
  m_CoeffAmortissement = src.m_CoeffAmortissement
  m_AgeMiniDepartRetraite = src.m_AgeMiniDepartRetraite
  m_AgeDepartRetraiteTauxPlein = src.m_AgeDepartRetraiteTauxPlein
  
  m_Commentaire = src.m_Commentaire

End Sub



'---------------------------------------------------
'##ModelId=5C90E97201E1
Public Function Load(DB As DataAccess, RECNO As Long, bForReformeRetraite As Boolean) As Boolean

   Dim theData As ADODB.Recordset
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Load

   Set theData = DB.OpenRecordset("SELECT * FROM " & IIf(bForReformeRetraite, "Assure_Retraite", "Assure") & " WHERE  RECNO=" & RECNO, Snapshot)
   If theData.EOF = False Then

      If IsNull(theData.Fields("RECNO")) Then
         m_RECNO = 0
      Else
         m_RECNO = theData.Fields("RECNO")
      End If

      If IsNull(theData.Fields("POGPECLE")) Then
         m_POGPECLE = 0
      Else
         m_POGPECLE = theData.Fields("POGPECLE")
      End If

      If IsNull(theData.Fields("POPERCLE")) Then
         m_POPERCLE = 0
      Else
         m_POPERCLE = theData.Fields("POPERCLE")
      End If

      If IsNull(theData.Fields("POSTECLE")) Then
         m_POSTECLE = 0
      Else
         m_POSTECLE = theData.Fields("POSTECLE")
      End If

      If IsNull(theData.Fields("PONumParamCalcul")) Then
         m_PONumParamCalcul = 0
      Else
         m_PONumParamCalcul = theData.Fields("PONumParamCalcul")
      End If

      If IsNull(theData.Fields("PONUMCLE")) Then
         m_PONUMCLE = ""
      Else
         m_PONUMCLE = theData.Fields("PONUMCLE")
      End If

      If IsNull(theData.Fields("POGARCLE")) Then
         m_POGARCLE = 0
      Else
         m_POGARCLE = theData.Fields("POGARCLE")
      End If

      If IsNull(theData.Fields("POCONVENTION")) Then
         m_POCONVENTION = ""
      Else
         m_POCONVENTION = theData.Fields("POCONVENTION")
      End If

      If IsNull(theData.Fields("PONOM")) Then
         m_PONOM = ""
      Else
         m_PONOM = theData.Fields("PONOM")
      End If

      If IsNull(theData.Fields("POSEXE")) Then
         m_POSEXE = ""
      Else
         m_POSEXE = theData.Fields("POSEXE")
      End If

      If IsNull(theData.Fields("POCSP")) Then
         m_POCSP = ""
      Else
         m_POCSP = theData.Fields("POCSP")
      End If

      If IsNull(theData.Fields("PONAIS")) Then
         m_PONAIS = 0
      Else
         m_PONAIS = theData.Fields("PONAIS")
      End If

      If IsNull(theData.Fields("POEFFET")) Then
         m_POEFFET = 0
      Else
         m_POEFFET = theData.Fields("POEFFET")
      End If

      If IsNull(theData.Fields("POTERME")) Then
         m_POTERME = 0
      Else
         m_POTERME = theData.Fields("POTERME")
      End If

      If IsNull(theData.Fields("POARRET")) Then
         m_POARRET = 0
      Else
         m_POARRET = theData.Fields("POARRET")
      End If

      If IsNull(theData.Fields("POREPRISE")) Then
         m_POREPRISE = 0
      Else
         m_POREPRISE = theData.Fields("POREPRISE")
      End If

      If IsNull(theData.Fields("POCAUSE")) Then
         m_POCAUSE = ""
      Else
         m_POCAUSE = theData.Fields("POCAUSE")
      End If

      If IsNull(theData.Fields("POTYPEF")) Then
         m_POTYPEF = 0
      Else
         m_POTYPEF = theData.Fields("POTYPEF")
      End If

      If IsNull(theData.Fields("PODELAI")) Then
         m_PODELAI = 0
      Else
         m_PODELAI = theData.Fields("PODELAI")
      End If

      If IsNull(theData.Fields("POPREMIER_PAIEMENT")) Then
         m_POPREMIER_PAIEMENT = 0
      Else
         m_POPREMIER_PAIEMENT = theData.Fields("POPREMIER_PAIEMENT")
      End If

      If IsNull(theData.Fields("POPRESTATION")) Then
         m_POPRESTATION = 0
      Else
         m_POPRESTATION = theData.Fields("POPRESTATION")
      End If

      If IsNull(theData.Fields("POPRESTATIONTOTAL")) Then
         m_POPRESTATIONTOTAL = 0
      Else
         m_POPRESTATIONTOTAL = theData.Fields("POPRESTATIONTOTAL")
      End If

      If IsNull(theData.Fields("POPRESTATION_AN_PREC")) Then
         m_POPRESTATION_AN_PREC = 0
      Else
         m_POPRESTATION_AN_PREC = theData.Fields("POPRESTATION_AN_PREC")
      End If

      If IsNull(theData.Fields("POPRESTATION_AN")) Then
         m_POPRESTATION_AN = 0
      Else
         m_POPRESTATION_AN = theData.Fields("POPRESTATION_AN")
      End If

      If IsNull(theData.Fields("POSIT")) Then
         m_POSIT = 0
      Else
         m_POSIT = theData.Fields("POSIT")
      End If

      If IsNull(theData.Fields("POPM_INCAP_1F")) Then
         m_POPM_INCAP_1F = 0
      Else
         m_POPM_INCAP_1F = theData.Fields("POPM_INCAP_1F")
      End If

      If IsNull(theData.Fields("POPM_PASS_1F")) Then
         m_POPM_PASS_1F = 0
      Else
         m_POPM_PASS_1F = theData.Fields("POPM_PASS_1F")
      End If

      If IsNull(theData.Fields("POPM_INVAL_1F")) Then
         m_POPM_INVAL_1F = 0
      Else
         m_POPM_INVAL_1F = theData.Fields("POPM_INVAL_1F")
      End If

      If IsNull(theData.Fields("POPM")) Then
         m_POPM = 0
      Else
         m_POPM = theData.Fields("POPM")
      End If

      If IsNull(theData.Fields("POPM_SORTIE")) Then
         m_POPM_SORTIE = 0
      Else
         m_POPM_SORTIE = theData.Fields("POPM_SORTIE")
      End If

      If IsNull(theData.Fields("POPM_VAR")) Then
         m_POPM_VAR = 0
      Else
         m_POPM_VAR = theData.Fields("POPM_VAR")
      End If

      If IsNull(theData.Fields("PODERNIERPAIEMENT")) Then
         m_PODERNIERPAIEMENT = 0
      Else
         m_PODERNIERPAIEMENT = theData.Fields("PODERNIERPAIEMENT")
      End If

      If IsNull(theData.Fields("PODEBUT")) Then
         m_PODEBUT = 0
      Else
         m_PODEBUT = theData.Fields("PODEBUT")
      End If

      If IsNull(theData.Fields("POFIN")) Then
         m_POFIN = 0
      Else
         m_POFIN = theData.Fields("POFIN")
      End If

      If IsNull(theData.Fields("PODEBUTTOTAL")) Then
         m_PODEBUTTOTAL = 0
      Else
         m_PODEBUTTOTAL = theData.Fields("PODEBUTTOTAL")
      End If

      If IsNull(theData.Fields("POFINTOTAL")) Then
         m_POFINTOTAL = 0
      Else
         m_POFINTOTAL = theData.Fields("POFINTOTAL")
      End If

      If IsNull(theData.Fields("POPSAP")) Then
         m_POPSAP = 0
      Else
         m_POPSAP = theData.Fields("POPSAP")
      End If

      If IsNull(theData.Fields("POPM_RI")) Then
         m_POPM_RI = 0
      Else
         m_POPM_RI = theData.Fields("POPM_RI")
      End If

      If IsNull(theData.Fields("POPM_REVALO")) Then
         m_POPM_REVALO = 0
      Else
         m_POPM_REVALO = theData.Fields("POPM_REVALO")
      End If

      If IsNull(theData.Fields("POCOT_REVALO")) Then
         m_POCOT_REVALO = 0
      Else
         m_POCOT_REVALO = theData.Fields("POCOT_REVALO")
      End If

      If IsNull(theData.Fields("PODATEIMPORT")) Then
         m_PODATEIMPORT = 0
      Else
         m_PODATEIMPORT = theData.Fields("PODATEIMPORT")
      End If

      If IsNull(theData.Fields("POPM_INCAP_1R")) Then
         m_POPM_INCAP_1R = 0
      Else
         m_POPM_INCAP_1R = theData.Fields("POPM_INCAP_1R")
      End If

      If IsNull(theData.Fields("POPM_PASS_1R")) Then
         m_POPM_PASS_1R = 0
      Else
         m_POPM_PASS_1R = theData.Fields("POPM_PASS_1R")
      End If

      If IsNull(theData.Fields("POPM_INVAL_1R")) Then
         m_POPM_INVAL_1R = 0
      Else
         m_POPM_INVAL_1R = theData.Fields("POPM_INVAL_1R")
      End If

      If IsNull(theData.Fields("POCATEGORIE")) Then
         m_POCATEGORIE = ""
      Else
         m_POCATEGORIE = theData.Fields("POCATEGORIE")
      End If

      If IsNull(theData.Fields("POPM_RCJT_1F")) Then
         m_POPM_RCJT_1F = 0
      Else
         m_POPM_RCJT_1F = theData.Fields("POPM_RCJT_1F")
      End If

      If IsNull(theData.Fields("POPM_REDUC_1F")) Then
         m_POPM_REDUC_1F = 0
      Else
         m_POPM_REDUC_1F = theData.Fields("POPM_REDUC_1F")
      End If

      If IsNull(theData.Fields("POPRESTA_RASSUR")) Then
         m_POPRESTA_RASSUR = 0
      Else
         m_POPRESTA_RASSUR = theData.Fields("POPRESTA_RASSUR")
      End If

      If IsNull(theData.Fields("POPM_RASSUR")) Then
         m_POPM_RASSUR = 0
      Else
         m_POPM_RASSUR = theData.Fields("POPM_RASSUR")
      End If

      If IsNull(theData.Fields("POPSAP_RASSUR")) Then
         m_POPSAP_RASSUR = 0
      Else
         m_POPSAP_RASSUR = theData.Fields("POPSAP_RASSUR")
      End If

      If IsNull(theData.Fields("POTRAITE_RASSUR")) Then
         m_POTRAITE_RASSUR = ""
      Else
         m_POTRAITE_RASSUR = theData.Fields("POTRAITE_RASSUR")
      End If

      If IsNull(theData.Fields("POTRAITE_RASSUR2")) Then
         m_POTRAITE_RASSUR2 = ""
      Else
         m_POTRAITE_RASSUR2 = theData.Fields("POTRAITE_RASSUR2")
      End If

      If IsNull(theData.Fields("POTRAITE_RASSUR_TAUX")) Then
         m_POTRAITE_RASSUR_TAUX = 0
      Else
         m_POTRAITE_RASSUR_TAUX = theData.Fields("POTRAITE_RASSUR_TAUX")
      End If

      If IsNull(theData.Fields("PODATEPAIEMENTESTIMEE")) Then
         m_PODATEPAIEMENTESTIMEE = False
      Else
         m_PODATEPAIEMENTESTIMEE = theData.Fields("PODATEPAIEMENTESTIMEE")
      End If

      If IsNull(theData.Fields("POPSAPCAPMOYEN")) Then
         m_POPSAPCAPMOYEN = False
      Else
         m_POPSAPCAPMOYEN = theData.Fields("POPSAPCAPMOYEN")
      End If

      If IsNull(theData.Fields("PODOSSIERCLOS")) Then
         m_PODOSSIERCLOS = False
      Else
         m_PODOSSIERCLOS = theData.Fields("PODOSSIERCLOS")
      End If

      If IsNull(theData.Fields("POGARCLE_NEW")) Then
         m_POGARCLE_NEW = ""
      Else
         m_POGARCLE_NEW = theData.Fields("POGARCLE_NEW")
      End If

      If IsNull(theData.Fields("POCATEGORIE_NEW")) Then
         m_POCATEGORIE_NEW = ""
      Else
         m_POCATEGORIE_NEW = theData.Fields("POCATEGORIE_NEW")
      End If

      If IsNull(theData.Fields("POREGROUPEMENT")) Then
         m_POREGROUPEMENT = ""
      Else
         m_POREGROUPEMENT = theData.Fields("POREGROUPEMENT")
      End If

      If IsNull(theData.Fields("POCODENATURE")) Then
         m_POCODENATURE = ""
      Else
         m_POCODENATURE = theData.Fields("POCODENATURE")
      End If

      If IsNull(theData.Fields("POCCN")) Then
         m_POCCN = 0
      Else
         m_POCCN = theData.Fields("POCCN")
      End If

      If IsNull(theData.Fields("POCODERISQUE")) Then
         m_POCODERISQUE = 0
      Else
         m_POCODERISQUE = theData.Fields("POCODERISQUE")
      End If

      If IsNull(theData.Fields("POPMANNULEE")) Then
         m_POPMANNULEE = 0
      Else
         m_POPMANNULEE = theData.Fields("POPMANNULEE")
      End If

      If IsNull(theData.Fields("POPSAPANNULEE")) Then
         m_POPSAPANNULEE = 0
      Else
         m_POPSAPANNULEE = theData.Fields("POPSAPANNULEE")
      End If

      If IsNull(theData.Fields("POCOMMENTANNUL")) Then
         m_POCOMMENTANNUL = ""
      Else
         m_POCOMMENTANNUL = theData.Fields("POCOMMENTANNUL")
      End If

      If IsNull(theData.Fields("PODATEENTREEINVAL")) Then
         m_PODATEENTREEINVAL = 0
      Else
         m_PODATEENTREEINVAL = theData.Fields("PODATEENTREEINVAL")
      End If

      If IsNull(theData.Fields("POPM_RCJT_1R")) Then
         m_POPM_RCJT_1R = 0
      Else
         m_POPM_RCJT_1R = theData.Fields("POPM_RCJT_1R")
      End If

      If IsNull(theData.Fields("POPM_REDUC_1R")) Then
         m_POPM_REDUC_1R = 0
      Else
         m_POPM_REDUC_1R = theData.Fields("POPM_REDUC_1R")
      End If

      If IsNull(theData.Fields("POIsCadre")) Then
         m_POIsCadre = False
      Else
         m_POIsCadre = theData.Fields("POIsCadre")
      End If

      If IsNull(theData.Fields("POSalaireAnnuel")) Then
         m_POSalaireAnnuel = 0
      Else
         m_POSalaireAnnuel = theData.Fields("POSalaireAnnuel")
      End If

      If IsNull(theData.Fields("POCleSituationFamille")) Then
         m_POCleSituationFamille = 0
      Else
         m_POCleSituationFamille = theData.Fields("POCleSituationFamille")
      End If

      If IsNull(theData.Fields("POTauxGarantieDC")) Then
         m_POTauxGarantieDC = 0
      Else
         m_POTauxGarantieDC = theData.Fields("POTauxGarantieDC")
      End If

      If IsNull(theData.Fields("PONbEnfant")) Then
         m_PONbEnfant = 0
      Else
         m_PONbEnfant = theData.Fields("PONbEnfant")
      End If

      If IsNull(theData.Fields("POAgeMoyenEnfant")) Then
         m_POAgeMoyenEnfant = 0
      Else
         m_POAgeMoyenEnfant = theData.Fields("POAgeMoyenEnfant")
      End If

      If IsNull(theData.Fields("POMajoEnfant")) Then
         m_POMajoEnfant = 0
      Else
         m_POMajoEnfant = theData.Fields("POMajoEnfant")
      End If

      If IsNull(theData.Fields("PORegimeDeces")) Then
         m_PORegimeDeces = 0
      Else
         m_PORegimeDeces = theData.Fields("PORegimeDeces")
      End If

      If IsNull(theData.Fields("POCategorieDeces")) Then
         m_POCategorieDeces = 0
      Else
         m_POCategorieDeces = theData.Fields("POCategorieDeces")
      End If

      If IsNull(theData.Fields("PORegimeRenteConjointTempo")) Then
         m_PORegimeRenteConjointTempo = 0
      Else
         m_PORegimeRenteConjointTempo = theData.Fields("PORegimeRenteConjointTempo")
      End If

      If IsNull(theData.Fields("POCategorieRenteConjointTempo")) Then
         m_POCategorieRenteConjointTempo = 0
      Else
         m_POCategorieRenteConjointTempo = theData.Fields("POCategorieRenteConjointTempo")
      End If

      If IsNull(theData.Fields("PORegimeRenteConjointViagere")) Then
         m_PORegimeRenteConjointViagere = 0
      Else
         m_PORegimeRenteConjointViagere = theData.Fields("PORegimeRenteConjointViagere")
      End If

      If IsNull(theData.Fields("POCategorieRenteConjointViagere")) Then
         m_POCategorieRenteConjointViagere = 0
      Else
         m_POCategorieRenteConjointViagere = theData.Fields("POCategorieRenteConjointViagere")
      End If

      If IsNull(theData.Fields("PORegimeRenteEduc")) Then
         m_PORegimeRenteEduc = 0
      Else
         m_PORegimeRenteEduc = theData.Fields("PORegimeRenteEduc")
      End If

      If IsNull(theData.Fields("POCategorieRenteEduc")) Then
         m_POCategorieRenteEduc = 0
      Else
         m_POCategorieRenteEduc = theData.Fields("POCategorieRenteEduc")
      End If

      If IsNull(theData.Fields("POPourcentLissage")) Then
         m_POPourcentLissage = 0
      Else
         m_POPourcentLissage = theData.Fields("POPourcentLissage")
      End If

      If IsNull(theData.Fields("POCoeffBCAC")) Then
         m_POCoeffBCAC = 0
      Else
         m_POCoeffBCAC = theData.Fields("POCoeffBCAC")
      End If

      If IsNull(theData.Fields("POContractant")) Then
         m_POContractant = ""
      Else
         m_POContractant = theData.Fields("POContractant")
      End If

      If IsNull(theData.Fields("POSituConv")) Then
         m_POSituConv = ""
      Else
         m_POSituConv = theData.Fields("POSituConv")
      End If

      If IsNull(theData.Fields("POEffetSitu")) Then
         m_POEffetSitu = 0
      Else
         m_POEffetSitu = theData.Fields("POEffetSitu")
      End If

      If IsNull(theData.Fields("POCode_Option_Contrat")) Then
         m_POCode_Option_Contrat = ""
      Else
         m_POCode_Option_Contrat = theData.Fields("POCode_Option_Contrat")
      End If

      If IsNull(theData.Fields("POEtablissement")) Then
         m_POEtablissement = ""
      Else
         m_POEtablissement = theData.Fields("POEtablissement")
      End If

      If IsNull(theData.Fields("PONumSinistre")) Then
         m_PONumSinistre = 0
      Else
         m_PONumSinistre = theData.Fields("PONumSinistre")
      End If

      If IsNull(theData.Fields("POTypeMvt")) Then
         m_POTypeMvt = ""
      Else
         m_POTypeMvt = theData.Fields("POTypeMvt")
      End If

      If IsNull(theData.Fields("POIJ")) Then
         m_POIJ = 0
      Else
         m_POIJ = theData.Fields("POIJ")
      End If

      If IsNull(theData.Fields("POMontantBase")) Then
         m_POMontantBase = 0
      Else
         m_POMontantBase = theData.Fields("POMontantBase")
      End If

      If IsNull(theData.Fields("POMontantRevalo")) Then
         m_POMontantRevalo = 0
      Else
         m_POMontantRevalo = theData.Fields("POMontantRevalo")
      End If

      If IsNull(theData.Fields("POTypeReglement")) Then
         m_POTypeReglement = ""
      Else
         m_POTypeReglement = theData.Fields("POTypeReglement")
      End If

      If IsNull(theData.Fields("PODebutIndemnisation")) Then
         m_PODebutIndemnisation = 0
      Else
         m_PODebutIndemnisation = theData.Fields("PODebutIndemnisation")
      End If

      If IsNull(theData.Fields("PONbJourIndemn")) Then
         m_PONbJourIndemn = 0
      Else
         m_PONbJourIndemn = theData.Fields("PONbJourIndemn")
      End If

      If IsNull(theData.Fields("POCreationDossier")) Then
         m_POCreationDossier = 0
      Else
         m_POCreationDossier = theData.Fields("POCreationDossier")
      End If

      If IsNull(theData.Fields("PODebutDossier")) Then
         m_PODebutDossier = 0
      Else
         m_PODebutDossier = theData.Fields("PODebutDossier")
      End If

      If IsNull(theData.Fields("POFinDossier")) Then
         m_POFinDossier = 0
      Else
         m_POFinDossier = theData.Fields("POFinDossier")
      End If

      If IsNull(theData.Fields("POMotifCloture")) Then
         m_POMotifCloture = ""
      Else
         m_POMotifCloture = theData.Fields("POMotifCloture")
      End If

      If IsNull(theData.Fields("PODebutRefSalaire")) Then
         m_PODebutRefSalaire = 0
      Else
         m_PODebutRefSalaire = theData.Fields("PODebutRefSalaire")
      End If

      If IsNull(theData.Fields("POFinRefSalaire")) Then
         m_POFinRefSalaire = 0
      Else
         m_POFinRefSalaire = theData.Fields("POFinRefSalaire")
      End If

      If IsNull(theData.Fields("POFamilleComptable")) Then
         m_POFamilleComptable = 0
      Else
         m_POFamilleComptable = theData.Fields("POFamilleComptable")
      End If

      If IsNull(theData.Fields("POInspecteur")) Then
         m_POInspecteur = ""
      Else
         m_POInspecteur = theData.Fields("POInspecteur")
      End If

      If IsNull(theData.Fields("POInsp")) Then
         m_POInsp = ""
      Else
         m_POInsp = theData.Fields("POInsp")
      End If

      If IsNull(theData.Fields("POApport")) Then
         m_POApport = ""
      Else
         m_POApport = theData.Fields("POApport")
      End If

      If IsNull(theData.Fields("POApport2")) Then
         m_POApport2 = ""
      Else
         m_POApport2 = theData.Fields("POApport2")
      End If

      If IsNull(theData.Fields("POGestionnaire")) Then
         m_POGestionnaire = ""
      Else
         m_POGestionnaire = theData.Fields("POGestionnaire")
      End If

      If IsNull(theData.Fields("POIndicCC")) Then
         m_POIndicCC = ""
      Else
         m_POIndicCC = theData.Fields("POIndicCC")
      End If

      If IsNull(theData.Fields("POCompagnie")) Then
         m_POCompagnie = ""
      Else
         m_POCompagnie = theData.Fields("POCompagnie")
      End If

      If IsNull(theData.Fields("POAppli")) Then
         m_POAppli = 0
      Else
         m_POAppli = theData.Fields("POAppli")
      End If

      If IsNull(theData.Fields("PONbIntervenant")) Then
         m_PONbIntervenant = 0
      Else
         m_PONbIntervenant = theData.Fields("PONbIntervenant")
      End If

      If IsNull(theData.Fields("POPM_AvecCorrectif")) Then
         m_POPM_AvecCorrectif = 0
      Else
         m_POPM_AvecCorrectif = theData.Fields("POPM_AvecCorrectif")
      End If

      If IsNull(theData.Fields("POPM_SansCorrectif")) Then
         m_POPM_SansCorrectif = 0
      Else
         m_POPM_SansCorrectif = theData.Fields("POPM_SansCorrectif")
      End If

      If IsNull(theData.Fields("POCorrectif")) Then
         m_POCorrectif = 0
      Else
         m_POCorrectif = theData.Fields("POCorrectif")
      End If

      If IsNull(theData.Fields("POPMReassAvecCorrectif")) Then
         m_POPMReassAvecCorrectif = False
      Else
         m_POPMReassAvecCorrectif = theData.Fields("POPMReassAvecCorrectif")
      End If

      If IsNull(theData.Fields("POPMAvecCorrectif")) Then
         m_POPMAvecCorrectif = False
      Else
         m_POPMAvecCorrectif = theData.Fields("POPMAvecCorrectif")
      End If

      If IsNull(theData.Fields("POPM_X")) Then
         m_POPM_X = 0
      Else
         m_POPM_X = theData.Fields("POPM_X")
      End If

      If IsNull(theData.Fields("POPM_XTERME")) Then
         m_POPM_XTERME = 0
      Else
         m_POPM_XTERME = theData.Fields("POPM_XTERME")
      End If

      If IsNull(theData.Fields("POPM_ANC")) Then
         m_POPM_ANC = 0
      Else
         m_POPM_ANC = theData.Fields("POPM_ANC")
      End If

      If IsNull(theData.Fields("POPM_DUREE")) Then
         m_POPM_DUREE = 0
      Else
         m_POPM_DUREE = theData.Fields("POPM_DUREE")
      End If

      If IsNull(theData.Fields("POCaptive")) Then
         m_POCaptive = False
      Else
         m_POCaptive = theData.Fields("POCaptive")
      End If

      If IsNull(theData.Fields("POBaseRevalo")) Then
         m_POBaseRevalo = 0
      Else
         m_POBaseRevalo = theData.Fields("POBaseRevalo")
      End If

      If IsNull(theData.Fields("POFRACT")) Then
         m_POFRACT = 0
      Else
         m_POFRACT = theData.Fields("POFRACT")
      End If

      If IsNull(theData.Fields("POECHU")) Then
         m_POECHU = 0
      Else
         m_POECHU = theData.Fields("POECHU")
      End If

      If IsNull(theData.Fields("POMontantMajore")) Then
         m_POMontantMajore = 0
      Else
         m_POMontantMajore = theData.Fields("POMontantMajore")
      End If

      If IsNull(theData.Fields("NUENRP3I")) Then
         m_NUENRP3I = 0
      Else
         m_NUENRP3I = theData.Fields("NUENRP3I")
      End If

      If IsNull(theData.Fields("TXTECHN")) Then
         m_TXTECHN = 0
      Else
         m_TXTECHN = theData.Fields("TXTECHN")
      End If

      If IsNull(theData.Fields("TXFRAIS")) Then
         m_TXFRAIS = 0
      Else
         m_TXFRAIS = theData.Fields("TXFRAIS")
      End If

      
      
      ' Evol 2010 - Lot 2
      If IsNull(theData.Fields("POCDCONTENTIEUX")) Then
         m_POCDCONTENTIEUX = ""
      Else
         m_POCDCONTENTIEUX = theData.Fields("POCDCONTENTIEUX")
      End If
      
      If IsNull(theData.Fields("PONUSINISTRE")) Then
         m_PONUSINISTRE = ""
      Else
         m_PONUSINISTRE = theData.Fields("PONUSINISTRE")
      End If

      If IsNull(theData.Fields("POCategorieInval")) Then
         m_POCategorieInval = 0
      Else
         m_POCategorieInval = theData.Fields("POCategorieInval")
      End If
      
      If IsNull(theData.Fields("POMontantCapConstit")) Then
         m_POMontantCapConstit = 0
      Else
         m_POMontantCapConstit = theData.Fields("POMontantCapConstit")
      End If
      
      If IsNull(theData.Fields("POMontantCapSousRisque")) Then
         m_POMontantCapSousRisque = 0
      Else
         m_POMontantCapSousRisque = theData.Fields("POMontantCapSousRisque")
      End If
      
      ' Réforme des retraites
      If bForReformeRetraite Then
      
        If IsNull(theData.Fields("POIdAssure")) Then
           m_IdAssure = 0
        Else
           m_IdAssure = theData.Fields("POIdAssure")
        End If
      
        If IsNull(theData.Fields("POTopAmortissable")) Then
           m_TopAmortissable = 0
        Else
           m_TopAmortissable = theData.Fields("POTopAmortissable")
        End If
      
        If IsNull(theData.Fields("POCoeffAmortissement")) Then
           m_CoeffAmortissement = 0
        Else
           m_CoeffAmortissement = theData.Fields("POCoeffAmortissement")
        End If
      
        If IsNull(theData.Fields("POAgeMiniDepartRetraite")) Then
           m_AgeMiniDepartRetraite = 0
        Else
           m_AgeMiniDepartRetraite = theData.Fields("POAgeMiniDepartRetraite")
        End If
      
        If IsNull(theData.Fields("POCoeffAmortissement")) Then
           m_CoeffAmortissement = 0
        Else
           m_CoeffAmortissement = theData.Fields("POCoeffAmortissement")
        End If
      
      End If
      ' Réforme des retraites
      
      If IsNull(theData.Fields("Commentaire")) Then
         m_Commentaire = ""
      Else
         m_Commentaire = theData.Fields("Commentaire")
      End If
      

   Else
      ret = False
   End If

   Load = ret

   Exit Function

err_Load:
   MsgBox "Erreur dans clsAssure::Load()", vbCritical
   Load = False

End Function

'##ModelId=5C90E9720292
Public Function SaveToCSV()

Dim line As String
 
  'line = m_CDCOMPAGNIE & ";" & m_CDAPPLI & ";" & _
  'ine = line & m_IDASSUREAGI & ";" & _



  
  Print #1, line

End Function

'---------------------------------------------------
'##ModelId=5C90E97202B9
Public Function Save(DB As DataAccess, bForReformeRetraite As Boolean) As Boolean
  Dim sql As String, ret As Boolean
  ret = True
  
  On Error GoTo err_Save
  
  If (IsNull(m_RECNO) Or m_RECNO = 0) Then
  
    sql = "INSERT INTO " & IIf(bForReformeRetraite, "Assure_Retraite", "Assure") & "(POGPECLE, POPERCLE, POSTECLE, PONumParamCalcul, PONUMCLE, POGARCLE, POCONVENTION, PONOM, POSEXE, POCSP, PONAIS, POEFFET, POTERME, POARRET, POREPRISE, POCAUSE, POTYPEF, PODELAI, POPREMIER_PAIEMENT, POPRESTATION, POPRESTATIONTOTAL, POPRESTATION_AN_PREC, POPRESTATION_AN, " _
        & "POSIT, POPM_INCAP_1F, POPM_PASS_1F, POPM_INVAL_1F, POPM, POPM_SORTIE, POPM_VAR, PODERNIERPAIEMENT, PODEBUT, POFIN, PODEBUTTOTAL, POFINTOTAL, POPSAP, POPM_RI, POPM_REVALO, POCOT_REVALO, PODATEIMPORT, POPM_INCAP_1R, POPM_PASS_1R, POPM_INVAL_1R, POCATEGORIE, POPM_RCJT_1F, POPM_REDUC_1F, POPRESTA_RASSUR, " _
        & "POPM_RASSUR, POPSAP_RASSUR, POTRAITE_RASSUR, POTRAITE_RASSUR2, POTRAITE_RASSUR_TAUX, PODATEPAIEMENTESTIMEE, POPSAPCAPMOYEN, PODOSSIERCLOS, POGARCLE_NEW, POCATEGORIE_NEW, POREGROUPEMENT, POCODENATURE, POCCN, POCODERISQUE, POPMANNULEE, POPSAPANNULEE, POCOMMENTANNUL, PODATEENTREEINVAL, POPM_RCJT_1R, " _
        & "POPM_REDUC_1R, POIsCadre, POSalaireAnnuel, POCleSituationFamille, POTauxGarantieDC, PONbEnfant, POAgeMoyenEnfant, POMajoEnfant, PORegimeDeces, POCategorieDeces, PORegimeRenteConjointTempo, POCategorieRenteConjointTempo, PORegimeRenteConjointViagere, POCategorieRenteConjointViagere, PORegimeRenteEduc, " _
        & "POCategorieRenteEduc, POPourcentLissage, POCoeffBCAC, POContractant, POSituConv, POEffetSitu, POCode_Option_Contrat, POEtablissement, PONumSinistre, POTypeMvt, POIJ, POMontantBase, POMontantRevalo, POTypeReglement, PODebutIndemnisation, PONbJourIndemn, POCreationDossier, PODebutDossier, " _
        & "POFinDossier, POMotifCloture, PODebutRefSalaire, POFinRefSalaire, POFamilleComptable, POInspecteur, POInsp, POApport, POApport2, POGestionnaire, POIndicCC, POCompagnie, POAppli, PONbIntervenant, POPM_AvecCorrectif, POPM_SansCorrectif, POCorrectif, POPMReassAvecCorrectif, POPMAvecCorrectif, " _
        & "POPM_X, POPM_XTERME, POPM_ANC, POPM_DUREE, POCaptive, POBaseRevalo, POFRACT, POECHU, POMontantMajore, NUENRP3I, TXTECHN, TXFRAIS, " _
        & "POCDCONTENTIEUX, PONUSINISTRE, POCategorieInval, POMontantCapConstit, POMontantCapSousRisque, Commentaire "
    
    If bForReformeRetraite Then
      sql = sql & ", POIdAssure, POTopAmortissable, POCoeffAmortissement,POAgeMiniDepartRetraite,POAgeDepartRetraiteTauxPlein"
    End If
    
    sql = sql & ") VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?"
    
    If bForReformeRetraite Then
      sql = sql & ", ?, ?, ?, ?, ?"
    End If
    
    sql = sql & ")"
  
  'MsgBox "Erreur dans clsAssure::Save() : " & sql, vbCritical
  Else
    
    sql = "UPDATE " & IIf(bForReformeRetraite, "Assure_Retraite", "Assure") & " SET POGPECLE=?, POPERCLE=?, POSTECLE=?, PONumParamCalcul=?, PONUMCLE=?, POGARCLE=?, POCONVENTION=?, PONOM=?, POSEXE=?, POCSP=?, PONAIS=?, POEFFET=?, POTERME=?, POARRET=?, POREPRISE=?, POCAUSE=?, POTYPEF=?, PODELAI=?, POPREMIER_PAIEMENT=?, POPRESTATION=?, POPRESTATIONTOTAL=?, " _
        & "POPRESTATION_AN_PREC=?, POPRESTATION_AN=?, POSIT=?, POPM_INCAP_1F=?, POPM_PASS_1F=?, POPM_INVAL_1F=?, POPM=?, POPM_SORTIE=?, POPM_VAR=?, PODERNIERPAIEMENT=?, PODEBUT=?, POFIN=?, PODEBUTTOTAL=?, POFINTOTAL=?, POPSAP=?, POPM_RI=?, POPM_REVALO=?, POCOT_REVALO=?, PODATEIMPORT=?, POPM_INCAP_1R=?, " _
        & "POPM_PASS_1R=?, POPM_INVAL_1R=?, POCATEGORIE=?, POPM_RCJT_1F=?, POPM_REDUC_1F=?, POPRESTA_RASSUR=?, POPM_RASSUR=?, POPSAP_RASSUR=?, POTRAITE_RASSUR=?, POTRAITE_RASSUR2=?, POTRAITE_RASSUR_TAUX=?, PODATEPAIEMENTESTIMEE=?, POPSAPCAPMOYEN=?, PODOSSIERCLOS=?, POGARCLE_NEW=?, POCATEGORIE_NEW=?, " _
        & "POREGROUPEMENT=?, POCODENATURE=?, POCCN=?, POCODERISQUE=?, POPMANNULEE=?, POPSAPANNULEE=?, POCOMMENTANNUL=?, PODATEENTREEINVAL=?, POPM_RCJT_1R=?, POPM_REDUC_1R=?, POIsCadre=?, POSalaireAnnuel=?, POCleSituationFamille=?, POTauxGarantieDC=?, PONbEnfant=?, POAgeMoyenEnfant=?, POMajoEnfant=?, " _
        & "PORegimeDeces=?, POCategorieDeces=?, PORegimeRenteConjointTempo=?, POCategorieRenteConjointTempo=?, PORegimeRenteConjointViagere=?, POCategorieRenteConjointViagere=?, PORegimeRenteEduc=?, POCategorieRenteEduc=?, POPourcentLissage=?, POCoeffBCAC=?, POContractant=?, POSituConv=?, POEffetSitu=?, " _
        & "POCode_Option_Contrat=?, POEtablissement=?, PONumSinistre=?, POTypeMvt=?, POIJ=?, POMontantBase=?, POMontantRevalo=?, POTypeReglement=?, PODebutIndemnisation=?, PONbJourIndemn=?, POCreationDossier=?, PODebutDossier=?, POFinDossier=?, POMotifCloture=?, PODebutRefSalaire=?, POFinRefSalaire=?, " _
        & "POFamilleComptable=?, POInspecteur=?, POInsp=?, POApport=?, POApport2=?, POGestionnaire=?, POIndicCC=?, POCompagnie=?, POAppli=?, PONbIntervenant=?, POPM_AvecCorrectif=?, POPM_SansCorrectif=?, POCorrectif=?, POPMReassAvecCorrectif=?, POPMAvecCorrectif=?, POPM_X=?, POPM_XTERME=?, POPM_ANC=?, " _
        & "POPM_DUREE=?, POCaptive=?, POBaseRevalo=?, POFRACT=?, POECHU=?, POMontantMajore=?, NUENRP3I=?, TXTECHN=?, TXFRAIS=?, " _
        & "POCDCONTENTIEUX=?,  PONUSINISTRE=?,  POCategorieInval=?,  POMontantCapConstit=?,  POMontantCapSousRisque=?, Commentaire=? "
    
    If bForReformeRetraite Then
      sql = sql & ", POIdAssure=?, POTopAmortissable=?, POCoeffAmortissement=?,POAgeMiniDepartRetraite=?,POAgeDepartRetraiteTauxPlein=?"
    End If
        
    sql = sql & " WHERE  RECNO=" & m_RECNO
    
  End If
  
  Dim theCommand As New ADODB.Command, prm As ADODB.Parameter
  theCommand.ActiveConnection = DB.Connection
  theCommand.CommandType = adCmdText
  theCommand.CommandText = sql
  
  Set prm = theCommand.CreateParameter("@POGPECLE", adInteger, adParamInput, 0, m_POGPECLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPERCLE", adInteger, adParamInput, 0, m_POPERCLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POSTECLE", adInteger, adParamInput, 0, m_POSTECLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONumParamCalcul", adInteger, adParamInput, 0, m_PONumParamCalcul)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONUMCLE", adVarChar, adParamInput, IIf(IsNull(m_PONUMCLE), 1, Len(m_PONUMCLE)), m_PONUMCLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POGARCLE", adInteger, adParamInput, 0, m_POGARCLE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCONVENTION", adVarChar, adParamInput, IIf(IsNull(m_POCONVENTION), 1, Len(m_POCONVENTION)), m_POCONVENTION)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONOM", adVarChar, adParamInput, IIf(IsNull(m_PONOM) Or m_PONOM = "", 1, Len(m_PONOM)), m_PONOM)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POSEXE", adVarChar, adParamInput, IIf(IsNull(m_POSEXE), 1, Len(m_POSEXE)), m_POSEXE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCSP", adVarChar, adParamInput, IIf(IsNull(m_POCSP) Or m_POCSP = "", 1, Len(m_POCSP)), m_POCSP)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONAIS", adDate, adParamInput, 0, IIf(m_PONAIS = 0, Null, m_PONAIS))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POEFFET", adDate, adParamInput, 0, IIf(m_POEFFET = 0 Or m_POEFFET = "", Null, m_POEFFET))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTERME", adDate, adParamInput, 0, IIf(m_POTERME = 0 Or m_POTERME = "", Null, m_POTERME))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POARRET", adDate, adParamInput, 0, IIf(m_POARRET = 0, Null, m_POARRET))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POREPRISE", adDate, adParamInput, 0, IIf(m_POREPRISE = 0, Null, m_POREPRISE))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCAUSE", adVarChar, adParamInput, IIf(IsNull(m_POCAUSE) Or m_POCAUSE = "", 1, Len(m_POCAUSE)), m_POCAUSE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTYPEF", adInteger, adParamInput, 0, m_POTYPEF)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODELAI", adInteger, adParamInput, 0, m_PODELAI)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPREMIER_PAIEMENT", adDate, adParamInput, 0, IIf(m_POPREMIER_PAIEMENT = 0, Null, m_POPREMIER_PAIEMENT))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPRESTATION", adDouble, adParamInput, 0, m_POPRESTATION)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPRESTATIONTOTAL", adDouble, adParamInput, 0, m_POPRESTATIONTOTAL)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPRESTATION_AN_PREC", adDouble, adParamInput, 0, m_POPRESTATION_AN_PREC)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPRESTATION_AN", adDouble, adParamInput, 0, m_POPRESTATION_AN)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POSIT", adInteger, adParamInput, 0, m_POSIT)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_INCAP_1F", adDouble, adParamInput, 0, m_POPM_INCAP_1F)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_PASS_1F", adDouble, adParamInput, 0, m_POPM_PASS_1F)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_INVAL_1F", adDouble, adParamInput, 0, m_POPM_INVAL_1F)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM", adDouble, adParamInput, 0, m_POPM)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_SORTIE", adDouble, adParamInput, 0, m_POPM_SORTIE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_VAR", adDouble, adParamInput, 0, m_POPM_VAR)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODERNIERPAIEMENT", adDate, adParamInput, 0, IIf(m_PODERNIERPAIEMENT = 0, Null, m_PODERNIERPAIEMENT))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODEBUT", adDate, adParamInput, 0, IIf(m_PODEBUT = 0, Null, m_PODEBUT))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFIN", adDate, adParamInput, 0, IIf(m_POFIN = 0, Null, m_POFIN))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODEBUTTOTAL", adDate, adParamInput, 0, IIf(m_PODEBUTTOTAL = 0, Null, m_PODEBUTTOTAL))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFINTOTAL", adDate, adParamInput, 0, IIf(m_POFINTOTAL = 0, Null, m_POFINTOTAL))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPSAP", adDouble, adParamInput, 0, m_POPSAP)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_RI", adDouble, adParamInput, 0, m_POPM_RI)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_REVALO", adDouble, adParamInput, 0, m_POPM_REVALO)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCOT_REVALO", adDouble, adParamInput, 0, m_POCOT_REVALO)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODATEIMPORT", adDate, adParamInput, 0, IIf(m_PODATEIMPORT = 0, Null, m_PODATEIMPORT))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_INCAP_1R", adDouble, adParamInput, 0, m_POPM_INCAP_1R)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_PASS_1R", adDouble, adParamInput, 0, m_POPM_PASS_1R)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_INVAL_1R", adDouble, adParamInput, 0, m_POPM_INVAL_1R)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCATEGORIE", adVarChar, adParamInput, IIf(IsNull(m_POCATEGORIE), 1, Len(m_POCATEGORIE)), m_POCATEGORIE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_RCJT_1F", adDouble, adParamInput, 0, m_POPM_RCJT_1F)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_REDUC_1F", adDouble, adParamInput, 0, m_POPM_REDUC_1F)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPRESTA_RASSUR", adDouble, adParamInput, 0, m_POPRESTA_RASSUR)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_RASSUR", adDouble, adParamInput, 0, m_POPM_RASSUR)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPSAP_RASSUR", adDouble, adParamInput, 0, m_POPSAP_RASSUR)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTRAITE_RASSUR", adVarChar, adParamInput, IIf(IsNull(m_POTRAITE_RASSUR), 1, Len(m_POTRAITE_RASSUR)), m_POTRAITE_RASSUR)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTRAITE_RASSUR2", adVarChar, adParamInput, IIf(IsNull(m_POTRAITE_RASSUR2), 1, Len(m_POTRAITE_RASSUR2)), m_POTRAITE_RASSUR2)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTRAITE_RASSUR_TAUX", adDouble, adParamInput, 0, m_POTRAITE_RASSUR_TAUX)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODATEPAIEMENTESTIMEE", adBoolean, adParamInput, 0, m_PODATEPAIEMENTESTIMEE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPSAPCAPMOYEN", adBoolean, adParamInput, 0, m_POPSAPCAPMOYEN)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODOSSIERCLOS", adBoolean, adParamInput, 0, m_PODOSSIERCLOS)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POGARCLE_NEW", adVarChar, adParamInput, IIf(IsNull(m_POGARCLE_NEW), 1, Len(m_POGARCLE_NEW)), m_POGARCLE_NEW)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCATEGORIE_NEW", adVarChar, adParamInput, IIf(IsNull(m_POCATEGORIE_NEW), 1, Len(m_POCATEGORIE_NEW)), m_POCATEGORIE_NEW)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POREGROUPEMENT", adVarChar, adParamInput, IIf(IsNull(m_POREGROUPEMENT), 1, Len(m_POREGROUPEMENT)), m_POREGROUPEMENT)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCODENATURE", adVarChar, adParamInput, IIf(IsNull(m_POCODENATURE), 1, Len(m_POCODENATURE)), m_POCODENATURE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCCN", adInteger, adParamInput, 0, m_POCCN)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCODERISQUE", adInteger, adParamInput, 0, m_POCODERISQUE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPMANNULEE", adDouble, adParamInput, 0, m_POPMANNULEE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPSAPANNULEE", adDouble, adParamInput, 0, m_POPSAPANNULEE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCOMMENTANNUL", adVarChar, adParamInput, IIf(IsNull(m_POCOMMENTANNUL), 1, Len(m_POCOMMENTANNUL)), m_POCOMMENTANNUL)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODATEENTREEINVAL", adDate, adParamInput, 0, IIf(m_PODATEENTREEINVAL = 0 Or m_PODATEENTREEINVAL = "", Null, m_PODATEENTREEINVAL))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_RCJT_1R", adDouble, adParamInput, 0, m_POPM_RCJT_1R)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_REDUC_1R", adDouble, adParamInput, 0, m_POPM_REDUC_1R)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POIsCadre", adBoolean, adParamInput, 0, m_POIsCadre)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POSalaireAnnuel", adDouble, adParamInput, 0, m_POSalaireAnnuel)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCleSituationFamille", adInteger, adParamInput, 0, m_POCleSituationFamille)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTauxGarantieDC", adDouble, adParamInput, 0, m_POTauxGarantieDC)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONbEnfant", adInteger, adParamInput, 0, m_PONbEnfant)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POAgeMoyenEnfant", adInteger, adParamInput, 0, m_POAgeMoyenEnfant)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMajoEnfant", adDouble, adParamInput, 0, m_POMajoEnfant)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PORegimeDeces", adInteger, adParamInput, 0, m_PORegimeDeces)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCategorieDeces", adInteger, adParamInput, 0, m_POCategorieDeces)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PORegimeRenteConjointTempo", adInteger, adParamInput, 0, m_PORegimeRenteConjointTempo)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCategorieRenteConjointTempo", adInteger, adParamInput, 0, m_POCategorieRenteConjointTempo)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PORegimeRenteConjointViagere", adInteger, adParamInput, 0, m_PORegimeRenteConjointViagere)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCategorieRenteConjointViagere", adInteger, adParamInput, 0, m_POCategorieRenteConjointViagere)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PORegimeRenteEduc", adInteger, adParamInput, 0, m_PORegimeRenteEduc)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCategorieRenteEduc", adInteger, adParamInput, 0, m_POCategorieRenteEduc)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPourcentLissage", adDouble, adParamInput, 0, m_POPourcentLissage)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCoeffBCAC", adDouble, adParamInput, 0, m_POCoeffBCAC)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POContractant", adVarChar, adParamInput, IIf(IsNull(m_POContractant), 1, Len(m_POContractant)), m_POContractant)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POSituConv", adVarChar, adParamInput, IIf(IsNull(m_POSituConv), 1, Len(m_POSituConv)), m_POSituConv)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POEffetSitu", adDate, adParamInput, 0, IIf(m_POEffetSitu = 0, Null, m_POEffetSitu))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCode_Option_Contrat", adVarChar, adParamInput, IIf(IsNull(m_POCode_Option_Contrat), 1, Len(m_POCode_Option_Contrat)), m_POCode_Option_Contrat)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POEtablissement", adVarChar, adParamInput, IIf(IsNull(m_POEtablissement), 1, Len(m_POEtablissement)), m_POEtablissement)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONumSinistre", adInteger, adParamInput, 0, m_PONumSinistre)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTypeMvt", adVarChar, adParamInput, IIf(IsNull(m_POTypeMvt), 1, Len(m_POTypeMvt)), m_POTypeMvt)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POIJ", adDouble, adParamInput, 0, m_POIJ)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMontantBase", adDouble, adParamInput, 0, m_POMontantBase)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMontantRevalo", adDouble, adParamInput, 0, m_POMontantRevalo)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POTypeReglement", adVarChar, adParamInput, IIf(IsNull(m_POTypeReglement), 1, Len(m_POTypeReglement)), m_POTypeReglement)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODebutIndemnisation", adDate, adParamInput, 0, IIf(m_PODebutIndemnisation = 0, Null, m_PODebutIndemnisation))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONbJourIndemn", adInteger, adParamInput, 0, m_PONbJourIndemn)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCreationDossier", adDate, adParamInput, 0, IIf(m_POCreationDossier = 0, Null, m_POCreationDossier))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODebutDossier", adDate, adParamInput, 0, IIf(m_PODebutDossier = 0, Null, m_PODebutDossier))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFinDossier", adDate, adParamInput, 0, IIf(m_POFinDossier = 0, Null, m_POFinDossier))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMotifCloture", adVarChar, adParamInput, IIf(IsNull(m_POMotifCloture), 1, Len(m_POMotifCloture)), m_POMotifCloture)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PODebutRefSalaire", adDate, adParamInput, 0, IIf(m_PODebutRefSalaire = 0, Null, m_PODebutRefSalaire))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFinRefSalaire", adDate, adParamInput, 0, IIf(m_POFinRefSalaire = 0, Null, m_POFinRefSalaire))
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFamilleComptable", adInteger, adParamInput, 0, m_POFamilleComptable)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POInspecteur", adVarChar, adParamInput, IIf(IsNull(m_POInspecteur), 1, Len(m_POInspecteur)), m_POInspecteur)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POInsp", adVarChar, adParamInput, IIf(IsNull(m_POInsp), 1, Len(m_POInsp)), m_POInsp)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POApport", adVarChar, adParamInput, IIf(IsNull(m_POApport), 1, Len(m_POApport)), m_POApport)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POApport2", adVarChar, adParamInput, IIf(IsNull(m_POApport2), 1, Len(m_POApport2)), m_POApport2)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POGestionnaire", adVarChar, adParamInput, IIf(IsNull(m_POGestionnaire), 1, Len(m_POGestionnaire)), m_POGestionnaire)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POIndicCC", adVarChar, adParamInput, IIf(IsNull(m_POIndicCC), 1, Len(m_POIndicCC)), m_POIndicCC)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCompagnie", adVarChar, adParamInput, IIf(IsNull(m_POCompagnie), 1, Len(m_POCompagnie)), m_POCompagnie)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POAppli", adInteger, adParamInput, 0, m_POAppli)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONbIntervenant", adInteger, adParamInput, 0, m_PONbIntervenant)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_AvecCorrectif", adDouble, adParamInput, 0, m_POPM_AvecCorrectif)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_SansCorrectif", adDouble, adParamInput, 0, m_POPM_SansCorrectif)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCorrectif", adDouble, adParamInput, 0, m_POCorrectif)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPMReassAvecCorrectif", adBoolean, adParamInput, 0, m_POPMReassAvecCorrectif)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPMAvecCorrectif", adBoolean, adParamInput, 0, m_POPMAvecCorrectif)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_X", adInteger, adParamInput, 0, m_POPM_X)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_XTERME", adInteger, adParamInput, 0, m_POPM_XTERME)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_ANC", adInteger, adParamInput, 0, m_POPM_ANC)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POPM_DUREE", adInteger, adParamInput, 0, m_POPM_DUREE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCaptive", adBoolean, adParamInput, 0, m_POCaptive)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POBaseRevalo", adInteger, adParamInput, 0, m_POBaseRevalo)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POFRACT", adInteger, adParamInput, 0, m_POFRACT)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POECHU", adInteger, adParamInput, 0, m_POECHU)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMontantMajore", adDouble, adParamInput, 0, m_POMontantMajore)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@NUENRP3I", adInteger, adParamInput, 0, m_NUENRP3I)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@TXTECHN", adDouble, adParamInput, 0, m_TXTECHN)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@TXFRAIS", adDouble, adParamInput, 0, m_TXFRAIS)
  theCommand.Parameters.Append prm
  
  
  
  ' Evol 2010 - Lot 2
  Set prm = theCommand.CreateParameter("@POCDCONTENTIEUX", adChar, adParamInput, 1, m_POCDCONTENTIEUX)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@PONUSINISTRE", adVarChar, adParamInput, IIf(IsNull(m_PONUSINISTRE), 1, Len(m_PONUSINISTRE)), m_PONUSINISTRE)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POCategorieInval", adDouble, adParamInput, 0, m_POCategorieInval)
'  If m_POCATEGORIE_NEW > 4 Then
'    prm.Value = Null
'  End If
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMontantCapConstit", adDouble, adParamInput, 0, m_POMontantCapConstit)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@POMontantCapSousRisque", adDouble, adParamInput, 0, m_POMontantCapSousRisque)
  theCommand.Parameters.Append prm
  
  Set prm = theCommand.CreateParameter("@Commentaire", adVarChar, adParamInput, IIf(IsNull(m_Commentaire), 1, Len(m_Commentaire)), m_Commentaire)
  theCommand.Parameters.Append prm
  
  ' Champs pour la reforme des retraites
  If bForReformeRetraite Then
    
    Set prm = theCommand.CreateParameter("@POIdAssure", adInteger, adParamInput, 0, m_IdAssure)
    theCommand.Parameters.Append prm
    
    Set prm = theCommand.CreateParameter("@POTopAmortissable", adBoolean, adParamInput, 0, m_TopAmortissable)
    theCommand.Parameters.Append prm
  
    Set prm = theCommand.CreateParameter("@POCoeffAmortissement", adDouble, adParamInput, 0, m_CoeffAmortissement)
    theCommand.Parameters.Append prm
  
    Set prm = theCommand.CreateParameter("POAgeMiniDepartRetraite", adDouble, adParamInput, 0, m_AgeMiniDepartRetraite)
    theCommand.Parameters.Append prm
  
    Set prm = theCommand.CreateParameter("POAgeDepartRetraiteTauxPlein", adDouble, adParamInput, 0, m_AgeDepartRetraiteTauxPlein)
    theCommand.Parameters.Append prm
    
  End If
  
 
  theCommand.Execute
  
  If (IsNull(m_RECNO) Or m_RECNO = 0) Then
    Dim rs As ADODB.Recordset
    
    Set theCommand = Nothing
    Set theCommand = New ADODB.Command
    
    theCommand.ActiveConnection = DB.Connection
    theCommand.CommandType = adCmdText
    theCommand.CommandText = "SELECT @@IDENTITY"
    
    Set rs = theCommand.Execute
    If rs.EOF = False Then
      m_RECNO = rs.Fields(0)
    End If
    
    rs.Close
  
  End If
  Save = ret
  
  Exit Function
  
err_Save:
  MsgBox "Erreur dans clsAssure::Save() : " & Err & vbLf & Err.Description, vbCritical
  Save = False
  Exit Function
  
  Resume Next
End Function


'---------------------------------------------------
'##ModelId=5C90E972034A
Public Function Delete(DB As DataAccess) As Boolean
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Delete

   Dim theCommand As New ADODB.Command
   theCommand.ActiveConnection = DB.Connection
   theCommand.CommandType = adCmdText
   theCommand.CommandText = "DELETE FROM Assure WHERE  RECNO= " & m_RECNO

   theCommand.Execute

   m_RECNO = 0


   Delete = ret

   Exit Function

err_Delete:
   MsgBox "Erreur dans clsAssure::Delete()", vbCritical
   Delete = False

End Function


