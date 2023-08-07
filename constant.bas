Attribute VB_Name = "constant"
Option Explicit

' extended properties pour l'import de feuille excel
' Excel 8.0        : format Excel 2000 et +
' IMEX=1           : traite les colonnes contenant des types mixées comme du text (evite NULL)
' TypeGuessRows=50 : nb de ligne à scanner pour déterminer le type d'une colonne
Public Const cdExcelExtendedProperties As String = """Excel 8.0;IMEX=1;HDR=YES;TypeGuessRows=1000;ImportMixedTypes=Text"""
Public Const cdExcelExtendedPropertiesDAO As String = "Excel 8.0;IMEX=1;HDR=YES;TypeGuessRows=1000;ImportMixedTypes=Text"

Public Const cdExcelExtendedPropertiesXls = """Excel 8.0;IMEX=1;TypeGuessRows=1000"""
Public Const cdExcelExtendedPropertiesXlsx = """Excel 12.0 Xml;HDR=YES;"""

Public Const ConnectionStringXls = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                         & "Data Source=%1;" _
                         & "Extended Properties=" & cdExcelExtendedPropertiesXls & ";" _
                         & "Persist Security Info=False"
                         
Public Const ConnectionStringXlsx = " Provider=Microsoft.ACE.OLEDB.12.0;" _
                         & "Data Source=%1; " _
                         & "Extended Properties=" & cdExcelExtendedPropertiesXlsx & ";" _
                        & "Persist Security Info=False"

' type des Erreur
Public Const errCollectionAbsent As Double = 3265  ' élément absent dans la collection

' type de garanties
Public Const cdTypeIncapacite As Integer = 1
Public Const cdTypeInvalidite As Integer = 2
Public Const cdTypeMaintienDeces As Integer = 90

' type de tables
Public Const cdTypeTable_LoiMaintienIncapacite As Integer = 1
Public Const cdTypeTable_LoiPassage As Integer = 2
Public Const cdTypeTable_LoiMaintienInvalidite As Integer = 3

Public Const cdTypeTableMortalite As Integer = 4
Public Const cdTypeTableGeneration As Integer = 5

Public Const cdTypeTableCoeffBCACInval As Integer = 14
Public Const cdTypeTableCoeffBCACIncap As Integer = 15

Public Const cdTypeTable_MortaliteIncap As Integer = 21
Public Const cdTypeTable_MortaliteInval As Integer = 23

Public Const cdTypeTable_LoiDependance As Integer = 30

Public Const cdTypeTable_BaremeAnneeStatutaire As Integer = 40

' code societe
Public Const cdSoc1 As Integer = 1
Public Const cdSoc2 As Integer = 2
Public Const cdSoc3 As Integer = 3

' code garantie
Public Const cdGar11 As Integer = 11 ' incap
Public Const cdGar12 As Integer = 12
Public Const cdGar13 As Integer = 13
Public Const cdGar14 As Integer = 14 ' upese : incap 1 an
Public Const cdGar15 As Integer = 15 ' upese : incap 90 jours

Public Const cdGar21 As Integer = 21
Public Const cdGar22 As Integer = 22
Public Const cdGar23 As Integer = 23

Public Const cdGar31 As Integer = 31
Public Const cdGar32 As Integer = 32
Public Const cdGar33 As Integer = 33
' garantie invalidite indexee
Public Const cdGar41 As Integer = 41
Public Const cdGar42 As Integer = 42
Public Const cdGar43 As Integer = 43

' risque ISICA
Public Const cdGar51 As Integer = 51 ' rente de conjoint
Public Const cdGar510 As Integer = 60  ' rente de conjoint temporaire
Public Const cdGar511 As Integer = 61 ' rente de conjoint viagère
Public Const cdGar53 As Integer = 53 ' rente education
Public Const cdGar56 As Integer = 56 ' deces
Public Const cdGar57 As Integer = 57 ' inaptitude a la conduite
Public Const cdGar59 As Integer = 59 ' arret de travail
Public Const cdGar80 As Integer = 80 ' Dépendance (=50+30)

' risque RENTES
Public Const cdGarRente As Integer = 70                           ' regime pour toute les rentes (1 ligne=1 palier)

' Régimes pour l'import
Public Const cdGarDeces_Import As Integer = 6                     ' deces
Public Const cdGarInaptitudeConduite_Import As Integer = 7        ' inaptitude a la conduite
Public Const cdGarIncapInval_Import As Integer = 9                ' arret de travail
Public Const cdGarRente_Import As Integer = 20                    ' regime IMPORT pour toute les rentes (1 ligne=1 palier)
Public Const cdGarDependance_Import As Integer = 30                    ' regime IMPORT pour la dépendance

Public Const cdGarRenteConjointTemporaire As Integer = 60 ' rente de conjoint temporaire
Public Const cdGarRenteConjointViagère As Integer = 61 ' rente de conjoint viagère

Public Const cdOffsetGar_DC As Integer = 90 ' offset entre garantie et garantie DC

Public Const cdGarDeces_DC As Integer = 96 ' rente de conjoint temporaire
Public Const cdGarRenteEducation_DC As Integer = 93 ' rente éducation
Public Const cdGarRenteConjointTemporaire_DC As Integer = 100 ' rente de conjoint temporaire
Public Const cdGarRenteConjointViagère_DC As Integer = 101 ' rente de conjoint viagère



Public Const cdM0 As Integer = 2
Public Const cdM1 As Integer = 4
Public Const cdCVD0 As Integer = 1
Public Const cdCVD1 As Integer = 3

Public ret_code As Integer

' filtre
Public Const FILTER_VALUE_NULL As String = "(null)"

' Valeurs Code Provisions

'1001 - AvecIncapacité  AvecPassage  AvecInvalidité  Avec    Avec    Avec
'1002 - AvecIncapacité  AvecPassage  SansInvalidité  Avec    Avec    Sans
'1003 - AvecIncapacité  SansPassage  AvecInvalidité  Avec    Sans    Avec
'1004 - AvecIncapacité  SansPassage  SansInvalidité  Avec    Sans    Sans
'1005 - SansIncapacité  AvecPassage  AvecInvalidité  Sans    Avec    Avec
'1006 - SansIncapacité  AvecPassage  SansInvalidité  Sans    Avec    Sans
'1007 - SansIncapacité  SansPassage  AvecInvalidité  Sans    Sans    Avec
'1008 - SansIncapacité  SansPassage  SansInvalidité  Sans    Sans    Sans
            
Public Const cdProvision_AvecIncap_AvecPassage_AvecInval = 1001
Public Const cdProvision_AvecIncap_AvecPassage_SansInval = 1002
Public Const cdProvision_AvecIncap_SansPassage_AvecInval = 1003
Public Const cdProvision_AvecIncap_SansPassage_SansInval = 1004
Public Const cdProvision_SansIncap_AvecPassage_AvecInval = 1005
Public Const cdProvision_SansIncap_AvecPassage_SansInval = 1006
Public Const cdProvision_SansIncap_SansPassage_AvecInval = 1007
Public Const cdProvision_SansIncap_SansPassage_SansInval = 1008

Public Const cdProvision_Incap_AvecPassage_Viager = 6001
Public Const cdProvision_Inval_Viager = 6003

Public Const cdProvision_DecesIndividuel = 2000             ' décès individuel
Public Const cdProvision_dependanceIndividuel = 3000        ' dépendance
Public Const cdProvision_CapitalInvaliditeProbable = 4000   ' capital invalidité probable
Public Const cdProvision_Statutaire = 5000   ' risque statutaire
    
' Valeur du champ POSIT (type de calcul des provisions)
Public Const cdPosit_IncapAvecPassage As Integer = 1  ' incap avec passage
Public Const cdPosit_Inval As Integer = 2             ' inval
Public Const cdPosit_IncapSansPassage As Integer = 3  ' incap sans passage sans INVAL

'Public Const cdPosit_SansIncap_AvecPassage_SansInval As Integer = 96  ' Sans Incap Avec Passage Sans Inval
'Public Const cdPosit_SansIncap_AvecPassage_AvecInval As Integer = 97  ' Sans Incap Avec Passage Avec Inval
'Public Const cdPosit_AvecIncap_AvecPassage_SansInval As Integer = 98  ' Avec Incap Sans Passage Sans Inval
'Public Const cdPosit_AvecIncap_AvecPassage_AvecInval As Integer = 99  ' Avec Incap Sans Passage Avec Inval Possible
'Public Const cdPosit_AvecIncap_SansPassage_SansInval As Integer = 98  ' Avec Incap Sans Passage Sans Inval
'Public Const cdPosit_AvecIncap_SansPassage_AvecInval As Integer = 99  ' Avec Incap Sans Passage Avec Inval Possible
         
'Public Const cdPosit_RenteConjoint As Integer = 4     ' rente conjoint
'Public Const cdPosit_RenteEduc As Integer = 5         ' rente education

Public Const cdPosit_Deces As Integer = 6             ' décès
Public Const cdPosit_Maternite As Integer = 7         ' maternité
Public Const cdPosit_Mensualisation As Integer = 8    ' mensualisation
Public Const cdPosit_Chomage As Integer = 9           ' chômage
'Public Const cdPosit_Dependance As Integer = 30       ' rente dépendance
'Public Const cdPosit_CapitalInvaliditeProbable = 4000   ' capital invalidité probable


Public Const cdCategorieInvalParDefaut As Integer = 2


' Nouveaux codes POSITION pour les rentes, régime unique 20
Public Const cdPosit_RenteCertaine As Integer = 21                ' rente certaine (proba=100% via "fausse" table)
Public Const cdPosit_RenteEducationTemporaire As Integer = 22     ' rente education temporaire
Public Const cdPosit_RenteEducationViagere As Integer = 23        ' rente education viagere
Public Const cdPosit_RenteConjointTemporaire As Integer = 24      ' rente conjoint temporaire
Public Const cdPosit_RenteConjointViagere As Integer = 25         ' rente conjoint viagere
Public Const cdPosit_RenteRetraiteTemporaire As Integer = 26      ' rente retraite temporaire
Public Const cdPosit_RenteRetraiteViagere As Integer = 27         ' rente retraite viagere
Public Const cdPosit_RenteAutreTemporaire As Integer = 28         ' autre rente temporaire
Public Const cdPosit_RenteAutreViagere As Integer = 29            ' autre rente viagere

Public Const cdPosit_Statutaire As Integer = 5000            ' Statutaire


Public Const cdPositImport_RenteCertaine As Integer = 21                ' rente certaine (proba=100% via "fausse" table)
Public Const cdPositImport_RenteEducation As Integer = 22     ' rente education temporaire
Public Const cdPositImport_RenteConjoint As Integer = 24      ' rente conjoint temporaire
Public Const cdPositImport_RenteRetraite As Integer = 26      ' rente retraite temporaire
Public Const cdPositImport_RenteAutre As Integer = 28         ' autre rente temporaire
Public Const cdPositImport_RenteEducation_Handicape  As Integer = 29         ' rente éducation handicapé


Public Const cdPositImport_IncapProf As Integer = 1
Public Const cdPositImport_IncapNonProf As Integer = 2

'Public Const cdPositImport_SansIncap_AvecPassage_SansInval As Integer = 96  ' Sans Incap Avec Passage Sans Inval
'Public Const cdPositImport_SansIncap_AvecPassage_AvecInval As Integer = 97  ' Sans Incap Avec Passage Avec Inval
'Public Const cdPositImport_AvecIncap_SansPassage_SansInval As Integer = 98  ' Avec Incap Sans Passage Sans Inval
'Public Const cdPositImport_AvecIncap_SansPassage_AvecInval As Integer = 99  ' Avec Incap Sans Passage Avec Inval Possible
'Public Const cdPositImport_Incap_SansPassage_SansInval As Integer = 98
'Public Const cdPositImport_Incap_SansPassage_AvecInval As Integer = 99

Public Const cdPositImport_InvalProf As Integer = 3
Public Const cdPositImport_InvalNonProf As Integer = 4

Public Const cdPositImport_Maternite As Integer = 5
Public Const cdPositImport_Mensualisation As Integer = 6

Public Const cdPositImport_Chomage As Integer = 4
'Public Const cdPositImport_Dependance  As Integer = 30         ' rente dépendance

'Display - All Fields
Public Const cDispAllFieldsID As Integer = 9999
Public Const cDispAllFieldsName As String = "&ALLFIELDS&"
Public Const cDefaultDisplayName As String = "Tous-Les-Champs"
Public Const cAllUsersUserName As String = "ALL-USERS"

' valeur de l'interpolation de l'inval
Public Enum InterpolationInval
  eInterpolationInval_NON = 0
  eInterpolationInval_Age = 1
  eInterpolationInval_CorrectionDuree = 2
  eInterpolationInval_AgeDuree = 3
End Enum

' Fractionnement Paiement
Public Enum Fractionnement
  eFractionnementAnnuel = 1
  eFractionnementSemestriel = 2
  eFractionnementTrimestriel = 4
  eFractionnementMensuel = 12
End Enum

' Echeance Paiement
Public Enum EcheancePaiement
  ePaiementAvance = 1
  ePaiementEchu = 2
End Enum

' Méthode de calcul Provisions DC
Public Enum MethodeCalculMaintienDC
  eCapitauxConstitutifs = 1
  ePrimeExoneree = 2
  ePctProvisionCalculee = 3
  ePasDeCalcul = 99
End Enum

' version des données (Champ DataVersion)
Public Enum tagDataVersion
  eInitiale = 0
  eModifie = 1
  eSupprimer = 2
  eDoublon = 3
  eUndelete = 4
  eAjouter = 5
End Enum


' Type de période
Public Enum tagTypePeriode
  eProvision = 1
  eCapitalConstitutifRente = 2
  eRevalo = 3
  eProvisionRetraite = 4
  eProvisionRetraiteRevalo = 5
  eStatutaire = 6
End Enum

Public Enum OperationStatus
  efailure = 0
  eSuccess = 1
  eNoData = 2
End Enum

Public Enum ArchiveRestoreErrors
  errDBConnectionProblem
  errBulkInsertNoRecordsInserted
  
End Enum


' Couleurs
Public Const LTGRAY As Long = &HC0C0C0
Public Const GRAY As Long = &H808080

Public Const CYAN As Long = &HFFFF00
Public Const LTCYAN As Long = &HFFFFC0
Public Const DKCYAN As Long = &H804000

Public Const LTYELLOW As Long = &HC0FFFF
Public Const YELLOW As Long = &HFFFF

Public Const LTGREEN As Long = &H80FF80
Public Const GREEN As Long = &HFF00

Public Const LTRED As Long = &H8080FF
Public Const RED As Long = &HFF
Public Const DKRED As Long = &H80

Public Const PINK As Long = &HC0C0FF


'
' Couleur VB : Bleu Vert Rouge
'
Public Const bleu As Long = &HFFC893
Public Const bleu_clair As Long = &HFFFFC0
Public Const vert_clair As Long = &HC0FFC0
Public Const jaune_clair As Long = &HC0FFFF
Public Const lavande_clair As Long = &HFFC0DC
Public Const orange_clair As Long = &HC0E6FF
Public Const noir As Long = &H0
Public Const blanc As Long = &HFFFFFF


'Import Statutaire - Non-Statutaire
Public Const cImportStat As String = "ImpStat"
Public Const cImportStandard As String = "ImpStandard"
Public Const cPeriodeStat As String = "PerStat"
Public Const cPeriodeStandard As String = "PerStandard"


'Enums for Import
Public Enum eTypeImport
  eImportComplet = 1
  eImportDonneesSeules = 2
  eImportTablesParametresSeules = 3
End Enum

Public Enum eTypeDelaiInactivite
  eDatePaiement = 1
  eDateFinPeriodePaiement = 2
End Enum

Public Enum eTypeCalculAnnualisation
  eEnsemblePaiement = 1
  eDernierPaiement = 2
End Enum




