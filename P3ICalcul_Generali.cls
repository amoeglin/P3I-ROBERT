VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "P3ICalcul_Generali"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75E600C5"
Option Explicit


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Calcul des Provisions suivant les spécifications Generali
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements iP3ICalcul


' pour le traitement des maintien deces
Private Enum TypeGarantieDeces
  Deces
  RenteConjoint
  RenteEducation
End Enum

' pour la méthode de calcul des maintien deces
Public Enum MethodeCalcul
  CapitauxConstitutifsSousRisques = 0
  CotisationsExonerees = 1
End Enum

'##ModelId=5C8A75E601EE
Private autoMode As Boolean
'##ModelId=5C8A75E6021C
Private avantReforme As Boolean
'##ModelId=5C8A75E6023B
Private apresReforme As Boolean

' membres
'##ModelId=5C8A75E60258
Private m_NumPeriode As Long           ' numéro de la période concernée
'##ModelId=5C8A75E6026A
Private m_NumGroupe As Long            ' numéro du groupe concerné
'##ModelId=5C8A75E60298
Private m_bRevalo As Boolean           ' si true, on revalorise

'##ModelId=5C8A75E602CC
Private m_Logger As clsLogger          ' fichier log
'##ModelId=5C8A75E602D5
Private m_fWait As frmWait


' variables utilisées pendant les calculs
'##ModelId=5C8A75E602D8
Private m_rsPeriode As ADODB.Recordset ' recordset Période
'Private m_rsAssure As ADODB.Recordset  ' recordset Assuré
'Private m_rsAssureTemp As ADODB.Recordset ' recordset AssuréTemp (Maintiens Décès)
  
'##ModelId=5C8A75E602E5
Private m_rsParamCalcul As ADODB.Recordset

'##ModelId=5C8A75E602E6
Private m_Posit As Integer             ' Position de l'assuré

'##ModelId=5C8A75E60316
Private m_rsStatProv As ADODB.Recordset                 ' required for Statutaire Calculation
'##ModelId=5C8A75E60324
Private m_rsStatCat As ADODB.Recordset                 ' required for Statutaire Calculation
'##ModelId=5C8A75E60335
Private m_rsStatutaireGarantie As ADODB.Recordset
'Private m_rsStatParams As ADODB.Recordset               ' required for
'Statutaire Calculation
'##ModelId=5C8A75E60343
Private m_rsCatR9 As ADODB.Recordset                    ' recherche dans la table CATR9
'##ModelId=5C8A75E60355
Private m_rsNCA As ADODB.Recordset                      ' recherche dans la table PassageNCA
'##ModelId=5C8A75E60362
Private m_rsPMRetenue As ADODB.Recordset                ' recherche dans la table PM_Retenue
'##ModelId=5C8A75E60374
Private m_rsTBQREGA As ADODB.Recordset                  ' recherche dans la table TBQREGA
'Private m_rsREGA01 As ADODB.Recordset                   ' recherche dans la
'table REGA01
'##ModelId=5C8A75E60382
Private m_rsCodesCat As ADODB.Recordset                 ' recherche dans la table CodesCat
'##ModelId=5C8A75E60393
Private m_rsCorrespondance_CatOption As ADODB.Recordset ' recherche dans la table Correspondance_CatOption
'Private m_rsProduits As ADODB.Recordset                 ' recherche dans la
'table Produits
'##ModelId=5C8A75E603A1
Private m_rsCapitaux_Moyens As ADODB.Recordset          ' recherche dans la table Capitaux_Moyens
'##ModelId=5C8A75E603B2
Private m_rsCDSITUAT As ADODB.Recordset                 ' recherche dans la table CDSITUAT

' RecordSet deconnecté pour accelerer les recherches
'##ModelId=5C8A75E603C0
Private m_rsTaux_FindTauxRenteEducation As ADODB.Recordset
'##ModelId=5C8A75E603D2
Private m_rsTauxRenteEducation As ADODB.Recordset
'##ModelId=5C8A75E70007
Private m_rsTauxRenteConjoint As ADODB.Recordset
'##ModelId=5C8A75E7000A
Private m_rsParamRentes As ADODB.Recordset
'##ModelId=5C8A75E70019
Private m_rsDonneesSociales As ADODB.Recordset

'##ModelId=5C8A75E70026
Private m_rsCorrespondanceRegimeCategorie As ADODB.Recordset
'##ModelId=5C8A75E70029
Private m_rsFindRegimeCategorieMaintienDeces As ADODB.Recordset
'##ModelId=5C8A75E70038
Private m_rsPlafondSS As ADODB.Recordset
'##ModelId=5C8A75E70045
Private m_rsIndemnisationSalaireSS As ADODB.Recordset
'##ModelId=5C8A75E70048
Private m_rsAgeSituationFamille As ADODB.Recordset
'##ModelId=5C8A75E70057
Private m_rsLissageProvision As ADODB.Recordset

'##ModelId=5C8A75E70067
Private m_rsProvisionBCAC As ADODB.Recordset
  
'##ModelId=5C8A75E70074
Private m_rsTaux_Reassurance As ADODB.Recordset

' Reforme des retraites
'##ModelId=5C8A75E70077
Private m_rsCoeffAmortissement As ADODB.Recordset
'##ModelId=5C8A75E70086
Private m_rsAgeDepartRetraite As ADODB.Recordset
Private m_rsAgeDepartRetraiteInval1 As ADODB.Recordset


'##ModelId=5C8A75E700A3
Private m_rsAnnualisation As ADODB.Recordset
'##ModelId=5C8A75E700A6
Private m_rsBaseRevalo As ADODB.Recordset

' Variables des fonctions de calcul de la revalorisation
'##ModelId=5C8A75E700B3
Private m_TMO As Double                ' Taux d'indexation
'##ModelId=5C8A75E700C2
Private m_DureRevalo As Single         ' Durée d'indexation en année
'##ModelId=5C8A75E700E2
Private m_TauxRevalo As Double         ' Taux de revalorisation
'##ModelId=5C8A75E70101
Private m_DureeLissage As Single


' Statutaire
'##ModelId=5C8A75E70120
Private m_CleTable_Statutaire As Long
'##ModelId=5C8A75E7013F
Private m_NomTable_Statutaire As String

' BCAC Incap
'##ModelId=5C8A75E7015F
Private m_BCAC_CleTable_Incap As Long
'##ModelId=5C8A75E7016E
Private m_BCAC_NomTable_Incap As String
'##ModelId=5C8A75E7018E
Private m_BCAC_AgeMin_Incap As Integer
'##ModelId=5C8A75E701AD
Private m_BCAC_AgeMax_Incap As Integer
'##ModelId=5C8A75E701CC
Private m_BCAC_CleTable_Incap_Retraite As Long
'##ModelId=5C8A75E701EB
Private m_BCAC_NomTable_Incap_Retraite As String
'##ModelId=5C8A75E70215
Private m_BCAC_AgeMin_Incap_Retraite As Integer
'##ModelId=5C8A75E70227
Private m_BCAC_AgeMax_Incap_Retraite As Integer
' BCAC Inval
'##ModelId=5C8A75E70246
Private m_BCAC_CleTable_Inval As Long
'##ModelId=5C8A75E7026B
Private m_BCAC_NomTable_Inval As String
'##ModelId=5C8A75E7028A
Private m_BCAC_AgeMin_Inval As Integer
'##ModelId=5C8A75E702A9
Private m_BCAC_AgeMax_Inval As Integer
'##ModelId=5C8A75E702CF
Private m_BCAC_CleTable_Inval_Retraite As Long
'##ModelId=5C8A75E702E2
Private m_BCAC_NomTable_Inval_Retraite As String
'##ModelId=5C8A75E70302
Private m_BCAC_AgeMin_Inval_Retraite As Integer
'##ModelId=5C8A75E70330
Private m_BCAC_AgeMax_Inval_Retraite As Integer

' Table Incap
'##ModelId=5C8A75E70350
Private m_CleTable_Incap As String
'##ModelId=5C8A75E7037F
Private m_AgeMin_Incap As Integer
'##ModelId=5C8A75E7039E
Private m_AgeMax_Incap As Integer
'##ModelId=5C8A75E703BD
Private m_CleTable_Incap_Retraite As String
'##ModelId=5C8A75E703DC
Private m_AgeMin_Incap_Retraite As Integer
'##ModelId=5C8A75E80014
Private m_AgeMax_Incap_Retraite As Integer
' Table Passage
'##ModelId=5C8A75E80042
Private m_CleTable_Pass As String
'##ModelId=5C8A75E80071
Private m_AgeMin_Pass As Integer
'##ModelId=5C8A75E80091
Private m_AgeMax_Pass As Integer
'##ModelId=5C8A75E800A0
Private m_CleTable_Pass_Retraite As String
'##ModelId=5C8A75E800BF
Private m_AgeMin_Pass_Retraite As Integer
'##ModelId=5C8A75E800DF
Private m_AgeMax_Pass_Retraite As Integer
' Table Inval
'##ModelId=5C8A75E8010E
Private m_CleTable_Inval As String
'##ModelId=5C8A75E8012D
Private m_AgeMin_Inval As Integer
'##ModelId=5C8A75E8014C
Private m_AgeMax_Inval As Integer
'##ModelId=5C8A75E8016B
Private m_CleTable_Inval_Retraite As String
'##ModelId=5C8A75E8018B
Private m_AgeMin_Inval_Retraite As Integer
'##ModelId=5C8A75E801AA
Private m_AgeMax_Inval_Retraite As Integer


' Situation Famille
'##ModelId=5C8A75E801C9
Private m_AgeSituationFamille_AgeMini_Cadre As Integer
'##ModelId=5C8A75E801E8
Private m_AgeSituationFamille_AgeMaxi_Cadre As Integer
'##ModelId=5C8A75E80219
Private m_AgeSituationFamille_AgeMini_NonCadre As Integer
'##ModelId=5C8A75E80238
Private m_AgeSituationFamille_AgeMaxi_NonCadre As Integer

' calcul P3I Individuel
'##ModelId=5C8A75E80257
Dim m_bP3I_Individuel As Boolean
  
' Date de calcul
'##ModelId=5C8A75E80276
Private m_DateCalcul As Date

'Type Periode
'##ModelId=5C8A75E802B5
Private m_IDTypeCalcul As Integer


' age de l'assuré pour les calculs
'##ModelId=5C8A75E802D7
Private m_xArret As Integer       ' âge à la date d'arrêt
'##ModelId=5C8A75E80306
Private m_xArretMois As Integer   ' âge à la date d'arrêt en mois
'##ModelId=5C8A75E80335
Private m_xEffet As Integer       ' âge à la date d'effet du contrat
'##ModelId=5C8A75E80354
Private m_xCalcul As Integer      ' âge à la date du calcul (=date de fin de période)
'##ModelId=5C8A75E80383
Private m_xTerme As Integer       ' âge à la date du terme du contrat (=de départ en retraite) (entier)
'##ModelId=5C8A75E803B2
Private m_xTermeDouble As Double  ' âge à la date du terme du contrat (=de départ en retraite) (double)

' Variables pour le stockage de la liste des catégories non trouvé dans
'TauxRenteEducation
'##ModelId=5C8A75E803E3
Private m_CatNotFoundRenteEducation As clsListCategorieIntrouvable
'##ModelId=5C8A75E90095
Private m_CatNotFoundRenteConjoint As clsListCategorieIntrouvable
'##ModelId=5C8A75E90104
Private m_CatNotFoundGarantieDC As clsListCategorieIntrouvable
'##ModelId=5C8A75E90170
Private m_CatNotFoundCorrespondanceGarantie As clsListCategorieIntrouvable

'##ModelId=5C8A75E90214
Private m_PlafondSSNotFound As clsListCategorieIntrouvable
'##ModelId=5C8A75E9026F
Private m_IndemnisationNotFound As clsListCategorieIntrouvable
'##ModelId=5C8A75E902D6
Private m_AgeSituationFamilleNotFound As clsListCategorieIntrouvable

'##ModelId=5C8A75E90353
Private m_ProduitsNotFound As clsListCategorieIntrouvable
'##ModelId=5C8A75E903C0
Private m_CodesCatNotFound As clsListCategorieIntrouvable
'##ModelId=5C8A75EA0038
Private m_Correspondance_CatOptionNotFound As clsListCategorieIntrouvable
'##ModelId=5C8A75EA00CB
Private m_Capitaux_MoyensNotFound As clsListCategorieIntrouvable

'##ModelId=5C8A75EA0139
Private m_MessageUnique As clsListeMessageUnique

' module de calcul des provisions
'##ModelId=5C8A75EA0146
Private m_ModuleCalcul_Provisions As clsCalcul_Provisions
'##ModelId=5C8A75EA01B4
Private m_recordsetList As clsRecordsetList

' module de calcul des provisions avec revalorisation
'##ModelId=5C8A75EA0204
Private m_ModuleCalcul_Provisions_Revalo As clsCalcul_Provisions_Revalo

' module de calcul des PM maintien Inval (recalcul BCAC)
'##ModelId=5C8A75EA0284
Private m_ModuleCalcul_PM_MaintienDeces As clsCalcul_PM_MaintienDeces

' Variables pour le stockage des types de garantie et de calcul du maintien décès
Private Type garantie
  CleGarantie As Long
  TypeGarantie As Long
  Libelle As String
  CalculMaintienDeces As Boolean
  CalculPassage As Boolean
End Type

'##ModelId=5C8A75EA02FE
Private m_NbGarantie As Integer
'##ModelId=5C8A75EA032F
Private m_Garantie() As garantie

' Méthode de calcul de Age et Ancienneté
'##ModelId=5C8A75EA0330
Private m_bMethodeVincent As Boolean

' type pour le stockage de la définition d'une garantie maintien décès
Private Type GarantieDC
  RegimeIncap                       As Long
  CategorieIncap                    As String
  RegimeDeces                       As Long
  CategorieDeces                    As String
  RegimeRenteConjointTemporaire     As Long
  CategorieRenteConjointTemporaire  As String
  RegimeRenteConjointViagere        As Long
  CategorieRenteConjointViagere     As String
  RegimeRenteEduc                   As Long
  CategorieRenteEduc                As String
  CVD0                              As Double
  M0                                As Double
  CVD1                              As Double
  M1                                As Double
  MajorationEnfant                  As Double
  MajorationAccident                As Double
  LissageProvision                  As Boolean
End Type

'##ModelId=5C8A75EA035C
Private m_BridageAge As Boolean
'##ModelId=5C8A75EA037B
Private m_NumParametreCalculAbsent As Boolean
' P3I Individuel
'Private m_bP3I_Individuel As Boolean
' proportion des invalidités 1ère catégorie
' Private m_Inval1Cat As Double

'

'##ModelId=5C8A75EA039B
Private Sub Class_Initialize()
  m_NumPeriode = -1
  m_NumGroupe = -1
  m_Posit = 0
  m_BridageAge = False
  
  m_AgeMin_Inval = 0
  m_AgeMax_Inval = 0
  m_AgeMin_Pass = 0
  m_AgeMax_Pass = 0
  m_AgeMin_Incap = 0
  m_AgeMax_Incap = 0
  
  m_AgeMin_Inval_Retraite = 0
  m_AgeMax_Inval_Retraite = 0
  m_AgeMin_Pass_Retraite = 0
  m_AgeMax_Pass_Retraite = 0
  m_AgeMin_Incap_Retraite = 0
  m_AgeMax_Incap_Retraite = 0
  
  'm_bP3I_Individuel = IIf(GetSettingIni(CompanyName, SectionName, "P3I_Individuel", "0") <> "0", True, False)
  
  ' m_Inval1Cat = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "Inval1Cat", "30"))
End Sub
  

'##ModelId=5C8A75EA03BA
Private Sub OpenDisconnected()
  On Error GoTo err_open
  
  Dim defaultOpenType As eRecordsetMode
  
  defaultOpenType = Disconnected
  
  ' on n'ouvre les recordsets qu'une seul fois et on travaille ensuite en mémoire (recordsets déconnectés)
  Set m_rsTauxRenteConjoint = m_dataSource.OpenRecordset("SELECT * FROM TauxRenteConjoint", defaultOpenType)
  Set m_rsTauxRenteEducation = m_dataSource.OpenRecordset("SELECT * FROM TauxRenteEducation", defaultOpenType)
  Set m_rsTaux_FindTauxRenteEducation = m_dataSource.OpenRecordset("SELECT * FROM TauxRenteEducation", defaultOpenType)
  Set m_rsCorrespondanceRegimeCategorie = m_dataSource.OpenRecordset("SELECT * FROM CorrespondanceGarantie", defaultOpenType)
  Set m_rsFindRegimeCategorieMaintienDeces = m_dataSource.OpenRecordset("SELECT * FROM GarantieDC", defaultOpenType)
  Set m_rsAgeSituationFamille = m_dataSource.OpenRecordset("SELECT * FROM AgeSituationFamille", defaultOpenType)
  Set m_rsPlafondSS = m_dataSource.OpenRecordset("SELECT * FROM PlafondSS", defaultOpenType)
  Set m_rsIndemnisationSalaireSS = m_dataSource.OpenRecordset("SELECT * FROM IndemnisationIncapInval", defaultOpenType)
  Set m_rsLissageProvision = m_dataSource.OpenRecordset("SELECT * FROM LissageProvision", defaultOpenType)
  
  Set m_rsAgeDepartRetraite = m_dataSource.OpenRecordset("SELECT * FROM AgeDepartRetraite WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, defaultOpenType)
  Set m_rsAgeDepartRetraiteInval1 = m_dataSource.OpenRecordset("SELECT * FROM AgeDepartRetraiteInval1 WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, defaultOpenType)
  Set m_rsCoeffAmortissement = m_dataSource.OpenRecordset("SELECT * FROM CoeffAmortissement WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode & " AND IdTypeCalcul=" & m_rsPeriode.fields("IdTypeCalcul"), defaultOpenType)
  
  ' Tables de référence
  LoadReferenceTables "ParamCalcul", "PELMINCAP"
  LoadReferenceTables "ParamCalcul", "PELMPASSAGE"
  LoadReferenceTables "ParamCalcul", "PELMINVAL"
  LoadReferenceTables "ParamCalcul", "PELMINCAP_Retraite"
  LoadReferenceTables "ParamCalcul", "PELMPASSAGE_Retraite"
  LoadReferenceTables "ParamCalcul", "PELMINVAL_Retraite"
  LoadReferenceTables "ParamCalcul", "PETRENTEEDUC"
  LoadReferenceTables "ParamCalcul", "PETRENTECONJOINT"
    
  ' pour les rentes
  Set m_rsParamRentes = m_dataSource.OpenRecordset("SELECT * FROM ParamRentes WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, defaultOpenType)
  
  LoadReferenceTables "ParamRentes", "TableHomme"
  LoadReferenceTables "ParamRentes", "TableFemme"
  
  ' Pour l'annualisation passage
  Set m_rsDonneesSociales = m_dataSource.OpenRecordset("SELECT * FROM DonneesSociales WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, defaultOpenType)
  
  
  ' Tables de référence Mortalité des assurés en Incap ou en Inval
  m_recordsetList.LoadRecordset "MortInval"
  m_recordsetList.LoadRecordset "MortIncap"
  'm_recordsetList.LoadRecordset "MortInval", m_rsPeriode.fields("PECleTableMortBCACInval_MDC")
  'm_recordsetList.LoadRecordset "MortIncap", m_rsPeriode.fields("PECleTableMortBCACIncap_MDC")
  
  
  ' Table Provisions BCAC précalculées
  Dim sSql As String
  
  sSql = BuildListID("ParamCalcul", "PECleTableBCACIncap_MDC", "ProvisionBCAC", "CleTable")
  sSql = sSql & " OR " & BuildListID("ParamCalcul", "PECleTableBCACInval_MDC", "ProvisionBCAC", "CleTable")
  sSql = sSql & " OR " & BuildListID("ParamCalcul", "PECleTableBCACIncap_MDC_Retraite", "ProvisionBCAC", "CleTable")
  sSql = sSql & " OR " & BuildListID("ParamCalcul", "PECleTableBCACInval_MDC_Retraite", "ProvisionBCAC", "CleTable")
  
  Set m_rsProvisionBCAC = m_dataSource.OpenRecordset("SELECT * FROM ProvisionBCAC WHERE " & sSql, defaultOpenType)
  
  Exit Sub
  
err_open:
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Sub


'##ModelId=5C8A75EA03D9
Private Function BuildListID(sNomTableParam As String, sNomChampParam As String, sNomTable As String, sNomChampId As String) As String
  Dim rs As ADODB.Recordset, sSql As String
  
  Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT " & sNomChampParam & " FROM " & sNomTableParam _
                                      & " WHERE PEGPECLE=" & m_NumGroupe & " AND PENUMCLE=" & m_NumPeriode, Snapshot)
  sSql = ""
  Do Until rs.EOF
    sSql = sSql & sNomChampId & "=" & rs.fields(0) & " OR "
    
    rs.MoveNext
  Loop
  
  rs.Close
  
  BuildListID = Left(sSql, Len(sSql) - 4)
End Function


'##ModelId=5C8A75EB007D
Private Sub LoadReferenceTables(sNomTableParam As String, sNomChampParam As String)
  Dim rs As ADODB.Recordset, sSql As String
  
  If sNomTableParam = "ParamRentes" Then
    Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT " & sNomChampParam & " FROM " & sNomTableParam _
                                        & " WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Snapshot)
  Else
    Set rs = m_dataSource.OpenRecordset("SELECT DISTINCT " & sNomChampParam & " FROM " & sNomTableParam _
                                        & " WHERE PEGPECLE=" & m_NumGroupe & " AND PENUMCLE=" & m_NumPeriode, Snapshot)
  End If
  
  Do Until rs.EOF
    If Not IsNull(rs.fields(0)) Then
      m_recordsetList.LoadRecordset rs.fields(0)
    End If
    
    rs.MoveNext
  Loop
  
  rs.Close
End Sub


'##ModelId=5C8A75EB00CC
Private Sub CloseAllDisconnected()
  CloseDisconnected m_rsTaux_FindTauxRenteEducation
  CloseDisconnected m_rsTauxRenteEducation
  CloseDisconnected m_rsTauxRenteConjoint
  CloseDisconnected m_rsFindRegimeCategorieMaintienDeces
  CloseDisconnected m_rsCorrespondanceRegimeCategorie
  CloseDisconnected m_rsPlafondSS
  CloseDisconnected m_rsIndemnisationSalaireSS
  CloseDisconnected m_rsAgeSituationFamille
  CloseDisconnected m_rsLissageProvision
  
  CloseDisconnected m_rsProvisionBCAC

  CloseDisconnected m_rsParamRentes
  CloseDisconnected m_rsDonneesSociales
  
  ' Reforme des retarite
  CloseDisconnected m_rsCoeffAmortissement
  CloseDisconnected m_rsAgeDepartRetraite
End Sub


'##ModelId=5C8A75EB00EB
Private Sub LogAllCategorie()
  m_MessageUnique.LogMessage
  
  m_CatNotFoundRenteEducation.LogCategorie
  m_CatNotFoundRenteConjoint.LogCategorie
  m_CatNotFoundGarantieDC.LogCategorie
  m_CatNotFoundCorrespondanceGarantie.LogCategorie
  
  m_PlafondSSNotFound.LogCategorie
  m_IndemnisationNotFound.LogCategorie
  m_AgeSituationFamilleNotFound.LogCategorie
  
  m_ProduitsNotFound.LogCategorie
  m_CodesCatNotFound.LogCategorie
  m_Correspondance_CatOptionNotFound.LogCategorie
  m_Capitaux_MoyensNotFound.LogCategorie
End Sub


'##ModelId=5C8A75EB010A
Private Sub CloseDisconnected(ByRef theRS As ADODB.Recordset)
  If Not theRS Is Nothing Then
    theRS.Close
    Set theRS = Nothing
  End If
End Sub


'##ModelId=5C8A75EB0139
Private Sub LoadGarantie()
  Dim rs As ADODB.Recordset
  Dim rq As String
  
  On Error GoTo err_LoadGarantie
  
  rq = "SELECT Garantie.GAGARCLE, Garantie.GALIB, Garantie.GACALPASS, TypeGarantie.TYPEGARCLE, TypeGarantie.CalculProvMaintien" _
      & " FROM Garantie Garantie INNER JOIN TypeGarantie TypeGarantie ON Garantie.GATYPEGAR = TypeGarantie.TYPEGARCLE " _
      & " WHERE Garantie.GAGROUPCLE=" & m_NumGroupe
  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  m_NbGarantie = 0
  
  Do Until rs.EOF
    ReDim Preserve m_Garantie(m_NbGarantie + 1) As garantie
    
    m_Garantie(m_NbGarantie).CleGarantie = rs.fields("GAGARCLE")
    m_Garantie(m_NbGarantie).TypeGarantie = rs.fields("TYPEGARCLE")
    m_Garantie(m_NbGarantie).Libelle = rs.fields("GALIB")
    m_Garantie(m_NbGarantie).CalculMaintienDeces = rs.fields("CalculProvMaintien")
    m_Garantie(m_NbGarantie).CalculPassage = rs.fields("GACALPASS")
    
    m_NbGarantie = m_NbGarantie + 1
    
    rs.MoveNext
  Loop
  
  rs.Close
  
  Exit Sub
  
err_LoadGarantie:
  
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
End Sub


'##ModelId=5C8A75EB0158
Private Sub FreeGarantie()
  m_NbGarantie = 0
  Erase m_Garantie
End Sub


'##ModelId=5C8A75EB0178
Private Function FindGarantie(CleGarantie As Long) As garantie
  Dim i As Integer
  
  For i = 0 To m_NbGarantie - 1
    If m_Garantie(i).CleGarantie = CleGarantie Then
      FindGarantie = m_Garantie(i)
      Exit Function
    End If
  Next
  
  Err.Raise -1, "FindGarantie", "Garantie introuvable : clé=" & CleGarantie
End Function


' fabrique la date anniversaire
'##ModelId=5C8A75EB01A6
Private Function DateAnniversaire(dateNais As Date, ByVal Age As Integer) As Date
  DateAnniversaire = DateSerial(Year(dateNais) + Age, Month(dateNais), Day(dateNais))
End Function


' calcul de l'age à la date de calcul
'##ModelId=5C8A75EB01F5
Private Function AgeAnniversaire(dateNais As Date, dateCalcul As Date) As Integer
  Dim i As Integer
  
  AgeAnniversaire = 0
  While DateSerial(Year(dateNais) + AgeAnniversaire, Month(dateNais), Day(dateNais)) < dateCalcul
    AgeAnniversaire = AgeAnniversaire + 1
  Wend
  AgeAnniversaire = AgeAnniversaire - 1
  If AgeAnniversaire < 0 Then
    AgeAnniversaire = 0
  End If
End Function

'##ModelId=5C8A75EB0252
Private Function FindTauxRenteEducation(Categorie As String, dateNaissance As Date, dateFinPeriodePaiement As Date, bMaintienDeces As Boolean, Age As Double) As Double
  ' recherche du taux
  Dim age_b1 As Integer
  Dim taux_b1 As Double
  Dim age_b2 As Integer
  Dim taux_b2 As Double
  Dim age_b3 As Integer
  Dim taux_b3 As Double
  Dim age_b4 As Integer
  Dim taux_b4 As Double
  
  Dim date_b1 As Date
  Dim date_b2 As Date
  Dim date_b3 As Date
  Dim date_b4 As Date
  
  
  m_dataHelper.Multi_Find m_rsTaux_FindTauxRenteEducation, "Categorie='" & Categorie & "'"
  
  If Not m_rsTaux_FindTauxRenteEducation.EOF Then
    age_b1 = m_rsTaux_FindTauxRenteEducation.fields("B1")
    taux_b1 = m_rsTaux_FindTauxRenteEducation.fields("B1_PCT")
    age_b2 = m_rsTaux_FindTauxRenteEducation.fields("B2")
    taux_b2 = m_rsTaux_FindTauxRenteEducation.fields("B2_PCT")
    age_b3 = m_rsTaux_FindTauxRenteEducation.fields("B3")
    taux_b3 = m_rsTaux_FindTauxRenteEducation.fields("B3_PCT")
    age_b4 = m_rsTaux_FindTauxRenteEducation.fields("B4")
    taux_b4 = m_rsTaux_FindTauxRenteEducation.fields("B4_PCT")
  Else
    m_CatNotFoundRenteEducation.AddCategorie 0, Categorie
    
    FindTauxRenteEducation = 0
    
    Exit Function
  End If
  
  ' age de l'enfant
  If IsNull(dateFinPeriodePaiement) Then
    FindTauxRenteEducation = 0
    Exit Function
  End If
    
  If bMaintienDeces Then
    ' recherche du palier
    If Age < age_b1 Then
      FindTauxRenteEducation = taux_b1
    ElseIf Age < age_b2 Then
      FindTauxRenteEducation = taux_b2
    ElseIf Age < age_b3 Then
      FindTauxRenteEducation = taux_b3
    ElseIf Age < age_b4 Then
      FindTauxRenteEducation = taux_b4
    Else
      FindTauxRenteEducation = 0
    End If
  Else
    ' palier en date anniversaire
    date_b1 = DateAnniversaire(dateNaissance, age_b1)
    date_b2 = DateAnniversaire(dateNaissance, age_b2)
    date_b3 = DateAnniversaire(dateNaissance, age_b3)
    date_b4 = DateAnniversaire(dateNaissance, age_b4)
    
    If dateNaissance > dateFinPeriodePaiement Then
      ' ne devrait pas se produire sauf en corse ou dans le Veme arrondissement de paris
      FindTauxRenteEducation = 0
      Exit Function
    End If
  
    ' recherche du palier
    If dateFinPeriodePaiement < date_b1 Then
      FindTauxRenteEducation = taux_b1
    ElseIf dateFinPeriodePaiement < date_b2 Then
      FindTauxRenteEducation = taux_b2
    ElseIf dateFinPeriodePaiement < date_b3 Then
      FindTauxRenteEducation = taux_b3
    ElseIf dateFinPeriodePaiement < date_b4 Then
      FindTauxRenteEducation = taux_b4
    Else
      FindTauxRenteEducation = 0
    End If
  End If
  
  If FindTauxRenteEducation < 0 Then
    FindTauxRenteEducation = 0
  End If
End Function

'##ModelId=5C8A75EB02EC
Private Function PassageCATR9(Categorie As String, NCA As String) As Boolean
  PassageCATR9 = False
  
  m_dataHelper.Multi_Find m_rsCatR9, "Categorie='" & Categorie & "'"
  
  If Not m_rsCatR9.EOF Then
    ' existe donc Incap+Passage ou Incap seul
    If CBool(m_rsCatR9.fields("PassageSuivantNCA")) = True Then
      m_dataHelper.Multi_Find m_rsNCA, "NCA=""" & NCA & """"
      
      If m_rsNCA.EOF Then
        ' si pas d'entrée dans PassageNCA, c'est de l'incap avec passage
        PassageCATR9 = True ' incap avec passage
      Else
        ' decode l'entrée
        If CBool(m_rsNCA.fields("Passage")) = True Then
          PassageCATR9 = True ' incap avec passage
        Else
          PassageCATR9 = False ' incap sans passage
        End If
      End If
    Else
      ' decode l'entree dans CATR9
      If CBool(m_rsCatR9.fields("Passage")) = True Then
        PassageCATR9 = True ' incap avec passage
      Else
        PassageCATR9 = False ' incap sans passage
      End If
    End If
  End If
End Function


'##ModelId=5C8A75EB033A
Private Sub InitListNotFound(ByRef c As clsListCategorieIntrouvable, nomTable As String, Optional NomCategorie As String = "Catégorie", Optional NomRegime As String = "Régime")
  Set c = New clsListCategorieIntrouvable
  
  If autoMode Then
    c.autoMode = True
  End If
  
  Set c.Logger = m_Logger
  c.tableName = nomTable
  If Not IsMissing(NomCategorie) Then
    c.NomCategorie = NomCategorie
  End If
  If Not IsMissing(NomRegime) Then
    c.NomRegime = NomRegime
  End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calcul des Provisions suivant les spécifications ISICA
'
'##ModelId=5C8A75EB03D6
Public Sub iP3ICalcul_CalculProvisionsAssures(bRevalo As Boolean, numPeriode As Long, NumGroupe As Long, Logger As clsLogger, Optional automatMode As Boolean = False, Optional Avant As Boolean = True, Optional Apres As Boolean = False)
  
  Dim rq As String, query_assure_retraite As String
  
  Dim StartTime As Double, EndTime As Double
  Dim debut As Date, fin As Date
  
  'StartTime = Timer
  debut = Now
  autoMode = automatMode
  avantReforme = Avant
  apresReforme = Apres
    
  On Error GoTo err_calcul
  
  Screen.MousePointer = vbHourglass
  
  m_bMethodeVincent = False
  If m_bMethodeVincent Then
    If Not autoMode Then
      MsgBox "Age et Ancienneté méthode Vincent !", vbCritical
    End If
  End If
  
  ' init des membres
  m_NumPeriode = numPeriode
  m_NumGroupe = NumGroupe
  m_bRevalo = bRevalo
  
  ' preparation du fichier log
  Set m_Logger = Logger
  If Not autoMode Then
    m_Logger.CreateLog "Lancement du calcul des provisions de la Période " & m_NumPeriode & " du groupe " & m_NumGroupe
  End If
    
  ' verification des paramètres
  If iP3ICalcul_CheckParametresAssures(numPeriode, NumGroupe, m_Logger) = False Then
        
    If Not autoMode Then
      Screen.MousePointer = vbDefault
      m_Logger.AfficheErreurLog False
    End If
    
    Dim Msg As String
    
    Msg = "Des erreurs ont été détectées dans les tables de configuration." & vbLf
    If m_NumParametreCalculAbsent Then
      Msg = Msg & vbLf & "!!! ATTENTION CERTAINS NUMEROS DE PARAMETRES DE CALCUL NE SONT PAS SPECIFIES DANS LES PARAMETRES DE LA PERIODE => VOIR LA TABLE CODESCAT" & vbLf
      Msg = Msg & "!!! LES LIGNES CORRESPODANTES NE SERONT PAS CALCULEES" & vbLf & vbLf
    End If
    Msg = Msg & "Voulez-vous continuer les calculs ?"
    
    If Not autoMode Then
      If MsgBox(Msg, vbQuestion + vbYesNo) = vbNo Then
        m_Logger.EcritTraceDansLog "Calcul annulé par l'utilisateur suite aux erreurs de paramétrage"
        Exit Sub
      End If
    Else
      Msg = "Des erreurs ont été détectées dans les tables de configuration." & vbLf
      If m_NumParametreCalculAbsent Then
        Msg = Msg & vbLf & "!!! ATTENTION CERTAINS NUMEROS DE PARAMETRES DE CALCUL NE SONT PAS SPECIFIES DANS LES PARAMETRES DE LA PERIODE => VOIR LA TABLE CODESCAT" & vbLf
        Msg = Msg & "!!! LES LIGNES CORRESPODANTES NE SERONT PAS CALCULEES" & vbLf & vbLf
      End If
      m_Logger.EcritTraceDansLog Msg
    End If
  End If

  m_Logger.EcritTraceDansLog "  "
  m_Logger.EcritTraceDansLog "##########    Calcul des provisions"
  m_Logger.EcritTraceDansLog "  "
  
  ' lancement des calculs
  Screen.MousePointer = vbHourglass
  
  Set m_fWait = New frmWait
  
  m_fWait.Caption = "Calcul des provisions en cours..."
  
  m_fWait.ProgressBar1.Min = 0
  m_fWait.ProgressBar1.Value = 0
  m_fWait.ProgressBar1.Max = 100
  
  ' init des listes de paramètres introuvables
  InitListNotFound m_CatNotFoundRenteEducation, "ParamRentes", "Catégorie", "Année survenance" ' PHM 28/11/2007 : anciennement table "TauxRenteEducation"
  InitListNotFound m_CatNotFoundRenteConjoint, "TauxRenteConjoint"
  InitListNotFound m_CatNotFoundGarantieDC, "GarantieDC"
  InitListNotFound m_CatNotFoundCorrespondanceGarantie, "CorrespondanceGarantie"
  
  InitListNotFound m_PlafondSSNotFound, "PlafondSS", "Année"
  InitListNotFound m_IndemnisationNotFound, "IndemnisationIncapInval"
  InitListNotFound m_AgeSituationFamilleNotFound, "AgeSituationFamille", "Age", "IsCadre"
  
  InitListNotFound m_ProduitsNotFound, "Produits", "Famille Comptable", "Catégorie"
  InitListNotFound m_CodesCatNotFound, "CodesCat", "Famille Comptable", "Catégorie"
  InitListNotFound m_Correspondance_CatOptionNotFound, "Correspondance_CatOption", "Code Option"
  InitListNotFound m_Capitaux_MoyensNotFound, "Capitaux_Moyens", "Code Produit", "Code Option Synthetique"
  
  Set m_MessageUnique = New clsListeMessageUnique
  Set m_MessageUnique.Logger = m_Logger
  
  If autoMode Then
    m_MessageUnique.autoMode = True
  End If
  
  ' 10/05/2004 : ancienne position
'  m_dataSource.BeginTrans
  
  If Not autoMode Then
    m_fWait.Show vbModeless
  End If
  
  ' charge les définitions de garantie
  LoadGarantie
  
  ' remplace les NULL par '' dans la table Reassurance
  m_dataSource.Execute m_dataHelper.ValidateSQL("UPDATE Reassurance SET NCA='' WHERE NCA is Null")
  
  ' efface les PMGD du calcul précédent
  m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode & " AND POGARCLE>90")
  m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_Retraite WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode & " AND POGARCLE>90")
  
  ' creer la requete : on parcours en testant DATE_COMPT comprise entre DateDebut et DateFin
  '*** Modif UPESE : rq = "SELECT * FROM Assure WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
  '*** PHM 25/09/2004 : rq = "SELECT * FROM Assure Assure INNER JOIN Garantie Garantie ON Assure.POGARCLE = Garantie.GAGARCLE WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
  rq = "SELECT * FROM Assure WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
  query_assure_retraite = "SELECT * FROM Assure_Retraite WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode & " AND POIdAssure="
  
  
  Dim m_rsAssure As ADODB.Recordset  ' recordset Assuré
  Dim m_rsAssureTemp As ADODB.Recordset ' recordset AssuréTemp (Maintiens Décès)
  
  Dim m_rsAssure_Retraite As ADODB.Recordset ' recordset AssuréTemp (Maintiens Décès)
  Dim m_rsAssureTemp_Retraite As ADODB.Recordset ' recordset AssuréTemp (Maintiens Décès)

  ' parcours les articles correspondant
  Set m_rsAssure = m_dataSource.OpenRecordset(rq, Dynamic)
  
  If Not m_rsAssure.EOF Then
    m_rsAssure.MoveLast
    
    m_fWait.ProgressBar1.Max = m_rsAssure.RecordCount
    
    m_rsAssure.MoveFirst
  Else
    m_fWait.ProgressBar1.Max = 1
  End If
  
  Set m_rsStatCat = m_dataSource.OpenRecordset("SELECT * From Statutaire_Categorie", Disconnected)
  Set m_rsStatutaireGarantie = m_dataSource.OpenRecordset("SELECT * FROM P3IUser.Statutaire_Garantie INNER JOIN P3IUser.Statutaire_Garantie_Code ON P3IUser.Statutaire_Garantie.ID = P3IUser.Statutaire_Garantie_Code.IDGarantie", Disconnected)
  
  Set m_rsCatR9 = m_dataSource.OpenRecordset("SELECT * From CATR9 WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  Set m_rsNCA = m_dataSource.OpenRecordset("SELECT * From PassageNCA WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  
'PHM 01/06/2006
  Set m_rsPMRetenue = m_dataSource.OpenRecordset("SELECT * From PM_Retenue WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  Set m_rsTBQREGA = m_dataSource.OpenRecordset("SELECT * From TBQREGA WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
'  Set m_rsREGA01 = m_dataSource.OpenRecordset("SELECT * From REGA01 WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  Set m_rsCodesCat = m_dataSource.OpenRecordset("SELECT * From CodesCat WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  Set m_rsCorrespondance_CatOption = m_dataSource.OpenRecordset("SELECT * From Correspondance_CatOption WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
'  Set m_rsProduits = m_dataSource.OpenRecordset("SELECT * From Produits WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  Set m_rsCapitaux_Moyens = m_dataSource.OpenRecordset("SELECT * From Capitaux_Moyens WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
'PHM 01/06/2006
  Set m_rsCDSITUAT = m_dataSource.OpenRecordset("SELECT * FROM CDSITUAT WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  
  m_dataSource.Execute "UPDATE Reassurance SET Resiliation=" & m_dataHelper.BuildDateString(#6/1/2079#) & " WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode & " AND Resiliation IS NULL"
  
  Set m_rsTaux_Reassurance = m_dataSource.OpenRecordset("SELECT * FROM Reassurance WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode, Disconnected)
  
  '*****   lecture des caractéristiques de la période traitée dans la table PERIODE    *****
  '*****   et des paramètres de calcul dans la table ParamCalcul
  'rq = "SELECT * FROM Periode WHERE PEGPECLE=" & m_NumGroupe & " AND PENUMCLE=" & m_NumPeriode
  rq = "SELECT PE.*, PA.* FROM Periode PE INNER JOIN ParamCalcul PA " _
      & " ON PA.PEGPECLE = PE.PEGPECLE AND PA.PENUMCLE = PE.PENUMCLE " _
      & " WHERE PE.PEGPECLE=" & m_NumGroupe & " AND PE.PENUMCLE=" & m_NumPeriode _
      & " ORDER BY PA.PENUMPARAMCALCUL"
  
'  Set m_rsPeriode = m_dataSource.OpenRecordset(rq, Dynamic)
  Set m_rsPeriode = m_dataSource.OpenRecordset(rq, Disconnected)
  
  If m_rsPeriode.EOF Then
  
    If Not autoMode Then
      Screen.MousePointer = vbDefault
      MsgBox "Les informations du groupe " & m_NumGroupe & " période " & m_NumPeriode & vbLf & _
           " sont absentes des Tables Periode et ParamCalcul"
    Else
      m_Logger.EcritTraceDansLog "Les informations du groupe " & m_NumGroupe & " période " & m_NumPeriode & vbLf & _
           " sont absentes des Tables Periode et ParamCalcul"
    End If
    
    Exit Sub
  End If
  
  
  '####
  
  'This is only required for periodes of type statutaire
  
  'If m_rsPeriode.fields("PETYPEPERIODE") = eStatutaire Then
  
    Set m_rsParamCalcul = m_dataSource.OpenRecordset("Select * from ParamCalcul Where PENUMCLE = " & m_NumPeriode & " And PEGPECLE = " & m_NumGroupe, Disconnected)
      
    If Not m_rsParamCalcul.EOF Then
      m_CleTable_Statutaire = IIf(IsNull(m_rsParamCalcul.fields("PECleTableStatutaire")), 0, m_rsParamCalcul.fields("PECleTableStatutaire"))
      
      ' Load the Statutaire_Provision Table to get all the coeffs for the Statutaire calculation
      'Set m_rsStatProv = m_dataSource.OpenRecordset("SELECT * From Statutaire_Provision", Disconnected)
    
      If m_CleTable_Statutaire <> 0 Then
        m_NomTable_Statutaire = m_dataHelper.GetParameter("SELECT NOMTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_CleTable_Statutaire)
        Set m_rsStatProv = m_dataSource.OpenRecordset("SELECT * From " & m_NomTable_Statutaire, Disconnected)
      End If
    End If
            
  'End If
  
  '###
  
  ' P3I Individuel
  If Not IsNull(m_rsPeriode.fields("PEP3I_INDIVIDUEL")) Then
    m_bP3I_Individuel = m_rsPeriode.fields("PEP3I_INDIVIDUEL")
  Else
    m_bP3I_Individuel = False
  End If
  
  
  ' Date de calcul
  m_DateCalcul = m_rsPeriode.fields("PEDATEFIN")
  
  'Type Periode
  m_IDTypeCalcul = m_rsPeriode.fields("IdTypeCalcul")
  
  m_Logger.EcritTraceDansLog "!!!! Calcul des provisions de la Période " & m_NumPeriode & " du groupe " & m_NumGroupe & " au " & Format(m_DateCalcul, "dd/MM/yyyy")
  m_Logger.EcritTraceDansLog "  "
  
  m_BridageAge = IIf(m_rsPeriode.fields("PEBRIDAGE_AGE") = True, True, False)
  
  ' init des valeurs Min et Max des tables de paramètres
  InitMinMax
  
  ' charge les modules de calcul
  Set m_recordsetList = New clsRecordsetList
  Set m_ModuleCalcul_Provisions = New clsCalcul_Provisions
  Set m_ModuleCalcul_Provisions_Revalo = New clsCalcul_Provisions_Revalo
  Set m_ModuleCalcul_PM_MaintienDeces = New clsCalcul_PM_MaintienDeces
  
  If autoMode Then
    m_recordsetList.autoMode = True
    m_ModuleCalcul_Provisions.autoMode = True
    m_ModuleCalcul_Provisions_Revalo.autoMode = True
    m_ModuleCalcul_PM_MaintienDeces.autoMode = True
    
    Set m_recordsetList.m_Logger = m_Logger
    Set m_ModuleCalcul_Provisions.m_Logger = m_Logger
    Set m_ModuleCalcul_Provisions_Revalo.m_Logger = m_Logger
    Set m_ModuleCalcul_PM_MaintienDeces.m_Logger = m_Logger
  End If
  
  Set m_ModuleCalcul_Provisions.recordsetList = m_recordsetList
  Set m_ModuleCalcul_Provisions_Revalo.recordsetList = m_recordsetList
  Set m_ModuleCalcul_Provisions_Revalo.ModuleCalcul_Provisions = m_ModuleCalcul_Provisions
  Set m_ModuleCalcul_PM_MaintienDeces.recordsetList = m_recordsetList
  
  m_recordsetList.LoadRecordset "TableMortalite"
  
  '***************
  
  ' ouvre tous les recordset des tables de références
  OpenDisconnected
  
  ' Annualisation par assuré
  rq = "SELECT POSTECLE, POGARCLE, POCONVENTION, POGARCLE_NEW, POCATEGORIE, POARRET, SUM(POPRESTATION_AN) as SUM_POPRESTATION_AN FROM ASSURE WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode _
       & " GROUP BY POSTECLE, POGARCLE, POCONVENTION, POGARCLE_NEW, POCATEGORIE, POARRET"
  
  Set m_rsAnnualisation = m_dataSource.OpenRecordset(rq, Disconnected)
  
  
  ' Présence Base/Revalo par assuré
  rq = "SELECT DISTINCT POSTECLE, PONUMCLE, PONOM, POARRET, PONAIS, POBaseRevalo, POPRESTATION_AN FROM ASSURE WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode _
       & " ORDER BY POSTECLE, PONUMCLE, PONOM, POARRET, PONAIS, POBaseRevalo"
  
  Set m_rsBaseRevalo = m_dataSource.OpenRecordset(rq, Disconnected)
  
  
  ' efface les traces du calcul précédent
  m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM AssureTemp WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode)
  m_dataSource.Execute m_dataHelper.ValidateSQL("DELETE FROM Assure_RetraiteTemp WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode)
  
  ' préparation du calcul des maintiens décès
  Set m_rsAssureTemp = m_dataSource.OpenRecordset("AssureTemp", table)
  Set m_rsAssureTemp_Retraite = m_dataSource.OpenRecordset("Assure_RetraiteTemp", table)
'  Set m_rsAssureTemp = m_dataSource.OpenRecordset("SELECT * FROM AssureTemp WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode, Disconnected)
  
  ' déclaration des variables des fonctions fctcalcprovinvalRevalo, fctcalcProvIncapRevalo
  m_DureeLissage = m_rsPeriode.fields("PEDUREELISSAGE")
  m_TMO = m_rsPeriode.fields("PETMO") / 100#
  m_DureRevalo = m_rsPeriode.fields("PEDUREEREVALO")
  m_TauxRevalo = m_rsPeriode.fields("PETXREVALO") / 100#

  ' code retour : si false, on annule les calculs (rollback)
  Dim codeRetour As Boolean
  
  codeRetour = True
  
  
  ' 10/05/2004 : nouvelle position
  m_dataSource.BeginTrans "CALCUL_PERIODE_" & m_NumPeriode
  
  Do While Not m_rsAssure.EOF
  
    'If m_rsAssure.fields("NUENRP3I") = 42975 Then Stop
  
    'write to log if the type of periode is STAT and the Assure is not of type STAT
    If m_rsPeriode.fields("PETYPEPERIODE") = eStatutaire Then
      If cdPosit_Statutaire <> m_rsAssure.fields("POSIT") Then
        m_Logger.EcritTraceDansLog "Le type de la période en traitement est du type 'STATUTAIRE' mais l'assuré numéro " & m_rsAssure.fields("NUENRP3I") & " n'est pas du type 'STATUTAIRE' !"
      End If
    Else
      If cdPosit_Statutaire = m_rsAssure.fields("POSIT") Then
        m_Logger.EcritTraceDansLog "Le type de la période en traitement n'est pas du type 'STATUTAIRE' mais l'assuré numéro " & m_rsAssure.fields("NUENRP3I") & " est du type 'STATUTAIRE' !"
      End If
    End If
    
    'StartTime = Timer
    
    If (m_rsAssure.AbsolutePosition Mod 9) = 0 Then
      ' affiche la position
      m_fWait.ProgressBar1.Value = m_rsAssure.AbsolutePosition
      m_fWait.Label1(0).Caption = "Article n°" & m_rsAssure.AbsolutePosition & " / " & m_fWait.ProgressBar1.Max
      m_fWait.Refresh
      DoEvents
    End If
    
    If m_fWait.fTravailAnnule = True Then
      codeRetour = False
      Exit Do
    End If
    
    
    
    '
    ' Paramètres de calcul
    '
    m_dataHelper.Multi_Find m_rsPeriode, "PENumParamCalcul=" & m_rsAssure.fields("PONumParamCalcul")
    If m_rsPeriode.EOF Then
      ' Paramètre introuvable
      m_rsPeriode.filter = adFilterNone
      m_Logger.EcritTraceDansLog "Période n°" & m_NumPeriode & " : Jeux de Paramètres de calcul n°" & m_rsAssure.fields("PONumParamCalcul") & " introuvables !"
      GoTo assure_suivant
    End If
    
    
    '
    ' Ligne Assure_Retraite associé si "portefeuille salaries"
    '
    If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
        And m_rsPeriode.fields("PEPortefeuilleSalaries") And m_rsAssure.fields("POGARCLE") = cdGar59 Then
      Set m_rsAssure_Retraite = m_dataSource.OpenRecordset(query_assure_retraite & m_rsAssure.fields("RECNO"), Dynamic)
      If m_rsAssure_Retraite.EOF = True Then
        m_Logger.EcritTraceDansLog "!!!! Assuré n°" & m_rsAssure.fields("RECNO") & " introuvables dans Assure_Retraite. Veuillez refaire l'import ! CALCUL INTERROMPU !!!!"
        codeRetour = False
        Exit Do
      End If
    End If
    

' tables BCAC
    ' Incap
    If m_BCAC_CleTable_Incap <> m_rsPeriode.fields("PECleTableBCACIncap_MDC") Then
      m_BCAC_CleTable_Incap = m_rsPeriode.fields("PECleTableBCACIncap_MDC")
      m_BCAC_NomTable_Incap = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))
      m_BCAC_AgeMin_Incap = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))
      m_BCAC_AgeMax_Incap = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))
    End If
    
    ' Inval
    If m_BCAC_CleTable_Inval <> m_rsPeriode.fields("PECleTableBCACInval_MDC") Then
      m_BCAC_CleTable_Inval = m_rsPeriode.fields("PECleTableBCACInval_MDC")
      m_BCAC_NomTable_Inval = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
      m_BCAC_AgeMin_Inval = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
      m_BCAC_AgeMax_Inval = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
    End If
' tables BCAC
  
' Table Incap, Passage, Inval
    If m_CleTable_Incap <> m_rsPeriode.fields("PELMINCAP") Then
      m_CleTable_Incap = m_rsPeriode.fields("PELMINCAP")
      m_AgeMin_Incap = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINCAP"))
      m_AgeMax_Incap = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINCAP"))
    End If
    
    If m_CleTable_Pass <> m_rsPeriode.fields("PELMPASSAGE") Then
      m_CleTable_Pass = m_rsPeriode.fields("PELMPASSAGE")
      m_AgeMin_Pass = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE"))
      m_AgeMax_Pass = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE"))
    End If
      
    If m_CleTable_Inval <> m_rsPeriode.fields("PELMINVAL") Then
      m_CleTable_Inval = m_rsPeriode.fields("PELMINVAL")
      m_AgeMin_Inval = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINVAL"))
      m_AgeMax_Inval = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINVAL"))
    End If
' Table Incap, Passage, Inval
    
    
    
  ' Reforme des Retraites
  ' tables BCAC
    ' Incap
    If m_BCAC_CleTable_Incap_Retraite <> m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite") Then
      m_BCAC_CleTable_Incap_Retraite = m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite")
      m_BCAC_NomTable_Incap_Retraite = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))
      m_BCAC_AgeMin_Incap_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))
      m_BCAC_AgeMax_Incap_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))
    End If
    
    ' Inval
    If m_BCAC_CleTable_Inval_Retraite <> m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite") Then
      m_BCAC_CleTable_Inval_Retraite = m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite")
      m_BCAC_NomTable_Inval_Retraite = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
      m_BCAC_AgeMin_Inval_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
      m_BCAC_AgeMax_Inval_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
    End If
  ' tables BCAC
  
  ' Table Incap, Passage, Inval
    If m_CleTable_Incap_Retraite <> m_rsPeriode.fields("PELMINCAP_Retraite") Then
      m_CleTable_Incap_Retraite = m_rsPeriode.fields("PELMINCAP_Retraite")
      m_AgeMin_Incap_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINCAP_Retraite"))
      m_AgeMax_Incap_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINCAP_Retraite"))
    End If
    
    If m_CleTable_Pass_Retraite <> m_rsPeriode.fields("PELMPASSAGE_Retraite") Then
      m_CleTable_Pass_Retraite = m_rsPeriode.fields("PELMPASSAGE_Retraite")
      m_AgeMin_Pass_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE_Retraite"))
      m_AgeMax_Pass_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE_Retraite"))
    End If
      
    If m_CleTable_Inval_Retraite <> m_rsPeriode.fields("PELMINVAL_Retraite") Then
      m_CleTable_Inval_Retraite = m_rsPeriode.fields("PELMINVAL_Retraite")
      m_AgeMin_Inval_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINVAL_Retraite"))
      m_AgeMax_Inval_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINVAL_Retraite"))
    End If
  ' Table Incap, Passage, Inval
    
    
    
    
    '
    ' traitement d'un assuré
    '
    'If m_rsAssure.fields("NUENRP3I") = 42975 Then Stop

    m_Posit = m_rsAssure.fields("POSIT")
    
    ' s'agit-il d'une rente ?
    Dim bIsRente As Boolean
    
    bIsRente = m_Posit = cdPosit_RenteCertaine Or m_Posit = cdPosit_RenteEducationTemporaire _
                Or m_Posit = cdPosit_RenteEducationViagere Or m_Posit = cdPosit_RenteConjointTemporaire _
                Or m_Posit = cdPosit_RenteConjointViagere Or m_Posit = cdPosit_RenteRetraiteTemporaire _
                Or m_Posit = cdPosit_RenteRetraiteViagere Or m_Posit = cdPosit_RenteAutreTemporaire _
                Or m_Posit = cdPosit_RenteAutreViagere
     
     
     
    '*********************
    '*********************
    '*********************
    '*********************
    
'    If Left(Trim(m_rsAssure.fields("PONOM")), 4) = "EDET" Then Stop
  
   
    '*********************
    '*********************
    '*********************
    '*********************
    
    
    
    ' âge à la date d'arrêt (millésime)
    'm_xArret = Year(m_rsAssure.fields("POARRET")) - Year(m_rsAssure.fields("PONAIS"))
    m_xArret = CalculAge(m_rsAssure.fields("PONAIS"), m_rsAssure.fields("POARRET"))
    
    ' âge à la date d'effet du contrat
    'm_xEffet = Year(m_rsAssure.fields("POEFFET")) - Year(m_rsAssure.fields("PONAIS"))
    m_xEffet = CalculAge(m_rsAssure.fields("PONAIS"), m_rsAssure.fields("POEFFET"))
    
    ' âge à la date du calcul (=date de fin de période)
    'm_xCalcul = Year(m_DateCalcul) - Year(m_rsAssure.fields("PONAIS"))
    m_xCalcul = CalculAge(m_rsAssure.fields("PONAIS"), m_DateCalcul)
          
    ' âge à la date du terme du contrat ( = âge de départ en retraite)
    'm_xTerme = Year(m_rsAssure.fields("POTERME")) - Year(m_rsAssure.fields("PONAIS"))
    If bIsRente = False And Not IsNull(m_rsAssure.fields("POTERME")) Then
      m_xTerme = CalculAge(m_rsAssure.fields("PONAIS"), m_rsAssure.fields("POTERME"))
      m_xTermeDouble = CalculAgeDouble(m_rsAssure.fields("PONAIS"), m_rsAssure.fields("POTERME"))
    Else
      m_xTerme = 120
      m_xTermeDouble = 120#
    End If
    
    
    ''''''''''''''''
    
    m_xArretMois = DateDiff("m", m_rsAssure.fields("PONAIS"), m_rsAssure.fields("POARRET"))
    If m_xArretMois < 0 Then
      m_xArretMois = 0
    End If
    
    
    ClearResultFields m_rsAssure, False
    If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
        And m_rsPeriode.fields("PEPortefeuilleSalaries") And m_rsAssure.fields("POGARCLE") = cdGar59 Then
      ClearResultFields m_rsAssure_Retraite, True
    End If
    
    ' test AM
    'If rsAssure.fields("NUENRP3I") = 14639 Then
    'Stop
    'End If
    
    '****
    '****
    '**** l'assuré est en INCAPACITE (Posit 1, 3, 1001 1002 1003 1004 1005 1006) ***
    '****
    '**** ATTENTION : CalculIncapacité peut passer l'assuré en Inval.
    '****             il faut donc que CalculInvalidité soit après CalculIncapacité
    
    ' si la date d'entrée en inval n'est pas renseignée alors Incap (sinon aller Inval)
    'If m_rsAssure.fields("NUENRP3I") = 5120 Then Stop
    
    If IsNull(m_rsAssure.fields("PODATEENTREEINVAL")) Then
    
        'code Provision                Variable                             Incap   Passage Inval
        '1001       cdProvision_AvecIncap_AvecPassage_AvecInval  = 1001     Avec    Avec    Avec
        '1002       cdProvision_AvecIncap_AvecPassage_SansInval  = 1002     Avec    Avec    Sans
        '1003       cdProvision_AvecIncap_SansPassage_AvecInval  = 1003     Avec    Sans    Avec
        '1004       cdProvision_AvecIncap_SansPassage_SansInval  = 1004     Avec    Sans    Sans
        '1005       cdProvision_SansIncap_AvecPassage_AvecInval  = 1005     Sans    Avec    Avec
        '1006       cdProvision_SansIncap_AvecPassage_SansInval  = 1006     Sans    Avec    Sans
        '1007       cdProvision_SansIncap_SansPassage_AvecInval  = 1007     Sans    Sans    Avec
        '1008       cdProvision_SansIncap_SansPassage_SansInval  = 1008     Sans    Sans    Sans
        '6001       cdProvision_Incap_AvecPassage_Viager  = 6001     Avec (Viagere)    Avec (Viagere)    Sans
        
        If m_Posit = cdPosit_IncapAvecPassage _
        Or m_Posit = cdPosit_IncapSansPassage _
        Or m_Posit = cdProvision_AvecIncap_AvecPassage_AvecInval _
        Or m_Posit = cdProvision_AvecIncap_AvecPassage_SansInval _
        Or m_Posit = cdProvision_AvecIncap_SansPassage_AvecInval _
        Or m_Posit = cdProvision_AvecIncap_SansPassage_SansInval _
        Or m_Posit = cdProvision_SansIncap_AvecPassage_AvecInval _
        Or m_Posit = cdProvision_SansIncap_AvecPassage_SansInval _
        Or m_Posit = cdProvision_SansIncap_SansPassage_AvecInval _
        Or m_Posit = cdProvision_SansIncap_SansPassage_SansInval _
        Or m_Posit = cdProvision_Incap_AvecPassage_Viager Then   ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
          
        Dim bForceCoeffProvisionIncapToZero As Boolean
          bForceCoeffProvisionIncapToZero = False
          
          If autoMode Then
            If avantReforme Then
              CalculIncapacité m_rsAssure, False, bForceCoeffProvisionIncapToZero
            End If
            
            If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
                And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
              If apresReforme Then
                CalculIncapacité m_rsAssure_Retraite, True, bForceCoeffProvisionIncapToZero
              End If
            End If
          Else
            CalculIncapacité m_rsAssure, False, bForceCoeffProvisionIncapToZero 'False
            ' ATTENTION : m_Posit peut passer à 2 Inval si l'assuré est en incap depuis plus de 36 mois
            'm _rsAssure.fields("POSIT") = m_Posit ' mise à jour du code position de l'assuré car il peut passer à 2
            
            If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
                And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
              CalculIncapacité m_rsAssure_Retraite, True, bForceCoeffProvisionIncapToZero ' False
            End If
          
          End If
        
        End If
    
    End If
    
    
    '****
    '****
    '**** STATUTAIRE (Posit 5000)
    '****
    '****
    
    If m_Posit = cdPosit_Statutaire Then
      If autoMode Then
        If avantReforme Then
          CalculStatutaire m_rsAssure, False
        End If
      
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculStatutaire m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculStatutaire m_rsAssure, False
      
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculStatutaire m_rsAssure_Retraite, True
        End If
      End If
      
    End If
    
    '****
    '****
    '**** l'assuré est INVALIDE (Posit = 2, 1007, 6003 ou devenu Invalide car Incapacité avec anc >= 36 mois) (Posit 2) ***
    '****
    '****
    
    
'    Dim test As Long
'    test = m_rsAssure.fields("NUENRP3I")
'    If m_rsAssure.fields("NUENRP3I") = 150 Then
'      m_Posit = 6003
'    End If
    
    
    If m_Posit <> cdPosit_Statutaire Then
      If m_Posit = cdPosit_Inval _
          Or m_Posit = cdProvision_SansIncap_SansPassage_AvecInval _
          Or m_Posit = cdProvision_Inval_Viager _
          Or Not IsNull(m_rsAssure.fields("PODATEENTREEINVAL")) Then  ' l'assuré est INVALIDE, utilisation de la fonction FctCalcProvInval
        
        If autoMode Then
          If avantReforme Then
            CalculInvalidité m_rsAssure, False
          End If
          
          If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
              And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
            If apresReforme Then
              CalculInvalidité m_rsAssure_Retraite, True
            End If
          End If
        Else
          CalculInvalidité m_rsAssure, False
          
          If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
              And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
            CalculInvalidité m_rsAssure_Retraite, True
          End If
        End If
        
      
      End If
    End If
    
    
   
    '****
    '****
    '**** Rente conjoint (Posit 4) *** désactivé
    '****
    '****
    
'    If m_Posit = cdPosit_RenteConjoint Then
'      CalculRenteConjoint m_rsAssure, False, 0, 0, 0, False, 0
'    End If
    
    
    
    '****
    '****
    '**** Rente education (Posit 5) *** désactivé
    '****
    '****
    
'    If m_Posit = cdPosit_RenteEduc Then
'      CalculRenteEducation m_rsAssure, False, 0, 0, 0, 0, 0
'    End If
    
    
    
    '****
    '****
    '**** deces (Posit 6) *** A VOIR AVEC FRANCIS
    '****
    '****
    
    If m_Posit = cdPosit_Deces Then
      CalculDeces m_rsAssure
    End If
    
    
    
    '****
    '****
    '**** maternité (Posit 7)
    '****
    '****
    
    If m_Posit = cdPosit_Maternite Then
      If autoMode Then
        If avantReforme Then
          CalculMaternite m_rsAssure, False
        End If
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculMaternite m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculMaternite m_rsAssure, False
      
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculMaternite m_rsAssure_Retraite, True
        End If
      End If
    End If
    
    
    
    '****
    '****
    '**** Mensualisation (Posit 8)
    '****
    '****
    
    If m_Posit = cdPosit_Mensualisation Then
      If autoMode Then
        If avantReforme Then
          CalculMensualisation m_rsAssure, False
        End If
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculMensualisation m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculMensualisation m_rsAssure, False
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculMensualisation m_rsAssure_Retraite, True
        End If
      End If
    End If
    
    
    '****
    '****
    '**** Chomage (Posit 9)
    '****
    '****
    
    If m_Posit = cdPosit_Chomage Then
      If autoMode Then
        If avantReforme Then
          CalculChomage m_rsAssure, False
        End If
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculChomage m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculChomage m_rsAssure, False
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculChomage m_rsAssure_Retraite, True
        End If
      End If
    End If

    
    '****
    '****
    '**** Rentes (Posit 20) ***
    '****
    '****
    
    'If m_rsAssure.fields("NUENRP3I") = 35 Then Stop
    
    If bIsRente = True Then
      Dim tp As tagTypePeriode
      
      If IsNull(m_rsPeriode.fields("PETYPEPERIODE")) Then
        tp = eProvision
      Else
        tp = m_rsPeriode.fields("PETYPEPERIODE")
      End If
        
      Select Case tp
        Case eProvision
          If autoMode Then
            If avantReforme Then
              CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
            End If
          Else
            CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
          End If
      
        Case eCapitalConstitutifRente
          If autoMode Then
            If avantReforme Then
              CalculCapitalConstitutifRentes m_rsAssure, False, 0, 0, 0, 0, 0
            End If
          Else
            CalculCapitalConstitutifRentes m_rsAssure, False, 0, 0, 0, 0, 0
          End If
      
        Case eRevalo
          If autoMode Then
            If avantReforme Then
              CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
            End If
          Else
            CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
          End If
      
        Case eProvisionRetraite, eProvisionRetraiteRevalo
          If autoMode Then
            If avantReforme Then
              CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
              ' correction RS + AM le 08 05 2018 non activée
              'CalculRentes m_rsAssure_Retraite, False, 0, 0, 0, 0, 0
            End If
          Else
            CalculRentes m_rsAssure, False, 0, 0, 0, 0, 0
            ' correction RS + AM le 08 05 2018 non activée
            'CalculRentes m_rsAssure_Retraite, False, 0, 0, 0, 0, 0
          End If

        
        Case Else
          Logger.EcritTraceDansLog "!!!! Calcul des provisions de la Période " & m_NumPeriode & " du groupe " & m_NumGroupe & " : Type de période inconnu !!!"
      End Select
    
    End If
    
    '****
    '****
    '**** dependance (Posit 30)
    '****
    '****
    'If m_rsAssure.fields("PONUMCLE") = 10030260 Then Stop

    If m_Posit = cdProvision_dependanceIndividuel Then
      If autoMode Then
        If avantReforme Then
          CalculDependance m_rsAssure, False
        End If
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculDependance m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculDependance m_rsAssure, False
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculDependance m_rsAssure_Retraite, True
        End If
      End If
    End If
    
    '****
    '****
    '**** CapitalInvaliditeProbable (Posit 40)
    '****
    '****

    If m_Posit = cdProvision_CapitalInvaliditeProbable Then
      If autoMode Then
        If avantReforme Then
          CalculCapitalInvaliditeProbable m_rsAssure, False
        End If
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculCapitalInvaliditeProbable m_rsAssure_Retraite, True
          End If
        End If
      Else
        CalculCapitalInvaliditeProbable m_rsAssure, False
        
        If (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
            And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculCapitalInvaliditeProbable m_rsAssure_Retraite, True
        End If
      End If
    End If
 
    '****
    '***  Calcul de la Provision pour sinistre a payer (PSAP)avec l'annualisation  UNIQUEMENT POUR P3I INDIVIDUEL
    '***  la PSAP de P3I COLLECTIVE est calculée dans l'IMPORT
    '**** pour P3I Individuel  = (PSAP * abattement * (1+ gestion))  ou bien ( = PSAP DECES * abattement pour le décès)
    '****
    
    ' If m_bP3I_Individuel = True Then
    ' test
'    If bForceCoeffProvisionIncapToZero = True Then
'    bForceCoeffProvisionIncapToZero = True
'    End If

    ' Test AM
'    If m_rsAssure.fields("NUENRP3I") = 155 Then
'    Stop
'    End If
        
        
    If bForceCoeffProvisionIncapToZero = False Then  ' si (bForceCoeffProvisionIncapToZero = True) la PSAP est déjà forcée à 0  -> CalculPSAP NON EFFECTUE (modif AM 03/02/2018)
        If autoMode Then
          If avantReforme Then
            CalculPSAP m_rsAssure, m_bP3I_Individuel, False
          End If
          
          If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
             And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
            If apresReforme Then
              CalculPSAP m_rsAssure_Retraite, m_bP3I_Individuel, True
            End If
          End If
        Else
          CalculPSAP m_rsAssure, m_bP3I_Individuel, False
          
          If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
             And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
            CalculPSAP m_rsAssure_Retraite, m_bP3I_Individuel, True
          End If
        End If
    End If
    
    ' End If
  
    '****
    '****
    '**** Code Situation  ***
    '****
    '****
    
    'Application du taux PM (Table CodeSituation) sur POPM et POPSAP)
    
    'If m_Posit <> cdPosit_Statutaire Then
    'If Not IsNull(m_rsPeriode.fields("PECALCULER")) And m_rsPeriode.fields("PECALCULER") = False Then
    ' Modif AM RS 02/05/2019 le Test est désactivé pour calculer toujours POPM ET POPSAP avec le coef TauxPM

      If autoMode Then
        If avantReforme Then
          CalculCodeSituation m_rsAssure
        End If
        
        If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
           And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          If apresReforme Then
            CalculCodeSituation m_rsAssure_Retraite
          End If
        End If
      Else
        CalculCodeSituation m_rsAssure
        
        If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
           And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
          CalculCodeSituation m_rsAssure_Retraite
        End If
      End If
    
    'fin du test désactivé AM RS 02/05/2019
    'End If
    
    
    '****
    '****
    '**** Reassurance  ***
    '****
    '****
    
    If autoMode Then
      If avantReforme Then
        CalculReassurance m_rsAssure
      End If
      
      If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
         And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
        If apresReforme Then
          CalculReassurance m_rsAssure_Retraite
        End If
      End If
    Else
      CalculReassurance m_rsAssure
      
      If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
         And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
        CalculReassurance m_rsAssure_Retraite
      End If
    End If
    
    
    
    '****
    '**** dossier clot  ***
    '****
    '****
    
    If autoMode Then
      If avantReforme Then
        DossierClot m_rsAssure
      End If
      
      If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
         And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
        If apresReforme Then
          DossierClot m_rsAssure_Retraite
        End If
      End If
    Else
      DossierClot m_rsAssure
      
      If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
         And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
        DossierClot m_rsAssure_Retraite
      End If
    End If
    
    
    
    '****
    '**** Maintien en garantie décès  ***
    '****
    '****
    
    If m_Posit <> cdPosit_Statutaire Then
      If m_bP3I_Individuel = False Then  ' (maintien décès pour P3I collective uniquement) 31/07/2013
      
          If (bIsRente = False) Then  ' (ce n'est pas une rente) alors calcul du maintien
            
            If autoMode Then
              If avantReforme Then
                CalculMaintienDeces m_rsAssure, m_rsAssureTemp, False, m_rsAssureTemp
              End If
              
              If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
                 And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
                If apresReforme Then
                  CalculMaintienDeces m_rsAssure_Retraite, m_rsAssureTemp_Retraite, True, m_rsAssureTemp
                End If
              End If
            Else
              CalculMaintienDeces m_rsAssure, m_rsAssureTemp, False, m_rsAssureTemp
              
              If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
                 And m_rsAssure.fields("POGARCLE") = cdGar59 And m_rsPeriode.fields("PEPortefeuilleSalaries") Then
                CalculMaintienDeces m_rsAssure_Retraite, m_rsAssureTemp_Retraite, True, m_rsAssureTemp
              End If
            End If
          End If
      End If
    End If


    ' valide les changements
    m_rsAssure.Update

    If m_rsPeriode.fields("PEPortefeuilleSalaries") And (eProvisionRetraite = m_rsPeriode.fields("PETYPEPERIODE") Or eProvisionRetraiteRevalo = m_rsPeriode.fields("PETYPEPERIODE")) _
       And m_rsPeriode.fields("PEPortefeuilleSalaries") And m_rsAssure.fields("POGARCLE") = cdGar59 Then
      m_rsAssure_Retraite.Update
      m_rsAssure_Retraite.Close
    End If
    
assure_suivant:
    ' assuré suivant
    m_rsAssure.MoveNext
    
    
    DoEvents
    'EndTime = Timer
    'Debug.Print "Calcul Loop: ", EndTime - StartTime

  Loop
  
  ' affiche les catégories introuvables
  LogAllCategorie
  
  ' ferme la boite d'attente
  If Not autoMode Then
    m_fWait.Hide
    Unload m_fWait
  End If
  
  Set m_fWait = Nothing
  
  ' ne valide que si il n'y a pas d'erreur
  If codeRetour = False Then
    ' annule les changements
    m_dataSource.RollbackTrans "CALCUL_PERIODE_" & m_NumPeriode
    
    m_Logger.EcritTraceDansLog "L'opération de calcul a été INTERROMPUE !   -   Aucun article n'a été modifié pour la période n°." & m_NumPeriode
    
    If Not autoMode Then
      MsgBox "L'opération de calcul a été INTERROMPUE !" & vbLf & "Aucun article n'a été modifié pour la période n°." & m_NumPeriode, vbExclamation
    End If
  Else
'    Set m_rsAssure.ActiveConnection = m_dataSource.Connection
'    m_rsAssure.UpdateBatch
    
'    Set m_rsAssureTemp.ActiveConnection = m_dataSource.Connection
'    m_rsAssureTemp.UpdateBatch
    
    ' recopie dans Assure les calculs des Maintiens en Garanties Décès qui sont dans AssureTemp
    RecopieMaintienDeces "Assure", "AssureTemp", False
    RecopieMaintienDeces "Assure_Retraite", "Assure_RetraiteTemp", True
    
    ' Connecte les lignes MDC d'Assure_Retraite à celles de Assure via Assure.POIdAssureMDC et Assure_Retraite.POIdAssure
    ConnecteLignesMDC
    
    ' purge des résultats intermédiares
    m_dataSource.Execute "DELETE FROM AssureTemp WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
    m_dataSource.Execute "DELETE FROM Assure_RetraiteTemp WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
    
    ' valide les changements
    m_dataSource.CommitTrans "CALCUL_PERIODE_" & m_NumPeriode
    
    'Modification 22/09/2013 : Mise à jour du PEBCALCULER = True (lorsqu'on clique sur le bouton calculer ou calculer&revaloriser)
    rq = "UPDATE Periode SET PECALCULER = 'True' WHERE PEGPECLE = " & m_NumGroupe & " AND PENUMCLE = " & m_NumPeriode
    m_dataSource.Execute rq
 
        
    Call m_Logger.EcritTraceDansLog("L'opération de calcul a été terminée avec succès")
     
  End If
  
  ' ferme les recordset deconnectés
  CloseAllDisconnected
  
  m_rsAnnualisation.Close
  m_rsBaseRevalo.Close
  
  m_rsAssureTemp.Close
  m_rsAssure.Close
  m_rsPeriode.Close
  
' PHM 15/01/2004
  m_rsCatR9.Close
  m_rsNCA.Close
' PHM 15/01/2004
  
' PHM 01/06/2006
  m_rsPMRetenue.Close
  m_rsTBQREGA.Close
'  m_rsREGA01.Close
  m_rsCodesCat.Close
  m_rsCorrespondance_CatOption.Close
'  m_rsProduits.Close
  m_rsCapitaux_Moyens.Close
' PHM 01/06/2006
  m_rsCDSITUAT.Close
  
  m_rsTaux_Reassurance.Close
  
  If Not m_rsStatProv Is Nothing Then
    m_rsStatProv.Close
  End If
  
  m_rsStatCat.Close
  m_rsStatutaireGarantie.Close
  
  m_dataSource.Execute "UPDATE Reassurance SET Resiliation=NULL WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode & " AND Resiliation=" & m_dataHelper.BuildDateString(#6/1/2079#)
  
  
  ' libere les definitions de garantie
  FreeGarantie
    
  ' libere les modules de calcul
  Set m_ModuleCalcul_Provisions = Nothing
  Set m_ModuleCalcul_Provisions_Revalo = Nothing
  Set m_ModuleCalcul_PM_MaintienDeces = Nothing
  Set m_recordsetList = Nothing
  
  ' controles de resultats
  iP3ICalcul_CheckResultatCalcul m_NumPeriode, m_NumGroupe, m_Logger
  
  
  'EndTime = Timer
  'MsgBox "Finish Time: ", Now ' EndTime - StartTime
  
  'fin = Now
  'MsgBox "Time in seconds : " & DateDiff("s", debut, fin) & " s"

  Exit Sub
  
err_calcul:
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & m_rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & m_rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Code Situation
'
'##ModelId=5C8A75EC00C9
Private Sub CalculCodeSituation(ByRef rsAssure As ADODB.Recordset)

  If Not IsNull(rsAssure.fields("POSituConv")) Then
    Dim CodeSitu As Long
    
    CodeSitu = CLng(rsAssure.fields("POSituConv"))
    
    m_dataHelper.Multi_Find m_rsCDSITUAT, "Code_CIE=" & rsAssure.fields("POCompagnie") & " AND Code_APP=" & rsAssure.fields("POAppli") & " AND CDSITUAT=" & CodeSitu
    If Not m_rsCDSITUAT.EOF Then
      Dim TauxPM As Double
      
      TauxPM = m_rsCDSITUAT.fields("TauxPM")
      'If TauxPM <> 100 Then Stop
      
      rsAssure.fields("POPM") = rsAssure.fields("POPM") * (TauxPM / 100#)
      rsAssure.fields("POPSAP") = rsAssure.fields("POPSAP") * (TauxPM / 100#)
      rsAssure.fields("POPourcentLissage") = TauxPM
    End If
          
  End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Incapacités
'
'##ModelId=5C8A75EC0107
Private Sub CalculIncapacité(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean, ByRef bForceCoeffProvisionIncapToZero As Boolean)
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer
  Dim franchise As Integer
  Dim feuilleLue As String
  Dim feuilleLueIncap As String
  Dim feuilleLuePassage As String
  Dim feuilleLuePassageViager As String
  Dim feuilleLueInval As String
  Dim iIncap As Double
  Dim fraisGestionIncap As Double
  Dim iInval As Double
  Dim fraisGestionInval As Double
  Dim ajoutCmt As String
  Dim rq As String
  Dim DateLimiteProvision As Date
  
  '
  ' AJOUT VARIABLES DE CALCUL POUR L ANNUALISATION PASSAGE   01/06/2013
  '
  
  '    variables de la table DONNEES SOCIALES
  Dim SmicHoraireBrut As Double
  Dim CoefficientSmic As Double
  Dim TauxIJNettes As Double
  Dim PlafondAnnuelSS As Double
  Dim TauxRembSS As Double
  Dim TauxGarantieAssureur As Double
  Dim BoutonAnnulisationPassage As Boolean
  
  
  '    variables pour le calcul des butées
  Dim PlafondJournalier18SMIC As Double
  Dim PlafondJournalierTA As Double
  Dim PrestaInferieure As Double
  Dim PrestaSuperieure As Double
  Dim DifferencePassSmic As Double
  Dim CoeffCorrection As Double
  
 
  '    variables pour les annualisations Incap et Passage
  Dim Annualisation As Double
  Dim AnnualisationPassage As Double
 
  '    variable code provision utilisée pour les nouveaux codes provisions 23/08/2013
  Dim CodeProvision As Integer
  Dim bAvecIncap As Boolean
  Dim bAvecPassage As Boolean
  Dim bAvecInval As Boolean
  Dim isNewMethodBCACAffine As Boolean
  
  Dim ageLimitViagere As Integer
 ' variables réferome retraite 2023 le 07/08/2023
  Dim PTPassageHorsInval1 As Double
  Dim PTPassageInval1 As Double
  Dim DateLimiteProvisionInval1 As Date
  Dim DureePassageInval1 As Integer
  Dim xTermePassageInval1 As Integer
  Dim PT1PassageHorsInval1 As Double
  Dim PT1PassageInval1 As Double
  Dim TauxIncapPassageSpecifiqueInval1 As Double
            
  On Error GoTo err_calcul
  
    'If rsAssure.fields("NUENRP3I") = 50021 Then Stop
    'If rsAssure.fields("NUENRP3I") = 6 Then Stop
    
    ageLimitViagere = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeLimitViagere", "120"))

    ' chargement des situations possibles
    bAvecIncap = False
    bAvecPassage = False
    bAvecInval = False
    
    CodeProvision = rsAssure.fields("POSIT")
    Select Case CodeProvision
    
    ' cdPosit_Incap_AvecPassage = 1
    Case cdPosit_IncapAvecPassage
       bAvecIncap = True
       bAvecPassage = True
       bAvecInval = True
    
    ' cdPosit_IncapSansPassage = 3
    Case cdPosit_IncapSansPassage
       bAvecIncap = True
       bAvecPassage = False
       bAvecInval = False
    
    ' cdProvision_AvecIncap_AvecPassage_AvecInval = 1001
    Case cdProvision_AvecIncap_AvecPassage_AvecInval
       bAvecIncap = True
       bAvecPassage = True
       bAvecInval = True
    
    ' cdProvision_AvecIncap_AvecPassage_SansInval = 1002
    Case cdProvision_AvecIncap_AvecPassage_SansInval
       bAvecIncap = True
       bAvecPassage = True
       bAvecInval = False
    
    ' cdProvision_AvecIncap_SansPassage_AvecInval = 1003
    Case cdProvision_AvecIncap_SansPassage_AvecInval
       bAvecIncap = True
       bAvecPassage = False
       bAvecInval = True
    
    ' cdProvision_AvecIncap_SansPassage_SansInval = 1004
    Case cdProvision_AvecIncap_SansPassage_SansInval
       bAvecIncap = True
       bAvecPassage = False
       bAvecInval = False
    
    ' cdProvision_SansIncap_AvecPassage_AvecInval = 1005
    Case cdProvision_SansIncap_AvecPassage_AvecInval
       bAvecIncap = False
       bAvecPassage = True
       bAvecInval = True
    
    ' cdProvision_SansIncap_AvecPassage_SansInval = 1006
    Case cdProvision_SansIncap_AvecPassage_SansInval
       bAvecIncap = False
       bAvecPassage = True
       bAvecInval = False
    
    ' cdProvision_SansIncap_SansPassage_AvecInval = 1007
    Case cdProvision_SansIncap_SansPassage_AvecInval
       bAvecIncap = False
       bAvecPassage = False
       bAvecInval = True
    
    ' cdProvision_SansIncap_SansPassage_AvecInval = 1008
    Case cdProvision_SansIncap_SansPassage_SansInval
       bAvecIncap = False
       bAvecPassage = False
       bAvecInval = False
       
    ' cdProvision_Incap_AvecPassage_Viager = 6001
    Case cdProvision_Incap_AvecPassage_Viager
       bAvecIncap = True
       bAvecPassage = True
       bAvecInval = False
    
    End Select
    
     
' chargement initial par défaut de l'annualisation passage = annualisation
  Annualisation = rsAssure.fields("POPRESTATION_AN")
  AnnualisationPassage = rsAssure.fields("POPRESTATION_AN")  '= annualisation
'
' calcul de l'annaulsation PASSAGE en fonction du bouton Check dans la période
 BoutonAnnulisationPassage = m_rsPeriode.fields("PEANNUALISATION_PASSAGE")
 If BoutonAnnulisationPassage = True Then   ' on recalcule l'annualisation réduite pour le PASSAGE et il existe des données sociales
     
     '  CHARGEMENT DES VARIABLES depuis la table DONNEES SOCIALES en fonction de la date de survenance de la presta Incap
      If FindDonneesSociales(rsAssure.fields("POARRET")) Then
        SmicHoraireBrut = m_rsDonneesSociales.fields("SmicHoraireBrut")
        CoefficientSmic = m_rsDonneesSociales.fields("CoefficientSmic")
        TauxIJNettes = m_rsDonneesSociales.fields("TauxIJNettes")
        PlafondAnnuelSS = m_rsDonneesSociales.fields("PlafondAnnuelSS")
        TauxRembSS = m_rsDonneesSociales.fields("TauxRembSS")
        TauxGarantieAssureur = m_rsDonneesSociales.fields("TauxGarantieAssureur")
        
        '  CALCUL DES BUTEES
        PlafondJournalier18SMIC = Arrondi(SmicHoraireBrut * CoefficientSmic * 35 / 7, 2)
        PlafondJournalierTA = Arrondi(PlafondAnnuelSS / 365.25, 2)
        
        PrestaInferieure = Arrondi((TauxGarantieAssureur * TauxIJNettes - TauxRembSS) * PlafondJournalier18SMIC * 365.25, 2)
        If PrestaInferieure < 0 Then
          PrestaInferieure = 0#
        End If
        
        PrestaSuperieure = Arrondi((TauxGarantieAssureur * TauxIJNettes * PlafondJournalierTA - TauxRembSS * PlafondJournalier18SMIC) * 365.25, 2)
        If PrestaSuperieure < 0 Then
          PrestaSuperieure = 0#
        End If
         
        DifferencePassSmic = Arrondi((PlafondJournalierTA - PlafondJournalier18SMIC) * TauxRembSS * 365.25, 2)
        If DifferencePassSmic < 0 Then
          DifferencePassSmic = 0#
        End If
        
         ' CALCUL du CoeffCorrection
          If Annualisation <= PrestaInferieure Then
             CoeffCorrection = 0#
          Else
               If Annualisation <= PrestaSuperieure Then
                 
                 If PrestaSuperieure - PrestaInferieure <> 0 Then
                   CoeffCorrection = (Annualisation - PrestaInferieure) / (PrestaSuperieure - PrestaInferieure)
                   
                   If CoeffCorrection < 0 Then
                   CoeffCorrection = 0
                   End If
                 
                 Else
                   CoeffCorrection = 0
                 End If
               Else  ' cas où Annualisation > PrestaSuperieure
                 CoeffCorrection = 1#
               End If
          End If
             
         ' CALCUL Annualisation de Passage
          AnnualisationPassage = Arrondi(Annualisation - DifferencePassSmic * CoeffCorrection, 2)
          If AnnualisationPassage < 0 Then
           AnnualisationPassage = 0#
          End If
      Else ' il n'y a pas de données sociales pour la période  (annualisaation passage= annualisation)
        SmicHoraireBrut = 0#
        CoefficientSmic = 0#
        TauxIJNettes = 0#
        PlafondAnnuelSS = 0#
        TauxRembSS = 0#
        TauxGarantieAssureur = 0#
      End If
 End If
  
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' annualisation AnnualisationPassage forcées à 0 en fonction des situations
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     'Annualisation
     If bAvecIncap = False Then
     Annualisation = 0
     End If
     
     'AnnualisationPassage
     If bAvecPassage = True Then
     rsAssure.fields("POPRESTATION_AN_PASSAGE") = AnnualisationPassage
     Else
     AnnualisationPassage = 0
     rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0
  
     End If
  
  ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
     
   ' paramètres de la fonction FctCalcProvInval
   'age x            = âge à la date d'entrée en incapacité
   'anc              = ancienneté dans l'état d'incapacité, en mois
   'duree            = entre date de l'ARRET et date TERME du contrat, en mois
   'franchise        = dans l'état d'incapacité, , en mois
   'feuilleLue       = loi de maintien en incapacité
   'taux technique     i
   'frais de gestion = fraisGestion
   
   
   ' Paramêtres supplémentaires nécessaires à la fonction fctCalcProvInvalRevalo
   ' m_TMO         = Taux d'indexation
   ' m_DureRevalo   = Durée d'indexation en année
   ' m_TauxRevalo  = Taux de revalorisation
   
   
   'If rsAssure.fields("NUENRP3I") = 1 Then Stop
   
   x = m_xArret
  
   If m_BridageAge = True Then
     x = Maximum(x, m_AgeMin_Incap)
     x = Minimum(x, m_AgeMax_Incap)
   End If
   
   ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
   'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
   anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
   If anc < 0 Then
     anc = 0
   End If

   If anc >= 36 Then
    'PHM 15/01/2004
     ' Assuré en arret depuis plus de 36 mois (3 ans)
     ' test si passage coché dans CATR9
     ' Si oui, on mets le gars en INVAL (m_Posit=2) et on saute au calcul m_Posit=2
     ' Si non, on force à 0

     If PassageCATR9(rsAssure.fields("POCATEGORIE"), rsAssure.fields("POCONVENTION")) Or (bAvecInval = True) Then
       m_Posit = cdPosit_Inval ' force le passage en invalidite
       rsAssure.fields("PODATEENTREEINVAL") = DateSerial(Year(rsAssure.fields("POARRET")), _
                                                   Month(rsAssure.fields("POARRET")) + 36, _
                                                   Day(rsAssure.fields("POARRET")))
       
       rsAssure.fields("POCategorieInval") = cdCategorieInvalParDefaut
       
       ajoutCmt = "Assuré en incapacité ancienneté= " & anc & " mois position= " & rsAssure.fields("POSIT") & " considéré comme invalide (Catégorie " & cdCategorieInvalParDefaut & ")"
       If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
         rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- " & ajoutCmt
       End If
     
     Else
       ajoutCmt = "PM annulée car assuré invalide en incapacité de durée>36 mois sans passage"
       If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
         rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- " & ajoutCmt
       End If
       
       ' on force les coeff de provisions incapacite à 0
       ForceCoeffProvisionIncapToZero rsAssure
'       rsAssure.fields("POPM_INVAL_1F") = 0
'       rsAssure.fields("POPM_INCAP_1F") = 0
'       rsAssure.fields("POPM_PASS_1F") = 0
'       rsAssure.fields("POPM") = 0
'       rsAssure.fields("POPM_REVALO") = 0
'       rsAssure.fields("POPM_SORTIE") = 0
'       rsAssure.fields("POPM_VAR") = 0
'       rsAssure.fields("POPM_INVAL_1R") = 0
'       rsAssure.fields("POPM_INCAP_1R") = 0
'       rsAssure.fields("POPM_PASS_1R") = 0
'       rsAssure.fields("POPM_REVALO") = 0
'
'       rsAssure.fields("POPM_x") = 0
'       rsAssure.fields("POPM_xTerme") = 0
'       rsAssure.fields("POPM_anc") = 0
'       rsAssure.fields("POPM_Duree") = 0
'
'       rsAssure.fields("POCorrectif") = 0
'       rsAssure.fields("TXFRAIS") = 0
       
       
    
     End If
'PHM 15/01/2004
   
   Else ' anc < 36 mois
     
     ' franchise en mois entre la date D'ARRET et la date d'EFFET du contrat
     'franchise = DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
     franchise = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
     If franchise < 0 Then
       franchise = 0
     End If
     
     ' franchise en mois ajout du délais de carence
     franchise = franchise + CInt(rsAssure.fields("PODELAI") / 30)
     If franchise < 0 Then
       franchise = 0
     End If
     
     
     If bForReformeRetraite And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
       
       ' Date limite précisée dans l'import
       '&&&&
       'rsAssure.fields("POTERME") = DateLimiteProvision
       DateLimiteProvision = ComputeDateLimiteApresReforme(rsAssure, rsAssure.fields("POCATEGORIE_NEW"), rsAssure.fields("POTERME"), rsAssure.fields("PONAIS"), rsAssure.fields("POARRET"), rsAssure.fields("NUENRP3I"))
       rsAssure.fields("POTERME") = DateLimiteProvision
       
       If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
         DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + ageLimitViagere, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
         rsAssure.fields("POTERME") = DateLimiteProvision  ' âge terme est forcé à 120 ans
       End If
       
       
       
       ' durée en mois entre date d'ARRET et date TERME du contrat
       Duree = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME"))
       If Duree < 0 Then
         Duree = 0
       End If
       
       
       ' âge à la date du terme du contrat ( = âge de départ en retraite)
       m_xTerme = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvision)
       m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), DateLimiteProvision)
       
       
       ' tables
       feuilleLue = m_rsPeriode.fields("PELMINCAP_Retraite")
       feuilleLueIncap = m_rsPeriode.fields("PELMINCAP_Retraite")
       feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE_Retraite")
       feuilleLueInval = m_rsPeriode.fields("PELMINVAL_Retraite")
       feuilleLuePassageViager = m_rsParamCalcul.fields("PELMINVAL_Viagere")
       
       'Modif: 11/07/2018
       If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
        'feuilleLuePassage = feuilleLuePassageViager
       End If
     
       ' Amortissement
       If rsAssure.fields("POTopAmortissable") = True Then
         rsAssure.fields("POCoeffAmortissement") = FindCoeffAmortissement(m_rsPeriode.fields("PEDATEEXT"), rsAssure.fields("PONAIS"))
       Else
         rsAssure.fields("POCoeffAmortissement") = 1
       End If
     
     Else
     
       If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
         DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + ageLimitViagere, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
         rsAssure.fields("POTERME") = DateLimiteProvision  ' âge terme est forcé à 120 ans
         m_xTerme = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvision)
         m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), DateLimiteProvision)

       End If

       
       ' durée en mois entre date d'ARRET et date TERME du contrat
       Duree = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME"))
       If Duree < 0 Then
         Duree = 0
       End If
      
       feuilleLue = m_rsPeriode.fields("PELMINCAP")
       feuilleLueIncap = m_rsPeriode.fields("PELMINCAP")
       feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE")
       feuilleLueInval = m_rsPeriode.fields("PELMINVAL")
       feuilleLuePassageViager = m_rsParamCalcul.fields("PELMINVAL_Viagere")
       
       
     End If
     
     '### Passage Avec Table Inval Viagère
     If CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
      If Not IsNull(m_rsPeriode.fields("PELMINVAL_Viagere")) Then
        feuilleLueInval = m_rsPeriode.fields("PELMINVAL_Viagere")
      End If
     End If
     
    
     
     
     iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
     fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100
     
     rsAssure.fields("TXTECHN") = iIncap * 100#
     rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
     
     iInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINVAL")) / 100
     fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINVAL")) / 100
     
     If bAvecIncap = True Then  'incap
      rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvIncap(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
     Else
      rsAssure.fields("POPM_INCAP_1F") = 0
     End If
     
     If bAvecPassage = True Then  'passage
     ' modification réforme retraite 2023 , le 07/08/2023
        If rsAssure.fields("POCATEGORIE_NEW") = 2 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then 'Incap Non Professionnelle
            'mode hors inval 1
            PTPassageHorsInval1 = m_ModuleCalcul_Provisions.FctCalcProvPassage(x, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE")) 'fraisGestionInval)
            
            'mode inval 1 nouvelle table spécifique inval 1 à 64 ans
            DateLimiteProvisionInval1 = ComputeDateLimiteApresReformeInval1(rsAssure, rsAssure.fields("POCATEGORIE_NEW"), rsAssure.fields("POTERME"), rsAssure.fields("PONAIS"), rsAssure.fields("POARRET"), rsAssure.fields("NUENRP3I"), True)
                   
            ' durée en mois entre date d'ARRET et date TERME du contrat
            DureePassageInval1 = NbMois(rsAssure.fields("POARRET"), DateLimiteProvisionInval1)
            If DureePassageInval1 < 0 Then
              DureePassageInval1 = 0
            End If
       
            xTermePassageInval1 = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvisionInval1)
            PTPassageInval1 = m_ModuleCalcul_Provisions.FctCalcProvPassage(x, xTermePassageInval1, anc, DureePassageInval1, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))  'fraisGestionInval)
            TauxIncapPassageSpecifiqueInval1 = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAPINVAL1")) / 100
            rsAssure.fields("POPM_PASS_1F") = Arrondi((1 - TauxIncapPassageSpecifiqueInval1) * PTPassageHorsInval1 + TauxIncapPassageSpecifiqueInval1 * PTPassageInval1, NbDecimalePM)
        
        Else ' cas hors (incap non professionnelle=2 et portefeuille salarié=true)
            rsAssure.fields("POPM_PASS_1F") = m_ModuleCalcul_Provisions.FctCalcProvPassage(x, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE")) 'fraisGestionInval)
        End If
     Else
      rsAssure.fields("POPM_PASS_1F") = 0
     End If
     
     ' stockage des paramètres pour l'incap / passage
     
     rsAssure.fields("POPM_x") = x
     rsAssure.fields("POPM_xTerme") = IIf(bForReformeRetraite, m_xTermeDouble, m_xTerme)
     rsAssure.fields("POPM_anc") = anc
     rsAssure.fields("POPM_Duree") = Duree
     
     If m_bRevalo Then
     
       If bAvecIncap = True Then  'incap
         rsAssure.fields("POPM_INCAP_1R") = m_ModuleCalcul_Provisions_Revalo.FctCalcProvIncapRevalo(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap, m_DureRevalo, m_TauxRevalo)
       Else
        rsAssure.fields("POPM_INCAP_1R") = 0
       End If
       
       If bAvecPassage = True Then  'passage
        If rsAssure.fields("POCATEGORIE_NEW") = 2 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then 'Incap Non Professionnelle
            'mode hors inval 1
            PTPassageHorsInval1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
            PTPassageInval1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x, xTermePassageInval1, anc, DureePassageInval1, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
            TauxIncapPassageSpecifiqueInval1 = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAPINVAL1")) / 100
            rsAssure.fields("POPM_PASS_1R") = Arrondi((1 - TauxIncapPassageSpecifiqueInval1) * PTPassageHorsInval1 + TauxIncapPassageSpecifiqueInval1 * PTPassageInval1, NbDecimalePM)
        Else
            rsAssure.fields("POPM_PASS_1R") = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
        End If
       Else
        rsAssure.fields("POPM_PASS_1R") = 0
       End If
     
     Else
      rsAssure.fields("POPM_INCAP_1R") = 0
      rsAssure.fields("POPM_PASS_1R") = 0
      rsAssure.fields("POPM_INVAL_1R") = 0
      rsAssure.fields("POPM_REVALO") = 0
      rsAssure.fields("POCOT_REVALO") = 0
     End If

#If P3I_GENERALI = 1 Then
' PHM 01/06/2006 : Pour le calcul du coefficient à l'âge x années et k mois, une interpolation entre les âges x et x+1 est effectuée.
    Dim PT As Double, PT1 As Double, k As Double, anc_regl As Integer, correctif As Double
    Dim POFIN As Variant ' Date
    
    POFIN = rsAssure.fields("POFIN")
    
    If Not IsNull(m_rsPeriode.fields("IsBCACNew")) Then
      isNewMethodBCACAffine = IIf(m_rsPeriode.fields("IsBCACNew") = True, True, False)
    End If

' PHM 16/11/2009
    ' Doit-on utiliser la PM avec ou sans correctif
    ' cas anc methode avec table PM_RETENUE
    
     ' RS AM 31/01/2018
     '### BCAC Affiné   NOUVELLE METHODE
     Dim nouvelleMethod As Boolean
     nouvelleMethod = False
     If isNewMethodBCACAffine = True And bAvecIncap = True And bAvecPassage = True And (m_IDTypeCalcul = 1 Or m_IDTypeCalcul = 3) Then
      nouvelleMethod = True
     End If
    
    If nouvelleMethod = False Then
      m_dataHelper.Multi_Find m_rsPMRetenue, "AnneeSurvenance=" & Year(rsAssure.fields("POARRET"))
        
      If Not m_rsPMRetenue.EOF Then
        If m_rsPMRetenue.fields("PMAvecCorrectif") = True Then
          If IsNull(rsAssure.fields("POFIN")) Then
            POFIN = m_DateCalcul ' date de calcul dans ce cas
          End If
        End If
      End If
    Else
      'nouvelle methode mois glissante
      POFIN = m_DateCalcul
    End If
    
' PHM 16/11/2009
    
    
    If Not IsNull(POFIN) Then
      ' Nombre de mois entre les 2 dates (date de fin de periode de la dernière période (ordre=PeriodeAu), date de calcul)
      'anc_regl = DateDiff("m", rsAssure.fields("PODERNIERPAIEMENT"), m_DateCalcul)
      anc_regl = DateDiff("m", rsAssure.fields("POARRET"), POFIN)
      If anc_regl < 0 Then
        anc_regl = 0
      End If
      
     'If rsAssure.fields("NUENRP3I") = 17087 Then Stop
     
     
     correctif = Arrondi(m_ModuleCalcul_Provisions.FctCalcCorrectifIncapacite(x, anc, anc_regl, feuilleLue, feuilleLuePassage), NbDecimalePM)
     
     ' calcul mois glissant dans la nouvelle méthode
     '### BCAC Affiné
      
     If nouvelleMethod = True Then
       Dim dureeMois As Integer
       Dim moisGlissant As Integer
       Dim dateSinistre As Date
       
       dureeMois = 0
       moisGlissant = 0
       dateSinistre = m_DateCalcul
             
       If Not IsNull(m_rsPeriode.fields("NumberOfMonth")) Then
         moisGlissant = m_rsPeriode.fields("NumberOfMonth")
       End If
      
       If Not IsNull(rsAssure.fields("POARRET")) Then
         dateSinistre = rsAssure.fields("POARRET")
       End If
      
       dureeMois = NbMois(dateSinistre, m_DateCalcul)
       If dureeMois < 0 Then
         dureeMois = 0
       End If
      
       'Don't apply correctif
       If isNewMethodBCACAffine = True And dureeMois > moisGlissant Then
        correctif = 1
       End If
     
     End If

      
      'k = m_xArretMois - Int(m_xArretMois / 12) * 12
      k = Arrondi(NbMois(rsAssure.fields("PONAIS"), rsAssure.fields("POARRET")) - x * 12, NbDecimalePM)
      If k < 0 Then
        k = 0
      End If
      
''''''''**** PHM : Ne pas modifier les coeff mais stocker le correctif
    If bAvecIncap = True Then 'incap
      PT = rsAssure.fields("POPM_INCAP_1F")
      PT1 = m_ModuleCalcul_Provisions.FctCalcProvIncap(x + 1, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
      rsAssure.fields("POPM_INCAP_1F") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
    Else
      rsAssure.fields("POPM_INCAP_1F") = 0
    End If
      
      
    If bAvecPassage = True Then 'passage
      PT = rsAssure.fields("POPM_PASS_1F")
      ' modification réforme retraite 2023 le 07/08/2023
        If rsAssure.fields("POCATEGORIE_NEW") = 2 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then 'Incap Non Professionnelle
        
            PT1PassageHorsInval1 = m_ModuleCalcul_Provisions.FctCalcProvPassage(x + 1, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE")) 'fraisGestionInval)
            PT1PassageInval1 = m_ModuleCalcul_Provisions.FctCalcProvPassage(x + 1, xTermePassageInval1, anc, DureePassageInval1, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE")) 'fraisGestionInval)
            PT1 = Arrondi((1 - TauxIncapPassageSpecifiqueInval1) * PT1PassageHorsInval1 + TauxIncapPassageSpecifiqueInval1 * PT1PassageInval1, NbDecimalePM)
        Else
        
            PT1 = m_ModuleCalcul_Provisions.FctCalcProvPassage(x + 1, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE")) 'fraisGestionInval)
        
        End If
      rsAssure.fields("POPM_PASS_1F") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
    Else
      rsAssure.fields("POPM_PASS_1F") = 0
    End If
    
      If m_bRevalo Then
      
        If bAvecIncap = True Then 'incap
          PT = rsAssure.fields("POPM_INCAP_1R")
          PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvIncapRevalo(x + 1, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap, m_DureRevalo, m_TauxRevalo)
          rsAssure.fields("POPM_INCAP_1R") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
        Else
          rsAssure.fields("POPM_INCAP_1R") = 0
        End If
        
        If bAvecPassage = True Then 'passage
          PT = rsAssure.fields("POPM_PASS_1R")
          ' modification réforme retraite 2023 le 07/08/2023
          If rsAssure.fields("POCATEGORIE_NEW") = 2 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then 'Incap Non Professionnelle
          
              PT1PassageHorsInval1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x + 1, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
              PT1PassageInval1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x + 1, xTermePassageInval1, anc, DureePassageInval1, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
              PT1 = Arrondi((1 - TauxIncapPassageSpecifiqueInval1) * PTPassageHorsInval1 + TauxIncapPassageSpecifiqueInval1 * PTPassageInval1, NbDecimalePM)
          Else
              PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvPassageRevalo(x + 1, m_xTerme, anc, Duree, franchise, feuilleLueIncap, feuilleLuePassage, feuilleLueInval, iIncap, fraisGestionIncap, iInval, 0, m_DureRevalo, m_TauxRevalo, m_rsPeriode.fields("PESANSLISSAGE_PASSAGE"))
          End If
          rsAssure.fields("POPM_PASS_1R") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
        Else
          rsAssure.fields("POPM_PASS_1R") = 0
        End If
        
      End If
    End If
' PHM 01/06/2006
#End If

'*** Modif ISICA  de mars 1999 : champs booleen dans la table garantie pour la calcul du passage
' PHM 25/09/2004      If m_Posit = 3 Or CBool(m_dataHelper.GetParameter("SELECT GACALPASS FROM Garantie WHERE GAGARCLE=" & rsAssure.Fields("POGARCLE"))) = True Then
     
     'Modif 11/7/2018
     If m_Posit <> cdProvision_Incap_AvecPassage_Viager Then
      'If m_Posit = cdPosit_IncapSansPassage Or FindGarantie(rsAssure.fields("POGARCLE")).CalculPassage = True Then
      ' Modif AM RS 02/05/2019
      If m_Posit = cdPosit_IncapSansPassage Or FindGarantie(rsAssure.fields("POGARCLE")).CalculPassage = False Then
        rsAssure.fields("POPM_PASS_1F") = 0
        rsAssure.fields("POPM_PASS_1R") = 0
      End If
     End If
'*** Modif ISICA
     
#If P2I_UPESE = 1 Then
'*** Modif UPESE  de nov 1999
     If rsAssure.fields("POGARCLE") = cdGar14 Or rsAssure.fields("POGARCLE") = cdGar15 Then
       rsAssure.fields("POPM_PASS_1F") = 0
       rsAssure.fields("POPM_PASS_1R") = 0
     End If
     
     ' coeff de passage ajuste a 300/365 (modif du montant de la rente)
     rsAssure.fields("POPM_PASS_1F") = Arrondi(rsAssure.fields("POPM_PASS_1F") * 300 / 365, NbDecimalePM)
     If m_bRevalo Then
       rsAssure.fields("POPM_PASS_1R") = Arrondi(rsAssure.fields("POPM_PASS_1R") * 300 / 365, NbDecimalePM)
     End If
     
'*** Modif UPESE
#End If

     rsAssure.fields("POPM_INVAL_1F") = 0
     rsAssure.fields("POPM_INVAL_1R") = 0
     
     '
     ' ancien calcul avant le 01/06/2013
     '
     'rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
     'rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
     'rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
     '
     'If m_bRevalo Then
     '  rsAssure.fields("POPM_REVALO") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1R") / 12) + rsAssure.fields("POPM_PASS_1R")) - rsAssure.fields("POPM")
     '  If rsAssure.fields("POPM_REVALO") < 0 Then
     '    rsAssure.fields("POPM_REVALO") = 0
     '  End If
     '  rsAssure.fields("POCOT_REVALO") = rsAssure.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO))
     'End If
  
  
     ' nouveau calcul APRES le 01/06/2013
    
     rsAssure.fields("POPM") = Arrondi(Annualisation * rsAssure.fields("POPM_INCAP_1F") / 12 + AnnualisationPassage * rsAssure.fields("POPM_PASS_1F"), 2)
     rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPM")
     
     ' attention variation uniquement avec l'annualisation de l'Incap (et non l'annualisation passage)
     rsAssure.fields("POPM_VAR") = Arrondi((Annualisation - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F")), 2)
          
     If m_bRevalo Then
       rsAssure.fields("POPM_REVALO") = Annualisation * rsAssure.fields("POPM_INCAP_1R") / 12 + AnnualisationPassage * rsAssure.fields("POPM_PASS_1R") - rsAssure.fields("POPM")
       If rsAssure.fields("POPM_REVALO") < 0 Then
         rsAssure.fields("POPM_REVALO") = 0
       End If
       rsAssure.fields("POCOT_REVALO") = Arrondi(rsAssure.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO)), 2)
     End If
 
     '
     '
    If IsNull(rsAssure.fields("POPRESTATION_AN")) Then
    
      rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0
      rsAssure.fields("POPM") = 0
      rsAssure.fields("POPM_REVALO") = 0
      rsAssure.fields("POCOT_REVALO") = 0
      rsAssure.fields("POPM_SORTIE") = 0
      rsAssure.fields("POPM_VAR") = 0
    End If
  
#If P3I_GENERALI = 1 Then
' PHM 10/07/2006
    ' Doit-on utiliser la PM avec ou sans correctif
' PHM 16/11/2009
    ' Voir plus haute : m_dataHelper.Multi_Find m_rsPMRetenue, "AnneeSurvenance=" & Year(rsAssure.fields("POARRET"))
' PHM 16/11/2009
      
    ' chargement des flags correctifs
    If (nouvelleMethod = False And Not m_rsPMRetenue.EOF) Then
        rsAssure.fields("POPMAvecCorrectif") = m_rsPMRetenue.fields("PMAvecCorrectif")
        rsAssure.fields("POPMReassAvecCorrectif") = m_rsPMRetenue.fields("PMReassAvecCorrectif")
    End If
      
    If (nouvelleMethod = True And dureeMois <= moisGlissant) Then
        rsAssure.fields("POPMAvecCorrectif") = True
        rsAssure.fields("POPMReassAvecCorrectif") = True
    End If

      
    ' RS AM 31/01/2018
    ' ANCIENNE méthode avec la table année de survenance
    ' OU NOUVELLE méthode avec mois glissant <= moisglissant

    ' calcul POPM avec et sans correctif
    If (nouvelleMethod = False And Not m_rsPMRetenue.EOF) Or (nouvelleMethod = True And dureeMois <= moisGlissant) Then
      rsAssure.fields("POPM_AvecCorrectif") = Arrondi(rsAssure.fields("POPM") * correctif, 2)
      rsAssure.fields("POPM_SansCorrectif") = Arrondi(rsAssure.fields("POPM"), 2)
      rsAssure.fields("POCorrectif") = correctif
      
      If (nouvelleMethod = False And m_rsPMRetenue.fields("PMAvecCorrectif") = True) Or (nouvelleMethod = True And dureeMois <= moisGlissant) Then
        rsAssure.fields("POPM") = rsAssure.fields("POPM_AvecCorrectif")
        rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE") * correctif, 2)
        rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR") * correctif, 2)
         
        If m_bRevalo Then
          rsAssure.fields("POPM_REVALO") = Arrondi(rsAssure.fields("POPM_REVALO") * correctif, 2)
          rsAssure.fields("POCOT_REVALO") = Arrondi(rsAssure.fields("POCOT_REVALO") * correctif, 2)
        End If
      End If
    End If
' PHM 10/07/2006
#End If

    rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
    rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
    rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
    If m_bRevalo Then
      rsAssure.fields("POPM_REVALO") = Arrondi(rsAssure.fields("POPM_REVALO"), 2)
      rsAssure.fields("POCOT_REVALO") = Arrondi(rsAssure.fields("POCOT_REVALO"), 2)
    End If
    
'*** Modif UPESE 27/11/1999
#If P2I_UPESE = 1 Then
    ' recherche si l'assure est egalement en invalidite
'        rq = "SELECT 1count(*) FROM Assure Assure INNER JOIN Garantie Garantie ON Assure.POGARCLE = Garantie.GAGARCLE " _
'             & " WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode _
'             & " AND Assure.PONUMCLE =""" & rsAssure.Fields("PONUMCLE") & """ AND Garantie.GATYPEGAR=" & cdTypeInvalidite
     rq = "SELECT 1 FROM Assure Assure INNER JOIN Garantie Garantie ON Assure.POGARCLE = Garantie.GAGARCLE " _
          & " WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode _
          & " AND Assure.PONUMCLE =""" & rsAssure.fields("PONUMCLE") & """ AND Garantie.GATYPEGAR=" & cdTypeInvalidite
     If m_dataHelper.GetParameterAsDouble(rq) > 0 Then
       ' sauvegarde les valeurs calculées
       rsAssure.fields("POPMANNULEE") = rsAssure.fields("POPM")
       If IsNull(rsAssure.fields("POCOMMENTANNUL")) Then
         rsAssure.fields("POCOMMENTANNUL") = " "
       Else
         If Left(rsAssure.fields("POCOMMENTANNUL"), 8) = "Import: " Then
           rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- "
         Else
           rsAssure.fields("POCOMMENTANNUL") = " "
         End If
       End If
       rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & "PM annulée car l'assuré est aussi en invalidité"
       
       ' on force les coeff de provisions incapacite à 0
       rsAssure.fields("POPM_INVAL_1F") = 0
       rsAssure.fields("POPM_INCAP_1F") = 0
       rsAssure.fields("POPM_PASS_1F") = 0
       rsAssure.fields("POPM") = 0
       
       If m_bRevalo Then rsAssure.fields("POPM_REVALO") = 0
       
       rsAssure.fields("POPM_SORTIE") = 0
       rsAssure.fields("POPM_VAR") = 0
       rsAssure.fields("POPM_INVAL_1R") = 0
       rsAssure.fields("POPM_INCAP_1R") = 0
       rsAssure.fields("POPM_PASS_1R") = 0
       rsAssure.fields("POPM_REVALO") = 0
     End If
#End If
'*** modif UPESE

' PHM 05/02/2004
     ' Assuré en arret depuis moins de 36 mois (3 ans)
     ' avec une date d'inval renseignée
     ' test si passage coché dans CATR9
     ' Si oui, le gars a ete mis en INVAL (m_Posit=2) dans l'import
     ' Si non, on force à 0
     If Not IsNull(rsAssure.fields("PODATEENTREEINVAL")) Then
       If Not PassageCATR9(rsAssure.fields("POCATEGORIE"), rsAssure.fields("POCONVENTION")) Then
         ajoutCmt = "PM annulée car assuré invalide en incapacité de durée<36 mois sans passage"
         If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
           rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & ajoutCmt
         End If
         
         ' on force les coeff de provisions incapacite à 0
         rsAssure.fields("POPM_INVAL_1F") = 0
         rsAssure.fields("POPM_INCAP_1F") = 0
         rsAssure.fields("POPM_PASS_1F") = 0
         rsAssure.fields("POPM") = 0
         rsAssure.fields("POPM_REVALO") = 0
         rsAssure.fields("POPM_SORTIE") = 0
         rsAssure.fields("POPM_VAR") = 0
         rsAssure.fields("POPM_INVAL_1R") = 0
         rsAssure.fields("POPM_INCAP_1R") = 0
         rsAssure.fields("POPM_PASS_1R") = 0
         rsAssure.fields("POPM_REVALO") = 0
       Else
         ajoutCmt = "Les tables CATR9 ou PassageNCA ont été modifiée : VOUS DEVEZ RELANCER L'IMPORT !"
         m_Logger.EcritTraceDansLog ajoutCmt
         If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
           rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & ajoutCmt
         End If
       End If
     End If
' PHM 05/02/2004
   End If ' anc >=36
   
   
   
  'New Functionality : Selection encours Collecteam : August 2017
  
  'this only applies for Position: 1 - Incapacité avec Passage
  'test also for : Type Periode : BILAN " ou " SIMUL
  'If CodeProvision = cdPosit_IncapAvecPassage And (m_IDTypeCalcul = 1 Or m_IDTypeCalcul = 3) Then
  If bAvecIncap = True And bAvecPassage = True And (m_IDTypeCalcul = 1 Or m_IDTypeCalcul = 3) Then
  
    Dim duree_Selection As Integer
    Dim nbj_Selection As Integer
    Dim dateFinPer As Date
    Dim codeCatContrat As String
    
    duree_Selection = 0
    nbj_Selection = -1
    dateFinPer = m_DateCalcul
    
    If Not IsNull(rsAssure.fields("POFIN")) Then
      dateFinPer = rsAssure.fields("POFIN")
    End If
    
    If Not IsNull(rsAssure.fields("POCATEGORIE")) Then
      codeCatContrat = rsAssure.fields("POCATEGORIE")
    End If
  
    m_rsCodesCat.MoveFirst
    m_rsCodesCat.filter = " Code_Cat_Contrat = '" & codeCatContrat & "'"
    If Not m_rsCodesCat.EOF Then
      If Not IsNull(m_rsCodesCat.fields("Nbj_Selection")) Then
        nbj_Selection = m_rsCodesCat.fields("Nbj_Selection")
      End If
    End If
    
    If nbj_Selection = -1 Then
      Exit Sub
    End If
  
    duree_Selection = DateDiff("d", dateFinPer, m_DateCalcul)
    
    If duree_Selection < 0 Then
      duree_Selection = 0
    End If
    
    'Force coeff Provision to 0
    
    '''''TEST AM 02/02/2018
'    If nbj_Selection = 200 Then
'    nbj_Selection = 200
'    End If
''
'    If rsAssure.fields("NUENRP3I") = 14639 Then
'    Stop
'    End If
'
'     If rsAssure.fields("NUENRP3I") = 15993 Then
'    Stop
'    End If
'
'     If rsAssure.fields("NUENRP3I") = 42444 Then
'    Stop
'    End If
'
'    If rsAssure.fields("NUENRP3I") = 42515 Then
'    Stop
'    End If
'
   '''''
    
    bForceCoeffProvisionIncapToZero = False
    
    If duree_Selection > 0 And duree_Selection > nbj_Selection And nbj_Selection > 0 Then
      ForceCoeffProvisionIncapToZero rsAssure
      bForceCoeffProvisionIncapToZero = True ' les coeff prov et PSAP sont forcés à 0 -> PAS de calcul de la PSAP (fonction CalculPSAP non appelée)
    End If
  
  End If
  
  Exit Sub
  
err_calcul:

  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
    
  Resume Next
End Sub

'##ModelId=5C8A75EC0165
Private Sub ForceCoeffProvisionIncapToZero(ByRef rsAssure As ADODB.Recordset)

  ' on force les coeff de provisions incapacite à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
       
    rsAssure.fields("POPM_x") = 0
    rsAssure.fields("POPM_xTerme") = 0
    rsAssure.fields("POPM_anc") = 0
    rsAssure.fields("POPM_Duree") = 0
    
    rsAssure.fields("POCorrectif") = 0
    rsAssure.fields("TXFRAIS") = 0
    
    'PSAP
    rsAssure.fields("POPSAP") = 0
    rsAssure.fields("POPSAP_RASSUR") = 0
    rsAssure.fields("POPSAPCAPMOYEN") = 0
    rsAssure.fields("POPSAPANNULEE") = 0
       
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Invalidités
'
'##ModelId=5C8A75EC01A3
Private Sub CalculInvalidité(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer, AgeLimiteInvalCat1 As Integer
  Dim franchise As Integer
  Dim feuilleLue As String
  Dim iInval As Double
  Dim fraisGestionInval As Double
  Dim DateMiseInval As Date
  Dim DateLimiteProvision As Date
  Dim AgeMax As Integer
  Dim dat As Date
  Dim CodeProvision As Integer
  Dim ageLimitViagere As Integer
    
  
  On Error GoTo err_calcul
  
   ageLimitViagere = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "AgeLimitViagere", "120"))
   CodeProvision = rsAssure.fields("POSIT")
   
   If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
    DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + ageLimitViagere, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
    rsAssure.fields("POTERME") = DateLimiteProvision  ' âge terme est forcé à 120 ans
   End If
  
   '
   'If rsAssure.fields("NUENRP3I") = 4605 Then Stop
   '
  
  
  ' l'assuré est INVALIDE, utilisation de la fonction FctCalcProvInval
  ' paramètres de la fonction FctCalcProvInval
  ' age x            = âge à la date d'entrée en invalidité
  ' anc              = ancienneté dans l'état d'invalidité
  ' duree            = entre âge au terme du contrat - âge à la date de l'ARRET
  ' franchise        = dans l'état d'invalidité
  ' feuilleLueInval  = loi de maintien en Invalidité
  ' taux technique     i
  ' frais de gestion = fraisGestionInval

  ' Mise à jour de la date d'entrée en invalidité
  If IsNull(rsAssure.fields("PODATEENTREEINVAL")) Then
    rsAssure.fields("PODATEENTREEINVAL") = rsAssure.fields("POARRET")
    
    ' commentaire et log = date d'entrée en inval forcée= date d'arrêt  (dans assuré, dans Log)
    m_Logger.EcritTraceDansLog "Assuré " & rsAssure.fields("PONOM") & ", Arrêt " & Format(rsAssure.fields("POARRET"), "dd/MM/yyyy") & " : Date d'entrée en inval forcée = date d'arrêt"
    
    If IsNull(rsAssure.fields("POCOMMENTANNUL")) Then
      rsAssure.fields("POCOMMENTANNUL") = " "
    Else
      If Left(rsAssure.fields("POCOMMENTANNUL"), 8) = "Import: " Then
        rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- "
      Else
        rsAssure.fields("POCOMMENTANNUL") = " "
      End If
    End If
    
    rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & "Date d'entrée en inval forcée = date d'arrêt"

  End If
  
  ' chargement date d'entrée en inval
  DateMiseInval = rsAssure.fields("PODATEENTREEINVAL")
  
    
  'âge à la date d'entrée en invalidité
  x = CalculAge(rsAssure.fields("PONAIS"), DateMiseInval)
    
  
  Dim InterpolationInval As Integer
  InterpolationInval = m_rsPeriode.fields("PEINTERPOLATION_INVAL")
     
 
  ' anciennté entre la (Date Entrée en Inval) et (date calcul)
' PHM 12/04/2007 : Pour le calcul du coefficient à l'âge x années et k mois, une interpolation entre les âges x et x+1 est effectuée.
  If InterpolationInval = eInterpolationInval_AgeDuree Then
    ' si interpolation Age/Duree, on ne tronque pas l'ancienneté on l'arrondi
    anc = CInt(NbMois(DateMiseInval, m_DateCalcul) / 12#)
  Else
    ' ancien code
    anc = Int(NbMois(DateMiseInval, m_DateCalcul) / 12#)
  End If
' PHM 12/04/2007 : Pour le calcul du coefficient à l'âge x années et k mois, une interpolation entre les âges x et x+1 est effectuée.
  If anc < 0 Then
    anc = 0
  End If
  

'**** PHM 26/07/2010 : Invalides 1ere catégorie du portefeuille SALARIE

  ' Pour les invalides de 1ère catégorie, au lieu de forcer l’âge limite de provisionnement à 65 ans
  ' nous souhaitons avoir la possibilité de choisir cet âge limite dans les paramètres de calcul de
  ' chaque période. Aussi ce critère deviendrait un paramètre au même titre que le taux technique
  ' pour le calcul des provisions.
  '
  ' Dans ce contexte, deux cas peuvent se présenter et la règle suivante doit être appliquée :
  ' 1-      Si âge au terme (DTLIMPRO - DTNAISSASS) de l’assuré < âge limite paramétré alors
  '         la DTLIMPRO devient (DTNAISS + âge limite paramétré)
  '
  ' 2-      Si âge au terme (DTLIMPRO - DTNAISSASS) de l’assuré > âge limite paramétré alors
  '         la DTLIMPRO reste inchangé
  
  If bForReformeRetraite And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
    
    feuilleLue = m_rsPeriode.fields("PELMINVAL_Retraite")
  
    ' Date limite précisée dans l'import
    'modif le 07/08/2023 appel dans le cas d'inval de cat 1
    If rsAssure.fields("POCategorieInval") = 1 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
        DateLimiteProvision = ComputeDateLimiteApresReformeInval1(rsAssure, rsAssure.fields("POCATEGORIE_NEW"), rsAssure.fields("POTERME"), rsAssure.fields("PONAIS"), DateMiseInval, rsAssure.fields("NUENRP3I"), False)
    Else
        DateLimiteProvision = ComputeDateLimiteApresReforme(rsAssure, rsAssure.fields("POCATEGORIE_NEW"), rsAssure.fields("POTERME"), rsAssure.fields("PONAIS"), DateMiseInval, rsAssure.fields("NUENRP3I"))
    End If
    
    If IsNull(m_rsPeriode.fields("PEAgeLimiteInvalCat1_Retraite")) Then
      AgeLimiteInvalCat1 = 65 ' age par défaut
    Else
      AgeLimiteInvalCat1 = m_rsPeriode.fields("PEAgeLimiteInvalCat1_Retraite") ' age indiqué dans les paramètres de calcul
    End If
    
    ' Invalides 1ere catégorie du portefeuille SALARIE ?
    If rsAssure.fields("POCategorieInval") = 1 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
      AgeMax = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvision)
      If AgeMax < AgeLimiteInvalCat1 Then
        DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + AgeLimiteInvalCat1, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
      End If
    End If
    
    'added 28-2-2018
    If CodeProvision = cdProvision_Inval_Viager Or CodeProvision = cdProvision_Incap_AvecPassage_Viager Then
      DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + ageLimitViagere, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
    End If
  
    rsAssure.fields("POTERME") = DateLimiteProvision
    
    If rsAssure.fields("POTopAmortissable") = True Then
      rsAssure.fields("POCoeffAmortissement") = FindCoeffAmortissement(m_rsPeriode.fields("PEDATEEXT"), rsAssure.fields("PONAIS"))
    Else
      rsAssure.fields("POCoeffAmortissement") = 1
    End If
  
  Else
    
    feuilleLue = m_rsPeriode.fields("PELMINVAL")

    ' Date limite précisée dans l'import
    DateLimiteProvision = rsAssure.fields("POTERME")
    If IsNull(m_rsPeriode.fields("PEAgeLimiteInvalCat1")) Then
      AgeLimiteInvalCat1 = 65 ' age par défaut
    Else
      AgeLimiteInvalCat1 = m_rsPeriode.fields("PEAgeLimiteInvalCat1") ' age indiqué dans les paramètres de calcul
    End If
    
    
    ' Invalides 1ere catégorie du portefeuille SALARIE ?
    If rsAssure.fields("POCategorieInval") = 1 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
      AgeMax = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvision)
      If AgeMax < AgeLimiteInvalCat1 Then
        DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + AgeLimiteInvalCat1, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
      End If
    End If
    
    'added 28-2-2018
    If CodeProvision = cdProvision_Inval_Viager Then
      DateLimiteProvision = DateSerial(Year(rsAssure.fields("PONAIS")) + ageLimitViagere, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
    End If
  
  End If
  
  
'**** PHM 26/07/2010 : Invalides 1ere catégorie du portefeuille SALARIE
  
  
  ' Durée entre la (date de mise en Invalidité) et la (terme du contrat)
  'Duree = Int(NbMois(DateMiseInval, rsAssure.fields("POTERME")) / 12#)
  Duree = Int(NbMois(DateMiseInval, DateLimiteProvision) / 12#)
  If Duree < 0 Then
    Duree = 0
  End If

  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  'm_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  m_xTerme = CalculAge(rsAssure.fields("PONAIS"), DateLimiteProvision)
  m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), DateLimiteProvision)
 
  ' franchise
  franchise = m_xEffet - x
  If franchise < 0 Then
    franchise = 0
  End If
  
  ' franchise en annee ajout du délais de carence
  franchise = franchise + CInt(rsAssure.fields("PODELAI") / 365)
  If franchise < 0 Then
    franchise = 0
  End If
  
  
  '### Invalidité Viagère
   If CodeProvision = cdProvision_Inval_Viager Then
    If Not IsNull(m_rsPeriode.fields("PELMINVAL_Viagere")) Then
      feuilleLue = m_rsPeriode.fields("PELMINVAL_Viagere")
    End If
   End If
  
    
  iInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINVAL")) / 100
  fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINVAL")) / 100
      
  rsAssure.fields("TXTECHN") = iInval * 100#
  rsAssure.fields("TXFRAIS") = fraisGestionInval * 100#
      
      
  rsAssure.fields("POPM_INCAP_1F") = 0
  rsAssure.fields("POPM_PASS_1F") = 0
  rsAssure.fields("POPM_INVAL_1F") = m_ModuleCalcul_Provisions.FctCalcProvInval(x, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval)

  rsAssure.fields("POPM_x") = x
  rsAssure.fields("POPM_xTerme") = IIf(bForReformeRetraite, m_xTermeDouble, m_xTerme)
  rsAssure.fields("POPM_anc") = anc
  
  rsAssure.fields("POCorrectif") = 1 ' chargement = 1  23/08/2013
  
  If bForReformeRetraite Then
    rsAssure.fields("POPM_Duree") = CDbl(NbMois(DateMiseInval, DateLimiteProvision)) / 12#
    If rsAssure.fields("POPM_Duree") < 0 Then
      rsAssure.fields("POPM_Duree") = 0
    End If
  Else
    rsAssure.fields("POPM_Duree") = Duree
  End If
 
  If m_bRevalo Then
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_INVAL_1R") = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
  Else
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = 0
  End If

#If P3I_GENERALI = 1 Then
' PHM 12/04/2007 : Pour le calcul du coefficient à l'âge x années et k mois, une interpolation entre les âges x et x+1 est effectuée.
  Dim PT As Double, PT1 As Double, k As Double
  
  ' interpolation Age
  If InterpolationInval = eInterpolationInval_Age Then
    
    PT = rsAssure.fields("POPM_INVAL_1F")
    PT1 = m_ModuleCalcul_Provisions.FctCalcProvInval(x + 1, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval)
    
'    k = Arrondi(NbMois(rsAssure.fields("PONAIS"), DateMiseInval) - x * 12, NbDecimalePM)
'    If k < 0 Then
'      k = 0
'    End If
'    rsAssure.fields("POPM_INVAL_1F") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
    
    k = NbMois(rsAssure.fields("PONAIS"), DateMiseInval)
    rsAssure.fields("POPM_INVAL_1F") = Arrondi(Interpolation(x, x + 1, PT, PT1, k / 12), NbDecimalePM)
    
    If m_bRevalo Then
      PT = rsAssure.fields("POPM_INVAL_1R")
      PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x + 1, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
      'rsAssure.fields("POPM_INVAL_1R") = Arrondi((PT - (k / 12#) * (PT - PT1)), NbDecimalePM)
      rsAssure.fields("POPM_INVAL_1R") = Arrondi(Interpolation(x, x + 1, PT, PT1, k / 12), NbDecimalePM)
    End If
  
  End If

  ' interpolation Age/Duree
  If InterpolationInval = eInterpolationInval_AgeDuree Then
    Dim PM1 As Double, PM2 As Double, PM3 As Double, PM4 As Double
    ' age
    PT = rsAssure.fields("POPM_INVAL_1F")
    PT1 = m_ModuleCalcul_Provisions.FctCalcProvInval(x + 1, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval)
    
    k = NbMois(rsAssure.fields("PONAIS"), DateMiseInval) / 12#
    PM1 = Interpolation(x, x + 1, PT, PT1, k)
    
    If m_bRevalo Then
      PT = rsAssure.fields("POPM_INVAL_1R")
      PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x + 1, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
      PM2 = Interpolation(x, x + 1, PT, PT1, k)
    End If
  
    ' duree + 1
    PT = m_ModuleCalcul_Provisions.FctCalcProvInval(x, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval)
    PT1 = m_ModuleCalcul_Provisions.FctCalcProvInval(x + 1, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval)
    PM3 = Interpolation(x, x + 1, PT, PT1, k)
    
    If m_bRevalo Then
      PT = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
      PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x + 1, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
      PM4 = Interpolation(x, x + 1, PT, PT1, k)
    End If
    
    ' duree
    'k = NbMois(DateMiseInval, rsAssure.fields("POTERME")) / 12#
    k = NbMois(DateMiseInval, DateLimiteProvision) / 12#
    rsAssure.fields("POPM_INVAL_1F") = Arrondi(Interpolation(Duree, Duree + 1, PM1, PM3, k), NbDecimalePM)
    If m_bRevalo Then
      rsAssure.fields("POPM_INVAL_1R") = Arrondi(Interpolation(Duree, Duree + 1, PM2, PM4, k), NbDecimalePM)
    End If
  End If
' PHM 12/04/2007


' PHM 01/06/2006 :  Generali : duree restante < 12 mois
  If InterpolationInval = eInterpolationInval_CorrectionDuree Then
    Dim duree_restant_mois As Integer
    
    'duree_restant_mois = DateDiff("m", m_DateCalcul, rsAssure.fields("POTERME"))
    duree_restant_mois = DateDiff("m", m_DateCalcul, DateLimiteProvision)
    
    If duree_restant_mois > 0 And duree_restant_mois < 12 Then
      PT = rsAssure.fields("POPM_INVAL_1F")
      PT1 = m_ModuleCalcul_Provisions.FctCalcProvInval(x, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval)
      
      'k = NbMois(DateMiseInval, rsAssure.fields("POTERME")) / 12#
      k = NbMois(DateMiseInval, DateLimiteProvision) / 12#
      rsAssure.fields("POPM_INVAL_1F") = Arrondi(Interpolation(Duree, Duree + 1, PT, PT1, k), NbDecimalePM)
      
      If m_bRevalo Then
        PT = rsAssure.fields("POPM_INVAL_1R")
        PT1 = m_ModuleCalcul_Provisions_Revalo.FctCalcProvInvalRevalo(x, anc, Duree + 1, franchise, feuilleLue, iInval, fraisGestionInval, m_DureRevalo, m_TauxRevalo)
        
        rsAssure.fields("POPM_INVAL_1R") = Arrondi(Interpolation(Duree, Duree + 1, PT, PT1, k), NbDecimalePM)
      End If
      
'      Dim correctif As Double
'
'      'k = m_xArretMois - (12 * m_xArret)
'      'correctif = 12# - k / 12#
'      ' Nb de mois restant à payer
'      k = Arrondi(DateDiff("d", m_DateCalcul, rsAssure.fields("POTERME")) / 30, 0)
'      correctif = k / 12#
'
'      rsAssure.fields("POPM_INVAL_1F") = correctif * rsAssure.fields("POPM_INVAL_1F")
'      rsAssure.fields("POPM_INVAL_1R") = correctif * rsAssure.fields("POPM_INVAL_1R")
    End If
  End If
' PHM 01/06/2006 :  Generali : duree restante < 12 mois

#End If
  
  rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_INVAL_1F")
  If IsNull(rsAssure.fields("POPM")) Then
    rsAssure.fields("POPM") = 0
  End If
  If m_bRevalo Then
    rsAssure.fields("POPM_REVALO") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_INVAL_1R") - rsAssure.fields("POPM")
    If IsNull(rsAssure.fields("POPM_REVALO")) Then
      rsAssure.fields("POPM_REVALO") = 0
    End If
    If rsAssure.fields("POPM_REVALO") < 0 Then rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = rsAssure.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO))
  End If
  rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_INVAL_1F")
  rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * rsAssure.fields("POPM_INVAL_1F")
     
  If IsNull(rsAssure.fields("POPRESTATION_AN")) Then
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
  End If
     
'*** Modif UPESE assuré invalide avec une garantie INCAPACITE
  ' PHM 25/09/2004 If rsAssure.Fields("GATYPEGAR") = cdTypeIncapacite Then
  If FindGarantie(rsAssure.fields("POGARCLE")).TypeGarantie = cdTypeIncapacite Then
    ' sauvegarde les valeurs calculées
    rsAssure.fields("POPMANNULEE") = rsAssure.fields("POPM")
    If IsNull(rsAssure.fields("POCOMMENTANNUL")) Then
      rsAssure.fields("POCOMMENTANNUL") = " "
    Else
      If Left(rsAssure.fields("POCOMMENTANNUL"), 8) = "Import: " Then
        rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- "
      Else
        rsAssure.fields("POCOMMENTANNUL") = " "
      End If
    End If
    rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & "PM annulée car assuré en invalidité mais garantie = incapacité"
    
    ' on force les coeff de provisions à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    If m_bRevalo Then
      rsAssure.fields("POPM_REVALO") = 0
    End If
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
'*** Modif UPESE
  Else
    rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
    If m_bRevalo Then
      rsAssure.fields("POPM_REVALO") = Arrondi(rsAssure.fields("POPM_REVALO"), 2)
      rsAssure.fields("POCOT_REVALO") = Arrondi(rsAssure.fields("POCOT_REVALO"), 2)
    End If
    rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
    rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
  End If

' PHM 18/11/2008 :  Generali : duree restante < 12 mois
  If rsAssure.fields("POPM") = 0 Then
    Dim duree_restant_jour As Double
    
    'duree_restant_jour = DateDiff("d", m_DateCalcul, rsAssure.fields("POTERME"))
    duree_restant_jour = DateDiff("d", m_DateCalcul, DateLimiteProvision)
    
    If duree_restant_jour > 0 And duree_restant_jour <= 365# Then
      rsAssure.fields("POPM_INVAL_1F") = Arrondi(duree_restant_jour / 365.25, NbDecimalePM)
      rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * duree_restant_jour / 365.25
      rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
      
      rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_INVAL_1F")
      rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * rsAssure.fields("POPM_INVAL_1F")
      If m_bRevalo Then
        rsAssure.fields("POPM_REVALO") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_INVAL_1R") - rsAssure.fields("POPM")
        If rsAssure.fields("POPM_REVALO") < 0 Then rsAssure.fields("POPM_REVALO") = 0
        rsAssure.fields("POCOT_REVALO") = rsAssure.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO))
      End If
    End If
  End If
' PHM 18/11/2008 :  Generali : duree restante < 12 mois
  
  Exit Sub
  
err_calcul:
   
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recherche les garanties pour les Rentes de Conjoint
'
'##ModelId=5C8A75EC01F2
Private Function FindTauxRenteConjoint(Regime As Long, Categorie As String) As Boolean
  ' recupere le taux de garantie
  m_dataHelper.Multi_Find m_rsTauxRenteConjoint, "Regime=" & Regime & " AND Categorie='" & Categorie & "'"
  
  FindTauxRenteConjoint = True
  If m_rsTauxRenteConjoint.EOF Then
    ' Ajoute à la liste des catégories non trouvées
    m_CatNotFoundRenteConjoint.AddCategorie Regime, Categorie
    FindTauxRenteConjoint = False
  End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Rentes de Conjoint
'
'##ModelId=5C8A75EC0234
Private Sub CalculRenteConjoint(ByRef rs As ADODB.Recordset, bMantienDeces As Boolean, salaireAnnuel As Double, coeffBCAC As Double, PctLissage As Double, bRenteTemporaire As Boolean, AgeAssure As Integer)
  Dim feuilleLue As String
  Dim iInval As Double
  Dim fraisGestionInval As Double
  Dim p As Byte
  Dim f As Byte
  Dim frac As Double
  
  Dim bTemporaire As Boolean, bCapitalMoyen As Boolean, capitalMoyen As Double, TauxRenteConjoint As Double
  Dim Duree As Integer, anneeNaissance As Integer, anneeTerme As Integer
  
  On Error GoTo err_calcul
  
  '
  ' on tiens compte de la garantie : 51 : sans distinction
  '                                  60 : temporaire : date de terme specifiée
  '                                  61 : viagère
  '
  
  ' parametres
  feuilleLue = m_rsPeriode.fields("PETRENTECONJOINT")
  iInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXRTCONJOINT")) / 100
  If bMantienDeces Then
    fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGestionConjoint_MDC")) / 100
  Else
    fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONCONJOINT")) / 100
  End If
  p = m_rsPeriode.fields("PEPAIEMENTCONJOINT")
  f = m_rsPeriode.fields("PEFRACTIONCONJOINT")


  rs.fields("TXTECHN") = iInval * 100#
  rs.fields("TXFRAIS") = fraisGestionInval * 100#
  
  
  ' recherche du taux de garanti
  If bMantienDeces Then
    ' années de calcul
    anneeNaissance = Year(rs.fields("PONAIS")) + IIf(rs.fields("POSEXE") = 2, -1 * m_rsPeriode.fields("PEAgeConjoint_MDC"), m_rsPeriode.fields("PEAgeConjoint_MDC"))
    anneeTerme = 0
    
    If FindTauxRenteConjoint(rs.fields("POGARCLE") - 50, rs.fields("POCATEGORIE")) = False Then
      If bRenteTemporaire Then
        anneeTerme = anneeNaissance + 60 ' terme à 60 ans pour éviter les erreurs de calcul
      End If
    Else
      If bRenteTemporaire Then
        anneeTerme = m_rsTauxRenteConjoint.fields("AgeTerme") + anneeNaissance
      End If
    End If
  Else
    anneeNaissance = Year(rs.fields("PONAIS"))
    If IsNull(rs.fields("POTERME")) Then
      anneeTerme = anneeNaissance + 60 ' terme à 60 ans pour éviter les erreurs de calcul
    Else
      anneeTerme = Year(rs.fields("POTERME"))
    End If
  End If
  
  
  ' init champs
  rs.fields("POPM_INCAP_1F") = 0
  rs.fields("POPM_PASS_1F") = 0
  rs.fields("POPM_INVAL_1F") = 0
  rs.fields("POPM_INCAP_1R") = 0
  rs.fields("POPM_PASS_1R") = 0
  rs.fields("POPM_INVAL_1R") = 0
  rs.fields("POPM_RCJT_1F") = 0
  rs.fields("POPM_REDUC_1F") = 0
  rs.fields("POPM_REDUC_1R") = 0
  rs.fields("POPM_REDUC_1R") = 0
  rs.fields("POPM_REVALO") = 0
  rs.fields("POPM_SORTIE") = 0
  rs.fields("POPM_VAR") = 0
  rs.fields("POPM_REVALO") = 0
  rs.fields("POCOT_REVALO") = 0
  
  ' âge à la date du calcul (=date de fin de période)
  m_xCalcul = Year(m_DateCalcul) - anneeNaissance
  If m_xCalcul < 0 Then
    m_xCalcul = 0
  End If
  
  ' f : 1 Annuel, 2 Semestriel, 3 Trimestriel, 4 Mensuel
  
  ' calcul de frac=(m-1)/2m
  Select Case f
    Case 1
      frac = 0
    Case 2
      frac = 0.25
    Case 3
      frac = 0.375
    Case 4
      frac = 11# / 24#
  End Select
  
  ' recherche le type de rente et le taux de garantie
  bTemporaire = False
  bCapitalMoyen = False
  
  If bMantienDeces Then
    bTemporaire = bRenteTemporaire
    
    TauxRenteConjoint = 0
    bCapitalMoyen = False
    
    If m_rsTauxRenteConjoint.EOF = False Then
      If Not IsNull(m_rsTauxRenteConjoint.fields("Taux_65x_x25")) Then
        If bTemporaire Then
          ' (x-25)*Taux
          TauxRenteConjoint = (AgeAssure - 25#) * m_rsTauxRenteConjoint.fields("Taux_65x_x25") / 100#
        Else
          ' (65-x)*Taux
          TauxRenteConjoint = (65# - AgeAssure) * m_rsTauxRenteConjoint.fields("Taux_65x_x25") / 100#
        End If
      ElseIf Not IsNull(m_rsTauxRenteConjoint.fields("Taux_k")) Then
        TauxRenteConjoint = m_rsTauxRenteConjoint.fields("Taux_k") / 100#
      ElseIf Not IsNull(m_rsTauxRenteConjoint.fields("CapitalMoyen")) Then
        TauxRenteConjoint = 0
        bCapitalMoyen = True
        capitalMoyen = m_rsTauxRenteConjoint.fields("CapitalMoyen")
      End If
    End If
    
    If TauxRenteConjoint < 0# Then
      TauxRenteConjoint = 0#
    End If
    
    ' capital forcé ?
    If m_rsPeriode.fields("PEForcerCapConstitMoyen_MDC") = True Then
      bCapitalMoyen = True
      capitalMoyen = m_rsPeriode.fields(IIf(bTemporaire, "PECapConstitMoyenTempo_MDC", "PECapConstitMoyenViagere_MDC"))
    End If
    
    If capitalMoyen < 0# Then
      capitalMoyen = 0
    End If
  Else
    bTemporaire = Not IsNull(rs.fields("POTERME"))
  End If
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  If bTemporaire = False Then
    '*** rente de conjoint viagere
    
    ' paiement
    If p = 1 Then ' paiement d'avance
      rs.fields("POPM_RCJT_1F") = m_ModuleCalcul_Provisions.FctCalc_äx(feuilleLue, iInval, m_xCalcul, anneeNaissance, m_fWait) - frac
    ElseIf p = 2 Then ' paiement à terme echu
      rs.fields("POPM_RCJT_1F") = m_ModuleCalcul_Provisions.FctCalc_ax(feuilleLue, iInval, m_xCalcul, anneeNaissance, m_fWait) + frac
    End If
    
    ' fractionnement deja inclus
    
    If bMantienDeces Then
      rs.fields("POTERME") = Null
    End If
  Else
    '*** rente de conjoint temporaire
    m_xTerme = anneeTerme - anneeNaissance
    If m_xTerme < 0 Then
      m_xTerme = 0
    End If
    m_xTermeDouble = m_xTerme
    
    'Duree = m_xTerme - m_xArret
    Duree = m_xTerme - m_xCalcul
    If Duree < 0 Then
      Duree = 0
    End If
  
    ' nEx
    Dim Correction As Double
    
    ' Correction=((m-1)/2m)*(1-nEx)
    Correction = frac * (1 - m_ModuleCalcul_Provisions.FctCalc_nEx(feuilleLue, iInval, m_xCalcul, Duree, anneeNaissance, m_fWait))
    
    ' paiement et fractionement
    If p = 1 Then ' paiement d'avance
      rs.fields("POPM_RCJT_1F") = m_ModuleCalcul_Provisions.FctCalc_äxn(feuilleLue, iInval, m_xCalcul, Duree, anneeNaissance, m_fWait) - Correction
    ElseIf p = 2 Then ' paiement à terme echu
      rs.fields("POPM_RCJT_1F") = m_ModuleCalcul_Provisions.FctCalc_axn(feuilleLue, iInval, m_xCalcul, Duree, anneeNaissance, m_fWait) + Correction
    End If
  
    If bMantienDeces Then
      rs.fields("POTERME") = DateSerial(anneeTerme, Month(rs.fields("PONAIS")), Day(rs.fields("PONAIS")))
    End If
  End If
  
  If bMantienDeces Then
    Dim RCTJ_1F As Double
    
    ' frais de gestion * coeff BCAC
    rs.fields("POCoeffBCAC") = coeffBCAC
    RCTJ_1F = Arrondi((1 + fraisGestionInval) * rs.fields("POPM_RCJT_1F") * coeffBCAC, NbDecimaleCalcul)
    rs.fields("POPM_RCJT_1F") = Arrondi((1 + fraisGestionInval) * rs.fields("POPM_RCJT_1F"), NbDecimaleCalcul)
    
    ' calcul
    If bCapitalMoyen Then
      ' pas de taux pour 1F
      rs.fields("POPM_RCJT_1F") = 0
      
      ' PM = CapitalSubstitutif * (1 + g) * coeffBCAC * PctLissage / 100
      rs.fields("POPM") = capitalMoyen * (1 + fraisGestionInval) * coeffBCAC * PctLissage / 100#
    Else
      ' rente temporaire ou viagere
      rs.fields("POPM") = salaireAnnuel * TauxRenteConjoint * RCTJ_1F * PctLissage / 100#
    End If
  Else
    ' frais de gestion
    rs.fields("POPM_RCJT_1F") = Arrondi((1 + fraisGestionInval) * rs.fields("POPM_RCJT_1F"), NbDecimalePM)
    
    ' calcul
    rs.fields("POPM") = rs.fields("POPRESTATION_AN") * rs.fields("POPM_RCJT_1F")
  End If

  ' arrondi
  rs.fields("POPM") = Arrondi(rs.fields("POPM"), 2)
  
  Exit Sub
  
err_calcul:

  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rs.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rs.fields("NUENRP3I") & vbLf & Err.Description
  End If
    
  'MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Rentes Education
'
'##ModelId=5C8A75EC0318
Private Sub CalculRenteEducation(ByRef rs As ADODB.Recordset, bMantienDeces As Boolean, salaireAnnuel As Double, coeffBCAC As Double, PctLissage As Double, nombreEnfant As Double, ageMoyenEnfant As Double)
  Dim age_b1 As Integer
  Dim age_b2 As Integer
  Dim age_b3 As Integer
  Dim age_b4 As Integer
  
  Dim taux_b1 As Double
  Dim taux_b2 As Double
  Dim taux_b3 As Double
  Dim taux_b4 As Double
  
  Dim palier1 As Double
  Dim palier2 As Double
  Dim palier3 As Double
  Dim palier4 As Double
  
  Dim feuilleLue As String
  Dim iInval As Double
  Dim fraisGestionInval As Double
  Dim p As Byte
  Dim f As Byte
  Dim frac As Double
  
  On Error GoTo err_calcul
  
  ' recherche du taux
  m_dataHelper.Multi_Find m_rsTauxRenteEducation, "Categorie='" & rs.fields("POCATEGORIE") & "'"
  
  If Not m_rsTauxRenteEducation.EOF Then
    age_b1 = m_rsTauxRenteEducation.fields("B1")
    taux_b1 = m_rsTauxRenteEducation.fields("B1_PCT")
    age_b2 = m_rsTauxRenteEducation.fields("B2")
    taux_b2 = m_rsTauxRenteEducation.fields("B2_PCT")
    age_b3 = m_rsTauxRenteEducation.fields("B3")
    taux_b3 = m_rsTauxRenteEducation.fields("B3_PCT")
    age_b4 = m_rsTauxRenteEducation.fields("B4")
    taux_b4 = m_rsTauxRenteEducation.fields("B4_PCT")
  Else
    age_b1 = 0
    taux_b1 = 0
    age_b2 = 0
    taux_b2 = 0
    age_b3 = 0
    taux_b3 = 0
    age_b4 = 0
    taux_b4 = 0
    
    ' Ajoute à la liste des catégories non trouvées
    m_CatNotFoundRenteEducation.AddCategorie 0, rs.fields("POCATEGORIE")
  End If
  
  ' recherche du terme (age_bx peut être egal à 0)
  palier1 = age_b1
  If palier1 < age_b2 Then
    palier1 = age_b2
  End If
  If palier1 < age_b3 Then
    palier1 = age_b3
  End If
  If palier1 < age_b4 Then
    palier1 = age_b4
  End If
  
  If bMantienDeces Then
    ' on calcul la date de naissance de l'enfant et on la stocke dans rs.Fields("PONAIS")
    ' a l'entree dans cette fonction, c'est la date de naissance de l'assuré qui est stockée dans rs.Fields("PONAIS")
    rs.fields("PONAIS") = DateSerial(Year(m_DateCalcul) - ageMoyenEnfant, 1, 1)
    If Not m_rsTauxRenteEducation.EOF Then
      ' affiche la date du dernier palier au lieu de la date de retraite dans POTERME
      rs.fields("POTERME") = DateSerial(Year(rs.fields("PONAIS")) + palier1, Month(rs.fields("PONAIS")), Day(rs.fields("PONAIS")))
    Else
      ' null car pas d'info
      rs.fields("POTERME") = Null
    End If
  Else
    ' affiche la date du dernier palier au lieu de la date de retraite dans POTERME
    rs.fields("POTERME") = DateSerial(Year(rs.fields("PONAIS")) + palier1, Month(rs.fields("PONAIS")), Day(rs.fields("PONAIS")))
  End If
  
  ' parametres
  feuilleLue = m_rsPeriode.fields("PETRENTEEDUC")
  iInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXRTEDUC")) / 100#
  If bMantienDeces Then
    fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGestionEduc_MDC")) / 100#
  Else
    fraisGestionInval = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONEDUC")) / 100#
  End If
  p = m_rsPeriode.fields("PEPAIEMENTEDUC")
  f = m_rsPeriode.fields("PEFRACTIONEDUC")

  rs.fields("TXTECHN") = iInval * 100#
  rs.fields("TXFRAIS") = fraisGestionInval * 100#


  ' init champs
  rs.fields("POPM_INCAP_1F") = 0
  rs.fields("POPM_PASS_1F") = 0
  rs.fields("POPM_INVAL_1F") = 0
  rs.fields("POPM_INCAP_1R") = 0
  rs.fields("POPM_PASS_1R") = 0
  rs.fields("POPM_INVAL_1R") = 0
  rs.fields("POPM_RCJT_1F") = 0
  rs.fields("POPM_REDUC_1F") = 0
  rs.fields("POPM_RCJT_1R") = 0
  rs.fields("POPM_REDUC_1R") = 0
  rs.fields("POPM_REVALO") = 0
  rs.fields("POPM_SORTIE") = 0
  rs.fields("POPM_VAR") = 0
  rs.fields("POPM_REVALO") = 0
  rs.fields("POCOT_REVALO") = 0
  
  ' âge anniversaire à la date du calcul
  If bMantienDeces Then
    m_xCalcul = ageMoyenEnfant
  Else
    m_xCalcul = AgeAnniversaire(rs.fields("PONAIS"), m_rsPeriode.fields("PEDATEEXT"))
  End If
  If m_xCalcul < 0 Then
    m_xCalcul = 0
  End If
  
  ' f : 1 Annuel, 2 Semestriel, 3 Trimestriel, 4 Mensuel
  
  ' calcul de frac=(m-1)/2m
  Select Case f
    Case 1
      frac = 0
    Case 2
      frac = 0.25
    Case 3
      frac = 0.375
    Case 4
      frac = 11# / 24#
  End Select


  palier1 = 0
  palier2 = 0
  palier3 = 0
  palier4 = 0

  ' m_xCalcul < age_b1
  If m_xCalcul < age_b1 Then
    
    If p = 1 Then ' paiement d'avance
              ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_äxn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) - Correction
       palier1 = (0 + taux_b1) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b1 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier2 = (0 + taux_b2) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b1 - m_xCalcul, frac, age_b2 - age_b1, Year(rs.fields("PONAIS")), m_fWait)
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b2 - m_xCalcul, frac, age_b3 - age_b2, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)
              
    ElseIf p = 2 Then ' paiement à terme echu
      ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_axn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) + Correction
    
       ' format (âge au calcul, différé, durée)
       palier1 = (0 + taux_b1) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b1 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier2 = (0 + taux_b2) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b1 - m_xCalcul, frac, age_b2 - age_b1, Year(rs.fields("PONAIS")), m_fWait)
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b2 - m_xCalcul, frac, age_b3 - age_b2, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)
    
    
    End If
  
  ElseIf m_xCalcul < age_b2 Then
  
    If p = 1 Then ' paiement d'avance
              ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_äxn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) - Correction
       palier1 = 0
       palier2 = (0 + taux_b2) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b2 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b2 - m_xCalcul, frac, age_b3 - age_b2, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)
    
    ElseIf p = 2 Then ' paiement à terme echu
      ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_axn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) + Correction
       palier1 = 0
       palier2 = (0 + taux_b2) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b2 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b2 - m_xCalcul, frac, age_b3 - age_b2, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)
    
    End If
  
  ElseIf m_xCalcul < age_b3 Then
  
    If p = 1 Then ' paiement d'avance
              ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_äxn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) - Correction
       palier1 = 0
       palier2 = 0
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b3 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)
    
              
    ElseIf p = 2 Then ' paiement à terme echu
      ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_axn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) + Correction
       palier1 = 0
       palier2 = 0
       palier3 = (0 + taux_b3) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b3 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, age_b3 - m_xCalcul, frac, age_b4 - age_b3, Year(rs.fields("PONAIS")), m_fWait)

    End If
  
  ElseIf m_xCalcul < age_b4 Then
  
    If p = 1 Then ' paiement d'avance
              ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_äxn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) - Correction
       palier1 = 0
       palier2 = 0
       palier3 = 0
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b4 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
              
    ElseIf p = 2 Then ' paiement à terme echu
      ''rs.Fields("POPM_REDUC_1F") = (0 + taux_b1) * FctCalc_axn(feuilleLue, iInval, m_xCalcul, Duree, Year(rs.Fields("PONAIS")), m_fWait) + Correction
       palier1 = 0
       palier2 = 0
       palier3 = 0
       palier4 = (0 + taux_b4) * m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, iInval, m_xCalcul, 0, frac, age_b4 - m_xCalcul, Year(rs.fields("PONAIS")), m_fWait)
  
    End If
  
  End If

  ' + frais de gestion

  rs.fields("POPM_REDUC_1F") = Arrondi((1 + fraisGestionInval) * (palier1 + palier2 + palier3 + palier4), NbDecimaleCalcul)

  ' calcul
  Dim tauxRteEduc As Double
  
  tauxRteEduc = FindTauxRenteEducation(rs.fields("POCATEGORIE"), rs.fields("PONAIS"), _
                                       IIf(IsNull(rs.fields("POFIN")), rs.fields("POFINTOTAL"), rs.fields("POFIN")), _
                                       bMantienDeces, ageMoyenEnfant)
  
  If tauxRteEduc <> 0 Then
    If bMantienDeces Then
      Dim REDUC_1F As Double
      
      ' on multiplie les PM par le nb d'enfants
      ' PM = nbEnfants * PM_pour_1_enfant
      REDUC_1F = Arrondi(rs.fields("POPM_REDUC_1F") * coeffBCAC, NbDecimaleCalcul)
      rs.fields("POPM_REDUC_1F") = Arrondi(rs.fields("POPM_REDUC_1F"), NbDecimaleCalcul)
      rs.fields("POPM") = nombreEnfant * salaireAnnuel * REDUC_1F * PctLissage / 100
      rs.fields("POCoeffBCAC") = coeffBCAC
      
      ' stocke les params de calcul
      rs.fields("PONbEnfant") = nombreEnfant
      rs.fields("POAgeMoyenEnfant") = ageMoyenEnfant
    Else
      rs.fields("POPM") = (rs.fields("POPRESTATION_AN") / (0 + tauxRteEduc)) * rs.fields("POPM_REDUC_1F")
    End If
  Else
    rs.fields("POPM") = 0
    If IsNull(rs.fields("POCOMMENTANNUL")) Then
      rs.fields("POCOMMENTANNUL") = " "
    Else
      If Left(rs.fields("POCOMMENTANNUL"), 8) = "Import: " Then
        rs.fields("POCOMMENTANNUL") = rs.fields("POCOMMENTANNUL") & " --- "
      Else
        rs.fields("POCOMMENTANNUL") = " "
      End If
    End If
    rs.fields("POCOMMENTANNUL") = rs.fields("POCOMMENTANNUL") & "PM annulée car TauxRenteEducation = 0"
  End If

  rs.fields("POPM") = Arrondi(rs.fields("POPM"), NbDecimalePM)
  
  Exit Sub
  
err_calcul:
    
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rs.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rs.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'##ModelId=5C8A75EC03E3
Private Function FindParamRente(dateSurvenance As Date) As Boolean
  FindParamRente = False
  
  m_rsParamRentes.MoveFirst
  
  Do Until m_rsParamRentes.EOF
    
    If IsNull(m_rsParamRentes.fields("Datedebut")) Then
      
      If dateSurvenance <= m_rsParamRentes.fields("DateFin") Then
        FindParamRente = True
        Exit Function
      End If
    
    Else
    
      If IsNull(m_rsParamRentes.fields("DateFin")) Then
      
        If m_rsParamRentes.fields("DateDebut") <= dateSurvenance Then
          FindParamRente = True
          Exit Function
        End If
          
      Else
      
        If m_rsParamRentes.fields("DateDebut") <= dateSurvenance And dateSurvenance <= m_rsParamRentes.fields("DateFin") Then
          FindParamRente = True
          Exit Function
        End If
          
      End If
    
    End If
    
    m_rsParamRentes.MoveNext
  Loop
  
End Function

'##ModelId=5C8A75ED0039
Private Function FindDonneesSociales(dateSurvenance As Date) As Boolean
  FindDonneesSociales = False
  
  m_rsDonneesSociales.MoveFirst
  
  Do Until m_rsDonneesSociales.EOF
    
    If IsNull(m_rsDonneesSociales.fields("Datedebut")) Then
      
      If dateSurvenance <= m_rsDonneesSociales.fields("DateFin") Then
        FindDonneesSociales = True
        Exit Function
      End If
    
    Else
    
      If IsNull(m_rsDonneesSociales.fields("DateFin")) Then
      
        If m_rsDonneesSociales.fields("DateDebut") <= dateSurvenance Then
          FindDonneesSociales = True
          Exit Function
        End If
          
      Else
      
        If m_rsDonneesSociales.fields("DateDebut") <= dateSurvenance And dateSurvenance <= m_rsDonneesSociales.fields("DateFin") Then
          FindDonneesSociales = True
          Exit Function
        End If
          
      End If
    
    End If
    
    m_rsDonneesSociales.MoveNext
  Loop
  
End Function





'##ModelId=5C8A75ED0078
Private Function FindCoeffAmortissement(dateCalcul As Date, dateNaissance As Date) As Double
  FindCoeffAmortissement = 1#
  
  Dim sDateCalcul As String, sDateNaissance As String
  
  sDateCalcul = Format(dateCalcul, "YYYYMM")
  sDateNaissance = " #" & Format(dateNaissance, "mm/dd/yyyy") & "# " 'm_dataHelper.BuildDateString(dateNaissance)
  
  m_dataHelper.Multi_Find m_rsCoeffAmortissement, "AnneeMoisCalcul=" & sDateCalcul & " AND DateDebut<=" & sDateNaissance & " AND DateFin>=" & sDateNaissance
  If Not m_rsCoeffAmortissement.EOF Then
    FindCoeffAmortissement = m_rsCoeffAmortissement.fields("CoeffAmortissement")
  Else
    ' Recherche sans date de calcul (999999)
    m_dataHelper.Multi_Find m_rsCoeffAmortissement, "AnneeMoisCalcul=999999 AND DateDebut<=" & sDateNaissance & " AND DateFin>=" & sDateNaissance
    If Not m_rsTaux_Reassurance.EOF Then
      FindCoeffAmortissement = m_rsCoeffAmortissement.fields("CoeffAmortissement")
    Else
      FindCoeffAmortissement = 1#
    End If
  End If
    
End Function


'##ModelId=5C8A75ED00C6
Private Function FindAgeDepartRetraiteInval1(dateNaissance As Date, bAgeTauxPlein As Boolean) As Double
  FindAgeDepartRetraiteInval1 = 1#
  
  Dim sDateNaissance As String
  
  sDateNaissance = " #" & Format(dateNaissance, "mm/dd/yyyy") & "# " 'm_dataHelper.BuildDateString(dateNaissance)
  
  m_dataHelper.Multi_Find m_rsAgeDepartRetraiteInval1, "DateDebut<=" & sDateNaissance & " AND DateFin>=" & sDateNaissance
  If Not m_rsAgeDepartRetraiteInval1.EOF Then
      FindAgeDepartRetraiteInval1 = IIf(bAgeTauxPlein, m_rsAgeDepartRetraiteInval1.fields("AgeMinimumTauxPlein"), m_rsAgeDepartRetraiteInval1.fields("AgeMinimum"))
  Else
    FindAgeDepartRetraiteInval1 = IIf(bAgeTauxPlein, 67, 64)
  End If
    
End Function
'##ModelId=5C8A75ED00C6
Private Function FindAgeDepartRetraite(dateNaissance As Date, bAgeTauxPlein As Boolean) As Double
  FindAgeDepartRetraite = 1#
  
  Dim sDateNaissance As String
  
  sDateNaissance = " #" & Format(dateNaissance, "mm/dd/yyyy") & "# " 'm_dataHelper.BuildDateString(dateNaissance)
  
  m_dataHelper.Multi_Find m_rsAgeDepartRetraite, "DateDebut<=" & sDateNaissance & " AND DateFin>=" & sDateNaissance
  If Not m_rsAgeDepartRetraite.EOF Then
      FindAgeDepartRetraite = IIf(bAgeTauxPlein, m_rsAgeDepartRetraite.fields("AgeMinimumTauxPlein"), m_rsAgeDepartRetraite.fields("AgeMinimum"))
  Else
    FindAgeDepartRetraite = IIf(bAgeTauxPlein, 67, 62)
  End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calcul de la date limite de provisionnement après réforme pour l'incap et
'l'inval
'
' 1 Incapacité professionnelle      MAX(DTLIMPRO ; DNA + DepartTxPlein(DNA))
' 2 Incapacité non professionnelle  MAX(DTLIMPRO ; Surv + 3 ans ; DNA +
'MinDepart(DNA))
' 3 Invalidité professionnelle      MAX(DTLIMPRO ; DNA + DepartTxPlein(DNA))
' 4 Invalidité non professionnelle  MAX(DTLIMPRO ; DNA + MinDepart(DNA)
'
'##ModelId=5C8A75ED0114
Private Function ComputeDateLimiteApresReforme(ByRef rsAssure As ADODB.Recordset, CodeProvision As Integer, dateLimite As Date, dateNaissance As Date, dateSurv As Date, ByVal NUENRP3I As Long) As Date
  Dim ageDepartMin As Double, ageDepartTxPlein As Double
  
  ageDepartMin = FindAgeDepartRetraite(dateNaissance, False)
  ageDepartTxPlein = FindAgeDepartRetraite(dateNaissance, True)
  
  rsAssure.fields("POAgeMiniDepartRetraite") = ageDepartMin
  rsAssure.fields("POAgeDepartRetraiteTauxPlein") = ageDepartTxPlein

  Dim DateDepartMin As Date, DateDepartTxPlein As Date, DateSurv3Ans As Date
  
  DateDepartMin = DateSerial(Year(dateNaissance) + Int(ageDepartMin), Month(dateNaissance) + ((ageDepartMin - Int(ageDepartMin)) * 12), Day(dateNaissance))
  DateSurv3Ans = DateAdd("yyyy", 3, dateSurv)
  DateDepartTxPlein = DateSerial(Year(dateNaissance) + Int(ageDepartTxPlein), Month(dateNaissance) + ((ageDepartTxPlein - Int(ageDepartTxPlein)) * 12), Day(dateNaissance))
  
  Select Case CodeProvision
    
    Case cdPositImport_IncapProf
      ComputeDateLimiteApresReforme = DateMax(dateLimite, DateDepartTxPlein)
      
    Case cdPositImport_IncapNonProf
      ComputeDateLimiteApresReforme = DateMax(DateMax(dateLimite, DateSurv3Ans), DateDepartMin)
      
    Case cdPositImport_InvalProf
      ComputeDateLimiteApresReforme = DateMax(dateLimite, DateDepartTxPlein)
      
    Case cdPositImport_InvalNonProf
      ComputeDateLimiteApresReforme = DateMax(dateLimite, DateDepartMin)
      
    Case Else
      m_Logger.EcritTraceDansLog "ComputeDateLimiteApresReforme() : Code provision " & CodeProvision & " inconnu ! - Ligne NUMENRP3I=" & NUENRP3I
      ComputeDateLimiteApresReforme = dateLimite
      
  End Select

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calcul de la date limite de provisionnement après réforme pour l'incap et
'l'inval
'
' 1 Incapacité professionnelle      MAX(DTLIMPRO ; DNA + DepartTxPlein(DNA))
' 2 Incapacité non professionnelle  MAX(DTLIMPRO ; Surv + 3 ans ; DNA +
'MinDepart(DNA))
' 3 Invalidité professionnelle      MAX(DTLIMPRO ; DNA + DepartTxPlein(DNA))
' 4 Invalidité non professionnelle  MAX(DTLIMPRO ; DNA + MinDepart(DNA)
'
'##ModelId=5C8A75ED0114
Private Function ComputeDateLimiteApresReformeInval1(ByRef rsAssure As ADODB.Recordset, CodeProvision As Integer, dateLimite As Date, dateNaissance As Date, dateSurv As Date, ByVal NUENRP3I As Long, bPassageIncapNonProf As Boolean) As Date
  Dim ageDepartMin As Double, ageDepartTxPlein As Double
  
  ageDepartMin = FindAgeDepartRetraiteInval1(dateNaissance, False)
  ageDepartTxPlein = FindAgeDepartRetraiteInval1(dateNaissance, True)
  
  'les 07/08/2023, Attention à ne pas écraser les données lors du nouveau calcul PM Passage
  If bPassageIncapNonProf = True Then
    ' rsAssure.fields("POAgeMiniDepartRetraite") = ageDepartMin => Déjà chargé avec ComputeDateLimiteApresReforme
    ' rsAssure.fields("POAgeDepartRetraiteTauxPlein") = ageDepartTxPlein => Déjà chargé avec ComputeDateLimiteApresReforme
  Else
    rsAssure.fields("POAgeMiniDepartRetraite") = ageDepartMin
    rsAssure.fields("POAgeDepartRetraiteTauxPlein") = ageDepartTxPlein
  End If

  Dim DateDepartMin As Date, DateDepartTxPlein As Date, DateSurv3Ans As Date
  
  DateDepartMin = DateSerial(Year(dateNaissance) + Int(ageDepartMin), Month(dateNaissance) + ((ageDepartMin - Int(ageDepartMin)) * 12), Day(dateNaissance))
  DateSurv3Ans = DateAdd("yyyy", 3, dateSurv)
  DateDepartTxPlein = DateSerial(Year(dateNaissance) + Int(ageDepartTxPlein), Month(dateNaissance) + ((ageDepartTxPlein - Int(ageDepartTxPlein)) * 12), Day(dateNaissance))
  
  Select Case CodeProvision
    
    Case cdPositImport_IncapProf
      ComputeDateLimiteApresReformeInval1 = DateMax(dateLimite, DateDepartTxPlein)
      
    Case cdPositImport_IncapNonProf
      ComputeDateLimiteApresReformeInval1 = DateMax(DateMax(dateLimite, DateSurv3Ans), DateDepartMin)
      
    Case cdPositImport_InvalProf
      ComputeDateLimiteApresReformeInval1 = DateMax(dateLimite, DateDepartTxPlein)
      
    Case cdPositImport_InvalNonProf
      ComputeDateLimiteApresReformeInval1 = DateMax(dateLimite, DateDepartMin)
      
    Case Else
      m_Logger.EcritTraceDansLog "ComputeDateLimiteApresReformeInval1() : Code provision " & CodeProvision & " inconnu ! - Ligne NUMENRP3I=" & NUENRP3I
      ComputeDateLimiteApresReformeInval1 = dateLimite
      
  End Select

End Function


'##ModelId=5C8A75ED01C0
Private Function CalculPMRente1Euros(TermePaiement As EcheancePaiement, feuilleLue As String, tx As Double, x As Double, k As Double, frac As Double, n As Double, anNais As Integer, g As Double, m_fWait As frmWait) As Double
                                     
  Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double
  Dim a As Double, b As Double, c As Double
  Dim t11 As Double, t22 As Double, t33 As Double, t44 As Double
  Dim aa As Double, bb As Double, cc As Double

  If TermePaiement = EcheancePaiement.ePaiementAvance Then
    
    ' paiement d'avance -> fct ä
    t1 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n), anNais, m_fWait)
    t2 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n), anNais, m_fWait)
    t3 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n) + 1, anNais, m_fWait)
    t4 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n) + 1, anNais, m_fWait)
     
    t11 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t22 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t33 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
    t44 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
      
  Else
    
    ' paiement échu -> fct a
    t1 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n), anNais, m_fWait)
    t2 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n), anNais, m_fWait)
    t3 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n) + 1, anNais, m_fWait)
    t4 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n) + 1, anNais, m_fWait)
     
    t11 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t22 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t33 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
    t44 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
  
  End If

  a = m_ModuleCalcul_Provisions.Interpolation(t1, t2, x)
  b = m_ModuleCalcul_Provisions.Interpolation(t3, t4, x)
  c = m_ModuleCalcul_Provisions.Interpolation(a, b, n)
   
  aa = m_ModuleCalcul_Provisions.Interpolation(t11, t22, x)
  bb = m_ModuleCalcul_Provisions.Interpolation(t33, t44, x)
  cc = m_ModuleCalcul_Provisions.Interpolation(aa, bb, n)
   
  ' interpolation du coeff sur l'age, la duree et le differé + Frais de Gestion
  CalculPMRente1Euros = Arrondi((1# + g) * m_ModuleCalcul_Provisions.Interpolation(c, cc, k), NbDecimalePM)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Rentes
'
'##ModelId=5C8A75ED02F6
Private Sub CalculRentes(ByRef rsAssure As ADODB.Recordset, bMantienDeces As Boolean, salaireAnnuel As Double, coeffBCAC As Double, PctLissage As Double, nombreEnfant As Double, ageMoyenEnfant As Double)

  Dim sTableH As String, sTableF As String, feuilleLue As String
  Dim i As Double, r As Double, g As Double, tx As Double, tx2 As Double
  Dim TermePaiement As EcheancePaiement, Fraction As Integer, anNais As Integer
  Dim bTemporaire As Boolean
  
  '   age          differe      duree
  Dim x As Double, k As Double, n As Double, dureeRevalo As Double, txRevalo As Double
  
  On Error GoTo err_CalculRentes
  
    
  '*****
  '***** DEBUG
  '*****
  
'If rsAssure.fields("NUENRP3I") = 152 Then Stop
  
  '*****
  '*****
  '*****
  
  
  ' recherche des paramètres
'  m_dataHelper.Multi_Find m_rsParamRentes, "NumParamCalcul=" & rsAssure.fields("PONumParamCalcul") _
'                                           & " AND Categorie='" & rsAssure.fields("POCATEGORIE") & "'" _
'                                           & " AND AnneeSurvenance=" & Year(rsAssure.fields("POARRET"))
'  Dim dSurv As String
'
'  dSurv = m_dataHelper.BuildDateString(rsAssure.fields("POARRET"))
'  m_dataHelper.Multi_Find m_rsParamRentes, "(DateDebut is null AND " & dSurv & "<=DateFin) " _
'                                           & " OR (DateDebut<=" & dSurv & " AND " & dSurv & "<=DateFin)" _
'                                           & " OR (DateDebut<=" & dSurv & " AND DateFin IS NULL)"
'  End If
'
'  If Not m_rsParamRentes.EOF Then
    
  If FindParamRente(rsAssure.fields("POARRET")) Then
    i = m_rsParamRentes.fields("TauxTechnique") / 100#
    g = m_rsParamRentes.fields("FraisGestion") / 100#
    r = m_rsParamRentes.fields("TauxRevalo") / 100#
    
    sTableH = m_rsParamRentes.fields("TableHomme")
    sTableF = m_rsParamRentes.fields("TableFemme")
  Else
    i = 0
    r = 0
    g = 0
    
    sTableH = ""
    sTableF = ""
    
    ' Ajoute à la liste des catégories non trouvées table ParamRentes : Année survenance, Catégorie
    m_CatNotFoundRenteEducation.AddCategorie Year(rsAssure.fields("POARRET")), rsAssure.fields("POCATEGORIE")
  End If
  
  
  ' parametres par defaut
  If i = 0 Then
    i = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXRTEDUC")) / 100#
  End If
  
  If g = 0 Then
    If bMantienDeces Then
      g = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGestionEduc_MDC")) / 100#
    Else
      g = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONEDUC")) / 100#
    End If
  End If
  
  
  rsAssure.fields("TXTECHN") = i * 100#
  rsAssure.fields("TXFRAIS") = g * 100#
  
  
  feuilleLue = IIf(rsAssure.fields("POSEXE") = 1, sTableH, sTableF)
  
  If feuilleLue = "" Then
    feuilleLue = m_rsPeriode.fields("PETRENTEEDUC")
  End If
  
  If IsNull(rsAssure.fields("POFRACT")) Then
    Select Case m_rsPeriode.fields("PEFRACTIONEDUC") ' 1: Annuel, 2: Semestriel, 3: Trimestriel, 4: Mensuel
      Case 1
        Fraction = Fractionnement.eFractionnementAnnuel
        
      Case 2
        Fraction = Fractionnement.eFractionnementSemestriel
        
      Case 3
        Fraction = Fractionnement.eFractionnementTrimestriel
      
      Case 4
        Fraction = Fractionnement.eFractionnementMensuel
    End Select
  Else
    Fraction = rsAssure.fields("POFRACT") ' 1: Annuel, 2: Semestriel, 4: Trimestriel, 12: Mensuel
  End If
  
  If IsNull(rsAssure.fields("POECHU")) Then
    TermePaiement = m_rsPeriode.fields("PEPAIEMENTEDUC") ' 1: Avance, 2: Echu
  Else
    TermePaiement = rsAssure.fields("POECHU") ' 1: Avance, 2: Echu
  End If
  
  
  ' init champs
  'rsAssure.fields("POPM_INCAP_1F") = 0
  'rsAssure.fields("POPM_PASS_1F") = 0
  'rsAssure.fields("POPM_INVAL_1F") = 0
  'rsAssure.fields("POPM_INCAP_1R") = 0
  'rsAssure.fields("POPM_PASS_1R") = 0
  'rsAssure.fields("POPM_INVAL_1R") = 0
  rsAssure.fields("POPM_RCJT_1F") = 0
  rsAssure.fields("POPM_RCJT_1R") = 0
  rsAssure.fields("POPM_REDUC_1R") = 0
  rsAssure.fields("POPM_REDUC_1F") = 0
  rsAssure.fields("POPM_REVALO") = 0
  rsAssure.fields("POPM_SORTIE") = 0
  rsAssure.fields("POPM_VAR") = 0
  rsAssure.fields("POPM_REVALO") = 0
  rsAssure.fields("POCOT_REVALO") = 0
  rsAssure.fields("POPM") = 0
  
  
  bTemporaire = (m_Posit = cdPosit_RenteCertaine Or m_Posit = cdPosit_RenteEducationTemporaire _
                 Or m_Posit = cdPosit_RenteConjointTemporaire _
                 Or m_Posit = cdPosit_RenteRetraiteTemporaire _
                 Or m_Posit = cdPosit_RenteAutreTemporaire)
  
  
  If Not IsNull(rsAssure.fields("POFIN")) Then
    If rsAssure.fields("POFIN") <= m_DateCalcul Then
      
      ' plus de PM
      Exit Sub
    
    End If
  End If
  
  
  ' x: âge à la date du calcul
  If bMantienDeces Then
    m_xCalcul = ageMoyenEnfant
  Else
    ' âge à la date du calcul (=date de fin de période)
    x = (m_DateCalcul - rsAssure.fields("PONAIS")) / 365.25
  End If
  If x < 0 Then
    x = 0
  End If
  m_xCalcul = x
  
  anNais = Year(rsAssure.fields("PONAIS"))
  
  ' k: differé à la date du calcul
  If m_DateCalcul < rsAssure.fields("PODEBUT") Then
    k = (rsAssure.fields("PODEBUT") - m_DateCalcul) / 365.25
  Else
    k = 0
  End If
  
  
  ' n: durée
  If bTemporaire Then
    If IsNull(rsAssure.fields("POFIN")) Or IsNull(rsAssure.fields("PODEBUT")) Then
      ' erreur de parametrage POFIN Null
      If Not IsNull(rsAssure.fields("POTERME")) Then
        m_Logger.EcritTraceDansLog "Erreurs de paramétrage : rente déclarée temporaire mais palier de rente POFIN=NULL, Durée POTERME-POFIN utilisée - Ligne NUMENRP3I=" & rsAssure.fields("NUENRP3I")
        n = (rsAssure.fields("POTERME") - rsAssure.fields("PODEBUT")) / 365.25
      Else
        m_Logger.EcritTraceDansLog "Erreurs de paramétrage : rente déclarée temporaire mais palier de rente POFIN=NULL et POTERME=NULL, Durée 10 ans utilisée - Ligne NUMENRP3I=" & rsAssure.fields("NUENRP3I")
        n = 10
      End If
    Else
      ' palier defini par POFIN - PODEBUT
      n = (rsAssure.fields("POFIN") - rsAssure.fields("PODEBUT")) / 365.25
    End If
  Else
    n = 200 ' valeur arbitraire
  End If
  
  
  ' taux equivalent
  tx = (1 + i) / (1 + r) - 1#   ' avec revalo
  tx2 = (1 + i) / (1 + 0) - 1#  ' sans revalo
  
  
  ' calcul de frac=(m-1)/2m
  Dim frac As Double
  
  Select Case Fraction
    Case Fractionnement.eFractionnementAnnuel
      frac = 0
    Case Fractionnement.eFractionnementSemestriel
      frac = 0.25
    Case Fractionnement.eFractionnementTrimestriel
      frac = 0.375
    Case Fractionnement.eFractionnementMensuel
      frac = 11# / 24#
    
    Case Else
      ' par defaut: normalement, il y a rejet lors de l'import
      frac = 0.375
  End Select


  '
  ' calcul avec interpollation age, durée, differé
  '

'''' ' interpolation du coeff sur l'age, la duree et le differé + Frais de Gestion
''''  If m_bRevalo Then
''''    rsAssure.fields("POPM_REDUC_1F") = CalculPMRente1Euros(TermePaiement, feuilleLue, tx2, x, k, frac, n, anNais, g, m_fWait)
''''    rsAssure.fields("POPM_REDUC_1R") = CalculPMRente1Euros(TermePaiement, feuilleLue, tx, x, k, frac, n, anNais, g, m_fWait)
''''  Else
''''    rsAssure.fields("POPM_REDUC_1F") = CalculPMRente1Euros(TermePaiement, feuilleLue, tx, x, k, frac, n, anNais, g, m_fWait)
''''    rsAssure.fields("POPM_REDUC_1R") = 0
''''  End If
'Modification ALG 10/03/2013
  
    
  If m_bRevalo Then

      dureeRevalo = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEDUREEREVALO"))
      txRevalo = m_rsPeriode.fields("PETXREVALO") / 100#
      tx = (1 + i) / (1 + txRevalo) - 1#    ' avec revalo
    
    If dureeRevalo < n Then
          rsAssure.fields("POPM_REDUC_1F") = Arrondi(CalculPMRente1Euros(TermePaiement, feuilleLue, tx2, x, k, frac, n, anNais, g, m_fWait), NbDecimalePM)
          
          
          rsAssure.fields("POPM_REDUC_1R") = 1 / (1 + txRevalo) ^ k * CalculPMRente1Euros(TermePaiement, feuilleLue, tx, x, k, frac, dureeRevalo, anNais, g, m_fWait) _
                                            + (1 + txRevalo) ^ dureeRevalo * CalculPMRente1Euros(TermePaiement, feuilleLue, tx2, x, k + dureeRevalo, frac, n - dureeRevalo, anNais, g, m_fWait)
                    
          rsAssure.fields("POPM_REDUC_1R") = Arrondi(rsAssure.fields("POPM_REDUC_1R"), NbDecimalePM)
    
    Else
      rsAssure.fields("POPM_REDUC_1F") = Arrondi(CalculPMRente1Euros(TermePaiement, feuilleLue, tx2, x, k, frac, n, anNais, g, m_fWait), NbDecimalePM)
      rsAssure.fields("POPM_REDUC_1R") = Arrondi(1 / (1 + txRevalo) ^ k * CalculPMRente1Euros(TermePaiement, feuilleLue, tx, x, k, frac, n, anNais, g, m_fWait), NbDecimalePM)
    End If

  Else
    rsAssure.fields("POPM_REDUC_1F") = Arrondi(CalculPMRente1Euros(TermePaiement, feuilleLue, tx, x, k, frac, n, anNais, g, m_fWait), NbDecimalePM)
    rsAssure.fields("POPM_REDUC_1R") = 0
  
  End If
    
  If bMantienDeces Then
    '
    ' PhM 20/09/2007 : A REVOIR...
    '
    Dim REDUC_1F As Double
    
    ' on multiplie les PM par le nb d'enfants
    ' PM = nbEnfants * PM_pour_1_enfant
    REDUC_1F = Arrondi(rsAssure.fields("POPM_REDUC_1F") * coeffBCAC, NbDecimaleCalcul)
    rsAssure.fields("POPM_REDUC_1F") = Arrondi(rsAssure.fields("POPM_REDUC_1F"), NbDecimaleCalcul)
    rsAssure.fields("POPM") = nombreEnfant * salaireAnnuel * REDUC_1F * PctLissage / 100
    rsAssure.fields("POCoeffBCAC") = coeffBCAC
    
    ' stocke les params de calcul
    rsAssure.fields("PONbEnfant") = nombreEnfant
    rsAssure.fields("POAgeMoyenEnfant") = ageMoyenEnfant
  
  Else
    
    rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_REDUC_1F")
    rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_REDUC_1F")
    rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * rsAssure.fields("POPM_REDUC_1F")
       
    If IsNull(rsAssure.fields("POPRESTATION_AN")) Then
      rsAssure.fields("POPM") = 0
      rsAssure.fields("POPM_REVALO") = 0
      rsAssure.fields("POCOT_REVALO") = 0
      rsAssure.fields("POPM_SORTIE") = 0
      rsAssure.fields("POPM_VAR") = 0
    End If
  
  End If

  If m_bRevalo Then
    rsAssure.fields("POPM_REVALO") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_REDUC_1R") - rsAssure.fields("POPM")
    If rsAssure.fields("POPM_REVALO") < 0 Then rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = rsAssure.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO))
  
    rsAssure.fields("POPM_REVALO") = Arrondi(rsAssure.fields("POPM_REVALO"), 2)
    rsAssure.fields("POCOT_REVALO") = Arrondi(rsAssure.fields("POCOT_REVALO"), 2)
  End If

  rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
  
  Exit Sub
  
err_CalculRentes:
    
    If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
 
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite le capital constitutif des Rentes
'
'##ModelId=5C8A75ED03C1
Private Sub CalculCapitalConstitutifRentes(ByRef rsAssure As ADODB.Recordset, bMantienDeces As Boolean, salaireAnnuel As Double, coeffBCAC As Double, PctLissage As Double, nombreEnfant As Double, ageMoyenEnfant As Double)
    
  Dim sTableH As String, sTableF As String, feuilleLue As String
  Dim i As Double, r As Double, g As Double, tx As Double
  Dim TermePaiement As Integer, Fraction As Integer, anNais As Integer
  Dim bTemporaire As Boolean
  
  '   age          differe      duree
  Dim x As Double, k As Double, n As Double
  
  On Error GoTo err_CalculCapitalConstitutifRentes
  
    
  '*****
  '***** DEBUG
  '*****
  
''''  If rsAssure.fields("NUENRP3I") = 33 Then Stop
  
  '*****
  '*****
  '*****
  
  
  ' recherche des paramètres
'  m_dataHelper.Multi_Find m_rsParamRentes, "NumParamCalcul=" & rsAssure.fields("PONumParamCalcul") _
'                                           & " AND Categorie='" & rsAssure.fields("POCATEGORIE") & "'" _
'                                           & " AND AnneeSurvenance=" & Year(rsAssure.fields("POARRET"))
'  Dim dSurv As String
'
'  dSurv = m_dataHelper.BuildDateString(rsAssure.fields("POARRET"))
'  m_dataHelper.Multi_Find m_rsParamRentes, "(DateDebut is null AND " & dSurv & "<=DateFin) " _
'                                           & " OR (DateDebut<=" & dSurv & " AND " & dSurv & "<=DateFin)" _
'                                           & " OR (DateDebut<=" & dSurv & " AND DateFin IS NULL)"
'  End If
'
'  If Not m_rsParamRentes.EOF Then
    
  If FindParamRente(rsAssure.fields("POARRET")) Then
    i = m_rsParamRentes.fields("TauxTechnique") / 100#
    g = m_rsParamRentes.fields("FraisGestion") / 100#
    r = m_rsParamRentes.fields("TauxRevalo") / 100#
    
    sTableH = m_rsParamRentes.fields("TableHomme")
    sTableF = m_rsParamRentes.fields("TableFemme")
  Else
    i = 0
    r = 0
    g = 0
    
    sTableH = ""
    sTableF = ""
    
    ' Ajoute à la liste des catégories non trouvées table ParamRentes : Année survenance, Catégorie
    m_CatNotFoundRenteEducation.AddCategorie Year(rsAssure.fields("POARRET")), rsAssure.fields("POCATEGORIE")
  End If
  
  
  ' parametres par defaut
  If i = 0 Then
    i = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXRTEDUC")) / 100#
  End If
  
  If g = 0 Then
    If bMantienDeces Then
      g = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGestionEduc_MDC")) / 100#
    Else
      g = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONEDUC")) / 100#
    End If
  End If
  
  
  rsAssure.fields("TXTECHN") = i * 100#
  rsAssure.fields("TXFRAIS") = g * 100#
  
  
  ' pour le capital constitutif, les assurés sont forcés à Femme lors de l'import
  'feuilleLue = IIf(rsAssure.fields("POSEXE") = 1, sTableH, sTableF)
  feuilleLue = sTableF
  
  If feuilleLue = "" Then
    feuilleLue = m_rsPeriode.fields("PETRENTEEDUC")
  End If
  
  If IsNull(rsAssure.fields("POFRACT")) Then
    Select Case m_rsPeriode.fields("PEFRACTIONEDUC") ' 1: Annuel, 2: Semestriel, 3: Trimestriel, 4: Mensuel
      Case 1
        Fraction = Fractionnement.eFractionnementAnnuel
        
      Case 2
        Fraction = Fractionnement.eFractionnementSemestriel
        
      Case 3
        Fraction = Fractionnement.eFractionnementTrimestriel
      
      Case 4
        Fraction = Fractionnement.eFractionnementMensuel
    End Select
  Else
    Fraction = rsAssure.fields("POFRACT") ' 1: Annuel, 2: Semestriel, 4: Trimestriel, 12: Mensuel
  End If
  
  If IsNull(rsAssure.fields("POECHU")) Then
    TermePaiement = m_rsPeriode.fields("PEPAIEMENTEDUC") ' 1: Avance, 2: Echu
  Else
    TermePaiement = rsAssure.fields("POECHU") ' 1: Avance, 2: Echu
  End If
  
  
  ' init champs
  'rsAssure.fields("POPM_INCAP_1F") = 0
  'rsAssure.fields("POPM_PASS_1F") = 0
  'rsAssure.fields("POPM_INVAL_1F") = 0
  'rsAssure.fields("POPM_INCAP_1R") = 0
  'rsAssure.fields("POPM_PASS_1R") = 0
  'rsAssure.fields("POPM_INVAL_1R") = 0
  rsAssure.fields("POPM_RCJT_1F") = 0
  rsAssure.fields("POPM_REDUC_1F") = 0
  rsAssure.fields("POPM_RCJT_1R") = 0
  rsAssure.fields("POPM_REDUC_1R") = 0
  rsAssure.fields("POPM_REVALO") = 0
  rsAssure.fields("POPM_SORTIE") = 0
  rsAssure.fields("POPM_VAR") = 0
  rsAssure.fields("POPM_REVALO") = 0
  rsAssure.fields("POCOT_REVALO") = 0
  rsAssure.fields("POPM") = 0
  
  
  bTemporaire = (m_Posit = cdPosit_RenteCertaine Or m_Posit = cdPosit_RenteEducationTemporaire _
                 Or m_Posit = cdPosit_RenteConjointTemporaire _
                 Or m_Posit = cdPosit_RenteRetraiteTemporaire _
                 Or m_Posit = cdPosit_RenteAutreTemporaire)
  
  
'  If Not IsNull(rsAssure.fields("POFIN")) Then
'    If rsAssure.fields("POFIN") <= m_DateCalcul Then
'
'      ' plus de PM
'      Exit Sub
'
'    End If
'  End If
'
'
  ' x: âge à la date du calcul
  If bMantienDeces Then
    m_xCalcul = ageMoyenEnfant
  Else
    ' âge à la date du calcul (=date de fin de période)
    'x = (m_DateCalcul - rsAssure.fields("PONAIS")) / 365.25
    x = (rsAssure.fields("POARRET") - rsAssure.fields("PONAIS")) / 365.25
  End If
  If x < 0 Then
    x = 0
  End If
  m_xCalcul = x
  
  anNais = Year(rsAssure.fields("PONAIS"))
  
  ' k: differé à la date du calcul
  If m_DateCalcul < rsAssure.fields("PODEBUT") Then
    'k = (rsAssure.fields("PODEBUT") - m_DateCalcul) / 365.25
    k = (rsAssure.fields("PODEBUT") - rsAssure.fields("POARRET")) / 365.25
  Else
    k = 0
  End If
  
  
  ' n: durée
  If bTemporaire Then
    If IsNull(rsAssure.fields("POFIN")) Or IsNull(rsAssure.fields("PODEBUT")) Then
      ' erreur de parametrage POFIN Null
      If Not IsNull(rsAssure.fields("POTERME")) Then
        m_Logger.EcritTraceDansLog "Erreurs de paramétrage : rente déclarée temporaire mais palier de rente POFIN=NULL, Durée POTERME-POFIN utilisée - Ligne NUMENRP3I=" & rsAssure.fields("NUENRP3I")
        n = (rsAssure.fields("POTERME") - rsAssure.fields("PODEBUT")) / 365.25
      Else
        m_Logger.EcritTraceDansLog "Erreurs de paramétrage : rente déclarée temporaire mais palier de rente POFIN=NULL et POTERME=NULL, Durée 10 ans utilisée - Ligne NUMENRP3I=" & rsAssure.fields("NUENRP3I")
        n = 10
      End If
    Else
      ' palier defini par POFIN - PODEBUT
      n = (rsAssure.fields("POFIN") - rsAssure.fields("PODEBUT")) / 365.25
    End If
  Else
    n = 200 ' valeur arbitraire
  End If
  
  
  ' taux equivalent
  tx = (1 + i) / (1 + r) - 1#
  
  ' calcul de frac=(m-1)/2m
  Dim frac As Double
  
  Select Case Fraction
    Case Fractionnement.eFractionnementAnnuel
      frac = 0
    Case Fractionnement.eFractionnementSemestriel
      frac = 0.25
    Case Fractionnement.eFractionnementTrimestriel
      frac = 0.375
    Case Fractionnement.eFractionnementMensuel
      frac = 11# / 24#
    
    Case Else
      ' par defaut: normalement, il y a rejet lors de l'import
      frac = 0.375
  End Select


  '
  ' calcul avec interpollation age, durée, differé
  '
  
  Dim xDebut As Integer

  xDebut = m_ModuleCalcul_Provisions.GetAgeMiniTableGeneration(feuilleLue, Int(x), anNais)
  If xDebut <> Int(x) Then
    m_Logger.EcritTraceDansLog "Age " & Int(x) & " à la date de l'arrêt " & Format(rsAssure.fields("POARRET"), "dd/MM/yyyy") & " introuvable dans la table " & feuilleLue & " pour le millesime " & anNais & " : l'âge " & xDebut & " est utilisé - Ligne NUMENRP3I=" & rsAssure.fields("NUENRP3I")
  End If
  
  Dim t1 As Double, t2 As Double, t3 As Double, t4 As Double
  Dim a As Double, b As Double, c As Double
  Dim t11 As Double, t22 As Double, t33 As Double, t44 As Double
  Dim aa As Double, bb As Double, cc As Double

  If TermePaiement = EcheancePaiement.ePaiementAvance Then
    
    ' paiement d'avance -> fct ä
    t1 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n), anNais, m_fWait)
    t2 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n), anNais, m_fWait)
    t3 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n) + 1, anNais, m_fWait)
    t4 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n) + 1, anNais, m_fWait)
     
    t11 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t22 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t33 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
    t44 = m_ModuleCalcul_Provisions.FctCalc_mäfract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
      
  Else
    
    ' paiement échu -> fct a
    t1 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n), anNais, m_fWait)
    t2 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n), anNais, m_fWait)
    t3 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k), frac, Int(n) + 1, anNais, m_fWait)
    t4 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k), frac, Int(n) + 1, anNais, m_fWait)
     
    t11 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t22 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n), anNais, m_fWait)
    t33 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x), Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
    t44 = m_ModuleCalcul_Provisions.FctCalc_mafract_xn(feuilleLue, tx, Int(x) + 1, Int(k) + 1, frac, Int(n) + 1, anNais, m_fWait)
  
  End If

  a = m_ModuleCalcul_Provisions.Interpolation(t1, t2, x)
  b = m_ModuleCalcul_Provisions.Interpolation(t3, t4, x)
  c = m_ModuleCalcul_Provisions.Interpolation(a, b, n)
   
  aa = m_ModuleCalcul_Provisions.Interpolation(t11, t22, x)
  bb = m_ModuleCalcul_Provisions.Interpolation(t33, t44, x)
  cc = m_ModuleCalcul_Provisions.Interpolation(aa, bb, n)
   
  ' interpolation du coeff sur l'age, la duree et le differé + Frais de Gestion
  rsAssure.fields("POPM_REDUC_1F") = Arrondi((1# + g) * m_ModuleCalcul_Provisions.Interpolation(c, cc, k), NbDecimalePM)
    
    
    
    
  If bMantienDeces Then
    
    '
    ' PhM 20/09/2007 : A REVOIR...
    '
    
    Dim REDUC_1F As Double
    
    ' on multiplie les PM par le nb d'enfants
    ' PM = nbEnfants * PM_pour_1_enfant
    REDUC_1F = Arrondi(rsAssure.fields("POPM_REDUC_1F") * coeffBCAC, NbDecimaleCalcul)
    rsAssure.fields("POPM_REDUC_1F") = Arrondi(rsAssure.fields("POPM_REDUC_1F"), NbDecimaleCalcul)
    rsAssure.fields("POPM") = nombreEnfant * salaireAnnuel * REDUC_1F * PctLissage / 100
    rsAssure.fields("POCoeffBCAC") = coeffBCAC
    
    ' stocke les params de calcul
    rsAssure.fields("PONbEnfant") = nombreEnfant
    rsAssure.fields("POAgeMoyenEnfant") = ageMoyenEnfant
  
  Else
    
    rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * rsAssure.fields("POPM_REDUC_1F")
  
  End If

  rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
  
  Exit Sub
  
err_CalculCapitalConstitutifRentes:
    
    If Not autoMode Then
      MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    Else
      m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
    End If
 
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les décès
'
'##ModelId=5C8A75EE00B4
Private Sub CalculDeces(ByRef rsAssure As ADODB.Recordset)
  On Error GoTo err_calcul
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  m_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
  If m_xTerme < 0 Then
    m_xTerme = 0
  End If
  m_xTermeDouble = m_xTerme
  
  rsAssure.fields("POPM") = 0
  
  Exit Sub
  
err_calcul:
   
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
 
  Resume Next
End Sub

'##ModelId=5C8A75EE00E2
Private Function CalculPSAPGeneric(ByRef rsAssure As ADODB.Recordset, garantie As String)

Dim rsBaremes As ADODB.Recordset
Dim surv_date As Date
Dim date_der_paiement As Date
Dim invent_date As Date
Dim date_ret As Date
Dim dateEndOfBareme As Date
Dim dateMaxIndem As Date

Dim numberDaysBareme As Integer
Dim numberDaysFromSurvToDP As Integer
Dim numberDaysFromDPToInvent As Integer
Dim numberDaysLeftFromSurvToDP As Integer
Dim numberDaysLeftFromDPToInvent As Integer
Dim numberDaysMaxFinPeriode As Integer

Dim alreadyAddedDays As Integer
Dim daysToAdd As Integer
Dim rate As Double
Dim Tj As Double
Dim first As Integer
Dim last As Integer

Dim AgeRet As Integer

Dim result As Double

On Error GoTo err_CalculPSAPGeneric


  '### test
'  Dim number As Integer
  'Number = rsAssure.fields("NUENRP3I")
  'If Number = 5 Then
  '  Number = 48
  'End If
  
  result = 0
  alreadyAddedDays = 0
  
  AgeRet = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "StatAgeRetraite", "62"))
  
  date_ret = DateAdd("yyyy", AgeRet, rsAssure.fields("PONAIS"))
  
  Tj = IIf(IsNull(rsAssure.fields("POPRESTATION_AN")), 0, rsAssure.fields("POPRESTATION_AN") / 365)
    
  surv_date = IIf(IsNull(rsAssure.fields("POARRET")), 0, rsAssure.fields("POARRET"))
  If surv_date = 0 Then
    CalculPSAPGeneric = 0
    Exit Function
  End If
  
  invent_date = IIf(IsNull(m_rsPeriode.fields("PEDATEEXT")), 0, m_rsPeriode.fields("PEDATEEXT"))
  If invent_date = 0 Then
    CalculPSAPGeneric = 0
    Exit Function
  End If
  
  'get info of all baremes from table
  Set rsBaremes = m_dataSource.OpenRecordset("SELECT * From PSAP_Baremes Where Garantie = '" & garantie & "' Order By FinPeriode", Disconnected)
  If Not rsBaremes.EOF Then
    rsBaremes.MoveLast
    numberDaysMaxFinPeriode = rsBaremes.fields("FinPeriode")
  Else
    GoTo err_CalculPSAPGeneric
  End If
  
  dateMaxIndem = DateAdd("d", numberDaysMaxFinPeriode, surv_date)
  
  If DateValue(invent_date) > DateValue(dateMaxIndem) Then
    invent_date = dateMaxIndem
  End If
  
  
  If DateValue(invent_date) > DateValue(date_ret) Then
    invent_date = date_ret
  End If
  
  'initialize number of days variables
  date_der_paiement = IIf(IsNull(rsAssure.fields("PODERNIERPAIEMENT")), DateAdd("d", -1, surv_date), rsAssure.fields("PODERNIERPAIEMENT"))
  
  Dim diffSurvDP As Integer
  diffSurvDP = Abs(DateDiff("d", surv_date, date_der_paiement) + 1)
  numberDaysFromSurvToDP = diffSurvDP
  
  numberDaysFromDPToInvent = DateDiff("d", date_der_paiement, invent_date)
  If numberDaysFromDPToInvent < 0 Then
    CalculPSAPGeneric = 0
    Exit Function
  End If
  
  numberDaysLeftFromSurvToDP = numberDaysFromSurvToDP
  numberDaysLeftFromDPToInvent = numberDaysFromDPToInvent

  
  If Not rsBaremes.EOF Then
    rsBaremes.MoveFirst
  Else
    GoTo err_CalculPSAPGeneric
  End If
  
  'iterate through all baremes for the desired garanty
  Do Until rsBaremes.EOF
  
    first = rsBaremes.fields("DebutPeriode")
    last = rsBaremes.fields("FinPeriode")
    rate = rsBaremes.fields("Taux")
    
    numberDaysBareme = last - first + 1
    If numberDaysBareme < 0 Then
      numberDaysBareme = 0
    End If
    
    If rate < 0 Or rate > 100 Then
      rate = 0
    Else
      rate = rate / 100
    End If

    'adjust how many days are left from Surv until DDP -- adjust also the number of days in the Bareme
    If numberDaysLeftFromSurvToDP > 0 Then
      If numberDaysBareme >= numberDaysLeftFromSurvToDP Then
        numberDaysBareme = numberDaysBareme - numberDaysLeftFromSurvToDP
        numberDaysLeftFromSurvToDP = 0
      Else
        numberDaysLeftFromSurvToDP = numberDaysLeftFromSurvToDP - numberDaysBareme
        numberDaysBareme = 0
      End If
    End If
    
    'if the number of days left in the bareme, nothing needs to be done
    If numberDaysBareme > 0 Then
      If numberDaysBareme >= numberDaysLeftFromDPToInvent Then
        daysToAdd = numberDaysLeftFromDPToInvent
        numberDaysLeftFromDPToInvent = 0
      Else
        numberDaysLeftFromDPToInvent = numberDaysLeftFromDPToInvent - numberDaysBareme
        daysToAdd = numberDaysBareme
      End If
      
      'update the calc
      result = result + daysToAdd * rate * Tj
    End If
  
    rsBaremes.MoveNext
    DoEvents
  Loop
  
  CalculPSAPGeneric = result
  
  Exit Function
  
err_CalculPSAPGeneric:
    
  CalculPSAPGeneric = 0
    
  If Not autoMode Then
    MsgBox "erreur dans CalculPSAPGeneric " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "erreur dans CalculPSAPGeneric " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite la PSAP
'
'##ModelId=5C8A75EE0140
Private Sub CalculPSAP(ByRef rsAssure As ADODB.Recordset, m_bP3I_Individuel As Boolean, bForReformeRetraite As Boolean)
Dim Duree As Double
Dim x As Integer
Dim anc As Integer
Dim anc_regl As Integer
Dim feuilleLue As String
Dim feuilleLuePassage As String
Dim correctif As Double

Dim CalculPSAP As String
Dim bCalculPSAP As Boolean

'Calcul Statutaire
Dim PSAP As Double
Dim typeSinistre As String

' If rsAssure.fields("NUENRP3I") = 10 Then Stop
 
' activation du calcul de la RAP (reste à payer)
bCalculPSAP = True ' calcul de la RAP actif par défaut

If m_bP3I_Individuel = True Then
  If Not IsNull(rsAssure.fields("POInspecteur")) Then
      CalculPSAP = UCase(Left(Trim(rsAssure.fields("POInspecteur")), 5))
      'CalculPSAP = Large(rsAssure.fields("POInspecteur"))
      
      CalculPSAP = rsAssure.fields("POInspecteur")
         If (CalculPSAP = "RAP=0" Or CalculPSAP = "RAP=o") Then
          bCalculPSAP = False
         End If
  End If
End If

  Select Case m_Posit
  
    Case cdPosit_Deces ' = 6
        
        'rsAssure.fields("POPSAP") inchangé déjà chargé dans l'import
        '
        'RECALCUL spécifique pour P3I Individuel
        '
        If m_bP3I_Individuel = True Then
            
            correctif = 1
            rsAssure.fields("POCorrectif") = correctif
            
            If rsAssure.fields("POPRESTATIONTOTAL") = 0 Then
                rsAssure.fields("POPSAP") = rsAssure.fields("POPRESTATION_AN")
            Else
                rsAssure.fields("POPSAP") = rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATIONTOTAL")
            End If
          
            If rsAssure.fields("POPSAP") < 0 Then
            rsAssure.fields("POPSAP") = 0
            End If
        
        Else
        
        ' PSAP décès P3I collective déjà chargé DANS LE MODULE IMPORT
        
        End If
        
    Case cdPosit_Statutaire ' = 5000
      
      'this is a numeric value that needs to be translated into the corresponding guarantee
      typeSinistre = Trim$(rsAssure.fields("POGARCLE_NEW")) ' P3IProvcoll : CDGARAN  => POGARCLE_NEW (in assure)
      
      If IsNumeric(typeSinistre) Then
        m_rsStatutaireGarantie.filter = " GarantieCode = " & typeSinistre
        If Not m_rsStatutaireGarantie.EOF Then
          typeSinistre = IIf(IsNull(m_rsStatutaireGarantie.fields("Garantie")), "ND", Trim$(m_rsStatutaireGarantie.fields("Garantie")))
        Else
          'Log: no garanty code available
          m_Logger.EcritTraceDansLog "Le code garantie : " & typeSinistre & " n'est pas défini !"
          Exit Sub
        End If
      End If
      
      '### add additional cases
      If typeSinistre = "ATMP" Then
        typeSinistre = "AT"
      End If
      
      PSAP = CalculPSAPGeneric(rsAssure, typeSinistre)
      
      rsAssure.fields("POPSAP") = PSAP
      
      If rsAssure.fields("POPSAP") < 0 Then
        rsAssure.fields("POPSAP") = 0
      End If
     
           
    Case Else ' dans tous les autres cas, calcul avec l'annualisation et 365.25
  
      If m_bP3I_Individuel = True Then ' P3I INDIVIDUEL calcul de la Duree restante avec date fin POFIN ou date d'ARRET PODERNIERPAIEMENT

          Duree = 0
          If Not IsNull(rsAssure.fields("POFIN")) Then ' date dernier réglement
            Duree = DateDiff("d", rsAssure.fields("POFIN"), m_DateCalcul)
          Else
              If Not IsNull(rsAssure.fields("POARRET")) Then ' date d'arrêt
                Duree = DateDiff("d", rsAssure.fields("POARRET"), m_DateCalcul)
              End If
          End If
          
          If Duree < 0# Then
            Duree = 0#
          End If
      
          If bCalculPSAP = True Then
           rsAssure.fields("POPSAP") = Arrondi(rsAssure.fields("POPRESTATION_AN") * Duree / 365.25, 2)
           If rsAssure.fields("POPSAP") < 0# Then
            rsAssure.fields("POPSAP") = 0#
           End If
          End If
      
      Else ' P3I COLLECTIVE calcul de la Duree restante avec date fin POFIN ou date d'ARRET pour 'incap Inval  sinon avec la date fin POFIN uniquement
          
          
          If m_Posit = cdPosit_IncapAvecPassage _
          Or m_Posit = cdPosit_Inval _
          Or m_Posit = cdPosit_IncapSansPassage _
          Or m_Posit = cdProvision_AvecIncap_AvecPassage_AvecInval _
          Or m_Posit = cdProvision_AvecIncap_AvecPassage_SansInval _
          Or m_Posit = cdProvision_AvecIncap_SansPassage_AvecInval _
          Or m_Posit = cdProvision_AvecIncap_SansPassage_SansInval _
          Or m_Posit = cdProvision_SansIncap_AvecPassage_AvecInval _
          Or m_Posit = cdProvision_SansIncap_AvecPassage_SansInval _
          Or m_Posit = cdProvision_SansIncap_SansPassage_AvecInval _
          Or m_Posit = cdProvision_SansIncap_SansPassage_SansInval _
          Or m_Posit = cdProvision_Incap_AvecPassage_Viager _
          Or m_Posit = cdProvision_Inval_Viager _
          Then ' calcul avec la durée POFIN

            Duree = 0
            
            If Not IsNull(rsAssure.fields("POFIN")) Then '
              Duree = DateDiff("d", rsAssure.fields("POFIN"), m_DateCalcul)
            
            Else
            
              If Not IsNull(rsAssure.fields("POARRET")) Then '
              Duree = DateDiff("d", rsAssure.fields("POARRET"), m_DateCalcul)
              End If
              
            End If
            
            If Duree < 0# Then
              Duree = 0#
            End If

            rsAssure.fields("POPSAP") = Arrondi(rsAssure.fields("POPRESTATION_AN") * Duree / 365.25, 2)
            If rsAssure.fields("POPSAP") < 0# Then
            rsAssure.fields("POPSAP") = 0#
            End If

          End If

      End If
''''

    ' pour P3I Individuel PSAP = PSAP * coef abattement * (1+gestion)
    
    If m_bP3I_Individuel = True Then
      
      If IsNull(rsAssure.fields("POPSAP")) Then
      rsAssure.fields("POPSAP") = 0
      End If
      
      If IsNull(rsAssure.fields("POCorrectif")) Then
      rsAssure.fields("POCorrectif") = 0
      End If
      
      If IsNull(rsAssure.fields("POPSAP")) Then
      rsAssure.fields("POPSAP") = 0
      End If
      
      If IsNull(rsAssure.fields("TXFRAIS")) Then
      rsAssure.fields("TXFRAIS") = 0
      End If
      If bCalculPSAP = True Then
       rsAssure.fields("POPSAP") = Arrondi(rsAssure.fields("POPSAP") * rsAssure.fields("POCorrectif") * (1 + rsAssure.fields("TXFRAIS") / 100), 2)
        If rsAssure.fields("POPSAP") < 0# Then
         rsAssure.fields("POPSAP") = 0#
        End If
      Else
       rsAssure.fields("POPSAP") = 0
      End If
    
    End If
  
  End Select

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite la réassurance
'
'##ModelId=5C8A75EE01AE
Private Sub CalculReassurance(ByRef rsAssure As ADODB.Recordset)
  '****
  '****
  '**** Reassurance  ***
  '****
  '****
  Dim tauxReass As Double
  Dim bFound As Boolean
  Dim rq As String
  
  On Error GoTo err_calcul
  
  ' init
  tauxReass = 0#
  bFound = False

  rsAssure.fields("POTRAITE_RASSUR") = ""
  rsAssure.fields("POTRAITE_RASSUR2") = ""
  rsAssure.fields("POTRAITE_RASSUR_TAUX") = 0#
  rsAssure.fields("POPRESTA_RASSUR") = 0
  rsAssure.fields("POPM_RASSUR") = 0
  rsAssure.fields("POPSAP_RASSUR") = 0

' PHM 06/02/2004
  ' cherche le taux de reassurance par regime/categorie/effet-resiliation
  ' NON ( effet <= premier paiement <= resiliation )
  ' ( effet <= date sinistre <= resiliation )
' PHM 06/02/2004

  'rq = "SELECT * FROM Reassurance WHERE GroupeCle=" & m_NumGroupe & " AND NumPeriode=" & m_NumPeriode
  rq = " Regime=" & IIf(rsAssure.fields("POGARCLE") > cdOffsetGar_DC, rsAssure.fields("POGARCLE") - cdOffsetGar_DC, rsAssure.fields("POGARCLE") - 50)
  If Not IsNull(rsAssure.fields("POCATEGORIE")) Then
    rq = rq & " AND Categorie='" & rsAssure.fields("POCATEGORIE") & "'"
  End If
'  rq = rq & " AND Effet <=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET")) _
'      & " AND (" & m_dataHelper.BuildIsNull("Resiliation") & " Or Resiliation >=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET")) & ")"
'  rq = rq & " AND Effet <=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET")) _
'      & " AND Resiliation >=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET"))
  rq = rq & " AND Effet <='" & Format(rsAssure.fields("POARRET"), "yyyy-mm-dd") & "'" _
      & " AND Resiliation >='" & Format(rsAssure.fields("POARRET"), "yyyy-mm-dd") & "'"

  'rq = rq & " ORDER BY Effet, Resiliation, NCA DESC, AnneeSinistre"

  'Set rsTaux_Reassurance = m_dataSource.OpenRecordset(rq, Disconnected)
'  Set rsTaux_Reassurance = m_dataSource.OpenRecordset(rq, Dynamic) ' ATTENTION : Snapshot est trop lent

  ' recherche NCA et annee du sinistre
  m_dataHelper.Multi_Find m_rsTaux_Reassurance, rq & " AND NCA='" & rsAssure.fields("POCONVENTION") & "' AND AnneeSinistre=" & Year(rsAssure.fields("POARRET"))
  If Not m_rsTaux_Reassurance.EOF Then
    m_rsTaux_Reassurance.Sort = "Effet, Resiliation, NCA DESC, AnneeSinistre"
    tauxReass = m_rsTaux_Reassurance.fields("Taux")
    bFound = True
  Else
    ' recherche NCA et annee du sinistre=0
    m_dataHelper.Multi_Find m_rsTaux_Reassurance, rq & " AND NCA='" & rsAssure.fields("POCONVENTION") & "' AND AnneeSinistre=0"
    If Not m_rsTaux_Reassurance.EOF Then
      m_rsTaux_Reassurance.Sort = "Effet, Resiliation, NCA DESC, AnneeSinistre"
      tauxReass = m_rsTaux_Reassurance.fields("Taux")
      bFound = True
    Else
      ' recherche NCA=Null et annee du sinistre
      m_dataHelper.Multi_Find m_rsTaux_Reassurance, rq & " AND NCA='' AND AnneeSinistre=" & Year(rsAssure.fields("POARRET"))
      If Not m_rsTaux_Reassurance.EOF Then
        m_rsTaux_Reassurance.Sort = "Effet, Resiliation, NCA DESC, AnneeSinistre"
        tauxReass = m_rsTaux_Reassurance.fields("Taux")
        bFound = True
      Else
        ' recherche NCA=Null et annee du sinistre=0
        m_dataHelper.Multi_Find m_rsTaux_Reassurance, rq & " AND NCA='' AND AnneeSinistre=0"
        If Not m_rsTaux_Reassurance.EOF Then
          m_rsTaux_Reassurance.Sort = "Effet, Resiliation, NCA DESC, AnneeSinistre"
          tauxReass = m_rsTaux_Reassurance.fields("Taux")
          bFound = True
'        Else
'          ' PhM 30/08/2006 : Recherche avec ReprisePassif
'          m_dataHelper.Multi_Find m_rsTaux_Reassurance, rq & " AND ReprisePassif=1 AND (" & m_dataHelper.BuildIsNull("Resiliation") & " Or Resiliation >=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET")) & ")"
'          If Not m_rsTaux_Reassurance.EOF Then
'            tauxReass = m_rsTaux_Reassurance.fields("Taux")
'            bFound = True
'          End If
        End If
      End If
    End If
  End If

  If bFound = True Then
    rsAssure.fields("POTRAITE_RASSUR") = m_rsTaux_Reassurance.fields("NomTraite") & " " & tauxReass * 100# & "%"
    rsAssure.fields("POTRAITE_RASSUR2") = m_rsTaux_Reassurance.fields("NomTraite")
    rsAssure.fields("POTRAITE_RASSUR_TAUX") = tauxReass
    rsAssure.fields("POCAPTIVE") = m_rsTaux_Reassurance.fields("Captive")
  End If

  'rsTaux_Reassurance.Close
  
  ' calcul quote-part
  If Not IsNull(rsAssure.fields("POPRESTATION")) Then
    rsAssure.fields("POPRESTA_RASSUR") = Arrondi(rsAssure.fields("POPRESTATION") * tauxReass, NbDecimalePM)
  End If
  rsAssure.fields("POPM_RASSUR") = Arrondi(IIf(IsNull(rsAssure.fields("POPM")), 0, rsAssure.fields("POPM") * tauxReass), 2)
  rsAssure.fields("POPSAP_RASSUR") = Arrondi(rsAssure.fields("POPSAP") * tauxReass, 2)

#If P3I_GENERALI = 1 Then
' PHM 10/07/2006
  If m_Posit = cdPosit_IncapAvecPassage _
     Or m_Posit = cdPosit_IncapSansPassage Then ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap

    If rsAssure.fields("POPMReassAvecCorrectif") = True Then
      If Not IsNull(rsAssure.fields("POPM_AvecCorrectif")) Then
        rsAssure.fields("POPM_RASSUR") = Arrondi(rsAssure.fields("POPM_AvecCorrectif") * tauxReass, NbDecimalePM)
      End If
    Else
      If Not IsNull(rsAssure.fields("POPM_SansCorrectif")) Then
        rsAssure.fields("POPM_RASSUR") = Arrondi(rsAssure.fields("POPM_SansCorrectif") * tauxReass, NbDecimalePM)
      End If
    End If

  End If
' PHM 10/07/2006
#End If
  
  Exit Sub
  
err_calcul:
   
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les dossiers clos
'
'##ModelId=5C8A75EE01DC
Private Sub DossierClot(ByRef rsAssure As ADODB.Recordset)
  Dim x As Integer
  
  On Error GoTo err_calcul
  
  If CBool(rsAssure.fields("PODOSSIERCLOS")) = True Then
    ' sauvegarde des valeurs
    rsAssure.fields("POPMANNULEE") = rsAssure.fields("POPM")
    
    If IsNull(rsAssure.fields("POCOMMENTANNUL")) Then
      rsAssure.fields("POCOMMENTANNUL") = "PM annulée car dossier clos"
    Else
      x = InStr(rsAssure.fields("POCOMMENTANNUL"), "PM annulée car dossier clos")
      If x = 0 Then
        rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- "
        rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & "PM annulée car dossier clos"
      End If
    End If
    
    rsAssure.fields("POPSAPANNULEE") = rsAssure.fields("POPSAP")
    
    ' annulations
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_RASSUR") = 0
    rsAssure.fields("POPSAP") = 0
    rsAssure.fields("POPSAP_RASSUR") = 0
  End If
  
  Exit Sub
  
err_calcul:
   
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recherche les regimes/categories maintiens deces à generer
' renvoi True si il y a quelque chose à calculer
'
'##ModelId=5C8A75EE0218
Private Function FindRegimeCategorieMaintienDeces(ByRef RegimeIncap As Long, ByRef CategorieIncap As String, ByRef gDC As GarantieDC) As Boolean
  Dim rq As String
  
  On Error GoTo err_FindRegimeCategorieMaintienDeces
  
  FindRegimeCategorieMaintienDeces = False
  
  If RegimeIncap = 0 Or CategorieIncap = 0 Then
    Err.Raise -1, "FindRegimeCategorieMaintienDeces", "Regime/Categorie ne peut être null"
    Exit Function
  End If
  
  ' table de correspondance des regimes/categories Incap -> regime/categorie DC, ret educ, rte cjt
  rq = "RegimeIncap=" & RegimeIncap & " AND CategorieIncap='" & CategorieIncap & "'"
  m_dataHelper.Multi_Find m_rsCorrespondanceRegimeCategorie, rq
  
  If m_rsCorrespondanceRegimeCategorie.EOF Then
    m_CatNotFoundCorrespondanceGarantie.AddCategorie RegimeIncap, CategorieIncap
  Else
    
    gDC.RegimeIncap = m_rsCorrespondanceRegimeCategorie.fields("RegimeIncap")
    gDC.CategorieIncap = m_rsCorrespondanceRegimeCategorie.fields("CategorieIncap")
    gDC.RegimeDeces = m_rsCorrespondanceRegimeCategorie.fields("RegimeDeces")
    gDC.CategorieDeces = m_rsCorrespondanceRegimeCategorie.fields("CategorieDeces")
    gDC.RegimeRenteConjointTemporaire = m_rsCorrespondanceRegimeCategorie.fields("RegimeRenteConjointTemporaire")
    gDC.CategorieRenteConjointTemporaire = m_rsCorrespondanceRegimeCategorie.fields("CategorieRenteConjointTemporaire")
    gDC.RegimeRenteConjointViagere = m_rsCorrespondanceRegimeCategorie.fields("RegimeRenteConjointViagere")
    gDC.CategorieRenteConjointViagere = m_rsCorrespondanceRegimeCategorie.fields("CategorieRenteConjointViagere")
    gDC.RegimeRenteEduc = m_rsCorrespondanceRegimeCategorie.fields("RegimeRenteEduc")
    gDC.CategorieRenteEduc = m_rsCorrespondanceRegimeCategorie.fields("CategorieRenteEduc")
    gDC.LissageProvision = m_rsCorrespondanceRegimeCategorie.fields("LissageProvision")
    
    FindRegimeCategorieMaintienDeces = True
  End If
  
  Exit Function

err_FindRegimeCategorieMaintienDeces:
  If Err = 94 Then ' Invalid use of Null
    Resume Next
  Else
    
    If Not autoMode Then
      MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
    Else
      m_Logger.EcritTraceDansLog "Erreur " & Err & vbLf & Err.Description
    End If
  End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' n'affecte la valeur que si le champ est non null
' si le champ vaut -1, on affecte 0 (pas de garantie)
'
'##ModelId=5C8A75EE027B
Private Sub AffecteSiNonNull(champ As field, ByRef var As Long)
  If Not IsNull(champ) Then
    If champ <> 0 Then
      If champ = -1 Then
        var = 0
      Else
        var = champ
      End If
    End If
  End If
End Sub

'##ModelId=5C8A75EE02D4
Private Sub AffecteSiNonNullString(champ As field, ByRef var As String)
  If Not IsNull(champ) Then
    If champ <> 0 Then
      If champ = -1 Then
        var = 0
      Else
        var = champ
      End If
    End If
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' rempli gDC avec les garanties souscrite par l'assure
'
'##ModelId=5C8A75EE0322
Private Sub FillGarantieMaintienDeces(ByRef rsAssure As ADODB.Recordset, ByRef RegimeIncap As Long, ByRef CategorieIncap As String, ByRef gDC As GarantieDC)
  Dim rq As String
  
  On Error GoTo err_FillGarantieMaintienDeces
  
  ' prend les données de la table GarantieDC
  FindRegimeCategorieMaintienDeces RegimeIncap, CategorieIncap, gDC
  
  ' place les données de l'import
  AffecteSiNonNull rsAssure.fields("PORegimeDeces"), gDC.RegimeDeces
  AffecteSiNonNullString rsAssure.fields("POCategorieDeces"), gDC.CategorieDeces
  
  AffecteSiNonNull rsAssure.fields("PORegimeRenteConjointTempo"), gDC.RegimeRenteConjointTemporaire
  AffecteSiNonNullString rsAssure.fields("POCategorieRenteConjointTempo"), gDC.CategorieRenteConjointTemporaire
  
  AffecteSiNonNull rsAssure.fields("PORegimeRenteConjointViagere"), gDC.RegimeRenteConjointViagere
  AffecteSiNonNullString rsAssure.fields("POCategorieRenteConjointViagere"), gDC.CategorieRenteConjointViagere
  
  AffecteSiNonNull rsAssure.fields("PORegimeRenteEduc"), gDC.RegimeRenteEduc
  AffecteSiNonNullString rsAssure.fields("POCategorieRenteEduc"), gDC.CategorieRenteEduc
  
  ' table de def des garanties DC
  rq = "Regime=" & gDC.RegimeDeces & " AND Categorie=" & gDC.CategorieDeces
  m_dataHelper.Multi_Find m_rsFindRegimeCategorieMaintienDeces, rq
  
  If m_rsFindRegimeCategorieMaintienDeces.EOF Then
    m_CatNotFoundGarantieDC.AddCategorie gDC.RegimeDeces, gDC.CategorieDeces
  Else
    gDC.CVD0 = m_rsFindRegimeCategorieMaintienDeces.fields("CVD0")
    gDC.M0 = m_rsFindRegimeCategorieMaintienDeces.fields("M0")
    gDC.CVD1 = m_rsFindRegimeCategorieMaintienDeces.fields("CVD1")
    gDC.M1 = m_rsFindRegimeCategorieMaintienDeces.fields("M1")
    gDC.MajorationEnfant = m_rsFindRegimeCategorieMaintienDeces.fields("MajorationEnfant")
    gDC.MajorationAccident = m_rsFindRegimeCategorieMaintienDeces.fields("MajorationAccident")
  End If
  
  Exit Sub

err_FillGarantieMaintienDeces:
  If Err = 94 Then ' Invalid use of Null
    Resume Next
  Else
    
    If Not autoMode Then
      MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
    Else
      m_Logger.EcritTraceDansLog "Erreur " & Err & vbLf & Err.Description
    End If
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' recherche la situation de famille, le nombre d'enfants et
' l'age moyen des enfants
'
'##ModelId=5C8A75EE03BE
Private Function FindSitutationFamille(Age As Integer, bIsCadre As Boolean) As Boolean
  FindSitutationFamille = True
  
  If bIsCadre = True Then
    If Age < m_AgeSituationFamille_AgeMini_Cadre Then
      Age = m_AgeSituationFamille_AgeMini_Cadre
    ElseIf Age > m_AgeSituationFamille_AgeMaxi_Cadre Then
      Age = m_AgeSituationFamille_AgeMaxi_Cadre
    End If
  Else
    If Age < m_AgeSituationFamille_AgeMini_NonCadre Then
      Age = m_AgeSituationFamille_AgeMini_NonCadre
    ElseIf Age > m_AgeSituationFamille_AgeMaxi_NonCadre Then
      Age = m_AgeSituationFamille_AgeMaxi_NonCadre
    End If
  End If
  
  m_dataHelper.Multi_Find m_rsAgeSituationFamille, "Age=" & Age & " AND IsCadre=" & IIf(bIsCadre, "1", "0")
  
  If m_rsAgeSituationFamille.EOF Then
    m_AgeSituationFamilleNotFound.AddCategorie IIf(bIsCadre, 1, 0), CLng(Age)
    FindSitutationFamille = False
  End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Calcul le taux de garantie en fontion de la situation de famille
' et du nombre d'enfants, avec estimation si les données sont absente.
'
'##ModelId=5C8A75EF0024
Private Function CalculTauxGaranti(ByRef rsAssureTemp As ADODB.Recordset, gDC As GarantieDC, Age As Integer, bIsCadre As Boolean) As Double
  CalculTauxGaranti = 0
  
  If IsNull(rsAssureTemp.fields("POCleSituationFamille")) Then
    If FindSitutationFamille(Age, bIsCadre) Then
      ' taux garanti = somme des probas situation de famille * taux table GarantieDC
      CalculTauxGaranti = (m_rsAgeSituationFamille.fields("CVD0") / 100#) * (gDC.CVD0 / 100#)
      CalculTauxGaranti = CalculTauxGaranti + (m_rsAgeSituationFamille.fields("CVD1") / 100#) * (gDC.CVD1 / 100# _
                                                                  + m_rsAgeSituationFamille.fields("NombreEnfant") _
                                                                    * gDC.MajorationEnfant / 100#)
      CalculTauxGaranti = CalculTauxGaranti + (m_rsAgeSituationFamille.fields("M0") / 100#) * (gDC.M0 / 100#)
      CalculTauxGaranti = CalculTauxGaranti + (m_rsAgeSituationFamille.fields("M1") / 100#) * (gDC.M1 / 100# _
                                                                  + m_rsAgeSituationFamille.fields("NombreEnfant") _
                                                                    * gDC.MajorationEnfant / 100#)
    End If
  Else
    ' taux garanti =  taux table GarantieDC pour la situation de famille de l'assuré
    Select Case rsAssureTemp.fields("POCleSituationFamille")
      Case 1  'CVD0
        If Not IsNull(rsAssureTemp.fields("POTauxGarantieDC")) Then
          CalculTauxGaranti = rsAssureTemp.fields("POTauxGarantieDC") / 100#
        Else
          CalculTauxGaranti = gDC.CVD0 / 100#
        End If
      Case 2 'M0
        If Not IsNull(rsAssureTemp.fields("POTauxGarantieDC")) Then
          CalculTauxGaranti = rsAssureTemp.fields("POTauxGarantieDC") / 100#
        Else
          CalculTauxGaranti = gDC.M0 / 100#
        End If
      Case 3 'CVD1
        If Not IsNull(rsAssureTemp.fields("POTauxGarantieDC")) Then
          CalculTauxGaranti = rsAssureTemp.fields("POTauxGarantieDC") / 100#
        Else
          CalculTauxGaranti = gDC.CVD1 / 100#
        End If
        If Not IsNull(rsAssureTemp.fields("PONbEnfant")) And Not IsNull(rsAssureTemp.fields("POMajoEnfant")) Then
          CalculTauxGaranti = CalculTauxGaranti + rsAssureTemp.fields("PONbEnfant") * rsAssureTemp.fields("POMajoEnfant") / 100#
        End If
      Case 4 'M1
        If Not IsNull(rsAssureTemp.fields("POTauxGarantieDC")) Then
          CalculTauxGaranti = rsAssureTemp.fields("POTauxGarantieDC") / 100#
        Else
          CalculTauxGaranti = gDC.M1 / 100#
        End If
        If Not IsNull(rsAssureTemp.fields("PONbEnfant")) And Not IsNull(rsAssureTemp.fields("POMajoEnfant")) Then
          CalculTauxGaranti = CalculTauxGaranti + rsAssureTemp.fields("PONbEnfant") * rsAssureTemp.fields("POMajoEnfant") / 100#
        End If
      
      Case Else
        
        If Not autoMode Then
          MsgBox "CalculTauxGaranti() : Situation de famille inconnu !", vbCritical
        Else
          m_Logger.EcritTraceDansLog "CalculTauxGaranti() : Situation de famille inconnu !"
        End If
  
        Err.Raise -1, "CalculTauxGaranti()", "Situation de famille inconnu !"
        Exit Function
    End Select
  End If
  
  ' proba accident
  CalculTauxGaranti = CalculTauxGaranti * (1 + (m_rsPeriode.fields("PEPCTACCIDENT") / 100#) * (gDC.MajorationAccident / 100#))

  CalculTauxGaranti = Arrondi(CalculTauxGaranti, NbDecimaleCalcul)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ajoute une ligne dans la table AssureTemp
' et recopie les infos de la table Assure
'
'##ModelId=5C8A75EF00B1
Private Sub CopieRecordAssure_AssureTemp(ByRef rsAssure As ADODB.Recordset, ByRef rsAssureTemp As ADODB.Recordset, NewRegime As Long, NewCategorie As String, bForceEnInval As Boolean, bForReformeRetraite As Boolean)
  Dim f As field
  
  On Error GoTo err_CopieRecord
  
  ' ajoute un record
  rsAssureTemp.AddNew
  
'  MsgBox "ATTENTION CopieRecordAssure_AssureTemp() : " & vbLf & "Faire copie des champs necessaire pour ne pas fausser les sommes !", vbCritical
  
  ' recopie les valeurs existantes
  For Each f In rsAssure.fields
    
    If f.Name = "RECNO" Then
    
    ElseIf f.Name = "POIdAssureMDC" Then
      rsAssureTemp.fields(f.Name) = 0
    
    ElseIf f.Name <> "POPRESTATION" And f.Name <> "POPRESTATION_AN" And f.Name <> "POPRESTATION_AN_PASSAGE" And f.Name <> "POPRESTATIONTOTAL" _
        And f.Name <> "POPM" And f.Name <> "POPM_INCAP_1F" And f.Name <> "POPM_PASS_1F" _
        And f.Name <> "POPM_INVAL_1F" And f.Name <> "POPM_RCJT_1F" And f.Name <> "POPM_REDUC_1F" _
        And f.Name <> "POPM_VAR" And f.Name <> "POPSAP" And f.Name <> "POPSAPCAPMOYEN" _
        And f.Name <> "POPM_RI" And f.Name <> "POPM_REVALO" And f.Name <> "POCOT_REVALO" _
        And f.Name <> "POPM_INCAP_1R" And f.Name <> "POPM_PASS_1R" And f.Name <> "POPM_INVAL_1R" _
        And f.Name <> "POPRESTA_RASSUR" And f.Name <> "POPM_RASSUR" And f.Name <> "POPSAP_RASSUR" _
        And f.Name <> "POSIT" And f.Name <> "POMontantBase" And f.Name <> "POMontantRevalo" _
        And f.Name <> "PONbJourIndemn" And f.Name <> "POPM_X" And f.Name <> "POPM_XTERME" _
        And f.Name <> "POPM_ANC" And f.Name <> "POPM_DUREE" _
      Then
      rsAssureTemp.fields(f.Name) = f
    
    Else
      rsAssureTemp.fields(f.Name) = 0
    End If
  
  Next
  
  ' champs inutiles
  rsAssureTemp.fields("POPRESTATION") = Null
  rsAssureTemp.fields("POPRESTATION_AN") = Null
  rsAssureTemp.fields("POPRESTATION_AN_PASSAGE") = Null
  rsAssureTemp.fields("POPRESTATIONTOTAL") = Null
  rsAssureTemp.fields("POGARCLE_NEW") = Null
  rsAssureTemp.fields("POCATEGORIE_NEW") = Null
  'rsAssureTemp.fields("POCODENATURE") = Null
  rsAssureTemp.fields("PODERNIERPAIEMENT") = Null
  rsAssureTemp.fields("POPREMIER_PAIEMENT") = Null
 
  ' traite reassurance
  rsAssureTemp.fields("POTRAITE_RASSUR") = ""

  ' change le regime/categorie
  rsAssureTemp.fields("POGARCLE") = NewRegime + 50 ' on calcul avec le regime normal puis on offset en 90
  rsAssureTemp.fields("POCATEGORIE") = NewCategorie
  rsAssureTemp.fields("POSIT") = cdTypeMaintienDeces
  
  rsAssureTemp.fields("POPM_INCAP_1F") = Null
  rsAssureTemp.fields("POPM_PASS_1F") = Null
  rsAssureTemp.fields("POPM_INVAL_1F") = Null
  rsAssureTemp.fields("POPM_RCJT_1F") = Null
  rsAssureTemp.fields("POPM_REDUC_1F") = Null
  
  rsAssureTemp.fields("POPM_INCAP_1R") = Null
  rsAssureTemp.fields("POPM_PASS_1R") = Null
  rsAssureTemp.fields("POPM_INVAL_1R") = Null
  rsAssureTemp.fields("POPM_RCJT_1R") = Null
  rsAssureTemp.fields("POPM_REDUC_1R") = Null
  
  rsAssureTemp.fields("POPM_AvecCorrectif") = Null
  rsAssureTemp.fields("POPM_SansCorrectif") = Null
  
  
  ' sert à connecter MDC retraite temp avec MDC temp
  rsAssureTemp.fields("POIdAssure") = rsAssure.fields("RECNO")
  

  ' commentaire
  If bForceEnInval Then
    rsAssureTemp.fields("POCOMMENTANNUL") = "Assuré forcé en Invalidité pour le calcul des provisions Maintien en Garantie Décès"
  Else
    rsAssureTemp.fields("POCOMMENTANNUL") = ""
  End If
  
  Exit Sub
    
err_CopieRecord:
  
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Sub
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Renvoi le coefficient de provision du BCAC (recalculé ou non)
'
'##ModelId=5C8A75EF0166
Private Sub InitMinMax()
  ' BCAC Incap
  m_BCAC_CleTable_Incap = m_rsPeriode.fields("PECleTableBCACIncap_MDC")
  m_BCAC_NomTable_Incap = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))
  m_BCAC_AgeMin_Incap = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))
  m_BCAC_AgeMax_Incap = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC"))

  ' BCAC Inval
  m_BCAC_CleTable_Inval = m_rsPeriode.fields("PECleTableBCACInval_MDC")
  m_BCAC_NomTable_Inval = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
  m_BCAC_AgeMin_Inval = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
  m_BCAC_AgeMax_Inval = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC"))
  
  ' AgeSituationFamille
  m_AgeSituationFamille_AgeMini_Cadre = m_dataHelper.GetParameter("SELECT Min(Age) FROM AgeSituationFamille WHERE IsCadre=1")
  m_AgeSituationFamille_AgeMaxi_Cadre = m_dataHelper.GetParameter("SELECT MAx(Age) FROM AgeSituationFamille WHERE IsCadre=1")
  
  m_AgeSituationFamille_AgeMini_NonCadre = m_dataHelper.GetParameter("SELECT Min(Age) FROM AgeSituationFamille WHERE IsCadre=0")
  m_AgeSituationFamille_AgeMaxi_NonCadre = m_dataHelper.GetParameter("SELECT MAx(Age) FROM AgeSituationFamille WHERE IsCadre=0")

  ' Table Incap, Passage, Inval
  m_CleTable_Incap = m_rsPeriode.fields("PELMINCAP")
  m_AgeMin_Incap = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINCAP"))
  m_AgeMax_Incap = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINCAP"))
  m_CleTable_Pass = m_rsPeriode.fields("PELMPASSAGE")
  m_AgeMin_Pass = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE"))
  m_AgeMax_Pass = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE"))
  m_CleTable_Inval = m_rsPeriode.fields("PELMINVAL")
  m_AgeMin_Inval = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINVAL"))
  m_AgeMax_Inval = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINVAL"))


' Reforme des retraites
  ' BCAC Incap
  m_BCAC_CleTable_Incap_Retraite = m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite")
  m_BCAC_NomTable_Incap_Retraite = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))
  m_BCAC_AgeMin_Incap_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))
  m_BCAC_AgeMax_Incap_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACIncap_MDC_Retraite"))

  ' BCAC Inval
  m_BCAC_CleTable_Inval_Retraite = m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite")
  m_BCAC_NomTable_Inval_Retraite = m_dataHelper.GetParameter("SELECT LIBTABLE FROM ListeTableLoi WHERE TABLECLE=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
  m_BCAC_AgeMin_Inval_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
  m_BCAC_AgeMax_Inval_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM ProvisionBCAC WHERE CleTable=" & m_rsPeriode.fields("PECleTableBCACInval_MDC_Retraite"))
  
  ' Table Incap, Passage, Inval
  m_CleTable_Incap_Retraite = m_rsPeriode.fields("PELMINCAP")
  m_AgeMin_Incap_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINCAP_Retraite"))
  m_AgeMax_Incap_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINCAP_Retraite"))
  m_CleTable_Pass_Retraite = m_rsPeriode.fields("PELMPASSAGE")
  m_AgeMin_Pass_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE_Retraite"))
  m_AgeMax_Pass_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMPASSAGE_Retraite"))
  m_CleTable_Inval_Retraite = m_rsPeriode.fields("PELMINVAL")
  m_AgeMin_Inval_Retraite = m_dataHelper.GetParameter("SELECT Min(Age) FROM " & m_rsPeriode.fields("PELMINVAL_Retraite"))
  m_AgeMax_Inval_Retraite = m_dataHelper.GetParameter("SELECT Max(Age) FROM " & m_rsPeriode.fields("PELMINVAL_Retraite"))
  
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Renvoi le coefficient de provision du BCAC (recalculé ou non)
'
'##ModelId=5C8A75EF0176
Private Sub ComputeAgeEtAnciennete(ByRef rsAssure As ADODB.Recordset, ByRef Posit As Integer, ByRef AgeArret As Integer, ByRef AgeCalcul As Integer, ByRef Anciennete As Integer, ByRef bForceEnInval As Boolean, InterpolationInval As Integer)
   
  On Error GoTo err_ComputeAgeEtAnciennete
    
  bForceEnInval = False
  
  If Posit <> cdPosit_Inval Then
    
    '
    ' Incap
    '
    
    ' Age à la date de sinistre
    AgeArret = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POARRET"))
    
    ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
    Anciennete = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
    If Anciennete > 36 And m_rsPeriode.fields("PEPMGDForcerInval") = True Then
      Posit = cdPosit_Inval
      bForceEnInval = True
    End If
  End If
  
  If Posit = cdPosit_Inval Then
  
    '
    ' Inval
    '
    
    Dim DateMiseInval As Date
    If IsNull(rsAssure.fields("PODATEENTREEINVAL")) Then
      DateMiseInval = rsAssure.fields("POARRET")
    Else
      DateMiseInval = rsAssure.fields("PODATEENTREEINVAL")
    End If
    
    ' Age à la date de sinistre
    AgeArret = CalculAge(rsAssure.fields("PONAIS"), DateMiseInval)
    
    If AgeArret < 0 Then
      AgeArret = 0
    End If
    
    ' anciennté entre la (Date Entrée en Inval) et (date calcul)
'    If InterpolationInval = eInterpolationInval_AgeDuree Then
'      ' si interpolation Age/Duree, on ne tronque pas l'ancienneté on l'arrondi
'      Anciennete = CInt(NbMois(DateMiseInval, m_DateCalcul) / 12#)
'    Else
'      ' ancien code
'      Anciennete = Int(NbMois(DateMiseInval, m_DateCalcul) / 12#)
'    End If
    
    Anciennete = m_xCalcul - AgeArret
    
    If Anciennete < 0 Then
      Anciennete = 0
    End If
  End If
  
  ' age à la date de calcul
  AgeCalcul = m_xCalcul
  
  Exit Sub
  
err_ComputeAgeEtAnciennete:
  
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Renvoi le coefficient de provision du BCAC (recalculé ou non)
'
'##ModelId=5C8A75EF0241
Private Function FindCoeffBCAC(ByRef rsAssure As ADODB.Recordset, ByRef Posit As Integer, ByRef bForceEnInval As Boolean, bForRevalo As Boolean, bForReformeRetraite As Boolean, ByRef txTech As Double, ByRef txFrais As Double, ByRef AgeArret As Integer, ByRef AgeTerme As Double, ByRef Anciennete As Integer, ByRef Duree As Double) As Double
  
  Dim rq As String, AgeCalcul As Integer, AgeLimite As Integer
  Dim DateAgeLimite As Date, DateAgeTerme As Date, dateLimite As Date, DateMiseInval As Date
  Dim tx As Double
  Dim InterpolationInval As Integer
  
  
  On Error GoTo err_FindCoeffBCAC




'If rsAssure.fields("NUENRP3I") = 10 Then Stop
    
    
    
    
  AgeArret = 0
  AgeCalcul = 0
  Anciennete = 0
  
  InterpolationInval = eInterpolationInval_AgeDuree     ' m_rsPeriode.fields("PEINTERPOLATION_INVAL")
  
  ComputeAgeEtAnciennete rsAssure, Posit, AgeArret, AgeCalcul, Anciennete, bForceEnInval, InterpolationInval
  'Evolution calcul maintien garantie décès AG le 28/10/2023
  If rsAssure.fields("POCategorieInval") = 1 And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
    'Ajout le 07/11/2023 AG pour gérer les calucls des anciennes périodes
    If IsNull(m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Inval1")) Or IsNull(m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Retraite_Inval1")) Then
      AgeLimite = IIf(bForReformeRetraite = True, m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Retraite"), m_rsPeriode.fields("PEAgeLimiteCalul_MDC"))
    Else
      AgeLimite = IIf(bForReformeRetraite = True, m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Retraite_Inval1"), m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Inval1"))
    End If
  Else
    AgeLimite = IIf(bForReformeRetraite = True, m_rsPeriode.fields("PEAgeLimiteCalul_MDC_Retraite"), m_rsPeriode.fields("PEAgeLimiteCalul_MDC"))
  End If
  
  If m_rsPeriode.fields("PERecalculBCAC") = True Then
    
    '
    ' on recalcule la valeur
    '
    
    If Posit <> cdPosit_Inval Then
      
      '
      ' Incap
      '
      
      If Not IsNull(rsAssure.fields("POTERME")) Then
        DateAgeTerme = rsAssure.fields("POTERME")
      Else
        DateAgeTerme = DateSerial(Year(rsAssure.fields("PONAIS")) + 200, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
      End If
      
      DateAgeLimite = DateSerial(Year(rsAssure.fields("PONAIS")) + AgeLimite, Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
      If DateAgeTerme < DateAgeLimite Then
        dateLimite = DateAgeTerme
      Else
        dateLimite = DateAgeLimite
      End If
      
      
      ' Cas particulier : Chomage, Maternite et Mensu
      If Not IsNull(rsAssure.fields("POTERME")) And (Posit = cdPosit_Chomage Or Posit = cdPosit_Maternite Or Posit = cdPosit_Mensualisation) Then
        dateLimite = DateAgeTerme
      End If
      
      
      ' durée en mois entre date d'ARRET et date limite
      Duree = NbMois(rsAssure.fields("POARRET"), dateLimite)
      If Duree < 0 Then
        Duree = 0
      End If
      
      AgeTerme = CalculAgeDouble(rsAssure.fields("PONAIS"), dateLimite)
      
      ' taux équivalent. m_TauxRevalo est déjà / 100
      If bForRevalo Then
        tx = (((1# + m_rsPeriode.fields("PETauxTechRecalculBCAC_MDC") / 100#) / (1# + m_TauxRevalo)) - 1#) * 100#
      Else
        tx = m_rsPeriode.fields("PETauxTechRecalculBCAC_MDC")
      End If
      
      ' dans ce cas, AgeArret = Age à la date d'entrée en inval
      FindCoeffBCAC = m_ModuleCalcul_PM_MaintienDeces.PMMaintienIncap(AgeArret, Anciennete, Duree, _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINCAP_Retraite", "PELMINCAP")), _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACIncap_MDC_Retraite", "PECleTableMortBCACIncap_MDC")), _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMPASSAGE_Retraite", "PELMPASSAGE")), _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                      tx, _
                                                                      m_rsPeriode.fields("PEGestionRecalculBCAC_MDC"), _
                                                                      tx, _
                                                                      m_rsPeriode.fields("PEGestionRecalculBCAC_MDC"), _
                                                                      AgeLimite _
                                                                     )
      txTech = tx
      txFrais = m_rsPeriode.fields("PEGestionRecalculBCAC_MDC")
      
    Else
    
      '
      ' Inval
      '
      
      DateAgeTerme = rsAssure.fields("POTERME")
      DateAgeLimite = DateSerial(Year(rsAssure.fields("PONAIS")) + AgeLimite, _
                                 Month(rsAssure.fields("PONAIS")), Day(rsAssure.fields("PONAIS")))
      If DateAgeTerme < DateAgeLimite Then
        dateLimite = DateAgeTerme
      Else
        dateLimite = DateAgeLimite
      End If
      
      If IsNull(rsAssure.fields("PODATEENTREEINVAL")) Then
        DateMiseInval = rsAssure.fields("POARRET")
      Else
        DateMiseInval = rsAssure.fields("PODATEENTREEINVAL")
      End If
      
      Duree = Int(NbMois(DateMiseInval, dateLimite) / 12#)
      If Duree < 0 Then
        Duree = 0
      End If
      
      AgeTerme = CalculAgeDouble(rsAssure.fields("PONAIS"), dateLimite)
      
      ' taux équivalent
      If bForRevalo Then
        tx = (((1# + m_rsPeriode.fields("PETauxTechRecalculBCAC_MDC") / 100#) / (1# + m_TauxRevalo)) - 1#) * 100#
      Else
        tx = m_rsPeriode.fields("PETauxTechRecalculBCAC_MDC")
      End If
      
      ' dans ce cas, AgeArret = Age à la date d'entrée en inval
      FindCoeffBCAC = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret, Anciennete, Duree, _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                      m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                      tx, _
                                                                      m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                     )
                                                                      
      txTech = tx
      txFrais = m_rsPeriode.fields("PEGestionRecalculBCAC_MDC")
    
    
      '
      ' Interpollation
      '
      #If P3I_GENERALI = 1 Then
        Dim PT As Double, PT1 As Double, k As Double
        
        ' interpolation Age
        If InterpolationInval = eInterpolationInval_Age Then
          
          PT = FindCoeffBCAC
          PT1 = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret + 1, Anciennete, Duree, _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                 tx, _
                                                                 m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                )
          
          k = NbMois(rsAssure.fields("PONAIS"), DateMiseInval)
          FindCoeffBCAC = Arrondi(Interpolation(AgeArret, AgeArret + 1, PT, PT1, k / 12), NbDecimalePM)
        
        End If
      
        ' interpolation Age/Duree
        If InterpolationInval = eInterpolationInval_AgeDuree Then
          Dim PM1 As Double, PM2 As Double, PM3 As Double, PM4 As Double
          
          ' age
          PT = FindCoeffBCAC
          PT1 = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret + 1, Anciennete, Duree, _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                 tx, _
                                                                 m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                )
          
          k = NbMois(rsAssure.fields("PONAIS"), DateMiseInval) / 12#
          PM1 = Interpolation(AgeArret, AgeArret + 1, PT, PT1, k)
          
          ' duree + 1
          PT = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret, Anciennete, Duree + 1, _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                 tx, _
                                                                 m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                )
          PT1 = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret + 1, Anciennete, Duree + 1, _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                 m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                 tx, _
                                                                 m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                )
          PM3 = Interpolation(AgeArret, AgeArret + 1, PT, PT1, k)
          
          ' duree
          k = NbMois(DateMiseInval, dateLimite) / 12#
          FindCoeffBCAC = Arrondi(Interpolation(Duree, Duree + 1, PM1, PM3, k), NbDecimalePM)
        End If
      
        ' Duree restante < 12 mois
        If InterpolationInval = eInterpolationInval_CorrectionDuree Then
          Dim duree_restant_mois As Integer
          
          duree_restant_mois = DateDiff("m", m_DateCalcul, dateLimite)
          
          If duree_restant_mois > 0 And duree_restant_mois < 12 Then
            PT = FindCoeffBCAC
            PT1 = m_ModuleCalcul_PM_MaintienDeces.PMMaintienInval(AgeArret, Anciennete, Duree + 1, _
                                                                   m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PELMINVAL_Retraite", "PELMINVAL")), _
                                                                   m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableMortBCACInval_MDC_Retraite", "PECleTableMortBCACInval_MDC")), _
                                                                   tx, _
                                                                   m_rsPeriode.fields("PEGestionRecalculBCAC_MDC") _
                                                                  )
            
            k = NbMois(DateMiseInval, dateLimite) / 12#
            FindCoeffBCAC = Arrondi(Interpolation(Duree, Duree + 1, PT, PT1, k), NbDecimalePM)
          End If
        End If
      #End If
    
    
      ' durée exacte
      If bForReformeRetraite Then
        Duree = CDbl(NbMois(DateMiseInval, dateLimite)) / 12#
        If Duree < 0 Then
          Duree = 0
        End If
      End If
    
    End If ' Inval
    
  
  Else ' Coeff précalculé
    
    ' on extrait la valeur de la table ProvisionBCAC
      
    txTech = 0
    txFrais = 0
    
    AgeTerme = 0
      
    If m_bRevalo = True Then
      m_Logger.EcritTraceDansLog "Calcul Maintien Décès : revalorisation = 0 car utilisation d'un coefficient BCAC précalculé - NUENRP3I=" & rsAssure.fields("NUENRP3I")
    End If

    
    If Posit = cdPosit_Inval Then
      ' Inval
      rq = "CleTable=" & m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableBCACInval_MDC_Retraite", "PECleTableBCACInval_MDC"))
    Else
      ' Incap
      rq = "CleTable=" & m_rsPeriode.fields(IIf(bForReformeRetraite = True, "PECleTableBCACIncap_MDC_Retraite", "PECleTableBCACIncap_MDC"))
    End If
          
    m_dataHelper.Multi_Find m_rsProvisionBCAC, rq & " AND Anciennete=" & Anciennete & " AND Age=" & AgeArret
    
    If m_rsProvisionBCAC.EOF Then
      Dim AgeMin As Integer, AgeMax As Integer
          
      ' age non disponible
      If Posit = cdPosit_Inval Then
        ' Inval
        AgeMin = IIf(bForReformeRetraite = True, m_BCAC_AgeMin_Inval_Retraite, m_BCAC_AgeMin_Inval)
        AgeMax = IIf(bForReformeRetraite = True, m_BCAC_AgeMax_Inval_Retraite, m_BCAC_AgeMax_Inval)
      Else
        ' Incap
        AgeMin = IIf(bForReformeRetraite = True, m_BCAC_AgeMin_Incap_Retraite, m_BCAC_AgeMin_Incap)
        AgeMax = IIf(bForReformeRetraite = True, m_BCAC_AgeMax_Incap_Retraite, m_BCAC_AgeMax_Incap)
      End If
      
      If AgeArret < AgeMin Then
        m_dataHelper.Multi_Find m_rsProvisionBCAC, rq & " AND Anciennete=" & Anciennete & " AND Age=" & AgeMin
      ElseIf AgeArret > AgeMax Then
        m_dataHelper.Multi_Find m_rsProvisionBCAC, rq & " AND Anciennete=" & Anciennete & " AND Age=" & AgeMax
      End If
    End If
    
    If Not m_rsProvisionBCAC.EOF Then
      FindCoeffBCAC = m_rsProvisionBCAC.fields("Provision")
    Else
      If Posit <> cdPosit_Inval Then
        If Anciennete > 36 Then
          If PassageCATR9(rsAssure.fields("POCATEGORIE"), rsAssure.fields("POCONVENTION")) = True Then
            ' message : assure en incap > 36 mois avec passage. aurait du passer en Posit=2
            m_Logger.EcritTraceDansLog "Coefficient BCAC absent de la table '" _
              & IIf(Posit = cdPosit_Inval, IIf(bForReformeRetraite = True, m_BCAC_NomTable_Inval_Retraite, m_BCAC_NomTable_Inval), IIf(bForReformeRetraite = True, m_BCAC_NomTable_Incap_Retraite, m_BCAC_NomTable_Incap)) _
              & "' : Age=" & AgeArret & " Ancienneté=" & Anciennete
          Else
            ' pas de message : assure en incap > 36 mois sans passage
          End If
        End If
      Else
        ' Posit=2
        m_Logger.EcritTraceDansLog "Coefficient BCAC absent de la table '" _
          & IIf(Posit = cdPosit_Inval, IIf(bForReformeRetraite = True, m_BCAC_NomTable_Inval_Retraite, m_BCAC_NomTable_Inval), IIf(bForReformeRetraite = True, m_BCAC_NomTable_Incap_Retraite, m_BCAC_NomTable_Incap)) _
          & "' : Age=" & AgeArret & " Ancienneté=" & Anciennete & " - NUENRP3I=" & rsAssure.fields("NUENRP3I")
      End If
      
      FindCoeffBCAC = 0#
    End If
  End If
  
  Exit Function
  
err_FindCoeffBCAC:
  
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Renvoi le salaire annuel
'
'##ModelId=5C8A75EF0386
Private Function FindSalaireAnnuel(ByRef rsAssure As ADODB.Recordset, Annee As Integer, Regime As Long, Categorie As String, bLissage As Boolean, ByRef PctLissage As Double) As Double
  Dim plafondSS As Double, PctIndemnisation As Double
  
  On Error GoTo err_FindSalaireAnnuel
  
  If IsNull(rsAssure.fields("POSalaireAnnuel")) Then
    ' estimation du salaire annuel
    
    ' recupere la valeur du plafond de l'année
    m_dataHelper.Multi_Find m_rsPlafondSS, "Annee=" & Annee
    If Not m_rsPlafondSS.EOF Then
      plafondSS = m_rsPlafondSS.fields("Montant")
    Else
      m_PlafondSSNotFound.AddCategorie 0, CLng(Annee)
    End If
      
    ' recupere la valeur du plafond de l'année
    m_dataHelper.Multi_Find m_rsIndemnisationSalaireSS, "Regime=" & Regime & " AND Categorie='" & Categorie & "'"
    If Not m_rsIndemnisationSalaireSS.EOF Then
      PctIndemnisation = m_rsIndemnisationSalaireSS.fields("PourcentIndemnisation")
    Else
      m_IndemnisationNotFound.AddCategorie Regime, Categorie
    End If
    
    If PctIndemnisation <= 50# Then
      ' compromis, normalement impossible à cause des CONSTRAINTs sur la base
      FindSalaireAnnuel = plafondSS
    Else
      ' <= PASS
      If rsAssure.fields("POPRESTATION_AN") / ((PctIndemnisation / 100#) - 0.5) <= plafondSS Then
        FindSalaireAnnuel = (rsAssure.fields("POPRESTATION_AN")) / ((PctIndemnisation / 100#) - 0.5)
      Else
        FindSalaireAnnuel = plafondSS + (rsAssure.fields("POPRESTATION_AN") - plafondSS * ((PctIndemnisation / 100#) - 0.5)) / (PctIndemnisation / 100#)
      End If
    End If
  Else
    ' valeur de l'import
    FindSalaireAnnuel = rsAssure.fields("POSalaireAnnuel")
  End If
  
  ' lissage ?
  PctLissage = 100
  If Annee < 2002 And m_rsPeriode.fields("PELissageProvision") = True And bLissage = True Then
    m_dataHelper.Multi_Find m_rsLissageProvision, "Annee=" & Year(m_DateCalcul)
    
    ' 100% si annee non trouvé
    If Not m_rsLissageProvision.EOF Then
      PctLissage = m_rsLissageProvision.fields("Pourcentage")
    End If
  End If
  
  ' arrondi
  FindSalaireAnnuel = Arrondi(FindSalaireAnnuel, NbDecimalePM)
  
  Exit Function
  
err_FindSalaireAnnuel:
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Function


'##ModelId=5C8A75F0006A
Private Function FindCapitalMoyen(Annee As Integer, Annualisation As Double, Categorie As String, FamilleComptable As Long, CodeOption As String, NUENRP3I As Long, MontantCapSousRisque As Double) As Double
  Dim bTNS As Boolean
  Dim Code_Produit As Long
  Dim CodeOptionSynthetique As String
  Dim PASS As Double
  
  
  On Error GoTo err_FindCapitalMoyen
  
  
  ' PASS
  m_dataHelper.Multi_Find m_rsPlafondSS, "Annee=" & Annee
  If m_rsPlafondSS.EOF = True Then
    m_PlafondSSNotFound.AddCategorie 0, CLng(Annee)
    m_dataHelper.Multi_Find m_rsPlafondSS, "Annee=" & Year(Now)
    If m_rsPlafondSS.EOF = True Then
      m_PlafondSSNotFound.AddCategorie 0, CLng(Annee)
      PASS = 34620 ' PASS 2010
      'm_Logger.EcritTraceDansLog "PASS pour l'année " & year(now) & " introuvable dans la table PlafondSS : 34620€ utilisé (=PASS 2010) "
    Else
      PASS = m_rsPlafondSS.fields("Montant")
    End If
  Else
    PASS = m_rsPlafondSS.fields("Montant")
  End If
  
  
  ' TNS ?
  bTNS = False
'  m_dataHelper.Multi_Find m_rsCodesCat, "Code_Cat_Contrat='" & Categorie & "' AND Famille_Comptable=" & FamilleComptable
'  If m_rsCodesCat.EOF Then
'    'm_Logger.EcritTraceDansLog "Categorie " & Categorie & ", FamilleComptable " & FamilleComptable & " introuvable dans CodesCat"
'    m_CodesCatNotFound.AddCategorie Categorie, FamilleComptable
'    bTNS = False
'  Else
'    bTNS = m_rsCodesCat.fields("Code_TNS")
'  End If
'
'  ' Produits
'  m_dataHelper.Multi_Find m_rsProduits, "Code_Cat_Contrat='" & Categorie & "' AND Code_Famille_Comptable=" & FamilleComptable
'  If m_rsProduits.EOF = True Then
'    'm_Logger.EcritTraceDansLog "Categorie " & Categorie & ", FamilleComptable " & FamilleComptable & " introuvable dans Produits"
'    m_ProduitsNotFound.AddCategorie Categorie, FamilleComptable
'
'    FindCapitalMoyen = 0
'    Exit Function
'  Else
'    Code_Produit = m_rsProduits.fields("Code_Produit")
'  End If
  
  
  ' Option Synthetique
  m_dataHelper.Multi_Find m_rsCorrespondance_CatOption, "CodeOption='" & CodeOption & "'"
  If m_rsCorrespondance_CatOption.EOF = True Then
    'm_Logger.EcritTraceDansLog "CodeOption " & CodeOption & " introuvable dans Correspondance_CatOption"
    m_Correspondance_CatOptionNotFound.AddCategorie 0, CodeOption
    
    If Annualisation < (PASS / 2#) Then
      CodeOptionSynthetique = "NC"
    Else
      CodeOptionSynthetique = "C"
    End If
  
  Else
    
    CodeOptionSynthetique = m_rsCorrespondance_CatOption.fields("Cat_Opt_Synth")
  
  End If
  
 
  ' AUTRES = Règle de 50% du PASS
  If UCase(CodeOptionSynthetique) = "AUTRES" Then
    
    If Annualisation < (PASS / 2#) Then
      CodeOptionSynthetique = "NC"
    Else
      CodeOptionSynthetique = "C"
    End If
  
  ElseIf UCase(CodeOptionSynthetique) = "MORNAY" Or UCase(CodeOptionSynthetique) = "DEXIA IS" Or UCase(CodeOptionSynthetique) = "DEXIA_IS" Then
    
    ' pour MORNAY ou DEXIA IS, le Montant est spécifié dans l'import
    FindCapitalMoyen = MontantCapSousRisque
    Exit Function
    
  End If

  
  
  ' Capitaux Moyens
  Dim errMsg As String
  
  errMsg = ""
  
'  m_dataHelper.Multi_Find m_rsCapitaux_Moyens, "Cat_Opt_Synth='" & IIf(bTNS, "TNS", CodeOptionSynthetique) & "' AND Code_Produit='" & Code_Produit & "'"
  m_dataHelper.Multi_Find m_rsCapitaux_Moyens, "Cat_Opt_Synth='" & IIf(bTNS, "TNS", CodeOptionSynthetique) & "'"
  If m_rsCapitaux_Moyens.EOF = True Then
    'm_Logger.EcritTraceDansLog "Cat_Opt_Synth=" & CodeOptionSynthetique & ", Code_Produit=" & Code_Produit & " introuvable dans Capitaux_Moyens"
    errMsg = "NUENRP3I=" & NUENRP3I & " : Cat_Opt_Synth='" & CodeOptionSynthetique & "' introuvable dans Capitaux_Moyens"
    
    m_Capitaux_MoyensNotFound.AddCategorie CodeOptionSynthetique, Code_Produit
    If CodeOptionSynthetique <> "C" And CodeOptionSynthetique <> "NC" Then
      If Annualisation < (PASS / 2#) Then
        CodeOptionSynthetique = "NC"
      Else
        CodeOptionSynthetique = "C"
      End If
    End If
    
    m_dataHelper.Multi_Find m_rsCapitaux_Moyens, "Cat_Opt_Synth='" & CodeOptionSynthetique & "'"
    If m_rsCapitaux_Moyens.EOF = True Then
      m_Capitaux_MoyensNotFound.AddCategorie CodeOptionSynthetique, Code_Produit
      FindCapitalMoyen = 0
    Else
      errMsg = errMsg & "  ==>  Valeur utilisée : Cat_Opt_Synth='" & CodeOptionSynthetique & "'"
      FindCapitalMoyen = m_rsCapitaux_Moyens.fields("Capital_Moyen")
    End If

'    m_dataHelper.Multi_Find m_rsCapitaux_Moyens, "Cat_Opt_Synth='" & CodeOptionSynthetique & "' AND Code_Produit='" & Code_Produit & "'"
'    If m_rsCapitaux_Moyens.EOF = True Then
'      m_Capitaux_MoyensNotFound.AddCategorie CodeOptionSynthetique, Code_Produit
'      FindCapitalMoyen = 0
'    Else
'      'm_Logger.EcritTraceDansLog "Valeur utilisée : Cat_Opt_Synth=" & CodeOptionSynthetique & ", Code_Produit=" & Code_Produit & " introuvable dans Capitaux_Moyens"
'      errMsg = errMsg & "  ==>  Valeur utilisée : Cat_Opt_Synth=" & CodeOptionSynthetique & ", Code_Produit=" & Code_Produit & " introuvable dans Capitaux_Moyens"
'      FindCapitalMoyen = m_rsCapitaux_Moyens.fields("Capital_Moyen")
'    End If
  Else
    FindCapitalMoyen = m_rsCapitaux_Moyens.fields("Capital_Moyen")
  End If
        
  If errMsg <> "" Then
      m_MessageUnique.AddMessage errMsg
  End If

  Exit Function

err_FindCapitalMoyen:
  If Not autoMode Then
    MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description
  End If
  
  Resume Next
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Maintiens en garantie decès
'
'##ModelId=5C8A75F00144
Private Sub CalculMaintienDeces(ByRef rsAssure As ADODB.Recordset, ByRef rsAssureTemp As ADODB.Recordset, bForReformeRetraite As Boolean, rsAssureMDC As ADODB.Recordset)
  Dim gDC As GarantieDC, i As Integer, bForceEnInval As Boolean
  Dim salaireAnnuel As Double, coeffBCAC As Double, coeffBCACR As Double, PctLissage As Double
  Dim nombreEnfant As Double, ageMoyenEnfant As Double, txTech As Double, txFrais As Double, txTechR As Double, txFraisR As Double
  Dim AgeArret As Integer, AgeTerme As Double, Anciennete As Integer, Duree As Double
  Dim sSql As String
  
  On Error GoTo err_CalculMaintienDeces
  
  
  ' arret pour test
  'If rsAssure.fields("NUENRP3I") = 3394 Then Stop


  ' Doit-on calculer un maintien décès ?
'  If rsAssure.fields("POPRESTATION_AN") <> 0 And rsAssure.fields("PODOSSIERCLOS") = False And FindGarantie(rsAssure.fields("POGARCLE")).CalculMaintienDeces = True Then
'  If rsAssure.fields("POPM") <> 0 And rsAssure.fields("PODOSSIERCLOS") = False Then
  If rsAssure.fields("PODOSSIERCLOS") = False Then
    
    '***
    '*** K.ROCUET 24-10-2006 : test si on a à calculer pour cette ligne (base/revalo = 1 seule fois)
    '***
    '==> S.VARATHARAJAN 28/07/2009 : on cµalcule le MGDC sur toutes les lignes
    '==> S.VARATHARAJAN 26/04/2010 : on calcule le MGDC sur 1 seules ligne en capitaux constitutifs
    '                                et sur toutes les lignes en % des PM
   
    ' parcours les articles correspondant
    Dim IsBase As Boolean, HasBase As Boolean
    
    IsBase = IIf(rsAssure.fields("POBaseRevalo") = 0, True, False)
    
    HasBase = IsBase
    
    If IsBase = False Then
      ' PHM 26/04/2010 : on ne calcul qu'une fois le MGDC clé:  NUDOSSIERPREST + LBASSURE + DTSURVSIN + DTNAISSASS
      '                                                  soit:  PONUMCLE       + PONOM    + POARRET   + PONAIS
      Dim sNom As String
      
      If Not IsNull(rsAssure.fields("PONOM")) Then
        sNom = rsAssure.fields("PONOM")
      End If
      sNom = Replace(sNom, "'", "''")
      
      sSql = " AND POSTECLE=" & rsAssure.fields("POSTECLE")
      sSql = sSql & " AND PONUMCLE=" & rsAssure.fields("PONUMCLE")
      'sSql = sSql & " AND PONOM='" & rsAssure.fields("PONOM") & "'"
      sSql = sSql & " AND PONOM='" & sNom & "'"
      sSql = sSql & " AND POARRET='" & rsAssure.fields("POARRET") & "'"
      sSql = sSql & " AND PONAIS='" & rsAssure.fields("PONAIS") & "'"

      m_dataHelper.Multi_Find m_rsBaseRevalo, sSql

      ' recherche si la ligne de base existe
      HasBase = False
      Do Until m_rsBaseRevalo.EOF
        
        ' PHM 07/11/2011 : on genere une ligne MDC pour la revalo si l'annualisation de la base est nulle
        If m_rsBaseRevalo.fields("POBaseRevalo") = 0 Then
          If IsNull(m_rsBaseRevalo.fields("POPRESTATION_AN")) Then
            HasBase = HasBase Or False
          ElseIf m_rsBaseRevalo.fields("POPRESTATION_AN") <= 0 Then
            HasBase = HasBase Or False
          Else
            HasBase = HasBase Or True
          End If
        End If

        'HasBase = HasBase Or IIf(m_rsBaseRevalo.fields("POBaseRevalo") = 0, True, False)

        m_rsBaseRevalo.MoveNext
      Loop

      If IsBase = False And HasBase = True And m_rsPeriode.fields("PEMethodeCalcul") = eCapitauxConstitutifs Then
        Exit Sub ' le MGDC est sur la base
      Else
        HasBase = HasBase
      End If

' PHM 28/07/2009 : on calcule le MGDC sur toutes les lignes
'
'      sSql = " AND POSTECLE=" & rsAssure.fields("POSTECLE")
'      sSql = sSql & " AND POGARCLE=" & rsAssure.fields("POGARCLE")
'      sSql = sSql & " AND POCONVENTION='" & rsAssure.fields("POCONVENTION") & "'"
'      sSql = sSql & " AND POGARCLE_NEW='" & rsAssure.fields("POGARCLE_NEW") & "'"
'      sSql = sSql & " AND POCATEGORIE='" & rsAssure.fields("POCATEGORIE") & "'"
'      sSql = sSql & " AND POARRET='" & Format(rsAssure.fields("POARRET"), "yyyy-mm-dd") & "'"
'
'      m_dataHelper.Multi_Find m_rsBaseRevalo, sSql
'
'      HasBase = False
'      Do Until m_rsBaseRevalo.EOF
'        HasBase = HasBase Or IIf(m_rsBaseRevalo.fields("POBaseRevalo") = 0, True, False)
'
'        m_rsBaseRevalo.MoveNext
'      Loop
'
'      If IsBase = False And HasBase = True Then
'        Exit Sub ' le MGDC est sur la base
'      Else
'        HasBase = HasBase
'      End If
'
' PHM 28/07/2009 : on calcule le MGDC sur toutes les lignes
    End If
    
    
    '
    ' methode de calcul
    '
    
    ' Capitaux constitutifs sous risque
    If m_rsPeriode.fields("PEMethodeCalcul") = eCapitauxConstitutifs Then
    
      ' recupere les droits de l'assure
      'FillGarantieMaintienDeces rsAssure.fields("POGARCLE") - 50, rsAssure.fields("POCATEGORIE"), gDC
      
      ' coeff provision BCAC
      coeffBCAC = FindCoeffBCAC(rsAssure, m_Posit, bForceEnInval, False, bForReformeRetraite, txTech, txFrais, AgeArret, AgeTerme, Anciennete, Duree)
      If m_bRevalo Then
        coeffBCACR = FindCoeffBCAC(rsAssure, m_Posit, bForceEnInval, True, bForReformeRetraite, txTechR, txFraisR, AgeArret, AgeTerme, Anciennete, Duree)
      End If
      
      ' Age et Ancienneté
      'ComputeAgeEtAnciennete m_Posit, AgeArret, AgeCalcul, anciennete, bForceEnInval
      
      ' salaire annuel à la date de survenance (= date d'arret)
      salaireAnnuel = FindSalaireAnnuel(rsAssure, Year(rsAssure.fields("POARRET")), rsAssure.fields("POGARCLE") - 50, rsAssure.fields("POCATEGORIE"), True, PctLissage)
      
      ' droit Maintien Deces ?
      'If gDC.RegimeDeces <> 0 And gDC.CategorieDeces <> 0 Then
      Dim capitalDC As Double, tauxGaranti As Double, Annualisation As Double
      
      ' ajout ligne
      'CopieRecordAssure_AssureTemp gDC.RegimeDeces, rsAssure.fields("POCATEGORIE"), bForceEnInval
      CopieRecordAssure_AssureTemp rsAssure, rsAssureTemp, 6, rsAssure.fields("POCATEGORIE"), bForceEnInval, False
      
      ' Pour la reforme des retraites : Stocke le Recno temporaire pour pouvoir connecter la ligne MDC dans Assure
      ' à la ligne MDC dans Assure_Retraite après recopie des tables temporaires vers les tables permanentes
      If bForReformeRetraite Then
        rsAssureTemp.fields("POIdAssure") = rsAssureMDC.fields("RECNO")
      End If

      rsAssureTemp.fields("TXTECHN") = txTech
      rsAssureTemp.fields("TXFRAIS") = txFrais
      rsAssureTemp.fields("TXTECHNR") = txTechR
      rsAssureTemp.fields("TXFRAISR") = txFraisR
  
      rsAssureTemp.fields("POPM_X") = AgeArret
      rsAssureTemp.fields("POPM_XTERME") = AgeTerme
      rsAssureTemp.fields("POPM_ANC") = Anciennete
      rsAssureTemp.fields("POPM_DUREE") = Duree
  
  
        '***
        '*** K.ROCUET 24-10-2006 : Annualisation sur la base et la revalo pour les 50% du PASS
        '***
        
        ' Champs IMPORT : NumSociete , Regime,   Sinistre,     GE,           Code_Cat_Contrat, Survenance
        ' Champs P3I    : POSTECLE ,   POGARCLE, POCONVENTION, POGARCLE_NEW, POCATEGORIE,      POARRET
        
  '      sSQL = "SELECT SUM(POPRESTATION_AN) FROM ASSURE WHERE POGPECLE=" & m_NumGroupe & " AND POPERCLE=" & m_NumPeriode
  '      sSQL = sSQL & " AND POSTECLE=" & rsAssure.fields("POSTECLE")
  '      sSQL = sSQL & " AND POGARCLE=" & rsAssure.fields("POGARCLE")
  '      sSQL = sSQL & " AND POCONVENTION='" & rsAssure.fields("POCONVENTION") & "'"
  '      sSQL = sSQL & " AND POGARCLE_NEW='" & rsAssure.fields("POGARCLE_NEW") & "'"
  '      sSQL = sSQL & " AND POCATEGORIE=" & rsAssure.fields("POCATEGORIE")
  '      sSQL = sSQL & " AND POARRET=" & m_dataHelper.BuildDateString(rsAssure.fields("POARRET"))
  '        Annualisation = m_dataHelper.GetParameterAsDouble(sSQL)
        
      sSql = " AND POSTECLE=" & rsAssure.fields("POSTECLE")
      sSql = sSql & " AND POGARCLE=" & rsAssure.fields("POGARCLE")
      sSql = sSql & " AND POCONVENTION='" & rsAssure.fields("POCONVENTION") & "'"
      sSql = sSql & " AND POGARCLE_NEW='" & rsAssure.fields("POGARCLE_NEW") & "'"
      sSql = sSql & " AND POCATEGORIE='" & rsAssure.fields("POCATEGORIE") & "'"
      sSql = sSql & " AND POARRET='" & Format(rsAssure.fields("POARRET"), "yyyy-mm-dd") & "'"
      
      m_dataHelper.Multi_Find m_rsAnnualisation, sSql
      
      If Not m_rsAnnualisation.EOF Then
        Annualisation = m_rsAnnualisation.fields("SUM_POPRESTATION_AN")
      Else
        Annualisation = rsAssure.fields("POPRESTATION_AN")
      End If
      
      
      capitalDC = FindCapitalMoyen(Year(rsAssure.fields("POARRET")), Annualisation, rsAssure.fields("POCATEGORIE"), _
                                   IIf(IsNull(rsAssure.fields("POFamilleComptable")), 0, rsAssure.fields("POFamilleComptable")), _
                                   IIf(IsNull(rsAssure.fields("POCode_Option_Contrat")), "", rsAssure.fields("POCode_Option_Contrat")), _
                                   rsAssure.fields("NUENRP3I"), _
                                   IIf(IsNull(rsAssure.fields("POMontantCapSousRisque")), 0, rsAssure.fields("POMontantCapSousRisque")))
      
      rsAssureTemp.fields("POMontantCapSousRisque") = capitalDC
      
      
      
      rsAssureTemp.fields("POPM") = capitalDC * (1 + m_rsPeriode.fields("PEGestionDeces_MDC") / 100#) * coeffBCAC * PctLissage / 100#
      rsAssureTemp.fields("POPM") = Arrondi(rsAssureTemp.fields("POPM"), 2)
             
      If m_bRevalo Then
        rsAssureTemp.fields("POPM_REVALO") = capitalDC * (1 + m_rsPeriode.fields("PEGestionDeces_MDC") / 100#) * coeffBCACR * PctLissage / 100#
        rsAssureTemp.fields("POPM_REVALO") = Arrondi(rsAssureTemp.fields("POPM_REVALO") - rsAssureTemp.fields("POPM"), 2)
        If rsAssureTemp.fields("POPM_REVALO") < 0 Then rsAssureTemp.fields("POPM_REVALO") = 0
        rsAssureTemp.fields("POCOT_REVALO") = rsAssureTemp.fields("POPM_REVALO") / (1 + ((1 - (1 + m_TMO) ^ -(m_DureeLissage - 1)) / m_TMO))
      Else
        rsAssureTemp.fields("POPM_REVALO") = 0
        rsAssureTemp.fields("POCOT_REVALO") = 0
      End If
      
#If P3I_GENERALI = 1 Then
    ' PHM 23/01/2012 : Ajout du correctif pour le calcul du MDC en incap
      ' Incap
      If rsAssure.fields("POSIT") = 1 Or rsAssure.fields("POSIT") = 3 _
         Or rsAssure.fields("POSIT") = 7 Or rsAssure.fields("POSIT") = 8 Or rsAssure.fields("POSIT") = 9 Then
        
        ' Doit-on utiliser la PM avec ou sans correctif
        m_dataHelper.Multi_Find m_rsPMRetenue, "AnneeSurvenance=" & Year(rsAssure.fields("POARRET"))
          
        If Not m_rsPMRetenue.EOF And Not IsNull(rsAssure.fields("POCorrectif")) Then
          Dim correctif As Double
         
          correctif = rsAssure.fields("POCorrectif")
        
          rsAssureTemp.fields("POPM_AvecCorrectif") = rsAssureTemp.fields("POPM") * correctif
          rsAssureTemp.fields("POPM_SansCorrectif") = rsAssureTemp.fields("POPM")
          rsAssureTemp.fields("POCorrectif") = correctif
          rsAssureTemp.fields("POPMAvecCorrectif") = m_rsPMRetenue.fields("PMAvecCorrectif")
          rsAssureTemp.fields("POPMReassAvecCorrectif") = m_rsPMRetenue.fields("PMReassAvecCorrectif")
           
          If m_rsPMRetenue.fields("PMAvecCorrectif") = True Then
            rsAssureTemp.fields("POPM") = rsAssureTemp.fields("POPM_AvecCorrectif")
            
            rsAssureTemp.fields("POPM_SORTIE") = rsAssureTemp.fields("POPM_SORTIE") * correctif
            rsAssureTemp.fields("POPM_VAR") = rsAssureTemp.fields("POPM_VAR") * correctif
             
            If m_bRevalo Then
              rsAssureTemp.fields("POPM_REVALO") = rsAssureTemp.fields("POPM_REVALO") * correctif
              rsAssureTemp.fields("POCOT_REVALO") = rsAssureTemp.fields("POCOT_REVALO") * correctif
            End If
          End If
        End If
     
     End If
#End If

      ' calcul provision
'      Select Case m_rsPeriode.fields("PEMethodeCalcul")
'        Case CapitauxConstitutifsSousRisques
'          tauxGaranti = CalculTauxGaranti(rsAssureTemp, gDC, AgeCalcul, rsAssureTemp.fields("POIsCadre"))
'          capitalDC = Arrondi(salaireAnnuel * tauxGaranti, 2)
'
'          rsAssureTemp.fields("POPM") = capitalDC * (1 + m_rsPeriode.fields("PEGestionDeces_MDC") / 100#) * coeffBCAC * PctLissage / 100#
'          rsAssureTemp.fields("POPM") = Arrondi(rsAssureTemp.fields("POPM"), NbDecimalePM)
'
'        Case CotisationsExonerees
'          rsAssureTemp.fields("POPM") = 0
'          rsAssureTemp.fields("POPM") = Arrondi(rsAssureTemp.fields("POPM"), NbDecimalePM)
'          rsAssureTemp.fields("POCOMMENTANNUL") = "PM Null car méthode des Cotisations Exonérées non implémentée !"
'      End Select
      
      ' reassurance
      CalculReassurance rsAssureTemp
      
      ' affecte la garantie DC (reassurance en code normal)
      rsAssureTemp.fields("POGARCLE") = cdGarDeces_DC
      
      ' variables de calcul
      'rsAssureTemp.fields("POSalaireAnnuel") = salaireAnnuel
      rsAssureTemp.fields("POCoeffBCAC") = Arrondi(coeffBCAC, NbDecimalePM)
      If m_bRevalo Then
        rsAssureTemp.fields("POCoeffBCACR") = Arrondi(coeffBCACR, NbDecimalePM)
      Else
        rsAssureTemp.fields("POCoeffBCACR") = 0
      End If
      
      rsAssureTemp.fields("POPourcentLissage") = 100
  
      
      If rsAssure.fields("POPM") = 0 Then
        
        ' pas de PM MDC si pas de PM garantie
        rsAssureTemp.fields("POPM") = 0
        rsAssureTemp.fields("POPM_REVALO") = 0
        rsAssureTemp.fields("POCOT_REVALO") = 0
      
        rsAssureTemp.fields("POCoeffBCAC") = 0
        rsAssureTemp.fields("POCoeffBCACR") = 0
      
      End If
      
      ' PHM 23/02/2010 : Code GE spécifique
      ' Incap
      If rsAssure.fields("POSIT") = 1 Or rsAssure.fields("POSIT") = 3 _
         Or rsAssure.fields("POSIT") = 7 Or rsAssure.fields("POSIT") = 8 Or rsAssure.fields("POSIT") = 9 Then
        rsAssureTemp.fields("POGARCLE_NEW") = 6442
      
      ' Inval
      ElseIf rsAssure.fields("POSIT") = 2 Then
        rsAssureTemp.fields("POGARCLE_NEW") = 6443
      
      Else
      
        ' Code générique
        rsAssureTemp.fields("POGARCLE_NEW") = 6125
      
      End If
      ' PHM 23/02/2010 : Code GE spécifique
      
      
      'rsAssureTemp.fields("TXTECHN") = 0
      'rsAssureTemp.fields("TXFRAIS") = 0
      
      ' valide les changements
      rsAssureTemp.Update
      'End If
      
    
    ' % des provisions calculées
    ElseIf m_rsPeriode.fields("PEMethodeCalcul") = ePctProvisionCalculee Then
    
      ' ajout ligne
      CopieRecordAssure_AssureTemp rsAssure, rsAssureTemp, 6, rsAssure.fields("POCATEGORIE"), bForceEnInval, False
  
      ' Pour la reforme des retraites : Stocke le Recno temporaire pour pouvoir connecter la ligne MDC dans Assure
      ' à la ligne MDC dans Assure_Retraite après recopie des tables temporaires vers les tables permanentes
      If bForReformeRetraite Then
        rsAssureTemp.fields("POIdAssure") = rsAssureMDC.fields("RECNO")
      End If

      ' PM
      rsAssureTemp.fields("POPM") = (m_rsPeriode.fields("PEPctPMCalculeeDC") / 100#) * rsAssure.fields("POPM")
      rsAssureTemp.fields("POPM") = Arrondi(rsAssureTemp.fields("POPM"), 2)
    
      ' reassurance
      CalculReassurance rsAssureTemp
      
      ' affecte la garantie DC (reassurance en code normal)
      rsAssureTemp.fields("POGARCLE") = cdGarDeces_DC
      
      
      ' PHM 23/04/2010 : Code GE spécifique
      ' Incap
      If rsAssure.fields("POSIT") = 1 Or rsAssure.fields("POSIT") = 3 _
         Or rsAssure.fields("POSIT") = 7 Or rsAssure.fields("POSIT") = 8 Or rsAssure.fields("POSIT") = 9 Then
        rsAssureTemp.fields("POGARCLE_NEW") = 6442
      
      ' Inval
      ElseIf rsAssure.fields("POSIT") = 2 Then
        rsAssureTemp.fields("POGARCLE_NEW") = 6443
      
      Else
      
        ' Code générique
        rsAssureTemp.fields("POGARCLE_NEW") = 6125
      
      End If
      ' PHM 23/04/2010 : Code GE spécifique
      
      
      ' variables de calcul
      rsAssureTemp.fields("POCoeffBCAC") = Null
      rsAssureTemp.fields("POCoeffBCACR") = Null
      rsAssureTemp.fields("POPourcentLissage") = 100
      
      rsAssureTemp.fields("TXTECHN") = Null
      rsAssureTemp.fields("TXFRAIS") = Null
      rsAssureTemp.fields("TXTECHNR") = Null
      rsAssureTemp.fields("TXFRAISR") = Null
      
      ' valide les changements
      rsAssureTemp.Update
    
    
    ' pas de calcul
    ElseIf m_rsPeriode.fields("PEMethodeCalcul") = ePasDeCalcul Then
    
      ' rien à faire
    
    End If
  
  End If ' Dossier clos ?
  
  Exit Sub
  
err_CalculMaintienDeces:
    
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
 
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Recopie les Maintiens en garantie decès de AssureTemp vers Assure
'
'##ModelId=5C8A75F001C1
Private Sub RecopieMaintienDeces(TableDest As String, TableSource As String, bForReformeRetraite As Boolean)
  Dim rq As String
  
  On Error GoTo err_RecopieMaintienDeces
  
  rq = "INSERT INTO " & TableDest _
        & "(POGPECLE, POPERCLE, POSTECLE, PONumParamCalcul, PONUMCLE, POGARCLE, POCONVENTION, PONOM, POSEXE, POCSP, PONAIS, POEFFET, POTERME, POARRET, " _
        & " POREPRISE, POCAUSE, POTYPEF, PODELAI, POPREMIER_PAIEMENT, POPRESTATION, POPRESTATIONTOTAL, POPRESTATION_AN_PREC, " _
        & " POPRESTATION_AN, POPRESTATION_AN_PASSAGE, POSIT, POPM_INCAP_1F, POPM_PASS_1F, POPM_INVAL_1F, POPM, POPM_SORTIE, POPM_VAR, PODERNIERPAIEMENT, " _
        & " PODEBUT, POFIN, PODEBUTTOTAL, POFINTOTAL, POPSAP, POPM_RI, POPM_REVALO, POCOT_REVALO, PODATEIMPORT, POPM_INCAP_1R, " _
        & " POPM_PASS_1R, POPM_INVAL_1R, POCATEGORIE, POPM_RCJT_1F, POPM_REDUC_1F, POPRESTA_RASSUR, POPM_RASSUR, POPSAP_RASSUR, " _
        & " POTRAITE_RASSUR, POTRAITE_RASSUR2, POTRAITE_RASSUR_TAUX, PODATEPAIEMENTESTIMEE, POPSAPCAPMOYEN, PODOSSIERCLOS, " _
        & " POGARCLE_NEW, POCATEGORIE_NEW, POREGROUPEMENT, POCODENATURE, POCCN, POCODERISQUE, POPMANNULEE, POPSAPANNULEE, " _
        & " POCOMMENTANNUL, PODATEENTREEINVAL, POPM_RCJT_1R, POPM_REDUC_1R, POIsCadre, POSalaireAnnuel, POCleSituationFamille, " _
        & " POTauxGarantieDC, PONbEnfant, POAgeMoyenEnfant, POMajoEnfant, PORegimeDeces, POCategorieDeces, PORegimeRenteConjointTempo, " _
        & " POCategorieRenteConjointTempo, PORegimeRenteConjointViagere, POCategorieRenteConjointViagere, PORegimeRenteEduc, POCategorieRenteEduc, " _
        & " POPourcentLissage, POCoeffBCAC, POContractant, POSituConv, POEffetSitu, POCode_Option_Contrat, POEtablissement, PONumSinistre, POTypeMvt, " _
        & " POIJ, POMontantBase, POMontantRevalo, POTypeReglement, PODebutIndemnisation, PONbJourIndemn, POCreationDossier, PODebutDossier, " _
        & " POFinDossier, POMotifCloture, PODebutRefSalaire, POFinRefSalaire, POFamilleComptable, POInspecteur, POInsp, POApport, POApport2, " _
        & " POGestionnaire, POIndicCC, POCompagnie, PONbIntervenant, POPM_AvecCorrectif, POPM_SansCorrectif, POCorrectif, POPMReassAvecCorrectif, " _
        & " POPMAvecCorrectif , POPM_X, POPM_XTERME, POPM_ANC, POPM_DUREE, POCaptive, POBaseRevalo, POMontantMajore, POFRACT, POECHU,NUENRP3I,TXTECHN,TXFRAIS," _
        & " POCDCONTENTIEUX, PONUSINISTRE, POCategorieInval, POMontantCapConstit, POMontantCapSousRisque, POCoeffBCACR, TXTECHNR, TXFRAISR, POAppli "
  
  If bForReformeRetraite Then
    rq = rq & ", POIdAssure, POTopAmortissable, POCoeffAmortissement,POAgeMiniDepartRetraite,POAgeDepartRetraiteTauxPlein "
  Else
    rq = rq & ", POIdAssureMDC "
  End If
        
  rq = rq & ") SELECT POGPECLE, POPERCLE, POSTECLE, PONumParamCalcul, PONUMCLE, POGARCLE, POCONVENTION, PONOM, POSEXE, POCSP, PONAIS, POEFFET, POTERME, POARRET, " _
        & " POREPRISE, POCAUSE, POTYPEF, PODELAI, POPREMIER_PAIEMENT, POPRESTATION, POPRESTATIONTOTAL, POPRESTATION_AN_PREC, " _
        & " POPRESTATION_AN, POPRESTATION_AN_PASSAGE, POSIT, POPM_INCAP_1F, POPM_PASS_1F, POPM_INVAL_1F, POPM, POPM_SORTIE, POPM_VAR, PODERNIERPAIEMENT, " _
        & " PODEBUT, POFIN, PODEBUTTOTAL, POFINTOTAL, POPSAP, POPM_RI, POPM_REVALO, POCOT_REVALO, PODATEIMPORT, POPM_INCAP_1R, " _
        & " POPM_PASS_1R, POPM_INVAL_1R, POCATEGORIE, POPM_RCJT_1F, POPM_REDUC_1F, POPRESTA_RASSUR, POPM_RASSUR, POPSAP_RASSUR, " _
        & " POTRAITE_RASSUR, POTRAITE_RASSUR2, POTRAITE_RASSUR_TAUX, PODATEPAIEMENTESTIMEE, POPSAPCAPMOYEN, PODOSSIERCLOS, " _
        & " POGARCLE_NEW, POCATEGORIE_NEW, POREGROUPEMENT, POCODENATURE, POCCN, POCODERISQUE, POPMANNULEE, POPSAPANNULEE, " _
        & " POCOMMENTANNUL, PODATEENTREEINVAL, POPM_RCJT_1R, POPM_REDUC_1R, POIsCadre, POSalaireAnnuel, POCleSituationFamille, " _
        & " POTauxGarantieDC, PONbEnfant, POAgeMoyenEnfant, POMajoEnfant, PORegimeDeces, POCategorieDeces, PORegimeRenteConjointTempo, " _
        & " POCategorieRenteConjointTempo, PORegimeRenteConjointViagere, POCategorieRenteConjointViagere, PORegimeRenteEduc, POCategorieRenteEduc, " _
        & " POPourcentLissage, POCoeffBCAC, POContractant, POSituConv, POEffetSitu, POCode_Option_Contrat, POEtablissement, PONumSinistre, POTypeMvt, " _
        & " POIJ, POMontantBase, POMontantRevalo, POTypeReglement, PODebutIndemnisation, PONbJourIndemn, POCreationDossier, PODebutDossier, " _
        & " POFinDossier, POMotifCloture, PODebutRefSalaire, POFinRefSalaire, POFamilleComptable, POInspecteur, POInsp, POApport, POApport2, " _
        & " POGestionnaire, POIndicCC, POCompagnie, PONbIntervenant, POPM_AvecCorrectif, POPM_SansCorrectif, POCorrectif, POPMReassAvecCorrectif, " _
        & " POPMAvecCorrectif , POPM_X, POPM_XTERME, POPM_ANC, POPM_DUREE, POCaptive, POBaseRevalo, POMontantMajore, POFRACT, POECHU,NUENRP3I,TXTECHN,TXFRAIS," _
        & " POCDCONTENTIEUX, PONUSINISTRE, POCategorieInval, POMontantCapConstit, POMontantCapSousRisque, POCoeffBCACR, TXTECHNR, TXFRAISR, POAppli "
        
  If bForReformeRetraite Then
    rq = rq & ", POIdAssure, POTopAmortissable, POCoeffAmortissement,POAgeMiniDepartRetraite,POAgeDepartRetraiteTauxPlein "
  Else
    ' Pour la reforme des retraites : Stocke le Recno temporaire pour pouvoir connecter la ligne MDC dans Assure
    ' à la ligne MDC dans Assure_Retraite après recopie des tables temporaires vers les tables permanentes
    rq = rq & ", RECNO "
  End If
  
  rq = rq & " FROM " & TableSource & " A WHERE A.POGPECLE=" & m_NumGroupe & " AND A.POPERCLE=" & m_NumPeriode
  
  
  m_dataSource.Execute rq
  
  Exit Sub
  
err_RecopieMaintienDeces:
  
  If Not autoMode Then
    MsgBox "Erreur " & Err & " pendant la recopie des maintiens en garanties décès : " & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " pendant la recopie des maintiens en garanties décès : " & vbLf & Err.Description
  End If
  
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Connecte les lignes MDC d'Assure_Retraite à celles de Assure via Assure.POIdA-
'ssureMDC
'et Assure_Retraite.POIdAssure
'
'##ModelId=5C8A75F00234
Private Sub ConnecteLignesMDC()
  Dim rq As String
  
  On Error GoTo err_RecopieMaintienDeces
  
  rq = "UPDATE Assure_Retraite SET POIdAssure = A.RECNO " _
     & " FROM Assure A INNER JOIN Assure_Retraite AR ON AR.POIdAssure=A.POIdAssureMDC " _
     & " WHERE A.POGPECLE=" & m_NumGroupe & " AND A.POPERCLE=" & m_NumPeriode
  
  m_dataSource.Execute rq
  
  Exit Sub
  
err_RecopieMaintienDeces:
  
  If Not autoMode Then
    MsgBox "Erreur " & Err & " pendant la recopie des maintiens en garanties décès : " & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " pendant la recopie des maintiens en garanties décès : " & vbLf & Err.Description
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' vérifie la présence des paramètres entre 2 tables
'
'##ModelId=5C8A75F0025F
Private Function CheckOne(titre As String, sSql As String, SrcTable As String, DestTable As String, LibRegime As String, LibCategorie As String, Logger As clsLogger, bAjoutTitre As Boolean) As Boolean
  Dim rs As ADODB.Recordset
  
  CheckOne = True
  Logger.EcritTraceDansLog titre & IIf(bAjoutTitre, " référencé dans '" & SrcTable & "' introuvable dans '" & DestTable & "'", "")
  
  Set rs = m_dataSource.OpenRecordset(sSql, Dynamic)
  
  Do Until rs.EOF
    If Not IsNull(rs.fields(0)) Then
      If rs.fields(0) <> -1 Or rs.fields(1) <> -1 Then
        Logger.EcritTraceDansLog "     " & LibRegime & " " & rs.fields(0) & " / " & LibCategorie & " " & rs.fields(1)
        CheckOne = False
      End If
    End If
    
    rs.MoveNext
  Loop
  
  If CheckOne = True Then
    Logger.EcritTraceDansLog "     Aucune erreur"
  End If
  
  Logger.EcritTraceDansLog "   "
  
  rs.Close
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' vérifie la présence des paramètres pour permettre les calculs sans erreur
'
'##ModelId=5C8A75F00351
Private Function iP3ICalcul_CheckParametresAssures(numPeriode As Long, NumGroupe As Long, Logger As clsLogger) As Boolean
  Dim sSql As String, bOk As Boolean
  
  bOk = True
  
  Logger.EcritTraceDansLog "  "
  Logger.EcritTraceDansLog "##########    Contrôle des paramètres (Régime/Catégorie, Age/Cadre, etc...)"
  Logger.EcritTraceDansLog "  "
  
  ' check NumParamCalcul dans ParamCalcul
  m_NumParametreCalculAbsent = False
  sSql = "SELECT DISTINCT A.PONumParamCalcul, A.PONumParamCalcul FROM Assure AS A LEFT OUTER JOIN ParamCalcul AS PA ON A.PONumParamCalcul = PA.PENUMPARAMCALCUL AND PA.PEGPECLE = A.POGPECLE AND PA.PENUMCLE = A.POPERCLE" _
        & " WHERE A.POGPECLE = " & NumGroupe & " AND A.POPERCLE = " & numPeriode & " AND PA.PENUMPARAMCALCUL IS NULL" _
        & " ORDER BY A.PONumParamCalcul"
  m_NumParametreCalculAbsent = Not CheckOne("Contrôle n° de Paramètre de calcul", sSql, "Assure", "ParamCalcul", "PONumParamCalcul", "PENumParamCalcul", Logger, True)
  If m_NumParametreCalculAbsent Then
    Logger.EcritTraceDansLog "          !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Logger.EcritTraceDansLog "          !!! ATTENTION CERTAINS NUMEROS DE PARAMETRES DE CALCUL NE SONT PAS SPECIFIES DANS LES PARAMETRES DE LA PERIODE => VOIR LA TABLE CODESCAT"
    Logger.EcritTraceDansLog "          !!! LES LIGNES CORRESPODANTES NE SERONT PAS CALCULEES"
    Logger.EcritTraceDansLog "          !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
    Logger.EcritTraceDansLog "  "
    Logger.EcritTraceDansLog "  "
  End If
  bOk = bOk And (Not m_NumParametreCalculAbsent)
  
  ' check regime categorie Incap dans IndemnisationIncapInval
'  sSQL = "SELECT DISTINCT A.POGARCLE-50, A.POCATEGORIE, I.Regime, I.Categorie" _
'         & " FROM Assure A LEFT JOIN IndemnisationIncapInval I " _
'         & "      ON (A.POGARCLE-50 = I.Regime) AND (A.POCATEGORIE = I.Categorie) " _
'         & " WHERE A.POGARCLE = 59 And I.Regime Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Incap/Inval", sSQL, "Assure", "IndemnisationIncapInval", "Régime", "Catégorie", Logger)
  
'  ' check regime categorie Incap dans CorrespondanceGarantie
'  sSQL = "SELECT DISTINCT A.POGARCLE-50, A.POCATEGORIE, I.RegimeIncap, I.CategorieIncap" _
'         & " FROM Assure A LEFT JOIN CorrespondanceGarantie I " _
'         & "      ON (A.POGARCLE-50 = I.RegimeIncap) AND (A.POCATEGORIE = I.CategorieIncap) " _
'         & " WHERE A.POGARCLE = 59 And I.RegimeIncap Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Incap/Inval", sSQL, "Assure", "CorrespondanceGarantie", "Régime", "Catégorie", Logger)
  
  ' check regime categorie DC Assure dans GarantieDC
'  sSQL = "SELECT DISTINCT A.PORegimeDeces, A.POCategorieDeces, I.Regime, I.Categorie" _
'         & " FROM Assure A LEFT JOIN GarantieDC I " _
'         & "      ON (A.PORegimeDeces = I.Regime) AND (A.POCategorieDeces = I.Categorie) " _
'         & " WHERE I.Regime Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Décès", sSQL, "Assure", "GarantieDC", "Régime", "Catégorie", Logger)
  
  ' check categorie rente educ Assure dans TauxRenteEducation
'  sSQL = "SELECT DISTINCT A.PORegimeRenteEduc, A.POCategorieRenteEduc, I.Categorie" _
'         & " FROM Assure A LEFT JOIN TauxRenteEducation I " _
'         & "      ON (A.POCategorieRenteEduc = I.Categorie) " _
'         & " WHERE I.Categorie Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Catégorie Rente Education", sSQL, "Assure", "TauxRenteEducation", "Régime", "Catégorie", Logger)
  
  ' check regime categorie rente conjoint temporaire Assure dans TauxRenteConjoint
'  sSQL = "SELECT DISTINCT A.PORegimeRenteConjointTempo, A.POCategorieRenteConjointTempo, I.Regime, I.Categorie" _
'         & " FROM Assure A LEFT JOIN TauxRenteConjoint I " _
'         & "      ON (A.PORegimeRenteConjointTempo = I.Regime) AND (A.POCategorieRenteConjointTempo = I.Categorie) " _
'         & " WHERE I.Regime Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Rente de Conjoint Temporaire", sSQL, "Assure", "TauxRenteConjoint", "Régime", "Catégorie", Logger)
  
  ' check regime categorie rente conjoint viagere Assure dans TauxRenteConjoint
'  sSQL = "SELECT DISTINCT A.PORegimeRenteConjointViagere, A.POCategorieRenteConjointViagere, I.Regime, I.Categorie" _
'         & " FROM Assure A LEFT JOIN TauxRenteConjoint I " _
'         & "      ON (A.PORegimeRenteConjointViagere = I.Regime) AND (A.POCategorieRenteConjointViagere = I.Categorie) " _
'         & " WHERE I.Regime Is Null AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Rente de Conjoint Viagere", sSQL, "Assure", "TauxRenteConjoint", "Régime", "Catégorie", Logger)

  ' check regime categorie DC CorrespondanceGarantie dans GarantieDC
'''''  sSql = "SELECT DISTINCT A.RegimeDeces, A.CategorieDeces, I.Regime, I.Categorie" _
'''''         & " FROM CorrespondanceGarantie A LEFT JOIN GarantieDC I " _
'''''         & "      ON (A.RegimeDeces = I.Regime) AND (A.CategorieDeces = I.Categorie) " _
'''''         & " WHERE I.Regime Is Null"
'''''  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Décès", sSql, "CorrespondanceGarantie", "GarantieDC", "Régime", "Catégorie", Logger, True)
  
  ' check regime categorie rente educ CorrespondanceGarantie dans TauxRenteEducation
'  sSQL = "SELECT DISTINCT A.RegimeRenteEduc, A.CategorieRenteEduc, I.Categorie" _
'         & " FROM CorrespondanceGarantie A LEFT JOIN TauxRenteEducation I " _
'         & "      ON (A.CategorieRenteEduc = I.Categorie) " _
'         & " WHERE I.Categorie Is Null"
'  bOk = bOk And CheckOne("Contrôle Catégorie Rente Education", sSQL, "CorrespondanceGarantie", "TauxRenteEducation", "Régime", "Catégorie", Logger)
  
  ' check regime categorie rente conjoint temporaire CorrespondanceGarantie dans TauxRenteConjoint
'  sSQL = "SELECT DISTINCT A.RegimeRenteConjointTemporaire, A.CategorieRenteConjointTemporaire, I.Regime, I.Categorie" _
'         & " FROM CorrespondanceGarantie A LEFT JOIN TauxRenteConjoint I " _
'         & "      ON (A.RegimeRenteConjointTemporaire = I.Regime) AND (A.CategorieRenteConjointTemporaire = I.Categorie) " _
'         & " WHERE I.Regime Is Null"
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Rente de Conjoint Temporaire", sSQL, "CorrespondanceGarantie", "TauxRenteConjoint", "Régime", "Catégorie", Logger)
  
  ' check regime categorie rente conjoint viagere CorrespondanceGarantie dans TauxRenteConjoint
'  sSQL = "SELECT DISTINCT A.RegimeRenteConjointViagere, A.CategorieRenteConjointViagere, I.Regime, I.Categorie" _
'         & " FROM CorrespondanceGarantie A LEFT JOIN TauxRenteConjoint I " _
'         & "      ON (A.RegimeRenteConjointViagere = I.Regime) AND (A.CategorieRenteConjointViagere = I.Categorie) " _
'         & " WHERE I.Regime Is Null"
'  bOk = bOk And CheckOne("Contrôle Régime/Catégorie Rente de Conjoint Viagere", sSQL, "CorrespondanceGarantie", "TauxRenteConjoint", "Régime", "Catégorie", Logger)
  
  ' check Age/Cadre dans AgeSituationFamille
'''''  sSql = "SELECT DISTINCT  (Year(GetDate())-Year(A.PONAIS)), A.POIsCadre, B.Age, B.IsCadre " _
'''''         & " FROM Assure A LEFT JOIN AgeSituationFamille B " _
'''''         & "      ON ( (Year(GetDate())-Year(A.PONAIS)) = B.Age) AND (A.POIsCadre = B.IsCadre) " _
'''''         & " WHERE A.POGARCLE = 59 AND B.Age IS NULL  AND A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & NumPeriode
'''''  bOk = bOk And CheckOne("Contrôle Age/Cadre", sSql, "Assure", "AgeSituationFamille", "Age", "IsCadre", Logger, True)

'  ' check regime categorie Incap dans IndemnisationIncapInval (pas erreur)
'  sSQL = "SELECT DISTINCT A.POGARCLE-50, A.POCATEGORIE, I.Regime, I.Categorie" _
'         & " FROM Assure A LEFT JOIN Reassurance I " _
'         & "      ON (A.POGARCLE-50 = I.Regime) AND (A.POCATEGORIE = I.Categorie) " _
'         & " WHERE A.POGARCLE < 90 AND I.Regime Is Null"
'  Call CheckOne("Contrôle Réassurance Régime/Catégorie (ne génère pas d'erreur = Régimes/Catégories non réassurés)", sSQL, "Assure", "Reassurance", "Régime", "Catégorie", Logger)
  
  '*** Controle Age <= 26 ans Rente
  sSql = "SELECT NUENRP3I, Year(PEDATEFIN)-Year(PONAIS) as Age " _
         & " FROM P3IUser.Assure AS A INNER JOIN P3IUser.Periode as P ON A.POGPECLE=P.PEGPECLE AND A.POPERCLE=P.PENUMCLE " _
         & " WHERE A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & numPeriode & " And (A.Posit = 22 Or A.Posit = 23) And (Year(PEDATEFIN) - Year(PONAIS)) > 26"
  bOk = bOk And CheckOne("Controle Age <= 26 ans pour Garantie Rente Education", sSql, "Assure", "Age", "NumEnrP3I", "Age", Logger, False)
  
  ' renvoie le resultat
  iP3ICalcul_CheckParametresAssures = bOk
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' controle des resultats
'
'##ModelId=5C8A75F003BE
Private Function iP3ICalcul_CheckResultatCalcul(numPeriode As Long, NumGroupe As Long, Logger As clsLogger) As Boolean
  Dim sSql As String, bOk As Boolean
  
  bOk = True
  
  Logger.EcritTraceDansLog "  "
  Logger.EcritTraceDansLog "##########    Contrôle des Résultats de calcul"
  Logger.EcritTraceDansLog "  "
  
  
  '*** Controle PM <= 1 000 000 Euros
  Dim sMaxPm As String
  
  sMaxPm = GetSettingIni(CompanyName, SectionName, "ControleMontantMaxPM", "1000000")
  sSql = "SELECT NUENRP3I, POPM " _
         & " FROM P3IUser.Assure AS A " _
         & " WHERE A.POGPECLE=" & NumGroupe & " AND A.POPERCLE=" & numPeriode & " And POPM>" & sMaxPm
  bOk = bOk And CheckOne("Controle PM >= " & sMaxPm & " Euros", sSql, "Assure", "PM", "NumEnrP3I", "PM", Logger, False)
  
  
  ' renvoie le resultat
  iP3ICalcul_CheckResultatCalcul = bOk
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Maternités
'
'##ModelId=5C8A75F10034
Private Sub CalculMaternite(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer
  Dim franchise As Integer
  Dim iIncap As Double
  Dim fraisGestionIncap As Double
  Dim ajoutCmt As String
  Dim rq As String
  
  On Error GoTo err_calcul
  
  rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0  'forcé à 0 le 03 06 2013  - non utilisé dans ce calcul -
  
  ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
    
  ' paramètres de la fonction FctCalcProvInval
  'age x            = âge à la date d'entrée en incapacité
  'anc              = ancienneté dans l'état d'incapacité, en mois
  'duree            = entre date de l'ARRET et date TERME du contrat, en mois
  'franchise        = dans l'état d'incapacité, , en mois
  'feuilleLue       = loi de maintien en incapacité
  'taux technique     i
  'frais de gestion = fraisGestion
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
  m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  If m_xTerme < 0 Then
    m_xTerme = 0
  End If
  m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))

  
  x = m_xArret
  
  ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
  'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
  anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
  If anc < 0 Then
    anc = 0
  End If
  
  If anc >= 10 Then ' 10 mois maxi
    ' Assuré en arret depuis plus de 10 mois
    ' on force à 0
    ajoutCmt = "PM annulée car assuré en maternité de durée>=10 mois"
    If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
      rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- " & ajoutCmt
    End If
    
    ' on force les coeff de provisions incapacite à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
       
    
  Else
    
    ' durée en mois entre date d'ARRET et date TERME du contrat : 10 mois max
    'Duree = Minimum(10, DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    Duree = Minimum(10, NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    If Duree < 0 Then
      Duree = 0
    End If
    
    ' franchise en mois entre la date D'ARRET et la date d'EFFET du contrat
    'franchise = DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    franchise = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    If franchise < 0 Then
      franchise = 0
    End If
    
    ' franchise en mois ajout du délais de carence
    franchise = franchise + CInt(rsAssure.fields("PODELAI") / 30)
    If franchise < 0 Then
      franchise = 0
    End If
    
    '
    iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
    fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100
    
    rsAssure.fields("TXTECHN") = iIncap * 100#
    rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
    
    rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvMensualiteCertaine(anc, Duree, franchise, iIncap, fraisGestionIncap)
    rsAssure.fields("POPM_PASS_1F") = 0
    
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    
    rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    
    rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
    rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
    rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
    
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = 0
    
    ' stockage des paramètres pour l'incap / passage
    rsAssure.fields("POPM_x") = x
    rsAssure.fields("POPM_xTerme") = m_xTerme
    rsAssure.fields("POPM_anc") = anc
    rsAssure.fields("POPM_Duree") = Duree
 
    If m_bRevalo = True Then
      m_Logger.EcritTraceDansLog "Calcul Maternité : pas de calcul de la revalorisation."
    End If
 
  End If ' anc < 10

  Exit Sub
  
err_calcul:
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    If Not autoMode Then
      MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    Else
      m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
    End If
   
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite les Mensualisations
'
'##ModelId=5C8A75F10092
Private Sub CalculMensualisation(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer
  Dim franchise As Integer
  Dim iIncap As Double
  Dim fraisGestionIncap As Double
  Dim ajoutCmt As String
  Dim rq As String
  Dim feuilleLue As String

  
  On Error GoTo err_calcul
  
  rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0  'forcé à 0 le 03 06 2013  - non utilisé dans ce calcul -

  
  ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
    
  ' paramètres de la fonction FctCalcProvInval
  'age x            = âge à la date d'entrée en incapacité
  'anc              = ancienneté dans l'état d'incapacité, en mois
  'duree            = entre date de l'ARRET et date TERME du contrat, en mois
  'franchise        = dans l'état d'incapacité, , en mois
  'feuilleLue       = loi de maintien en incapacité
  'taux technique     i
  'frais de gestion = fraisGestion
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
  m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  If m_xTerme < 0 Then
    m_xTerme = 0
  End If
  m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  
  x = m_xArret
  
  ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
  'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
  anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
  If anc < 0 Then
    anc = 0
  End If
  
  If anc >= 12 Then ' 12 mois maxi
    ' Assuré en arret depuis plus de 12 mois
    ' on force à 0
    ajoutCmt = "PM annulée car assuré en mensualisation de durée>=12 mois"
    If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
      rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- " & ajoutCmt
    End If
    
    ' on force les coeff de provisions incapacite à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
   
  Else
    
    ' durée en mois entre date d'ARRET et date TERME du contrat : 12 mois max
    'Duree = Minimum(12, DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    Duree = Minimum(12, NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    If Duree < 0 Then
      Duree = 0
    End If
    
    ' franchise en mois entre la date D'ARRET et la date d'EFFET du contrat
    'franchise = DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    franchise = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    If franchise < 0 Then
      franchise = 0
    End If
    
    ' franchise en mois ajout du délais de carence
    franchise = franchise + CInt(rsAssure.fields("PODELAI") / 30)
    If franchise < 0 Then
      franchise = 0
    End If
    
    '
    feuilleLue = m_rsPeriode.fields("PELMINCAP")
    iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
    fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100
    
    rsAssure.fields("TXTECHN") = iIncap * 100#
    rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
    
    ' PHM 02-11-2006
    ' si mensu = rente certaine max 12 mois : rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvMensualiteCertaine(anc, Duree, franchise, iIncap, fraisGestionIncap)
    ' si mensu = incap max 12 mois : rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvIncap(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
    rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvIncap(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
    rsAssure.fields("POPM_PASS_1F") = 0
    
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    
    rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    
    rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
    rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
    rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
    
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = 0
    
    ' stockage des paramètres pour l'incap / passage
    rsAssure.fields("POPM_x") = x
    rsAssure.fields("POPM_xTerme") = m_xTerme
    rsAssure.fields("POPM_anc") = anc
    rsAssure.fields("POPM_Duree") = Duree
 
    If m_bRevalo = True Then
      m_Logger.EcritTraceDansLog "Calcul Mensualisation : pas de calcul de la revalorisation."
    End If
 
  End If ' anc < 10

  Exit Sub
  
err_calcul:
    
    If Not autoMode Then
      MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    Else
      m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
    End If
  
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite le Chomage
'
'##ModelId=5C8A75F100E0
Private Sub CalculChomage(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer
  Dim franchise As Integer
  Dim iIncap As Double
  Dim fraisGestionIncap As Double
  Dim ajoutCmt As String
  Dim rq As String
  
  'ALG 31/07/2013
  Dim anc_regl As Integer
  Dim POFIN As Variant ' Date
  Dim feuilleLueIncap As String
  Dim feuilleLuePassage As String
  Dim correctif As Double
  Dim DureeChomage As Integer
  
  On Error GoTo err_calcul
  
  rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0  'forcé à 0 le 03 06 2013  - non utilisé dans ce calcul -

  
  ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
    
  ' paramètres de la fonction FctCalcProvInval
  'age x            = âge à la date d'entrée en incapacité
  'anc              = ancienneté dans l'état d'incapacité, en mois
  'duree            = entre date de l'ARRET et date TERME du contrat, en mois
  'franchise        = dans l'état d'incapacité, , en mois
  'feuilleLue       = loi de maintien en incapacité
  'taux technique     i
  'frais de gestion = fraisGestion
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
  m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  If m_xTerme < 0 Then
    m_xTerme = 0
  End If
  m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  
  x = m_xArret
  
  ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
  'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
  anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
  If anc < 0 Then
    anc = 0
  End If
 
    'recherche duree chomage
    
   If IsNull(m_rsPeriode.fields("PEDUREECHOMAGE")) Then
    DureeChomage = 24 ' valeur par défaut pour P3I collective
   Else
    DureeChomage = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEDUREECHOMAGE"))
   End If
    ' durée en mois entre date d'ARRET et date TERME du contrat : DureeChomage mois max
    'Duree = Minimum(24, DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    Duree = Minimum(DureeChomage, NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
    If Duree < 0 Then
      Duree = 0
    End If
    
    ' franchise en mois entre la date D'ARRET et la date d'EFFET du contrat
    'franchise = DateDiff("m", rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    franchise = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
    If franchise < 0 Then
      franchise = 0
    End If
    
    ' franchise en mois ajout du délais de carence
    franchise = franchise + CInt(rsAssure.fields("PODELAI") / 30)
    If franchise < 0 Then
      franchise = 0
    End If
    '
    iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
    fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100

    ' stockage des paramètres âge, durée...
    rsAssure.fields("POPM_x") = x
    rsAssure.fields("POPM_xTerme") = m_xTerme
    rsAssure.fields("POPM_anc") = anc
    rsAssure.fields("POPM_Duree") = Duree


  If anc >= DureeChomage Then ' DureeChomage mois maxi
    ' Assuré en arret depuis plus de DureeChomage mois
    ' on force à 0
    ajoutCmt = "PM annulée car assuré en mensualisation de durée>= " & DureeChomage & " mois"
    
    If InStr(rsAssure.fields("POCOMMENTANNUL"), ajoutCmt) = 0 Then
      rsAssure.fields("POCOMMENTANNUL") = rsAssure.fields("POCOMMENTANNUL") & " --- " & ajoutCmt
    End If
    
    ' on force les coeff de provisions incapacite à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
   
  Else  ' anc <  DureeChomage  calcul P3I individuel et collective
  
    ''''''''''''''''''''''''''''''''''''''''''''''
    '           P3I INDIVIDUEL 31 07 2013
    ''''''''''''''''''''''''''''''''''''''''''''''
    If m_bP3I_Individuel = True Then
        
        ' calcul de l'ancienneté à la date du dernier réglement
        POFIN = rsAssure.fields("POFIN") ' date du dernier réglement
        If IsNull(rsAssure.fields("POFIN")) Then
            POFIN = m_DateCalcul ' date de calcul dans ce cas
        End If
        
        anc_regl = 0
        anc_regl = DateDiff("m", rsAssure.fields("POARRET"), POFIN)
        If anc_regl < 0 Then
            anc_regl = 0
        End If

        If bForReformeRetraite And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
        
           ' tables
           feuilleLueIncap = m_rsPeriode.fields("PELMINCAP_Retraite")
           feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE_Retraite")
    
         Else
                 
           feuilleLueIncap = m_rsPeriode.fields("PELMINCAP")
           feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE")
           
         End If
            ' calcul correctif comme dans l'incap pas de correctif pour le chomage
         'correctif = m_ModuleCalcul_Provisions.FctCalcCorrectifIncapacite(x, anc, anc_regl, feuilleLueIncap, feuilleLuePassage)
         'correctif = Arrondi(correctif, NbDecimalePM)
         correctif = 1
         
         rsAssure.fields("TXTECHN") = iIncap * 100#
         rsAssure.fields("TXFRAIS") = 0#
       
         rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvMensualiteCertaine(anc, Duree, franchise, iIncap, 0)  'fraisGestionIncap=0
         rsAssure.fields("POPM_PASS_1F") = 0
         rsAssure.fields("POPM_INCAP_1R") = 0
         rsAssure.fields("POPM_PASS_1R") = 0
         rsAssure.fields("POPM_INVAL_1F") = 0
         rsAssure.fields("POPM_INVAL_1R") = 0
         
         rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
         rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
         rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
                  
         rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
         rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
         rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
         
         rsAssure.fields("POPM_REVALO") = 0
         rsAssure.fields("POCOT_REVALO") = 0
         
         If Not m_rsPMRetenue.EOF Then  'est que la pm retenue est avec ou sans correctif dans la table pmretenue
         
            rsAssure.fields("POPM_AvecCorrectif") = Arrondi(rsAssure.fields("POPM") * correctif, 2)
            rsAssure.fields("POPM_SansCorrectif") = Arrondi(rsAssure.fields("POPM"), 2)
            rsAssure.fields("POCorrectif") = correctif
            rsAssure.fields("POPMAvecCorrectif") = m_rsPMRetenue.fields("PMAvecCorrectif")
            rsAssure.fields("POPMReassAvecCorrectif") = m_rsPMRetenue.fields("PMReassAvecCorrectif")
           
            If m_rsPMRetenue.fields("PMAvecCorrectif") = True Then
                rsAssure.fields("POPM") = rsAssure.fields("POPM_AvecCorrectif")
                rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE") * correctif, 2)
                rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR") * correctif, 2)
    
            
            End If
         End If
    ''''''''''''''''''''''''''''''''''''''''''''''
    '           P3I COLLECTIVE
    ''''''''''''''''''''''''''''''''''''''''''''''
    Else ' = P3I COLLECTIVE
    
        rsAssure.fields("TXTECHN") = 0 ' 14/05/2008 : iIncap * 100#
        rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
        ' 14/05/2008 : Coeff Chomage = NbMoisRestant/12
        'rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcProvMensualiteCertaine(anc, Duree, franchise, iIncap, fraisGestionIncap)

        rsAssure.fields("POPM_INCAP_1F") = Arrondi((1 + fraisGestionIncap) * (Duree - Maximum(anc, franchise)), NbDecimalePM)
        rsAssure.fields("POPM_PASS_1F") = 0
        rsAssure.fields("POPM_INCAP_1R") = 0
        rsAssure.fields("POPM_PASS_1R") = 0
        rsAssure.fields("POPM_INVAL_1F") = 0
        rsAssure.fields("POPM_INVAL_1R") = 0
        
        rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
        rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
        rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
        
        rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
        rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
        rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
           
        rsAssure.fields("POPM_REVALO") = 0
        rsAssure.fields("POCOT_REVALO") = 0
    
    End If
  End If

  Exit Sub
  
err_calcul:
    
  If Not autoMode Then
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Else
    m_Logger.EcritTraceDansLog "Erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description
  End If
  
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite la Statutaire au 31/08/2016
'
'##ModelId=5C8A75F1014D
Private Sub CalculStatutaire(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
    
  Dim Duree As Integer
  Dim x As Integer
  Dim anc As Integer
  Dim franchise As Integer
  Dim iIncap As Double
  Dim fraisGestionIncap As Double
  Dim ajoutCmt As String
  Dim rq As String
  
  'Define fields to store required fields for the recovery of Statutaire coeffs
  Dim coeff As Double
  Dim calc As Double
  Dim calcVar As Double
  Dim Tj As Double
  Dim TjVAR As Double
  Dim daysSurvUntilToday As Integer
  Dim typeSinistre As String
  Dim collect As String
  Dim semaine As Integer
  Dim sexe As Integer
  Dim Age As Integer
  Dim tableYear As Integer
  
  'define all coeffs we need to recover
  Dim PM_AT As Double
  Dim PM_CLD As Double
  Dim PM_CLM As Double
  Dim PM_CLM_CLD As Double
  Dim PM_MO As Double
  Dim PM_MO_CLM As Double
  Dim PM_MO_CLD As Double
  
  Dim AgeRet As Integer
  Dim AgeMin As Integer
  Dim AgeMax As Integer
  Dim MOMaxSemaine As Integer
  Dim MATMaxSemaine As Integer
  Dim ATMaxSemaine As Integer
  Dim CLMMaxSemaine As Integer
  Dim CLDMaxSemaine As Integer
  Dim AnneeBareme As Integer
  Dim collectNumber As Integer

  
'  AgeRet = IIf(IsNull(m_rsStatParams.fields("AgeRetraite")), 62, m_rsStatParams.fields("AgeRetraite"))
'  AgeMin = IIf(IsNull(m_rsStatParams.fields("AgeMin")), 20, m_rsStatParams.fields("AgeMin"))
'  AgeMax = IIf(IsNull(m_rsStatParams.fields("AgeMax")), 64, m_rsStatParams.fields("AgeMax"))
'  MOMaxSemaine = IIf(IsNull(m_rsStatParams.fields("MO_Max_Semaine")), 52, m_rsStatParams.fields("MO_Max_Semaine"))
'  MATMaxSemaine = IIf(IsNull(m_rsStatParams.fields("MAT_Max_Semaine")), 23, m_rsStatParams.fields("MAT_Max_Semaine"))
'  ATMaxSemaine = IIf(IsNull(m_rsStatParams.fields("AT_Max_Semaine")), 366, m_rsStatParams.fields("AT_Max_Semaine"))
'  CLMMaxSemaine = IIf(IsNull(m_rsStatParams.fields("CLM_Max_Semaine")), 156, m_rsStatParams.fields("CLM_Max_Semaine"))
'  CLDMaxSemaine = IIf(IsNull(m_rsStatParams.fields("CLD_Max_Semaine")), 252, m_rsStatParams.fields("CLD_Max_Semaine"))
'  AnneeBareme = IIf(IsNull(m_rsStatParams.fields("AnneeBareme")), 2004, m_rsStatParams.fields("AnneeBareme"))

  AgeRet = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_AGERET")), 62, m_rsParamCalcul.fields("PESTAT_AGERET"))
  AgeMin = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_AGEMIN")), 20, m_rsParamCalcul.fields("PESTAT_AGEMIN"))
  AgeMax = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_AGEMAX")), 64, m_rsParamCalcul.fields("PESTAT_AGEMAX"))
  MOMaxSemaine = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_MO")), 52, m_rsParamCalcul.fields("PESTAT_MO"))
  CLMMaxSemaine = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_CLM")), 156, m_rsParamCalcul.fields("PESTAT_CLM"))
  CLDMaxSemaine = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_CLD")), 252, m_rsParamCalcul.fields("PESTAT_CLD"))
  MATMaxSemaine = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_MAT")), 23, m_rsParamCalcul.fields("PESTAT_MAT"))
  ATMaxSemaine = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_AT")), 366, m_rsParamCalcul.fields("PESTAT_AT"))
  AnneeBareme = IIf(IsNull(m_rsParamCalcul.fields("PESTAT_ANNBAR")), 2004, m_rsParamCalcul.fields("PESTAT_ANNBAR"))

  
  On Error GoTo err_calcul
  
  If m_rsStatProv Is Nothing Then
     m_Logger.EcritTraceDansLog "Calcul Statutaire : La table qui contienne les provisions statutaire ne peut pas être charge !"
     
     If MsgBox("Calcul Statutaire : La table qui contienne les provisions statutaire ne peut pas être chargé !" & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) <> vbYes Then
       End
     End If
     
     'Open parameters for periode and click save
     'Offer to stop the program
     
     'Resume Next
     Exit Sub
  End If
  
  '********************************
  '***** Generic Calculations *****
  '********************************
  
  ' paramètres de la fonction FctCalcProvInval
  'age x            = âge à la date d'entrée en incapacité
  'anc              = ancienneté dans l'état d'incapacité, en mois
  'duree            = entre date de l'ARRET et date TERME du contrat, en mois
  'franchise        = dans l'état d'incapacité, , en mois
  'feuilleLue       = loi de maintien en incapacité
  'taux technique     i
  'frais de gestion = fraisGestion
  
  ' âge à la date du terme du contrat ( = âge de départ en retraite)
  'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
  m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  
  If m_xTerme < 0 Then
    m_xTerme = 0
  End If
  
  m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
  x = m_xArret
  
  ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
  'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
  anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
  If anc < 0 Then
    anc = 0
  End If
    
  ' durée en mois entre date d'ARRET et date TERME du contrat : 10 mois max
  Duree = Minimum(10, NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME")))
  If Duree < 0 Then
    Duree = 0
  End If
  
  ' franchise en mois entre la date D'ARRET et la date d'EFFET du contrat
  franchise = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POEFFET"))
  If franchise < 0 Then
    franchise = 0
  End If
  
  ' franchise en mois ajout du délais de carence
  franchise = franchise + CInt(rsAssure.fields("PODELAI") / 30)
  If franchise < 0 Then
    franchise = 0
  End If
  
  ' stockage des paramètres pour l'incap / passage
  rsAssure.fields("POPM_x") = x
  rsAssure.fields("POPM_xTerme") = m_xTerme
  rsAssure.fields("POPM_anc") = anc
  rsAssure.fields("POPM_Duree") = Duree
  
  
  'Initialize Coeffs & Other Fields
  iIncap = 0
  fraisGestionIncap = 0
  
  rsAssure.fields("TXTECHN") = iIncap * 100#
  rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
  
  rsAssure.fields("POPM_INCAP_1F") = 0
  rsAssure.fields("POPM_INVAL_1F") = 0
  rsAssure.fields("POPM_PASS_1F") = 0
  
  rsAssure.fields("POPM_INCAP_1R") = 0
  rsAssure.fields("POPM_INVAL_1R") = 0
  rsAssure.fields("POPM_PASS_1R") = 0
  
  rsAssure.fields("POPM_REVALO") = 0
  rsAssure.fields("POCOT_REVALO") = 0
    
  
  '************************************************
  '***** Specific Calculations For Statutaire *****
  '************************************************
  
  'get all required variables
  Age = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POARRET"))  ' m_rsPeriode.fields("PEDATEEXT"))  Date_Naiss -- Date Sinistre (( Date_Invent ))
  
  If Age < AgeMin Then
    Age = AgeMin
  ElseIf Age > AgeMax Then
    Age = AgeMax
  End If
  
  Tj = rsAssure.fields("POPRESTATION_AN") / 365
  TjVAR = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) / 365
  sexe = rsAssure.fields("POSEXE") ' 1 == Homme ; 2 == Femme
  
  'IIf(IsNull(m_rsStatParams.fields("AgeRetraite")), 62, m_rsStatParams.fields("AgeRetraite"))
  
  'this is a numeric value that needs to be translated into the corresponding guarantee
  typeSinistre = Trim$(rsAssure.fields("POGARCLE_NEW")) ' P3IProvcoll : CDGARAN  => POGARCLE_NEW (in assure)
  
  If IsNumeric(typeSinistre) Then
    m_rsStatutaireGarantie.filter = " GarantieCode = " & typeSinistre
    If Not m_rsStatutaireGarantie.EOF Then
      typeSinistre = IIf(IsNull(m_rsStatutaireGarantie.fields("Garantie")), "ND", Trim$(m_rsStatutaireGarantie.fields("Garantie")))
    End If
  End If
  
  '### add additional cases
  If typeSinistre = "ATMP" Then
    typeSinistre = "AT"
  ElseIf typeSinistre = "CLM/CLD" Then
    typeSinistre = "CLD"
  End If
  
  If typeSinistre <> "MAT" And typeSinistre <> "AT" And typeSinistre <> "MO" And typeSinistre <> "CLM" And typeSinistre <> "CLD" Then
    m_Logger.EcritTraceDansLog "Le type de sinistre suivant n'existe pas : " & typeSinistre
    Exit Sub
  End If
  
  '### test
'  Dim number As Integer
'  number = rsAssure.fields("NUENRP3I")
'  If number = 21 Then
'    number = 22
'  End If

  semaine = Int(DateDiff("d", rsAssure.fields("POARRET"), m_rsPeriode.fields("PEDATEEXT")) / 7) + 1  'Date-Surv -- Date_Invent
  
  'adjust number of weeks to the allowed maximum if necessary
  If typeSinistre = "MO" And semaine > MOMaxSemaine Then
    semaine = MOMaxSemaine
  ElseIf (typeSinistre = "CLD" Or typeSinistre = "CLM/CLD") And semaine > CLDMaxSemaine Then
    semaine = CLDMaxSemaine
  ElseIf typeSinistre = "CLM" And semaine > CLMMaxSemaine Then
    semaine = CLMMaxSemaine
  ElseIf (typeSinistre = "AT" Or typeSinistre = "ATMP") And semaine > ATMaxSemaine Then
    semaine = ATMaxSemaine
  End If
  
  rsAssure.fields("POPM_ANC") = semaine
    
  collectNumber = IIf(IsNull(rsAssure.fields("POCATEGORIE")), 750, Trim$(rsAssure.fields("POCATEGORIE")))
  
  m_rsStatCat.MoveFirst
  m_rsStatCat.filter = " Categorie = '" & collectNumber & "'"
  If Not m_rsStatCat.EOF Then
    collect = m_rsStatCat.fields("TypeCollectivite")
  End If
 
  tableYear = AnneeBareme
  
  'get all coeffs from Stat table if the sinistre <> Maternité
  If typeSinistre <> "MAT" Then
    m_rsStatProv.MoveFirst
    
    On Error GoTo 0
    
'    m_rsStatProv.filter = " TableAnnee = " & tableYear & " And Collect = '" & collect & "'" & " And TypeSinistre = '" & typeSinistre & "'" _
'      & " And Sexe = " & sexe & " And AgeMalade = " & Age & " And Semaine = " & semaine
    m_rsStatProv.filter = " Collect = '" & collect & "'" & " And TypeSinistre = '" & typeSinistre & "'" _
      & " And Sexe = " & sexe & " And AgeMalade = " & Age & " And Semaine = " & semaine

    If Err.Number = 3265 Then
      MsgBox "the following item cannot be found in the selected Statutaire_Provision Table: " & " Collect = '" & collect & "'" & " TypeSinistre = '" & typeSinistre & "'" _
      & " Sexe = " & sexe & " AgeMalade = " & Age & " Semaine = " & semaine, vbCritical
      Resume Next
    ElseIf Err.Number <> 0 Then
      GoTo err_calcul
    End If
    
    On Error GoTo err_calcul
      
    If Not m_rsStatProv.EOF Then
      'we found a record - set all coeffs
      PM_AT = m_rsStatProv.fields("PM_AT")
      PM_CLD = m_rsStatProv.fields("PM_CLD")
      PM_CLM = m_rsStatProv.fields("PM_CLM")
      PM_MO = m_rsStatProv.fields("PM_MO")
      PM_CLM_CLD = m_rsStatProv.fields("PM_CLM_CLD")
      PM_MO_CLM = m_rsStatProv.fields("PM_MO_CLM")
      PM_MO_CLD = m_rsStatProv.fields("PM_MO_CLD")
    Else
      'nothing was found - there is a problem, log an error message
      m_Logger.EcritTraceDansLog "Aucune coefficient était trouvé dans la table 'Statutaire_Provision' pour la ligne NUMENRP3I=" & _
        rsAssure.fields("NUENRP3I") & "  et pour les valeurs suivants : TableAnnee = " & tableYear & " And Collect = '" & collect & "'" & " And TypeSisnistre = '" & typeSinistre & "'" _
      & " And Sexe = " & sexe & " And AgeMalade = " & Age & " And Semaine = " & semaine
      
      '### reset the filter
      m_rsStatProv.filter = adFilterNone
      
      Exit Sub
    End If
  End If
  
  'Calculate POPM => PM_MO , PM_CLD...
  If typeSinistre = "MO" Then
    
    coeff = PM_MO + PM_MO_CLM + PM_MO_CLD
    calc = Tj * coeff
    calcVar = TjVAR * coeff
    
    'transform coeff into annual coeff
    rsAssure.fields("POPM_INCAP_1F") = PM_MO
    rsAssure.fields("POPM_PASS_1F") = PM_MO_CLM
    rsAssure.fields("POPM_INVAL_1F") = PM_MO_CLD
    
  ElseIf typeSinistre = "CLM" Then
  
    coeff = PM_CLM + PM_CLM_CLD
    calc = Tj * coeff
    calcVar = TjVAR * coeff
    
    rsAssure.fields("POPM_INCAP_1F") = PM_CLM
    rsAssure.fields("POPM_PASS_1F") = PM_CLM_CLD
    
  ElseIf typeSinistre = "CLD" Or typeSinistre = "CLM/CLD" Then
  
    coeff = PM_CLD
    calc = Tj * coeff
    calcVar = TjVAR * coeff
    rsAssure.fields("POPM_INCAP_1F") = coeff
    
  ElseIf typeSinistre = "AT" Or typeSinistre = "ATMP" Then
  
    coeff = PM_AT
    calc = Tj * coeff
    calcVar = TjVAR * coeff
    rsAssure.fields("POPM_INCAP_1F") = coeff
    
  ElseIf typeSinistre = "MAT" Then
  
    calc = 0
    coeff = 0
    
    If sexe = 2 Then 'Femme
      If semaine < MATMaxSemaine Then
        'coeff = (161 - daysSurvUntilToday)
        coeff = (MATMaxSemaine - semaine)
        calc = Tj * coeff
        calcVar = TjVAR * coeff
      End If
    End If
    
    rsAssure.fields("POPM_INCAP_1F") = coeff
    
  Else
    'throw error - TypeSinistre doe snot exist
    Err.Raise 111, "Calcul Statutaire", "Le type de sinistre suivant n'existe pas : " & typeSinistre
  
  End If
  
  
  rsAssure.fields("POPM") = calc
  rsAssure.fields("POPM_SORTIE") = calc
  rsAssure.fields("POPM_VAR") = calcVar
  
  'rounding
  rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
  rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
  rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
  

  If m_bRevalo = True Then
    m_Logger.EcritTraceDansLog "Calcul Statutaire : pas de calcul de la revalorisation."
  End If

  Exit Sub
  
err_calcul:
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
  Resume Next
  
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite la Dépendance au 31/07/2013
'
'##ModelId=5C8A75F1019B
Private Sub CalculDependance(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
    Dim Duree As Integer
    Dim x As Integer
    Dim anc As Integer
    Dim franchise As Integer
    Dim iIncap As Double
    Dim fraisGestionIncap As Double
    Dim ajoutCmt As String
    Dim rq As String
    Dim anc_regl As Integer
    Dim POFIN As Variant ' Date
    Dim feuilleLue As String
    
    On Error GoTo err_calcul
    
    rsAssure.fields("POPRESTATION_AN_PASSAGE") = 0  'forcé à 0 le 03 06 2013  - non utilisé dans ce calcul -
    
    
    ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
      
    ' paramètres de la fonction FctCalcProvInval
    'age x            = âge à la date d'entrée en incapacité
    'anc              = ancienneté dans l'état d'incapacité, en mois
    'duree            = entre date de l'ARRET et date TERME du contrat, en mois
    'franchise        = dans l'état d'incapacité, , en mois
    'feuilleLue       = loi de maintien en incapacité
    'taux technique     i
    'frais de gestion = fraisGestion
    
    ' âge à la date du terme du contrat ( = 0 si viagère)
    'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
    If Not IsNull(rsAssure.fields("POTERME")) Then
    m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
    m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
    Else
    m_xTerme = 0 ' valeur arbitraire
    End If
    If m_xTerme < 0 Then
      m_xTerme = 0
    End If
        
    x = m_xArret
    
    ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
    'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
    anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
    If anc < 0 Then
      anc = 0
    End If
    

    
    ' durée en mois entre date d'ARRET et date TERME du contrat
    If Not IsNull(rsAssure.fields("POTERME")) Then
    Duree = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME"))
    Else
    Duree = 0
    End If
    If Duree < 0 Then
      Duree = 0
    End If

    '
    iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
    fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100
    ' stockage des paramètres âge, durée...
    rsAssure.fields("POPM_x") = x
    rsAssure.fields("POPM_xTerme") = m_xTerme
    rsAssure.fields("POPM_anc") = anc
    rsAssure.fields("POPM_Duree") = Duree
    
    rsAssure.fields("POCorrectif") = 1 ' chargement = 1  23/08/2013
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           P3I INDIVIDUEL et COLLECTIVE 31 07 2013
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
     If bForReformeRetraite And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
        ' tables
        feuilleLue = m_rsPeriode.fields("PELMDEPENDANCE")
     Else
        feuilleLue = m_rsPeriode.fields("PELMDEPENDANCE")
     End If
    
    
    'If rsAssure.fields("NUENRP3I") = 16 Then Stop
    
    rsAssure.fields("TXTECHN") = iIncap * 100#
    rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
    rsAssure.fields("POPM_INCAP_1F") = m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLue, fraisGestionIncap)
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    
    rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
    rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
             
    rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
    rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
    rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
    
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POCOT_REVALO") = 0
    
Exit Sub
 
err_calcul:
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    'MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Traite le Capital Invalidité Probable au 04/08/2013
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'##ModelId=5C8A75F101E9
Private Sub CalculCapitalInvaliditeProbable(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
    Dim Duree As Integer
    Dim x As Integer
    Dim anc As Integer
    Dim franchise As Integer
    Dim iIncap As Double
    Dim fraisGestionIncap As Double
    Dim ajoutCmt As String
    Dim rq As String
    Dim anc_regl As Integer
    Dim POFIN As Variant ' Date
    Dim feuilleLueIncap As String
    Dim feuilleLuePassage As String
    Dim correctif As Double
    Dim Inval1Cat As Double
    
    On Error GoTo err_calcul
    
    rsAssure.fields("POPRESTATION_AN_PASSAGE") = rsAssure.fields("POPRESTATION_AN")
    
    ' l'assuré est en INCAPACITE, utilisation de la fonction FctCalcProvIncap
      
    ' paramètres de la fonction FctCalcProvInval
    'age x            = âge à la date d'entrée en incapacité
    'anc              = ancienneté dans l'état d'incapacité, en mois
    'duree            = entre date de l'ARRET et date TERME du contrat, en mois
    'franchise        = dans l'état d'incapacité, , en mois
    'feuilleLue       = loi de maintien en incapacité
    'taux technique     i
    'frais de gestion = fraisGestion
    
    ' âge à la date du terme du contrat ( = 0 si viagère)
    'm_xTerme = Year(rsAssure.fields("POTERME")) - Year(rsAssure.fields("PONAIS"))
    If Not IsNull(rsAssure.fields("POTERME")) Then
    m_xTerme = CalculAge(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
    m_xTermeDouble = CalculAgeDouble(rsAssure.fields("PONAIS"), rsAssure.fields("POTERME"))
    Else
    m_xTerme = 0 ' valeur arbitraire
    End If
    If m_xTerme < 0 Then
      m_xTerme = 0
    End If
        
    x = m_xArret
    
    ' Nombre de mois entre les 2 dates (date d'ARRET, date de FIN DE PERIODE)
    'anc = DateDiff("m", rsAssure.fields("POARRET"), m_DateCalcul)
    anc = NbMois(rsAssure.fields("POARRET"), m_DateCalcul)
    If anc < 0 Then
      anc = 0
    End If
    
    ' durée en mois entre date d'ARRET et date TERME du contrat
    If Not IsNull(rsAssure.fields("POTERME")) Then
    Duree = NbMois(rsAssure.fields("POARRET"), rsAssure.fields("POTERME"))
    Else
    Duree = 0
    End If
    If Duree < 0 Then
      Duree = 0
    End If

    '
    iIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PETXINCAP")) / 100
    fraisGestionIncap = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEGESTIONINCAP")) / 100
    Inval1Cat = m_dataHelper.GetDouble2(m_rsPeriode.fields("PEINVAL1CAT")) / 100
    
    ' stockage des paramètres âge, durée...
    rsAssure.fields("POPM_x") = x
    rsAssure.fields("POPM_xTerme") = m_xTerme
    rsAssure.fields("POPM_anc") = anc
    rsAssure.fields("POPM_Duree") = Duree
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           P3I INDIVIDUEL et COLLECTIVE 31 07 2013
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
    If m_bP3I_Individuel = True Then
        
        ' calcul de l'ancienneté à la date du dernier réglement
        POFIN = rsAssure.fields("POFIN") ' date du dernier réglement
        If IsNull(rsAssure.fields("POFIN")) Then
            POFIN = m_DateCalcul ' date de calcul dans ce cas
        End If
        
        anc_regl = 0
        anc_regl = DateDiff("m", rsAssure.fields("POARRET"), POFIN)
        If anc_regl < 0 Then
            anc_regl = 0
        End If
        
        If bForReformeRetraite And m_rsPeriode.fields("PEPortefeuilleSalaries") = True Then
        
           ' tables
           feuilleLueIncap = m_rsPeriode.fields("PELMINCAP_Retraite")
           feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE_Retraite")
        
        Else
                 
           feuilleLueIncap = m_rsPeriode.fields("PELMINCAP")
           feuilleLuePassage = m_rsPeriode.fields("PELMPASSAGE")
           
        End If
            ' calcul correctif comme dans l'incap
         correctif = m_ModuleCalcul_Provisions.FctCalcCorrectifIncapacite(x, anc, anc_regl, feuilleLueIncap, feuilleLuePassage)
         correctif = Arrondi(correctif, NbDecimalePM)
    
        rsAssure.fields("TXTECHN") = iIncap * 100#
        rsAssure.fields("TXFRAIS") = fraisGestionIncap * 100#
        rsAssure.fields("POPM_INCAP_1F") = 0#
        If m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLueIncap, fraisGestionIncap) <> 0 Then
        
        ' test
        'Dim a1
        'Dim b1
        'a1 = rsAssure.fields("NUENRP3I")
        'b1 = rsAssure.fields("PONUMCLE")
        'a1 = m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLuePassage, fraisGestionIncap)
        'b1 = m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLueIncap, 0)
        
            rsAssure.fields("POPM_PASS_1F") = Arrondi((m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLuePassage, fraisGestionIncap) / m_ModuleCalcul_Provisions.FctCalcLanc(x, anc, feuilleLueIncap, 0)) * (1 - Inval1Cat), NbDecimalePM) ' avec fraisGestionIncap
        Else
            rsAssure.fields("POPM_PASS_1F") = 0#
        End If
        rsAssure.fields("POPM_INCAP_1R") = 0
        rsAssure.fields("POPM_PASS_1R") = 0
        rsAssure.fields("POPM_INVAL_1F") = 0
        rsAssure.fields("POPM_INVAL_1R") = 0
        
        rsAssure.fields("POPM") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
        rsAssure.fields("POPM_SORTIE") = rsAssure.fields("POPRESTATION_AN") * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
        rsAssure.fields("POPM_VAR") = (rsAssure.fields("POPRESTATION_AN") - rsAssure.fields("POPRESTATION_AN_PREC")) * ((rsAssure.fields("POPM_INCAP_1F") / 12) + rsAssure.fields("POPM_PASS_1F"))
                 
        rsAssure.fields("POPM") = Arrondi(rsAssure.fields("POPM"), 2)
        rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE"), 2)
        rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR"), 2)
        
        rsAssure.fields("POPM_REVALO") = 0
        rsAssure.fields("POCOT_REVALO") = 0

        If Not m_rsPMRetenue.EOF Then  'est que la pm retenue est avec ou sans correctif dans la table pmretenue
         
            rsAssure.fields("POPM_AvecCorrectif") = Arrondi(rsAssure.fields("POPM") * correctif, 2)
            rsAssure.fields("POPM_SansCorrectif") = Arrondi(rsAssure.fields("POPM"), 2)
            rsAssure.fields("POCorrectif") = correctif
            rsAssure.fields("POPMAvecCorrectif") = m_rsPMRetenue.fields("PMAvecCorrectif")
            rsAssure.fields("POPMReassAvecCorrectif") = m_rsPMRetenue.fields("PMReassAvecCorrectif")
           
            If m_rsPMRetenue.fields("PMAvecCorrectif") = True Then
                rsAssure.fields("POPM") = rsAssure.fields("POPM_AvecCorrectif")
                rsAssure.fields("POPM_SORTIE") = Arrondi(rsAssure.fields("POPM_SORTIE") * correctif, 2)
                rsAssure.fields("POPM_VAR") = Arrondi(rsAssure.fields("POPM_VAR") * correctif, 2)
            End If
        End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           P3I COLLECTIVE 31 07 2013
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
    ' pas de calcul
    ' on force les coeff de provisions incapacite à 0
    rsAssure.fields("POPM_INVAL_1F") = 0
    rsAssure.fields("POPM_INCAP_1F") = 0
    rsAssure.fields("POPM_PASS_1F") = 0
    rsAssure.fields("POPM") = 0
    rsAssure.fields("POPM_REVALO") = 0
    rsAssure.fields("POPM_SORTIE") = 0
    rsAssure.fields("POPM_VAR") = 0
    rsAssure.fields("POPM_INVAL_1R") = 0
    rsAssure.fields("POPM_INCAP_1R") = 0
    rsAssure.fields("POPM_PASS_1R") = 0
    rsAssure.fields("POPM_REVALO") = 0
   
    End If
Exit Sub
 
err_calcul:
    MsgBox "erreur " & Err & " à la ligne NUENRP3I = " & rsAssure.fields("NUENRP3I") & vbLf & Err.Description, vbCritical
    'MsgBox "erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A75F10247
Private Function CalculAge(DateInf As Date, DateSup As Date) As Integer
     
  If m_rsPeriode.fields("PEAGE_ANNIVERSAIRE") = False Then
    ' âge par millésime
    CalculAge = Year(DateSup) - Year(DateInf)
  Else
    ' âge anniversaire
    ' CalculAge = Int(DateDiff("m", DateInf, DateSup) / 12#)
    CalculAge = Int(NbMois(DateInf, DateSup) / 12#)
  End If
  
  If CalculAge < 0 Then
    CalculAge = 0
  End If
  
End Function


'##ModelId=5C8A75F1029D
Private Function CalculAgeDouble(DateInf As Date, DateSup As Date) As Double
     
  If m_rsPeriode.fields("PEAGE_ANNIVERSAIRE") = False Then
    ' âge par millésime
    CalculAgeDouble = Year(DateSup) - Year(DateInf)
  Else
    ' âge anniversaire
    ' CalculAge = Int(DateDiff("m", DateInf, DateSup) / 12#)
    CalculAgeDouble = Round(CDbl(NbMois(DateInf, DateSup)) / 12#, 2)
  End If
  
  If CalculAgeDouble < 0# Then
    CalculAgeDouble = 0#
  End If
  
End Function

'##ModelId=5C8A75F102FE
Private Function NbMois(DATE1 As Date, DATE2 As Date) As Integer

  NbMois = 0
  If Year(DATE1) < Year(DATE2) Then
    NbMois = 12 - Month(DATE1) + 12 * (Year(DATE2) - Year(DATE1) - 1) + Month(DATE2)
  ElseIf Year(DATE1) = Year(DATE2) Then
    NbMois = Month(DATE2) - Month(DATE1)
  End If
  
  If Day(DATE1) < Day(DATE2) Then
    If Day(DATE2) - Day(DATE1) > 15 Then NbMois = NbMois + 1
  Else
    If Day(DATE1) - Day(DATE2) > 15 Then NbMois = NbMois - 1
  End If
  
  If NbMois < 0 Then
    NbMois = 0
  End If

End Function

'##ModelId=5C8A75F1033C
Private Function NbSemaine(DATE1 As Date, DATE2 As Date) As Integer

  NbSemaine = Abs(DateDiff("w", DATE1, DATE2))

End Function


'##ModelId=5C8A75F1039A
Private Sub ClearResultFields(ByRef rsAssure As ADODB.Recordset, bForReformeRetraite As Boolean)
  rsAssure.fields("POPM_INVAL_1F") = Null
  rsAssure.fields("POPM_INCAP_1F") = Null
  rsAssure.fields("POPM_PASS_1F") = Null
  rsAssure.fields("POPM_RCJT_1F") = Null
  rsAssure.fields("POPM_REDUC_1F") = Null
  
  rsAssure.fields("POPM") = Null
  rsAssure.fields("POPM_REVALO") = Null
  rsAssure.fields("POPM_SORTIE") = Null
  rsAssure.fields("POPM_VAR") = Null
  
  rsAssure.fields("POPM_INVAL_1R") = Null
  rsAssure.fields("POPM_INCAP_1R") = Null
  rsAssure.fields("POPM_PASS_1R") = Null
  rsAssure.fields("POPM_REDUC_1R") = Null
  rsAssure.fields("POPM_RCJT_1R") = Null
  
  rsAssure.fields("POPM_REVALO") = Null
  rsAssure.fields("POCOT_REVALO") = Null
  
  rsAssure.fields("POCoeffBCAC") = Null
  rsAssure.fields("POCoeffBCACR") = Null
  
  If bForReformeRetraite = True Then
    rsAssure.fields("POCoeffAmortissement") = Null
    rsAssure.fields("POAgeMiniDepartRetraite") = Null
    rsAssure.fields("POAgeDepartRetraiteTauxPlein") = Null
  End If
End Sub
