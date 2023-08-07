VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iP3IGeneraliImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période dans la
' table Assuré
'
'FIXIT: Declare 'CommonDialog1' with an early-bound data type                              FixIT90210ae-R1672-R1B8ZE
Public Function DoImport(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Long, _
                         NumPeriode As Long, DateDebut As String, DateFin As String, _
                         nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, _
                         dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, _
                         Optional autoMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactivitA As Long, _
                         Optional typeCalculAnnualisationA As Long, Optional excelFileA As String, Optional logFileA As String) As Boolean


End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant au controle d'un lot de données
'
'FIXIT: Declare 'CommonDialog1' with an early-bound data type                              FixIT90210ae-R1672-R1B8ZE
Public Function DoControle(CommonDialog1 As Object, ByRef destDB As DataAccess, CleGroupe As Integer, _
                           NumeroLot As Long, NumPeriode As Long, sFichierIni As String) As Boolean

End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Methode servant à l'import des données d'une période
' depuis un lot présent dans P3ILOGTRAIT et P3IPROVCOLL
'
Public Function DoImportSASP3I(NumeroLot As Long, logPath As String, ByRef destDB As DataAccess, CleGroupe As Long, _
                               NumPeriode As Long, DateDebut As String, DateFin As String, _
                               nbJourMax As Long, nbJourDC As Long, AgeRetraite As Byte, _
                               dateCloture As Date, sFichierIni As String, m_bP3I_Individuel As Boolean, _
                               Optional autoMode As Boolean, Optional dateArreteA As Date, Optional typeDelaiInactivitA As Long, _
                               Optional typeCalculAnnualisationA As Long, Optional logFileA As String) As Boolean

End Function

'For Statutaire Functionality
Public Function SetStatutaireVariables(mNumPeriodeStat As Long, mNumPeriodeNonStat As Long, mPathSexFileExcel As String, _
    mCategoryCodeSTAT As String, mSexAllMale As Boolean, mTwoLotImport As Boolean) As Boolean

End Function







