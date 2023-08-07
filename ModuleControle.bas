Attribute VB_Name = "ModuleControle"
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function GetWinUser() As String

  Dim sDomainName As String * 255
  Dim lDomainNameLength As Long
  Dim sUserName As String
  Dim bUserSid(255) As Byte
  Dim lSIDType As Long
  
  sUserName = String(100, Chr$(0))
  GetUserName sUserName, 100
  GetWinUser = Left$(sUserName, InStr(sUserName, Chr$(0)) - 1)

End Function


Private Sub Dump(sTitre As String, szSQL As String, m_Logger As clsLogger, destDB As DataAccess, bStoredProcedure As Boolean, Optional NumeroLot As Long)

  On Error GoTo err_dump
  
  ' titre
  m_Logger.EcritTraceDansLog sTitre
  
  Dim rs As ADODB.Recordset, f As ADODB.Field, sz As String
  Dim nbRecord As Long
  
  ' requete de controle
  If Not bStoredProcedure Then
    Set rs = destDB.OpenRecordset(szSQL, Snapshot)
  Else
    Set rs = destDB.OpenRecordset(szSQL, StoredProcedure) ' doit contenir SET NOCOUNT ON
    'rs.MoveFirst
    
'    Set cmd = New ADODB.Command
'
'    Set cmd.ActiveConnection = destDB.Connection
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = szSQL
'
'    cmd.Parameters.Refresh
'
'    cmd.Parameters("@NumeroLot") = NumeroLot
'
'    Set rs = cmd.Execute
  End If
  
  ' liste des champs
  sz = ""
  For Each f In rs.Fields
    sz = sz & f.Name & "-"
  Next
  m_Logger.EcritTraceDansLog sz
  
  nbRecord = 0
  Do Until rs.EOF
    ' valeurs
    sz = ""
    nbRecord = nbRecord + 1
    
    For Each f In rs.Fields
      If IsNull(f.Value) Then
        sz = sz & "-"
      Else
        sz = sz & f.Value & "-"
      End If
    Next
    m_Logger.EcritTraceDansLog sz
    
    rs.MoveNext
  Loop
  
  m_Logger.EcritTraceDansLog "Nb lignes trouvees : " & nbRecord ' la proxcedure stockée doit utiliser SET NOCOUNT ON qui ne renvoi pas le nb de record
  m_Logger.EcritTraceDansLog ""
  
  rs.Close
  
  Exit Sub
  
err_dump:
  m_Logger.EcritTraceDansLog "Erreur " & Err & " - " & Err.Description
  Resume Next
End Sub


Public Function DoControle(CommonDialog1 As Object, destDB As DataAccess, CleGroupe As Integer, _
                           NumeroLot As Long, NumPeriode As Long, sFichierIni As String) As Boolean
  
  DoControle = False
  
  Dim sz As String, szSQL As String
  
  Dim m_Logger As clsLogger
  Set m_Logger = New clsLogger
  
  m_Logger.FichierLog = sReadIniFile("Dir", "LogPath", "##", 255, sFichierIni) & GetWinUser & "_ControleImport.log"
  m_Logger.CreateLog "Controle du lot " & NumeroLot
  m_Logger.EcritTraceDansLog ""
  m_Logger.EcritTraceDansLog "(Vous pouvez utiliser le bouton Excel pour visualiser ce fichier en colonne)"
  m_Logger.EcritTraceDansLog ""
  
  ' form de selection des controles a effectuer
  Dim frmCC As New frmChoixControle
  
  ret_code = 0
  
  
  '
  ' /!\ Pas d'accent à cause de l'export Excel qui génére des caractères chinois /!\
  '
    
  frmCC.sFichierIni = sFichierIni
  frmCC.Caption = "Contrôle du lot n°" & NumeroLot
  frmCC.Show vbModal
  
  If ret_code = 0 Then
    
    Screen.MousePointer = vbHourglass
    
    If frmCC.chkCodeProv.Value = vbChecked Then
      '*** Controle Code_Prov dans TBQREGA
      szSQL = "SELECT A.Code_GE, A.Lib_Long_GE, A.CODE_PROV" _
              & " FROM  TBQREGA AS A" _
              & " WHERE A.NumPeriode=" & NumPeriode & " AND (A.Code_GE IS NULL OR A.CODE_PROV NOT IN(SELECT CodeProv FROM CodeProvision) )"
      Dump "Controle Code_Prov invalide dans TBQREGA", szSQL, m_Logger, destDB, False
    End If
   
   
    If frmCC.chkCDPRODUIT.Value = vbChecked Then
      '*** Controle CDPRODUIT et NumParamCalcul dans CODESCAT
      szSQL = "SELECT DISTINCT P.CDCOMPAGNIE, P.CDAPPLI, P.CDPRODUIT" _
              & " FROM  P3IPROVCOLL AS P LEFT OUTER JOIN " _
              & " CODESCAT AS C ON (P.CDCOMPAGNIE = C.Code_CIE AND P.CDAPPLI = C.Code_APP AND P.CDPRODUIT = C.Code_Cat_Contrat AND C.NumPeriode=" & NumPeriode & ")" _
              & " WHERE   (P.NUTRAITP3I = " & NumeroLot & ") AND (C.Code_CIE IS NULL)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle CDPRODUIT absent dans CODESCAT", szSQL, m_Logger, destDB, False
      
      szSQL = "SELECT DISTINCT P.CDCOMPAGNIE, P.CDAPPLI, P.CDPRODUIT, C.NumParamCalcul" _
              & " FROM  P3IPROVCOLL AS P INNER JOIN " _
              & " CODESCAT AS C ON (P.CDCOMPAGNIE = C.Code_CIE AND P.CDAPPLI = C.Code_APP AND P.CDPRODUIT = C.Code_Cat_Contrat AND C.NumPeriode=" & NumPeriode & ") " _
              & " LEFT JOIN ParamCalcul PRM ON (PRM.PENUMPARAMCALCUL=C.NumParamCalcul AND PRM.PENUMCLE=" & NumPeriode & ") " _
              & " WHERE   (P.NUTRAITP3I = " & NumeroLot & ") AND (IsNull(C.NumParamCalcul,0)=0 OR IsNull(PRM.PENUMPARAMCALCUL,0)=0) " _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Parametre de calcul a renseigner", szSQL, m_Logger, destDB, False
    End If
    
    
    If frmCC.chkCodeGE.Value = vbChecked Then
      '*** Controle Code GE
      szSQL = "SELECT DISTINCT P.CDGARAN " _
              & " FROM    P3IPROVCOLL AS P LEFT OUTER JOIN" _
              & " TBQREGA AS A ON (P.CDGARAN = A.Code_GE AND A.NumPeriode=" & NumPeriode & ")" _
              & " WHERE   (P.NUTRAITP3I = " & NumeroLot & ") AND (A.Code_GE IS NULL)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Code GE absent dans TBQREGA", szSQL, m_Logger, destDB, False
    End If
    

    If frmCC.chkCATR9.Value = vbChecked Then
      '*** Controle CATR9 et CATR9INVAL
      szSQL = "SELECT DISTINCT P.CDPRODUIT, P.CDGARAN, T.Code_Prov, CI9.Categorie as CATR9INVAL, C9.Categorie AS CATR9 " _
              & " FROM  P3IPROVCOLL AS P" _
              & "       INNER JOIN TBQREGA AS T ON T.Code_CIE=P.CDCOMPAGNIE AND T.Code_APP=P.CDAPPLI AND T.Code_GE=P.CDGARAN " _
              & "       LEFT JOIN CATR9INVAL AS CI9 ON CI9.GroupeCle=T.GroupeCle AND CI9.NumPeriode=T.NumPeriode AND CI9.Categorie=P.CDPRODUIT " _
              & "       LEFT JOIN CATR9 AS C9 ON C9.GroupeCle=T.GroupeCle AND C9.NumPeriode=T.NumPeriode AND C9.Categorie=P.CDPRODUIT" _
              & " WHERE   (P.NUTRAITP3I = " & NumeroLot & ") AND (T.NumPeriode=" & NumPeriode & ") AND (T.Code_Prov BETWEEN 1 AND 4) " _
              & "         AND (CI9.Categorie IS NULL AND C9.Categorie IS NULL)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Categories absentes de CATR9 ou CATR9INVAL pour les arrets de travail", szSQL, m_Logger, destDB, False
    End If
    

    If frmCC.chkChoixPrest.Value = vbChecked Then
      '*** Controle CDCHOIXPREST
      szSQL = "SELECT  DISTINCT P.CDCHOIXPREST, C.CategorieInval " _
              & " FROM  P3IPROVCOLL AS P" _
              & "       INNER JOIN TBQREGA AS T ON T.Code_CIE=P.CDCOMPAGNIE AND T.Code_APP=P.CDAPPLI AND T.Code_GE=P.CDGARAN " _
              & "       LEFT JOIN CODECATINV AS C ON C.GroupeCle=T.GroupeCle AND C.NumPeriode=T.NumPeriode AND C.CDCHOIXPREST=P.CDCHOIXPREST " _
              & " WHERE   (P.NUTRAITP3I = " & NumeroLot & ") AND (T.NumPeriode=" & NumPeriode & ") AND (T.Code_Prov BETWEEN 1 AND 4) " _
              & "         AND (IsNull(C.CategorieInval,0) NOT BETWEEN 1 AND 3) " _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle CDCHOIXPREST inconnu ou categorie inval incorecte dans CODECATINV", szSQL, m_Logger, destDB, False
    End If


    If frmCC.chkDate.Value = vbChecked Then
      '*** Contrôle des Dates
      ControleDates NumeroLot, destDB, m_Logger
    End If
        
    
    
    If frmCC.chkMontant.Value = vbChecked Then
      '*** Controle Montant Prestation
      sz = sReadIniFile("P3I", "ControleMontantMax", "100000", 20, sFichierIni)
      szSQL = "SELECT NUENRP3I, MTPREANN, MTPREREV, MTPREMAJ " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND ((MTPREANN+MTPREREV+MTPREMAJ)>" & sz & " OR (IsNull(MTPREANN,0)+IsNull(MTPREREV,0)+IsNull(MTPREMAJ,0))<=0)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Montant Prestation <= 0 ou > " & sz, szSQL, m_Logger, destDB, False
    End If
    
    
    
    If frmCC.chkDateNaissSurv.Value = vbChecked Then
      '*** Controle Date de Naissance < Date de Survenance
      szSQL = "SELECT NUENRP3I, DTNAISSASS, DTSURVSIN " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND ((DTNAISSASS<>0 AND DTNAISSASS>=DTSURVSIN) OR (DTNAISSREN<>0 AND DTNAISSREN>=DTSURVSIN))" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Date de Naissance < Date de Survenance", szSQL, m_Logger, destDB, False
    End If
    
    
    
    If frmCC.chkDateInvalSurv.Value = vbChecked Then
      '*** Controle Date de mise en inval < Date de Survenance
      szSQL = "SELECT NUENRP3I, DTMISINV, DTSURVSIN " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND (DTMISINV<>0 AND DTMISINV<DTSURVSIN)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Date de Mise en Invalidite < Date de Survenance", szSQL, m_Logger, destDB, False
    End If
    
    
    
    If frmCC.chkNaiss1900.Value = vbChecked Then
      '*** Controle Date de Naissance < 01/01/1900
      szSQL = "SELECT NUENRP3I, DTNAISSASS " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND ((DTNAISSASS<>0 AND DTNAISSASS<19000101) OR (DTNAISSREN<>0 AND DTNAISSREN<19000101))" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Date de Naissance < 01/01/1900", szSQL, m_Logger, destDB, False
    End If
        
    
    
    If frmCC.chkNom.Value = vbChecked Then
      '*** Controle LBASSURE et LBRENTIER vide
      szSQL = "SELECT NUENRP3I, LBASSURE, LBRENTIER " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND (LBASSURE='' OR LBASSURE IS NULL) AND (LBRENTIER='' OR LBRENTIER IS NULL)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Nom Assure", szSQL, m_Logger, destDB, False
      
      '*** Controle LBRENTIER vide et rente
      szSQL = "SELECT P.NUENRP3I, A.Code_Prov, P.LBRENTIER  " _
              & " FROM    P3IPROVCOLL AS P INNER JOIN " _
              & "         TBQREGA AS A ON (P.CDGARAN = A.Code_GE AND A.NumPeriode = " & NumPeriode & "  AND A.Code_Prov BETWEEN 20 AND 30) " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ") AND (P.LBRENTIER='' OR P.LBRENTIER IS NULL) " _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Nom Rentier", szSQL, m_Logger, destDB, False
    End If
        
    
    If frmCC.chkAgeRente.Value = vbChecked Then
      '*** Controle Age >= 18 ans hors Rente
      szSQL = "SELECT P.NUENRP3I, P.CDGARAN, R.Lib_Long_GE, R.Code_PROV, P.DTNAISSASS, P.DTSURVSIN, P.AGESURVSIN " _
              & " FROM TBQREGA AS R RIGHT JOIN P3IPROVCOLL AS P ON R.Code_GE = P.CDGARAN And R.NumPeriode = " & NumPeriode & " AND R.GroupeCle = " & CleGroupe _
              & " WHERE (P.NUTRAITP3I = " & NumeroLot & ")  " _
              & "      AND (R.Code_PROV <> " & cdPositImport_RenteEducation & ") " _
              & "      AND (P.AGESURVSIN<>0 AND (P.AGESURVSIN<18 OR P.AGESURVSIN>65))" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Age >=18 ans et Age <=65 ans pour Garantie Hors Rente Education", szSQL, m_Logger, destDB, False
    End If
        
    
    If frmCC.chkRenteEduc.Value = vbChecked Then
      Dim dateCalcul As Date
      
      dateCalcul = destDB.CreateHelper().GetParameter("SELECT PEDATEEXT FROM Periode WHERE PENUMCLE = " & NumPeriode)
    
      '*** Controle Age <= 26 ans Rente education
      szSQL = "SELECT DISTINCT P.NUENRP3I, P.CDGARAN, R.Lib_Long_GE, R.Code_PROV, P.DTNAISSREN, P.DTSURVSIN, CAST(Round(P.DTSURVSIN/10000,0)-Round(P.DTNAISSREN/10000,0) AS int) AGESURVSIN, CAST(" & Year(dateCalcul) & "-Round(P.DTNAISSREN/10000,0) AS int)  [AGECALCUL] " _
              & " FROM TBQREGA AS R RIGHT JOIN P3IPROVCOLL AS P ON R.Code_GE = P.CDGARAN And R.NumPeriode = " & NumPeriode & " AND R.GroupeCle = " & CleGroupe _
              & " WHERE (P.NUTRAITP3I = " & NumeroLot & ")  " _
              & "      AND (R.Code_PROV = " & cdPositImport_RenteEducation & ") " _
              & "      AND ((" & Year(dateCalcul) & "-Round(P.DTNAISSREN/10000,0))>26)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Controle Age <=26 ans pour Garantie Rente Education", szSQL, m_Logger, destDB, False
    End If
    
    
    If frmCC.chkDoublon.Value = vbChecked Then
      '*** Controle des doublons
      sz = sReadIniFile("P3I", "ScriptRechercheDoublon", App.Path & "\Sql\P3I_Recherche_Doublon.sql", 255, sFichierIni)
    
      szSQL = FileToString(sz)
      'szSQL = Replace(szSQL, "@NumeroLot", NumeroLot)
      szSQL = Replace(szSQL, "<NUMEROLOT>", NumeroLot)
      
      Dump "Recherche des doublons", szSQL, m_Logger, destDB, False
    End If
    
    
    If frmCC.chkCapitauxConstitutif.Value = vbChecked Then
      '*** Liste des Capitaux Constitutif GA < 2006
      szSQL = "SELECT NUENRP3I, CDCOMPAGNIE, CDAPPLI, RTRIM(CDPRODUIT) as CDPRODUIT, DTSURVSIN, LBSOUSCR, LBRENTIER " _
              & " FROM P3IPROVCOLL AS P " _
              & " WHERE (NUTRAITP3I = " & NumeroLot & ")"
              
      Dim rs As ADODB.Recordset
              
      Set rs = destDB.OpenRecordset("SELECT Categorie FROM Categorie_GA_2006", Snapshot)
      Do Until rs.EOF
        szSQL = szSQL & " AND RTRIM(CDPRODUIT)<>'" & Trim(rs.Fields("Categorie")) & "' "
        rs.MoveNext
      Loop
      rs.Close
      
      szSQL = szSQL & " AND (CDCOMPAGNIE=14 AND DTSURVSIN<20060101 AND RTRIM(LBSOUSCR)<>'GROUPE ARPEGE' AND LBRENTIER IS NOT NULL)" _
              & " AND ( " _
              & "      P.DataVersion = 1 " _
              & "      OR (P.DataVersion=0 AND NOT EXISTS(SELECT 1 FROM P3IUser.P3IPROVCOLL P2 " _
              & "                                         WHERE P2.NUTRAITP3I=P.NUTRAITP3I AND P2.NUENRP3I=P.NUENRP3I AND P2.DataVersion>=1) ) " _
              & "     ) "
      Dump "Liste des Capitaux Constitutif GA < 2006", szSQL, m_Logger, destDB, False
    End If
    
    
    ' decharge la form
    Unload frmCC
      
    Screen.MousePointer = vbDefault
    
  Else
    ' controle annuler
    m_Logger.EcritTraceDansLog "Contrôle annuler !"
  
  End If
  
  Set frmCC = Nothing
    
  m_Logger.EcritTraceDansLog ""
  m_Logger.EcritTraceDansLog "(Vous pouvez utiliser le bouton Excel pour visualiser ce fichier en colonne)"
  
  m_Logger.AfficheErreurLog False
  
End Function


'Purpose     :  Returns the contents of a file as a single continuous string
'Inputs      :  sFileName               The path and file name of the file to open and read
'Outputs     :  The contents of the specified file
'Notes       :  Usually used for text files, but will load any file type.
'Revisions   :
Function FileLoad(ByVal sFileName As String) As String
    Dim iFileNum As Integer, lFileLen As Long

    On Error GoTo ErrFinish
    'Open File
    iFileNum = FreeFile
    'Read file
    Open sFileName For Binary Access Read As #iFileNum
    lFileLen = LOF(iFileNum)
    'Create output buffer
    FileLoad = String(lFileLen, " ")
    'Read contents of file
    Get iFileNum, 1, FileLoad

ErrFinish:
    Close #iFileNum
    On Error GoTo 0
End Function


Private Sub ControleDates(NumeroLot As Long, destDB As DataAccess, m_Logger As clsLogger)
  On Error GoTo err_ControleDates
  
  ' titre
  m_Logger.EcritTraceDansLog "Contrôle des Dates"
  
  Dim rs As ADODB.Recordset, f As ADODB.Field, sz As String
  Dim fErrorFound As Boolean, lErrorCount As Long
  
  ' requete de controle
'  Set rs = destDB.OpenRecordset("Select * from P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot, snapshot)
  Set rs = destDB.OpenRecordset("Select NUENRP3I, DTNAISSASS, DTNAISSREN, DTNAISSCOR, DTSURVSIN, DTEFFREN, DTLIMPRO, DTDERREG, DTDEBPER, DTFINPER, DTMISINV, DTCALCULPROV, " _
                                & " DTTRAITPROV, DTCREATI, DTSIGBIA, DTDECSIN, DTSITUATSIN, DTPREECH, DTDERECH, DTDEBPERECH1, DTFINPERECH1, DTDEBPERECH2, " _
                                & " DTFINPERECH2, DTDEBPERECH3, DTFINPERECH3, DTDEBPERPIP, DTFINPERPIP, DTSAISIEPERJUSTIF, DTDEBPERJUSTIF, DTFINPERJUSTIF, " _
                                & " DTDEBDERPERRGLTADA, DTFINDERPERRGLTADA, DTDERPERRGLTADA, DTDEBDERPERRGLTADC, DTFINDERPERRGLTADC, DTDERPERRGLTADC, " _
                                & " DTDEBPROV , DTFINPROV From P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot, Snapshot)
  
  lErrorCount = 0
  Do Until rs.EOF
    fErrorFound = False
    sz = "NUENRP3I=" & rs.Fields("NUENRP3I")
    
    For Each f In rs.Fields
      If Left(f.Name, 2) = "DT" Then
        If Not IsNull(f.Value) Then
          If CheckDate(f.Value) = False Then
            sz = sz & ", " & f.Name & "=" & f.Value
            fErrorFound = True
          End If
        End If
      End If
    Next
    
    If fErrorFound = True Then
      m_Logger.EcritTraceDansLog sz
      lErrorCount = lErrorCount + 1
    End If
    
    rs.MoveNext
  Loop
  
  m_Logger.EcritTraceDansLog "Nb lignes trouvées : " & lErrorCount
  m_Logger.EcritTraceDansLog ""
  
  rs.Close
  
  Exit Sub
  
err_ControleDates:
  m_Logger.EcritTraceDansLog "Erreur " & Err & " - " & Err.Description
  Resume Next
End Sub

Public Function CheckDate(lDate As Long) As Boolean
  Dim str As String, d As Date
  
  On Error GoTo err_CheckDate
  
  CheckDate = False
  
  If lDate = 0 Then
    CheckDate = True
    Exit Function
  End If
  
  '                    YYYYMMDD
  str = Format(lDate, "00000000")
  
  d = DateSerial(Left(str, 4), Mid(str, 5, 2), Right(str, 2))
  
  CheckDate = True
  
  Exit Function
  
err_CheckDate:
  CheckDate = False
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie une date sous forme de long en date
' format YYYYMMDD
'
Public Function ConvertDate(lDate As Long) As Date
  Dim str As String
  
  If lDate = 0 Then
    ConvertDate = Null
    Exit Function
  End If
  
  '                    YYYYMMDD
  str = Format(lDate, "00000000")
  
  ConvertDate = DateSerial(Left(str, 4), Mid(str, 5, 2), Right(str, 2))
 
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie une date sous forme de long en date
' format YYYYMMDD
'
Public Function ConvertTime(lTime As Long) As Date
  Dim str As String
  
  If lTime = 0 Then
    ConvertTime = Null
    Exit Function
  End If
  
  '                    HHMMSS
  str = Format(lTime, "000000")
  
  ConvertTime = TimeSerial(Left(str, 2), Mid(str, 3, 2), Right(str, 2))
 
End Function

Public Function CheckTime(lTime As Long) As Boolean
  Dim str As String, d As Date
  
  On Error GoTo err_CheckTime
  
  CheckTime = False
  
  If lTime = 0 Then
    CheckTime = True
    Exit Function
  End If
  
  '                    HHMMSS
  str = Format(lTime, "000000")
  
  d = TimeSerial(Left(str, 2), Mid(str, 3, 2), Right(str, 2))
  
  CheckTime = True
  
  Exit Function
  
err_CheckTime:
  CheckTime = False
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Lecture d'une date
'
Public Function ReadDate(f As ADODB.Field) As Variant
 
  If IsNull(f.Value) Then
    ReadDate = Null
    Exit Function
  ElseIf f.Value = 0 Then
    ReadDate = Null
    Exit Function
  ElseIf f.Value = "" Then
    ReadDate = Null
    Exit Function
  ElseIf CheckDate(f.Value) = False Then
    ReadDate = Null
    Exit Function
  Else
    ReadDate = ConvertDate(f.Value)
  End If

End Function

Public Function ReadDateXL(f As DAO.Field) As Variant
 
  If IsNull(f.Value) Then
    ReadDateXL = Null
  ElseIf f.Value = 0 Then
    ReadDateXL = Null
  ElseIf f.Value = "" Then
    ReadDateXL = Null
  ElseIf IsDate(f.Value) = False Then
    ReadDateXL = Null
  Else
    ReadDateXL = CDate(f.Value)
  End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Lecture d'une date DAO
'
Public Function ReadDateDAO(f As DAO.Field) As Variant
 
  If IsNull(f.Value) Then
    ReadDateDAO = Null
    Exit Function
  ElseIf f.Value = 0 Then
    ReadDateDAO = Null
    Exit Function
  ElseIf IsDate(f.Value) Then
    ReadDateDAO = f.Value
    Exit Function
  End If
  
  ReadDateDAO = Null

End Function


'**********************************************************
'PURPOSE:    Puts all lines of file into a string array
'PARAMETERS: FileName = FullPath of File
'            TheArray = StringArray to which contents
'                       Of File will be added.
'Example
'  Dim sArray() as String
'  FileToArray "C:\MyTextFile.txt", sArray
'  For lCtr = 0 to Ubound(sArray)
'  Debug.Print sArray(lCtr)
'  Next

'NOTES:
'  --  Requires a reference to Microsoft Scripting Runtime
'      Library
'  --  You can write this method in a number of different ways
'      For instance, you can take advantage of VB 6's ability to
'      return an array.
' --   You can also read all the contents of the file and use the
'      Split function with vbCrlf as the delimiter, but I
'      wanted to illustrate use of the ReadLine
'      and AtEndOfStream methods.
Public Function FileToString(ByVal filename As String) As String

  Dim oFSO As New FileSystemObject
  Dim oFSTR As Scripting.TextStream
  Dim ret As Long, lCtr As Long
  Dim sz As String

  FileToString = ""

  If Dir(filename) = "" Then Exit Function

  On Error GoTo ErrorHandler
  
  Set oFSTR = oFSO.OpenTextFile(filename)
   
  Do While Not oFSTR.AtEndOfStream
    
    sz = Trim(oFSTR.ReadLine)
    
    If Left(sz, 2) <> "--" Then
      FileToString = FileToString & sz & vbLf
    End If
      
  Loop
  
  oFSTR.Close
     

ErrorHandler:
  
  Set oFSTR = Nothing

End Function


