VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalcul_PM_MaintienDeces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75E20238"
Option Explicit

' acces aux tables
'##ModelId=5C8A75E20376
Private m_theRecordsetList As clsRecordsetList

'##ModelId=5C8A75E203C4
Public autoMode As Boolean
'##ModelId=5C8A75E203E6
Public m_Logger As clsLogger
'

'##ModelId=5C8A75E203E7
Public Property Set recordsetList(r As clsRecordsetList)
  Set m_theRecordsetList = r
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PMInval : calcul de la PM Maintien Deces pour un invalide
'
'##ModelId=5C8A75E3002A
Public Function PMMaintienInval(x As Integer, anc As Integer, ByVal Duree As Integer, sLoiMaintienInval As String, idMortaliteInval As Long, i As Double, fraisGestion As Double) As Double
  
  Dim xLoi As Integer, xMort As Integer
  Dim rsLoi As ADODB.Recordset, rsMort As ADODB.Recordset
  
  Dim k As Integer
  Dim rq As String
  
  Dim wfinIteration As Boolean
  
  Dim v As Double, PT As Double
  Dim Lanc_Loi As Double
  Dim lk_Loi As Double, lk_Mort As Double, lk1_Mort As Double
  
  On Error GoTo err_PMMaintienInval
  
  '
  ' cas de non-calcul
  '
  If Duree <= 0 Then
    PMMaintienInval = 0#
  End If
  
  '
  ' init variables
  '
  wfinIteration = False
  
  v = 1# / (1# + (i / 100#))
 
  '***
  ' ouvre la table de loi de maintien inval
  '***
  Set rsLoi = m_theRecordsetList.LoadRecordset(sLoiMaintienInval)
  
  xLoi = x
  
  ' vérification de l'age minimun de la table
  If xLoi < m_theRecordsetList.AgeMini(sLoiMaintienInval) Then ' 1ere ligne de la table = XXX ans ou plus
    xLoi = m_theRecordsetList.AgeMini(sLoiMaintienInval)
  End If
  
  ' vérification de l'age maximun de la table
  If xLoi > m_theRecordsetList.AgeMaxi(sLoiMaintienInval) Then ' 1ere ligne de la table = XXX ans ou moins
    xLoi = m_theRecordsetList.AgeMaxi(sLoiMaintienInval)
  End If
  
  '***
  ' se place à l'age x dans la table de loi de maintien inval
  '***
  m_dataHelper.Multi_Find rsLoi, "Age=" & xLoi
  If rsLoi.EOF Then
    PMMaintienInval = 0
    Exit Function
  End If
  
  
  '***
  ' ouvre la table de mortalite inval
  '***
  Set rsMort = m_theRecordsetList.LoadRecordset("MortInval", idMortaliteInval)
  
  xMort = x
  
  ' vérification de l'age minimun de la table
  If xMort < m_theRecordsetList.AgeMini("MortInval") Then ' 1ere ligne de la table = XXX ans ou plus
    xMort = m_theRecordsetList.AgeMini("MortInval")
  End If
  
  ' vérification de l'age maximun de la table
  If xMort > m_theRecordsetList.AgeMaxi("MortInval") Then ' 1ere ligne de la table = XXX ans ou moins
    xMort = m_theRecordsetList.AgeMaxi("MortInval")
  End If
  
  
  '***
  ' se place à l'age x dans la table de mortalite inval
  '***
  m_dataHelper.Multi_Find rsMort, "Age=" & xMort & " AND CleTable=" & idMortaliteInval
  If rsMort.EOF Then
    PMMaintienInval = 0#
    Exit Function
  End If
    
    
  '***
  ' lit Lx(x, anc) loi de maintien
  '***
  On Error Resume Next
  Lanc_Loi = rsLoi.fields("Anc" & anc)
  If Err <> 0 Then
    PMMaintienInval = 0#
    Exit Function
  End If
  
  If Lanc_Loi = 0# Then
    PMMaintienInval = 0#
    Exit Function
  End If


  PT = 0#
  
  ' sommation
  For k = 0 To Duree - anc - 1
    
    On Error Resume Next
    
    Err = 0
    lk_Loi = rsLoi.fields("Anc" & k + anc)
    If Err <> 0 Then
      lk_Loi = 0#
    End If
        
    Err = 0
    lk_Mort = rsMort.fields("Anc" & k + anc)
    If Err <> 0 Then
      lk_Mort = 0#
    End If
        
    Err = 0
    lk1_Mort = rsMort.fields("Anc" & k + anc + 1)
    If Err <> 0 Then
      lk1_Mort = 0#
      If errCollectionAbsent = Err.Number Then
        wfinIteration = True ' l'élément lk1_mort n'existe pas dans la ligne, on arrête l'itération
      End If
    End If
    
    On Error GoTo err_PMMaintienInval
    
    ' fin d'itération (à la première valeur de lk,lk1  nulle ou lk absent)
    If (lk_Loi = 0#) Or (lk_Mort = 0#) Or (lk1_Mort = 0#) Or (wfinIteration = True) Then
      Exit For
    End If
    
    PT = PT + (lk_Loi * ((lk_Mort - lk1_Mort) / lk_Mort) * v ^ (k + 0.5))
  
    'Debug.Print k, lk_Loi, lk_Mort, lk1_Mort
  
  Next k
  
  ' dénominateur
  PT = PT / Lanc_Loi
  
  ' ajoute les frais
  PT = PT * (1# + (fraisGestion / 100#))
  
  ' arrondi du resultat
  PMMaintienInval = Arrondi(PT, NbDecimaleCalcul)
  
  Exit Function

err_PMMaintienInval:
    
  If Not autoMode Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical, "PMMaintienInval"
  Else
    m_Logger.EcritTraceDansLog "Erreur in PMMaintienInval " & Err & " : " & Err.Description
  End If
  
  PMMaintienInval = 0#
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' PMIncap : calcul de la PM Maintien Deces pour un incapable
'
'##ModelId=5C8A75E300A7
Public Function PMMaintienIncap(x As Integer, anc As Integer, ByVal Duree As Integer, sLoiMaintienIncap As String, idMortaliteIncap As Long, sLoiPassage As String, sLoiMaintienInval As String, idMortaliteInval As Long, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, AgeLimite As Integer) As Double
    
  Dim xLoi As Integer, xMort As Integer, xPassage As Integer
  Dim rsLoi As ADODB.Recordset, rsMort As ADODB.Recordset, rsPassage As ADODB.Recordset
  
  Dim k As Integer
  Dim rq As String
  
  Dim wfinIteration As Boolean
  
  Dim v As Double, PT_Mort As Double, PT_Passage As Double, PT_Inval As Double
  Dim Lanc_Loi As Double, Lanc_Mort As Double
  Dim lk_Passage As Double, lk_Mort As Double, lk1_Mort As Double
  
  On Error GoTo err_PMMaintienIncap
  
  '
  ' cas de non-calcul
  '
  If Duree <= 0 Then
    PMMaintienIncap = 0#
  End If
  
  '
  ' init variables
  '
  wfinIteration = False
  
  v = 1# / (1# + (iIncap / 100#))
 
  '***
  ' ouvre la table de loi de maintien incap
  '***
  Set rsLoi = m_theRecordsetList.LoadRecordset(sLoiMaintienIncap)
  
  xLoi = x
  
  ' vérification de l'age minimun de la table
  If xLoi < m_theRecordsetList.AgeMini(sLoiMaintienIncap) Then ' 1ere ligne de la table = XXX ans ou plus
    xLoi = m_theRecordsetList.AgeMini(sLoiMaintienIncap)
  End If
  
  ' vérification de l'age maximun de la table
  If xLoi > m_theRecordsetList.AgeMaxi(sLoiMaintienIncap) Then ' 1ere ligne de la table = XXX ans ou moins
    xLoi = m_theRecordsetList.AgeMaxi(sLoiMaintienIncap)
  End If
  
  '***
  ' se place à l'age x dans la table de loi de maintien incap
  '***
  m_dataHelper.Multi_Find rsLoi, "Age=" & xLoi
  If rsLoi.EOF Then
    PMMaintienIncap = 0
    Exit Function
  End If
  
  
  '***
  ' ouvre la table de loi de passage
  '***
  Set rsPassage = m_theRecordsetList.LoadRecordset(sLoiPassage)
  
  xPassage = x
  
  ' vérification de l'age minimun de la table
  If xPassage < m_theRecordsetList.AgeMini(sLoiPassage) Then ' 1ere ligne de la table = XXX ans ou plus
    xPassage = m_theRecordsetList.AgeMini(sLoiPassage)
  End If
  
  ' vérification de l'age maximun de la table
  If xPassage > m_theRecordsetList.AgeMaxi(sLoiPassage) Then ' 1ere ligne de la table = XXX ans ou moins
    xPassage = m_theRecordsetList.AgeMaxi(sLoiPassage)
  End If
  
  '***
  ' se place à l'age x dans la table de loi de passage
  '***
  m_dataHelper.Multi_Find rsPassage, "Age=" & xPassage
  If rsPassage.EOF Then
    PMMaintienIncap = 0
    Exit Function
  End If
  
  
  '***
  ' ouvre la table de mortalite incap
  '***
  Set rsMort = m_theRecordsetList.LoadRecordset("MortIncap", idMortaliteIncap)
  
  xMort = x
  
  ' vérification de l'age minimun de la table
  If xMort < m_theRecordsetList.AgeMini("MortIncap") Then ' 1ere ligne de la table = XXX ans ou plus
    xMort = m_theRecordsetList.AgeMini("MortIncap")
  End If
  
  ' vérification de l'age maximun de la table
  If xMort > m_theRecordsetList.AgeMaxi("MortIncap") Then ' 1ere ligne de la table = XXX ans ou moins
    xMort = m_theRecordsetList.AgeMaxi("MortIncap")
  End If
  
  
  '***
  ' se place à l'age x dans la table de mortalite incap
  '***
  m_dataHelper.Multi_Find rsMort, "Age=" & xMort & " AND CleTable=" & idMortaliteIncap
  If rsMort.EOF Then
    PMMaintienIncap = 0#
    Exit Function
  End If
    
    
  '***
  ' lit Lx(x, anc) loi de maintien incap
  '***
  On Error Resume Next
  Lanc_Loi = rsLoi.fields("Anc" & anc)
  If Err <> 0 Then
    PMMaintienIncap = 0#
    Exit Function
  End If
  
  If Lanc_Loi = 0# Then
    PMMaintienIncap = 0#
    Exit Function
  End If
  
  '***
  ' lit Lx(x, anc) mortalite incap
  '***
  On Error Resume Next
  Lanc_Mort = rsMort.fields("Anc" & anc)
  If Err <> 0 Then
    PMMaintienIncap = 0#
    Exit Function
  End If
  
  If Lanc_Mort = 0# Then
    PMMaintienIncap = 0#
    Exit Function
  End If


  '***
  ' sommation mortalité
  '***
  PT_Mort = 0#
  wfinIteration = False
  
  For k = 0 To Duree - anc - 1
    
    On Error Resume Next
    
    Err = 0
    lk_Mort = rsMort.fields("Anc" & k + anc)
    If Err <> 0 Then
      lk_Mort = 0#
    End If
        
    Err = 0
    lk1_Mort = rsMort.fields("Anc" & k + anc + 1)
    If Err <> 0 Then
      lk1_Mort = 0#
      If errCollectionAbsent = Err.Number Then
        wfinIteration = True ' l'élément lk1_mort n'existe pas dans la ligne, on arrête l'itération
      End If
    End If
    
    On Error GoTo err_PMMaintienIncap
    
    ' fin d'itération (à la première valeur de lk,lk1  nulle ou lk absent)
    If (lk_Mort = 0#) Or (lk1_Mort = 0#) Or (wfinIteration = True) Then
      Exit For
    End If
    
    PT_Mort = PT_Mort + ((lk_Mort - lk1_Mort) * v ^ ((k + 1) / 12#))
  
  Next k
  
  ' dénominateur
  PT_Mort = PT_Mort / Lanc_Mort
  
  
  '***
  ' sommation passage
  '***
  PT_Passage = 0#
  wfinIteration = False
  
  For k = 1 To Duree - anc
    
    On Error Resume Next
    
    Err = 0
    lk_Passage = rsPassage.fields("Anc" & k + anc)
    If Err <> 0 Then
      lk_Passage = 0#
      If errCollectionAbsent = Err.Number Then
        wfinIteration = True ' l'élément lk_Passage n'existe pas dans la ligne, on arrête l'itération
      End If
    End If
    
    On Error GoTo err_PMMaintienIncap
    
    ' fin d'itération (à la première valeur de lk,lk1  nulle ou lk absent)
    If (wfinIteration = True) Then
      Exit For
    End If
    
    PT_Passage = PT_Passage + (lk_Passage * v ^ (k / 12#))
  
  Next k
  
  ' dénominateur
  PT_Passage = PT_Passage / Lanc_Loi
  
  
  '***
  ' PM Inval
  '***                             duree = calcul jusqu'à fin table
  'PHM 07/10/2011 : PT_Inval = PMMaintienInval(x, 0, 999, sLoiMaintienInval, idMortaliteInval, iInval, fraisGestionInval)
  PT_Inval = PMMaintienInval(x, 0, 999, sLoiMaintienInval, idMortaliteInval, iInval, fraisGestionInval)
  'PHM 02/11/2011 : PT_Inval = PMMaintienInval(x, 0, AgeLimite - x, sLoiMaintienInval, idMortaliteInval, iInval, fraisGestionInval)
  
  
  '***
  ' calcul final
  '***
  
  ' ajoute les frais
  PMMaintienIncap = (1# + (fraisGestionIncap / 100#)) * (PT_Mort + PT_Passage * PT_Inval)
  
  ' arrondi du resultat
  PMMaintienIncap = Arrondi(PMMaintienIncap, NbDecimaleCalcul)
  
  Exit Function

err_PMMaintienIncap:
  
  If Not autoMode Then
    MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical, "PMMaintienInval"
  Else
    m_Logger.EcritTraceDansLog "Erreur in PMMaintienInval " & Err & " : " & Err.Description
  End If
  PMMaintienIncap = 0#
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' mets à jour les min et max par rapport à la table passée
' Utilise sTableName ou idTable
'
'##ModelId=5C8A75E30182
Public Sub OpenTableAndGetAgeAndAncMinMax(sTableName As String, idTable As Long, ByRef AgeMin As Integer, ByRef AgeMax As Integer, ByRef AncMin As Integer, ByRef AncMax As Integer)
  Dim rs As ADODB.Recordset, f As ADODB.field
  Dim workAncMin As Integer, workAncMax As Integer
  
  ' recherche du max de la table
  workAncMax = 0
  workAncMin = 0
  
  If idTable = 0 Then
    ' table structure Age Anc0 Anc1 ... AncMax
    
    Set rs = m_theRecordsetList.LoadRecordset(sTableName)
    
    ' lit de l'ancienneté mini et maxi
    For Each f In rs.fields
      If UCase(Left(f.Name, 3)) = "ANC" Then
        If workAncMin > CInt(mID(f.Name, 4)) Then
          workAncMin = CInt(mID(f.Name, 4))
        End If
        
        If workAncMax < CInt(mID(f.Name, 4)) Then
          workAncMax = CInt(mID(f.Name, 4))
        End If
      End If
    Next
    
  Else
    ' table structure
    '   CleTable Age Anciennete Valeur
    ' OU
    '   CleTable Age Anc0 Anc1 ... AncMax
    Set rs = m_theRecordsetList.LoadRecordset(sTableName, idTable)
    
    ' test la structure de la table
    On Error Resume Next
    Err = 0
    Set f = rs.fields("anciennete")
    If Err <> 0 And Not f Is Nothing Then
      ' table structure CleTable Age Anciennete Valeur
      On Error GoTo 0
      
      workAncMin = m_dataHelper.GetParameter("SELECT MIN(Anciennete) FROM " & sTableName & " WHERE CleTable=" & idTable)
      workAncMax = m_dataHelper.GetParameter("SELECT MAX(Anciennete) FROM " & sTableName & " WHERE CleTable=" & idTable)
    Else
      ' table structure CleTable Age Anc0 Anc1 ... AncMax
      On Error GoTo 0
      
      ' lit de l'ancienneté mini et maxi
      For Each f In rs.fields
        If UCase(Left(f.Name, 3)) = "ANC" Then
          If workAncMin > CInt(mID(f.Name, 4)) Then
            workAncMin = CInt(mID(f.Name, 4))
          End If
          
          If workAncMax < CInt(mID(f.Name, 4)) Then
            workAncMax = CInt(mID(f.Name, 4))
          End If
        End If
      Next
    End If
  End If
  
  ' vérification de l'age minimun de la table
  If AgeMin < m_theRecordsetList.AgeMini(sTableName) Then
    AgeMin = m_theRecordsetList.AgeMini(sTableName)
  End If
  
  ' vérification de l'age maximun de la table
  If AgeMax > m_theRecordsetList.AgeMaxi(sTableName) Then
    AgeMax = m_theRecordsetList.AgeMaxi(sTableName)
  End If
  
  ' vérification de l'ancienneté minimun de la table
  If AncMin < workAncMin Then
    AncMin = workAncMin
  End If
  
  ' vérification de l'ancienneté maximun de la table
  If AncMax > workAncMax Then
    AncMax = workAncMax
  End If
End Sub



'##ModelId=5C8A75E301F0
Public Sub CalculTableCoeffPMMaintienDecesIncap(sLoiMaintienIncap As String, idMortaliteIncap As Long, sLoiPassage As String, sLoiMaintienInval As String, idMortaliteInval As Long, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, rsProvisionBCAC As ADODB.Recordset, idTableBCAC As Long, ProgressBar1 As ProgressBar, AgeLimite As Integer)
  Dim x As Integer, xMin As Integer, xMax As Integer
  Dim anc As Integer, AncMin As Integer, AncMax As Integer
  Dim PT As Double, pmBCAC As clsProvisionBCAC
    
  ' init des bornes
  xMin = 0
  xMax = 999
  
  AncMin = 0
  AncMax = 999
    
  ' recupere les bornes min et max en fonction de toutes les tables
  OpenTableAndGetAgeAndAncMinMax sLoiMaintienIncap, 0, xMin, xMax, AncMin, AncMax
  OpenTableAndGetAgeAndAncMinMax "MortIncap", idMortaliteIncap, xMin, xMax, AncMin, AncMax
  OpenTableAndGetAgeAndAncMinMax sLoiPassage, 0, xMin, xMax, AncMin, AncMax
  OpenTableAndGetAgeAndAncMinMax sLoiMaintienInval, 0, xMin, xMax, AncMin, AncMax
  OpenTableAndGetAgeAndAncMinMax "MortInval", idMortaliteInval, xMin, xMax, AncMin, AncMax
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = xMin
  ProgressBar1.Value = xMin
  ProgressBar1.Max = xMax + 1
  
  ' calcul
  For x = xMin To xMax
    ProgressBar1.Value = x
    ProgressBar1.Refresh
    
    For anc = AncMin To AncMax
      ' calcul la provision
      PT = PMMaintienIncap(x, anc, Minimum(35, Abs((AgeLimite - x) * 12 - anc)), sLoiMaintienIncap, idMortaliteIncap, sLoiPassage, _
                           sLoiMaintienInval, idMortaliteInval, iIncap, fraisGestionIncap, _
                           iInval, fraisGestionInval, AgeLimite)
      
      ' ajoute la valeur dans la table
      Set pmBCAC = New clsProvisionBCAC
      
      If autoMode Then
        pmBCAC.autoMode = True
        Set pmBCAC.m_Logger = m_Logger
      End If
      
      pmBCAC.m_CleTable = idTableBCAC
      pmBCAC.m_Age = x
      pmBCAC.m_Anciennete = anc
      pmBCAC.m_Provision = PT
      
      pmBCAC.Save m_dataSource
      
      Set pmBCAC = Nothing
    Next
  Next
  
  ProgressBar1.Visible = False
End Sub



'##ModelId=5C8A75E302E4
Public Sub CalculTableCoeffPMMaintienDecesInval(sLoiMaintienInval As String, idMortaliteInval As Long, iInval As Double, fraisGestionInval As Double, rsProvisionBCAC As ADODB.Recordset, idTableBCAC As Long, ProgressBar1 As ProgressBar, AgeLimite As Integer)
  Dim x As Integer, xMin As Integer, xMax As Integer
  Dim anc As Integer, AncMin As Integer, AncMax As Integer
  Dim PT As Double
    
  ' init des bornes
  xMin = 0
  xMax = 999
  
  AncMin = 0
  AncMax = 999
    
  ' recupere les bornes min et max en fonction de toutes les tables
  OpenTableAndGetAgeAndAncMinMax sLoiMaintienInval, 0, xMin, xMax, AncMin, AncMax
  OpenTableAndGetAgeAndAncMinMax "MortInval", idMortaliteInval, xMin, xMax, AncMin, AncMax
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = xMin
  ProgressBar1.Value = xMin
  ProgressBar1.Max = xMax + 1
  
  ' calcul
  For x = xMin To xMax
    ProgressBar1.Value = x
    ProgressBar1.Refresh
    
    For anc = AncMin To AncMax
      ' calcul la provision
      PT = PMMaintienInval(x, anc, AgeLimite - x, sLoiMaintienInval, idMortaliteInval, iInval, fraisGestionInval)
      
      ' ajoute la valeur dans la table
      rsProvisionBCAC.AddNew
      
      rsProvisionBCAC.fields("CleTable") = idTableBCAC
      rsProvisionBCAC.fields("Age") = x
      rsProvisionBCAC.fields("Anciennete") = anc
      rsProvisionBCAC.fields("Provision") = PT
      
      rsProvisionBCAC.Update
    Next
  Next
  
  ProgressBar1.Visible = False
End Sub

