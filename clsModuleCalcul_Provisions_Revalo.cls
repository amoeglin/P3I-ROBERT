VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalcul_Provisions_Revalo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75E00268"
Option Explicit

' acces aux tables
'##ModelId=5C8A75E0039B
Private m_theRecordsetList As clsRecordsetList

' module de calcul
'##ModelId=5C8A75E10022
Private m_ModuleCalcul_Provisions As clsCalcul_Provisions

'##ModelId=5C8A75E1008E
Public autoMode As Boolean
'##ModelId=5C8A75E100BE
Public m_Logger As clsLogger
'

'##ModelId=5C8A75E100CC
Public Property Set recordsetList(r As clsRecordsetList)
  Set m_theRecordsetList = r
End Property

'##ModelId=5C8A75E100EB
Public Property Set ModuleCalcul_Provisions(r As clsCalcul_Provisions)
  Set m_ModuleCalcul_Provisions = r
End Property


'*******************************
'  CALCUL PROVISION INCAPACITE *  FONCTION FctCalcProvIncapRevalo calcule UNE SEULE PROVISION pour les paramères passé
'*******************************  'FctCalcProvIncapRevalo(x, anc, duree, franchise, feuilleLue, i, fraisGestion)
                                
'##ModelId=5C8A75E1010B
Public Function FctCalcProvIncapRevalo(x As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, i As Double, fraisGestion As Double, DureRevalo As Single, TauxRevalo As Double) As Double
  Dim v As Double, m As Double, PT As Double, Lanc As Double, lk As Double, lk1 As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim wfinIteration As Boolean
  wfinIteration = False
  Dim FM As Double   ' Flux Mensuel pour calcul revalorisation
  Dim FM1 As Double  ' Flux Mensuel pour calcul revalorisation

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvIncap = 0
 
 ' sélection de la ligne pour l'age demandé dans la feuillelue
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
'  rq = "SELECT * FROM P3IUser." & feuilleLue & " WHERE Age = " & x
'  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLue)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvIncapRevalo = 0
'    rs.Close
    Exit Function
  End If
   
   
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))
  If Lanc = 0 Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvIncapRevalo = 0
'    rs.Close
    Exit Function
  End If
    


 On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvIncap = 0

  m = 12       ' fractionnement mensuel
  v = 1 / (1 + i)
  PT = 0
  
    
  If anc > franchise Then
    kdeb = anc
  Else
    kdeb = franchise
  End If
  
  For k = kdeb To Duree - 1                ' duree = 36 mois - 1
    
    On Error GoTo LkAbsent
    lk = val(rs.fields("Anc" & k))
    
    On Error GoTo Lk1Absent
    lk1 = val(rs.fields("Anc" & k + 1))
    
    On Error GoTo MsgError
    
    ' fin d'itération (à la première valeur de lk,lk1  nulle ou lk absent)
    If (lk = 0) Or (lk1 = 0) Or (wfinIteration = True) Then
      Exit For
    End If
    ' revalo depuis la date de l'arrêt
    'FM = FluxMensuel(k, DureRevalo, TauxRevalo)
    'FM1 = FluxMensuel(k + 1, DureRevalo, TauxRevalo)
    '
    ' revalo depuis la date de calcul
    FM = FluxMensuel(k - anc, DureRevalo, TauxRevalo)
    FM1 = FluxMensuel(k + 1 - anc, DureRevalo, TauxRevalo)
    
    PT = PT + (1 / (2 * Lanc)) * ((lk * FM) * (v ^ ((k - anc) / m)) + (lk1 * FM1) * (v ^ ((k + 1 - anc) / m)))
    ''''PT = PT + (1 / (1 * Lanc)) * ((lk * FM) * (v ^ ((k - anc) / m)))
    
  Next k

FIN_FctCALCUL_PT:
'  rs.Close

  FctCalcProvIncapRevalo = Arrondi(PT * (1 + fraisGestion), NbDecimalePM)  ' avec frais de gestion
  Exit Function
  

LxAbsent:
  Lanc = 0
  Resume Next

LkAbsent:
  lk = 0
  Resume Next

Lk1Absent:
  lk1 = 0
    If errCollectionAbsent = Err.Number Then
    wfinIteration = True ' l'élément Sk1 n'existe pas dans la ligne, on arrête l'itération
    End If
  Resume Next


MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & str(Err.Number) & " générée par " _
        & Err.Source & Chr(13) & Err.Description
        
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvIncapRevalo", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcProvIncapRevalo " & Err & " : " & Err.Description
    End If
    
    FctCalcProvIncapRevalo = 0
'    rs.Close
    Exit Function
    
    Resume Next
  End If
  
  
End Function

'*******************************
'  CALCUL PROVISION INVALIDITE *  FONCTION FctCalcProvInval calcule UNE SEULE PROVISION pour les paramères passé
'*******************************  'FctCalcProvInval(x, anc, duree, franchise, feuilleLue, i, fraisGestion)
                                
'##ModelId=5C8A75E101E5
Public Function FctCalcProvInvalRevalo(x As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, i As Double, fraisGestion As Double, DureRevalo As Single, TauxRevalo As Double) As Double
  Dim v As Double, m As Double, PT As Double, Lanc As Double, lk As Double, lk1 As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim wfinIteration As Boolean: wfinIteration = False
  Dim FA As Double  ' Flux annuel
  Dim FA1 As Double ' Flux annuel anneée + 1

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvInval = 0
 
 ' sélection de la ligne pour l'age demandé dans la feuillelue
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
'  rq = "SELECT * FROM P3IUser." & feuilleLue & " WHERE Age = " & x
'  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLue)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvInvalRevalo = 0
'    rs.Close
    Exit Function
  End If
   
   
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))
  If Lanc = 0 Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvInvalRevalo = 0
'    rs.Close
    Exit Function
  End If
    


 On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvInval = 0

  m = 1       ' fractionnement annuel
  v = 1 / (1 + i)
  PT = 0
  
    
  If anc > franchise Then
    kdeb = anc
  Else
    kdeb = franchise
  End If
  
  For k = kdeb To Duree - 1
    
    On Error GoTo LkAbsent
    lk = val(rs.fields("Anc" & k))
        
    On Error GoTo Lk1Absent
    lk1 = val(rs.fields("Anc" & k + 1))
    
    On Error GoTo MsgError
    
    ' fin d'itération (à la première valeur de lk,lk1  nulle ou lk absent)
    If (lk = 0) Or (lk1 = 0) Or (wfinIteration = True) Then
      Exit For
    End If
    ' revalo depuis la date de l'arrêt
    'FA = FluxAnnuel(k, DureRevalo, TauxRevalo)
    'FA1 = FluxAnnuel(k + 1, DureRevalo, TauxRevalo)
    
    ' revalo depuis la date de calcul
    FA = FluxAnnuel(k - anc, DureRevalo, TauxRevalo)
    FA1 = FluxAnnuel(k + 1 - anc, DureRevalo, TauxRevalo)
   
    PT = PT + (1 / (2 * Lanc)) * ((lk * FA) * (v ^ ((k - anc) / m)) + (lk1 * FA1) * (v ^ ((k + 1 - anc) / m)))
    
  Next k

FIN_FctCALCUL_PT:
'  rs.Close

  FctCalcProvInvalRevalo = Arrondi(PT * (1 + fraisGestion), NbDecimalePM)  ' avec frais de gestion
  Exit Function
  

LxAbsent:
  Lanc = 0
  Resume Next

LkAbsent:
  lk = 0
  Resume Next

Lk1Absent:
  lk1 = 0
  If errCollectionAbsent = Err.Number Then
    wfinIteration = True ' l'élément lk1 n'existe pas dans la ligne, on arrête l'itération
  End If
  Resume Next


MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & str(Err.Number) & " générée par " _
        & Err.Source & Chr(13) & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvInvalRevalo", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcProvInvalRevalo " & Err & " : " & Err.Description
    End If
        
    FctCalcProvInvalRevalo = 0
'    rs.Close
    Exit Function
    
    Resume Next
  End If
  
  
End Function

'**********************************************
'  CALCUL PROVISION   PASSAGE    Incap/Inval  *  FONCTION FctCalcProvPassage calcule UNE SEULE PROVISION pour les paramères passé
'**********************************************  'FctCalcProvPassage
                                
'##ModelId=5C8A75E102A3
Public Function FctCalcProvPassageRevalo(x As Integer, xdepartRetraite As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLueIncap As String, feuilleLuePassage As String, feuilleLueInval As String, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, DureRevalo As Single, TauxRevalo As Double, Optional bSansLissage As Boolean = False) As Double
                                         
  Dim v As Double, m As Double, PT As Double, Lanc As Double, Lduree As Double, Sk As Double, Sk1 As Double, DureeFA As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
    Dim dureeInval As Integer
  Dim PTInvalxDeb As Double, PTInvalx As Double, PTInvalx1 As Double, PTInval As Double, PTInval1 As Double
  Dim wfinIteration As Boolean
  wfinIteration = False
  Dim FM As Double ' Flux Mensuel
  Dim FA As Double ' Flux Annuel

  
  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
 
  ' sélection de la ligne pour l'age demandé dans la feuilleIncap
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
'  rq = "SELECT * FROM P3IUser." & feuilleLueIncap & " WHERE Age = " & x
'  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLueIncap)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvPassageRevalo = 0
'    rs.Close
    Exit Function
  End If
   
   
  
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))            ' effectifs pour anc mois en incap
  
  If Lanc = 0 Then
    FctCalcProvPassageRevalo = 0
'    rs.Close
    Exit Function
  End If
    
 
 ' sélection de la ligne pour l'age demandé dans la feuillePassage
  
  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
'  rq = "SELECT * FROM P3IUser." & feuilleLuePassage & " WHERE Age = " & x
'  Set rs2 = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs2 = m_theRecordsetList.LoadRecordset(feuilleLuePassage)
  m_dataHelper.Multi_Find rs2, "Age=" & x
  
  If rs2.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvPassageRevalo = 0
'    rs2.Close
    Exit Function
  End If
  
  
  m = 12       ' fractionnement mensuel
  v = 1 / (1 + iIncap)
    
  
  If anc > franchise Then
    kdeb = anc
  Else
    kdeb = franchise
  End If
  
      Dim xEntierPrec As Integer, xEntier1Prec As Integer, xEntier As Integer, kMois As Integer
      Dim xDouble As Double
      
      
      PTInvalx = 0           ' pour l'âge xEntier     (x+k)
      PTInvalx1 = 0          ' pour l'âge xEntier + 1 (x+k) + 1
      PTInval = 0            ' pour l'âge x+k par interpolation linéaire
      PTInval1 = 0            ' pour l'âge x+k+1 par interpolation linéaire
      PTInvalxDeb = 0        ' pour (x + kdeb)
      PT = 0                 ' RAZ PT
      
      
  
  For k = kdeb To Duree - 1
          
      
     '********  CALCUL PROVISION INVAL par interpolation linéaire pour les âges entiers (x+k) et (x+k) + 1 ***
      
      ' xDouble = (x * 12 + k) / 12 ' pour les autre options
      xDouble = (x * 12 + k + 1) / 12 ' pour l'option 0   à k+1
      xEntier = Int(xDouble)
      'kMois = (x * 12 + k) - xEntier * 12           ' nombre de mois courrus entre xentier et xEntier+1
      kMois = (x * 12 + k + 1) - xEntier * 12         ' ' pour l'option 0   à k+1
     
      dureeInval = xdepartRetraite - xEntier   ' durée pour l'âge xEntier
    
      If dureeInval < 0 Then
        dureeInval = 0
      Else
      End If
      
      
      ' calcul de PTInvalx à l'âge entier
      If (xEntierPrec <> xEntier) And (xEntier <> 0) Then
        PTInvalx = FctCalcProvInvalRevalo(xEntier, 0, dureeInval, 0, feuilleLueInval, iInval, fraisGestionInval, DureRevalo, TauxRevalo)
      Else
      ' PTInvalx sans changement
      End If
      xEntierPrec = xEntier
     
            
      ' calcul de PTInvalx1 à l'âge entier suivant (xEntier + 1)
      If (xEntier1Prec <> xEntier + 1) Then
      PTInvalx1 = FctCalcProvInvalRevalo(xEntier + 1, 0, dureeInval - 1, 0, feuilleLueInval, iInval, fraisGestionInval, DureRevalo, TauxRevalo)
    
      Else
      ' PTInvalx1 sans changement
      End If
      xEntier1Prec = xEntier + 1
      
      ' interpolation linéaire
      PTInval = PTInvalx - (kMois / 12) * (PTInvalx - PTInvalx1)
      PTInval1 = PTInvalx - ((kMois + 1) / 12) * (PTInvalx - PTInvalx1)

      
      ' stockage de la PROVISION INVAL INITIALE à l'âge entier initial (x+kdeb)
      If PTInvalxDeb = 0 Then
      PTInvalxDeb = PTInval
      End If
      
      '********  FIN INTERPOLATION INVAL  ***********
      
        Sk = 0
        Sk1 = 0
        
        On Error GoTo SkAbsent
        Sk = val(rs2.fields("Anc" & k))         ' lue dans feuilleLuePassage
        
        On Error GoTo Sk1Absent
        Sk1 = val(rs2.fields("Anc" & k + 1))    ' lue dans feuilleLuePassage
        
        On Error GoTo MsgError
        
        
        ''''''''''''''''' calcul option 0 + correctif ''''''''''''''''
        ' revalo depuis la date de l'arrêt
   
        ' revalo depuis la date de calcul
        ' FA = FluxAnnuel(DureeFA, DureRevalo, TauxRevalo)
        
        DureeFA = Int((k + 1 - anc) / 12)
        If DureeFA < 0 Then DureeFA = 0
        FA = FluxAnnuel(DureeFA, DureRevalo, TauxRevalo)
        
        PT = PT + (1 / Lanc) * ((FA * Sk) * (v ^ ((k + 1 - anc) / m))) * PTInval
  
        If wfinIteration = True Then
          Exit For
        End If
  
  Next k
        
  'ajout du correctif
  On Error GoTo LdureeAbsent
  Lduree = val(rs.fields("Anc" & k + 1))     ' effectifs pour duree mois en incap = k+1
  
  On Error GoTo MsgError
  ' revalo depuis la date de l'arrêt

  ' revalo depuis la date de calcul
  'FA = FluxAnnuel(DureeFA, DureRevalo, TauxRevalo)
   
  DureeFA = Int((k + 1 - anc) / 12)
  If DureeFA < 0 Then
    DureeFA = 0
  End If
  
  FA = FluxAnnuel(DureeFA, DureRevalo, TauxRevalo)
  
  
  ' PHM 16/04/2007 : possibilité de désactiver le terme correcteur
  If bSansLissage = False Then
    PT = PT + FA * (Lduree / Lanc) * (v ^ ((k + 1 - anc) / m)) * PTInval
  End If
  
FIN_FctCALCUL_PT:
'  rs.Close        ' loi incap
'  rs2.Close       ' loi passage incap/inval

  ' ajustement de PT qui doit être inférieur à la valeur initiale PTInvalxDdEB
  If PT > PTInvalxDeb Then
    PT = PTInvalxDeb
  Else
  End If
  
  FctCalcProvPassageRevalo = Arrondi(PT * (1 + fraisGestionIncap), NbDecimalePM)
  Exit Function
  

LxAbsent:
  Lanc = 0
  Resume Next



SkAbsent:
  Sk = 0
  Resume Next

Sk1Absent:
  Sk1 = 0
    If errCollectionAbsent = Err.Number Then
    wfinIteration = True ' l'élément Sk1 n'existe pas dans la ligne, on arrête l'itération
    
    End If
    Resume Next


LdureeAbsent:
  Lduree = 0
  Resume Next


MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & str(Err.Number) & " générée par " _
        & Err.Source & Chr(13) & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvPassageRevalo", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcProvPassageRevalo " & Err & " : " & Err.Description
    End If
    
    FctCalcProvPassageRevalo = 0
'    rs.Close
'    rs2.Close
    Exit Function
    
  End If
  
  
End Function


'*****************************************************************
'  CALCUL PROVISION D'une table COMPLETE  POUR UNE GARANTIE Incap ou Passage ou Inval   à partir de feuille lue
'*****************************************************************
'Public Sub CalcTableauProv(x, anciennete, duree, franchise, feuilleLue, iIncap, fraisGestionIncap, iInval, fraisGestionInval, xdepartRetraite, garantie, ProgressBar1)


'##ModelId=5C8A75E103D6
Public Sub CalcTableauProvRevalo(x As Integer, Anciennete As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, xdepartRetraite As Integer, garantie As String, ProgressBar1 As ProgressBar, feuilleLueIncap As String, feuilleLueInval As String)
 
  
  Dim PT As Double
  Dim xdeb As Integer, xfin As Integer, k As Integer, anc As Integer
  Dim feuilleMaj As String, rq As String
  Dim rs As ADODB.Recordset
  
  x = 0
  
  Dim Msg ' libellé message erreur
  
  Screen.MousePointer = vbHourglass
  
  ' construit le nom de la table DESTINATION
  
  If garantie = "Incapacité" Then
     feuilleMaj = "PROV_" & feuilleLue & "_" & (iIncap * 100) & "_" & (fraisGestionIncap * 100)
        
        ElseIf garantie = "Passage" Then
        feuilleMaj = "PROV_" & feuilleLue & "_" & (iIncap * 100) & "_" & (fraisGestionIncap * 100)
  
            ElseIf garantie = "Invalidité" Then
            feuilleMaj = "PROV_" & feuilleLue & "_" & (iInval * 100) & "_" & (fraisGestionInval * 100)
  End If
  
  
  ' remplace '.' par '_'
  Do
    x = InStr(1, feuilleMaj, ".")
    
    If x = 0 Then x = InStr(1, feuilleMaj, ",")
    
    If x <> 0 Then
      feuilleMaj = Left(feuilleMaj, x - 1) + Right(feuilleMaj, Len(feuilleMaj) - x)
    End If
  Loop Until x = 0
  
  Dim oCat As New ADOX.Catalog
  
  Set oCat.ActiveConnection = m_dataSource.Connection
  
  ' crée la table de destination (feuillemaj)
  On Error Resume Next  ' si la table n'existe pas pas erreur
  oCat.Tables.Delete feuilleMaj
  On Error GoTo 0 ' reset detection erreur
          
On Error GoTo ErrCalcTableauProv
  
  m_dataSource.Execute "SELECT * INTO " & feuilleMaj & " FROM " & feuilleLue & " WHERE Age=-999"

'  m_dataSource.Execute "CREATE INDEX " & feuilleMaj & "_Age ON " & feuilleMaj & "(Age) with PRIMARY"
  m_dataSource.Execute "ALTER TABLE " & feuilleMaj & " ADD CONSTRAINT PK_" & feuilleMaj & " PRIMARY KEY CLUSTERED (Age)"

On Error GoTo 0 ' reset detection erreur
  
  ' cherche l'age de début xdeb =min(age) dans la feuille lue
  rq = "SELECT min(Age) as AnDebut FROM " & feuilleLue
  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  If Not rs.EOF Then
    xdeb = rs.fields("AnDebut")
  Else
    
    If Not autoMode Then
      Screen.MousePointer = vbDefault
      MsgBox "Table " & feuilleLue & " INVALIDE !"
    Else
      m_Logger.EcritTraceDansLog "Table " & feuilleLue & " INVALIDE !"
    End If
    
    rs.Close
    Exit Sub
  End If
  
  rs.Close
  
  ' debut du calcul
  
  ' ouvre les tables de reference
  m_theRecordsetList.InitRecordsetList
  
  m_theRecordsetList.LoadRecordset feuilleLue
  m_theRecordsetList.LoadRecordset feuilleLueIncap
  m_theRecordsetList.LoadRecordset feuilleLueInval
  
  'reprise des paramètres suivants de Sub CalcTableauProvIncap
  'x, anc, duree, franchise, feuilleLue, i, fraisGestion, xdepartRetraite, ProgressBar1
  
  
  
  PT = 0
  k = 0
  x = 0
  xfin = xdepartRetraite
  
  ' affiche la boite d'attente
  If Not ProgressBar1 Is Nothing Then
   ProgressBar1.Visible = True
  
   ProgressBar1.Min = xdeb
   ProgressBar1.Max = xfin
   ProgressBar1.Value = xdeb
  End If
  
  rq = "SELECT * FROM " & feuilleMaj
  Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  ' sauve les parametres de calcul dans la table avec Age=-1
'  rs.AddNew
'  rs.fields("Age") = -1
'  rs.fields("Anc0") = "anc, duree, franchise, feuilleLue, iIncap , GestionIncap, iInval , GestionInval, xdepartRetraite)"
'  rs.fields("Anc1") = anciennete
'  rs.fields("Anc2") = Duree
'  rs.fields("Anc3") = franchise
'  rs.fields("Anc4") = feuilleLue
'  rs.fields("Anc5") = iIncap
'  rs.fields("Anc6") = fraisGestionIncap
'  rs.fields("Anc7") = iInval
'  rs.fields("Anc8") = fraisGestionInval
'  rs.fields("Anc9") = xdepartRetraite
'  rs.fields("Anc10") = garantie
'  rs.fields("Anc11") = feuilleLueIncap
'  rs.fields("Anc12") = feuilleLueInval
'  rs.Update

  rs.AddNew
  rs.fields("Age") = -2
  rs.fields("Comment") = "Anc, Duree, Franchise, Table, iIncap , GestionIncap, iInval , GestionInval, xDepartRetraite, Garantie, TableIncap, TableInval"
  rs.Update

  rs.AddNew
  rs.fields("Age") = -1
  rs.fields("Comment") = Anciennete & ", " & Duree & ", " & franchise & ", " & feuilleLue & ", " & iIncap & ", " & fraisGestionIncap _
                         & ", " & iInval & ", " & fraisGestionInval & ", " & xdepartRetraite & ", " & garantie & ", " & feuilleLueIncap & ", " & feuilleLueInval
  rs.Update
  
  ' calcul pour chaque valeur de x
  For x = xdeb To xfin
    
  ' calcul de la duree pour l'invalidite, pour chaque age x,  duree = xdepartretraite - x
    If garantie = "Invalidité" Then
    Duree = xdepartRetraite - x
    Else
    End If
    
    ' ajoute l'enregistrement
    rs.AddNew
    
    rs.fields("Age") = x
    
    If Not ProgressBar1 Is Nothing Then
      ProgressBar1.Value = x
    End If
    
    
    
    ' calcule et stocke les provisions dans une ligne pour un âge donnée x et les anciennetés de anc à durée-1
    For anc = Anciennete To Duree - 1
           
      If garantie = "Incapacité" Then
      PT = m_ModuleCalcul_Provisions.FctCalcProvIncap(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
        
        ElseIf garantie = "Passage" Then
                       
        PT = m_ModuleCalcul_Provisions.FctCalcProvPassage(x, xdepartRetraite, anc, Duree, franchise, feuilleLueIncap, feuilleLue, feuilleLueInval, iIncap, fraisGestionIncap, iInval, fraisGestionInval)
      
          ElseIf garantie = "Invalidité" Then
          PT = m_ModuleCalcul_Provisions.FctCalcProvInval(x, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval)
           
            Else
            
            If Not autoMode Then
              MsgBox "Table " & feuilleMaj & " NON CREEE pour la garantie " & garantie
            Else
              m_Logger.EcritTraceDansLog "Table " & feuilleMaj & " NON CREEE pour la garantie " & garantie
            End If
            Exit Sub
      
      End If
      
      If PT = -1 Then
        PT = 0
        If Not autoMode Then
          Screen.MousePointer = vbDefault
        End If
        rs.Close
        Exit Sub
      End If
      
      ' controle valeur de PT
      If PT < 0 Then
      Msg = "la provision calculée PT est négative = " & PT & " sa valeur est forcée à 0 " _
      & vbLf & " por x= " & x & " et anc= " & anc
      
      If Not autoMode Then
        MsgBox Msg, , "Code erreur dans CalcTableauProv "
      Else
        m_Logger.EcritTraceDansLog "Erreur dans CalcTableauProv " & Err & " : " & Err.Description
      End If
      
      PT = 0
      End If
      
      
      On Error Resume Next  ' si le champs n'existae pas dans la table
      rs.fields("Anc" & anc) = PT
      On Error GoTo 0 ' reset detection erreur
    
    Next anc
      
    ' valide la ligne
    rs.Update
  Next x
  
  rs.Close
    
  
  ' ferme les tables de reference
  m_theRecordsetList.CloseLoadedRecordset
  
  
  If Not ProgressBar1 Is Nothing Then
    ProgressBar1.Visible = False
  End If
  
  If Not autoMode Then
    Screen.MousePointer = vbDefault
  End If


MsgError:
  If Err.Number <> 0 Then
    Msg = "Erreur " & str(Err.Number) & " générée par " _
        & Err.Source & Chr(13) & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, , "Code erreur dans CalcTableauProv ", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans CalcTableauProv " & Err & " : " & Err.Description
    End If
  Else
  End If
  
  Exit Sub
   
ErrCalcTableauProv:
  On Error GoTo 0
  
  
  If Not autoMode Then
    Screen.MousePointer = vbDefault
    MsgBox "La table " & feuilleLue & " n'existe pas !" & vbLf & "Calcul annulé", vbCritical + vbOKOnly
  Else
    m_Logger.EcritTraceDansLog "La table " & feuilleLue & " n'existe pas !" & vbLf & "Calcul annulé"
  End If
  
  
End Sub

'                          dureeAn = k en années
'##ModelId=5C8A75E20117
Public Function FluxAnnuel(DureeAn, DureRevalo, TauxRevalo) As Double
 
  Select Case DureeAn
  Case Is <= DureRevalo
    FluxAnnuel = (1 + TauxRevalo) ^ (DureeAn)
  Case Is > DureRevalo
    FluxAnnuel = (1 + TauxRevalo) ^ DureRevalo
  End Select
End Function

'                           duree = k en mois
'##ModelId=5C8A75E20165
Public Function FluxMensuel(Duree, DureRevalo, TauxRevalo) As Double
  Select Case Duree
    Case Is <= DureRevalo * 12
      FluxMensuel = (1 + TauxRevalo) ^ Int((Duree + 6) / 12)
    Case Is > DureRevalo * 12
      FluxMensuel = (1 + TauxRevalo) ^ DureRevalo
  End Select
End Function

