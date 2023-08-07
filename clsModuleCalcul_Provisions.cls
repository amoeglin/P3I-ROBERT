VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCalcul_Provisions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75DC02E9"
Option Explicit

' acces aux tables
'##ModelId=5C8A75DD000B
Private m_theRecordsetList As clsRecordsetList

'##ModelId=5C8A75DD000C
Public autoMode As Boolean
'##ModelId=5C8A75DD002C
Public m_Logger As clsLogger
'

'##ModelId=5C8A75DD002D
Public Property Set recordsetList(r As clsRecordsetList)
  Set m_theRecordsetList = r
End Property

'*******************************
'  CALCUL PROVISION INCAPACITE *  FONCTION FctCalcProvIncap calcule UNE SEULE PROVISION pour les paramères passé
'*******************************  'FctCalcProvIncap(x, anc, duree, franchise, feuilleLue, i, fraisGestion)
                                
'##ModelId=5C8A75DD0059
Public Function FctCalcProvIncap(ByVal x As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, i As Double, fraisGestion As Double) As Double
  Dim v As Double, m As Double, PT As Double, Lanc As Double, lk As Double, lk1 As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim wfinIteration As Boolean
  wfinIteration = False
 

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvIncap = 0
 
 ' sélection de la ligne pour l'age demandé dans la feuillelue
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
  'rq = "SELECT * FROM " & feuilleLue & " WHERE Age = " & x
  'Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLue)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvIncap = 0
'    rs.Close
    Exit Function
  End If
   
   
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))
  If Lanc = 0 Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvIncap = 0
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

    PT = PT + (1 / (2 * Lanc)) * (lk * (v ^ ((k - anc) / m)) + lk1 * (v ^ ((k + 1 - anc) / m)))
    
  Next k

FIN_FctCALCUL_PT:
'  rs.Close

  FctCalcProvIncap = Arrondi(PT * (1 + fraisGestion), NbDecimalePM)  ' avec frais de gestion
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
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
       
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvIncap", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur in FctCalcProvIncap " & Err & " : " & Err.Description
    End If
    
    FctCalcProvIncap = 0
'    rs.Close
    Exit Function
    
    Resume Next
  End If
  
  
End Function

'*******************************
'  CALCUL PROVISION INVALIDITE *  FONCTION FctCalcProvInval calcule UNE SEULE PROVISION pour les paramères passé
'*******************************  'FctCalcProvInval(x, anc, duree, franchise, feuilleLue, i, fraisGestion)
                                
'##ModelId=5C8A75DD00E6
Public Function FctCalcProvInval(ByVal x As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, i As Double, fraisGestion As Double) As Double
  Dim v As Double, m As Double, PT As Double, Lanc As Double, lk As Double, lk1 As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset
  Dim wfinIteration As Boolean
  wfinIteration = False
 

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvInval = 0
 
 ' sélection de la ligne pour l'age demandé dans la feuillelue
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
  'rq = "SELECT * FROM " & feuilleLue & " WHERE Age = " & x
  'Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLue)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvInval = 0
'    rs.Close
    Exit Function
  End If
   
   
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))
  If Lanc = 0 Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvInval = 0
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
  
  If Duree < 0 Then
    Duree = Duree
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
    
    PT = PT + (1 / (2 * Lanc)) * (lk * (v ^ ((k - anc) / m)) + lk1 * (v ^ ((k + 1 - anc) / m)))
  
  Next k

FIN_FctCALCUL_PT:
'  rs.Close

  FctCalcProvInval = Arrondi(PT * (1 + fraisGestion), NbDecimalePM)  ' avec frais de gestion
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
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
      
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvInval", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur in FctCalcProvInval " & Err & " : " & Err.Description
    End If
        
    FctCalcProvInval = 0
'    rs.Close
    Exit Function

    
    Resume Next
  End If
  
  
End Function

'**********************************************
'  CALCUL PROVISION   PASSAGE    Incap/Inval  *  FONCTION FctCalcProvPassage calcule UNE SEULE PROVISION pour les paramères passé
'**********************************************  'FctCalcProvPassage
                                
'##ModelId=5C8A75DD0182
Public Function FctCalcProvPassage(ByVal x As Integer, xdepartRetraite As Integer, anc As Integer, Duree As Integer, franchise As Integer, feuilleLueIncap As String, feuilleLuePassage As String, feuilleLueInval As String, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, Optional bSansLissage As Boolean = False) As Double
                                   
  Dim v As Double, m As Double, PT As Double, Lanc As Double, Lduree As Double, Sk As Double, Sk1 As Double
  Dim k As Integer, kdeb As Integer
  Dim rq As String
  Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
  Dim dureeInval As Integer
  Dim PTInvalxDeb As Double, PTInvalx As Double, PTInvalx1 As Double, PTInval As Double, PTInval1 As Double
  Dim wfinIteration As Boolean
  
  wfinIteration = False
  
  
  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
 
  ' sélection de la ligne pour l'age demandé dans la feuilleIncap
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
  'If x = 37 Then ' 1ere ligne de la table = 20ans ou moins
  '  Stop
  'End If
  
  
  'rq = "SELECT * FROM " & feuilleLueIncap & " WHERE Age = " & x
  'Set rs = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLueIncap)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvPassage = 0
'    rs.Close
    Exit Function
  End If
   
   
  
  On Error GoTo LxAbsent
  Lanc = val(rs.fields("Anc" & anc))            ' effectifs pour anc mois en incap
  
  If Lanc = 0 Then
    FctCalcProvPassage = 0
'    rs.Close
    Exit Function
  End If
    
 
 ' sélection de la ligne pour l'age demandé dans la feuillePassage
  
  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
  
'  rq = "SELECT * FROM " & feuilleLuePassage & " WHERE Age = " & x
'  Set rs2 = m_dataSource.OpenRecordset(rq, Disconnected)
  
  Set rs2 = m_theRecordsetList.LoadRecordset(feuilleLuePassage)
  m_dataHelper.Multi_Find rs2, "Age=" & x
  
  If rs2.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcProvPassage = 0
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
    End If
    
    ' calcul de PTInvalx à l'âge entier
    If (xEntierPrec <> xEntier) And (xEntier <> 0) Then
      PTInvalx = FctCalcProvInval(xEntier, 0, dureeInval, 0, feuilleLueInval, iInval, fraisGestionInval)
    Else
      ' PTInvalx sans changement
    End If
    xEntierPrec = xEntier
    
          
    ' calcul de PTInvalx1 à l'âge entier suivant (xEntier + 1)
    If (xEntier1Prec <> xEntier + 1) Then
      PTInvalx1 = FctCalcProvInval(xEntier + 1, 0, dureeInval - 1, 0, feuilleLueInval, iInval, fraisGestionInval)
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
    
    PT = PT + (1 / Lanc) * (Sk * (v ^ ((k + 1 - anc) / m))) * PTInval
    
    
    If wfinIteration = True Then
      Exit For
    End If
  
  Next k
        
  
  'ajout du correctif
  On Error GoTo LdureeAbsent
  Lduree = val(rs.fields("Anc" & k + 1))     ' effectifs pour duree mois en incap = k+1
  
  On Error GoTo MsgError
  
  ' PHM 16/04/2007 : possibilité de désactiver le terme correcteur
  If bSansLissage = False Then
    PT = PT + (Lduree / Lanc) * (v ^ ((k + 1 - anc) / m)) * PTInval
  End If
        
       

FIN_FctCALCUL_PT:
  
  
  ' ajustement de PT qui doit être inférieur à la valeur initiale PTInvalxDdEB
  If PT > PTInvalxDeb Then
    PT = PTInvalxDeb
  Else
  End If
  
  FctCalcProvPassage = Arrondi(PT * (1 + fraisGestionIncap), NbDecimalePM)
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
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
  
    If Not autoMode Then
      MsgBox Msg, , "Code erreur fonction FctCalcProvPassage", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur in FctCalcProvPassage " & Err & " : " & Err.Description
    End If
    
    FctCalcProvPassage = 0
'    rs.Close
'    rs2.Close
    Exit Function
    
  End If
  
  
End Function


'*****************************************************************
'  CALCUL PROVISION D'une table COMPLETE  POUR UNE GARANTIE Incap ou Passage ou Inval   à partir de feuille lue
'*****************************************************************
'Public Sub CalcTableauProv(x, anciennete, duree, franchise, feuilleLue, iIncap, fraisGestionIncap, iInval, fraisGestionInval, xdepartRetraite, garantie, ProgressBar1)


'##ModelId=5C8A75DD02AD
Public Sub CalcTableauProv(ByVal x As Integer, Anciennete As Integer, Duree As Integer, franchise As Integer, feuilleLue As String, iIncap As Double, fraisGestionIncap As Double, iInval As Double, fraisGestionInval As Double, xdepartRetraite As Integer, garantie As String, ProgressBar1 As ProgressBar, feuilleLueIncap As String, feuilleLueInval As String)
 
  
  Dim PT As Double
  Dim xdeb As Integer, xfin As Integer, k As Integer, anc As Integer
  Dim feuilleMaj As String, rq As String
  Dim rs As ADODB.Recordset
  
  x = 0
  
  Dim Msg As String ' libellé message erreur
  
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
  
  'Dim oCat As New ADOX.Catalog
'  Dim newTable As ADOX.Table, tableOrg As ADOX.Table
'  Dim newfield As ADOX.Column, oKey As ADOX.key
'  Dim oProp As ADOX.Property
  
  'Set oCat.ActiveConnection = m_dataSource.Connection
  
  ' crée la table de destination (feuillemaj)
  On Error Resume Next  ' si la table n'existe pas pas erreur
  'oCat.Tables.Delete feuilleMaj
  m_dataSource.Execute "DROP TABLE " & feuilleMaj
  On Error GoTo 0 ' reset detection erreur
  
On Error GoTo ErrCalcTableauProv
  m_dataSource.Execute "SELECT * INTO " & feuilleMaj & " FROM " & feuilleLue & " WHERE Age=-999"
  
  'm_dataSource.Execute "CREATE INDEX " & feuilleMaj & "_Age ON " & feuilleMaj & "(Age) with PRIMARY"
  m_dataSource.Execute "ALTER TABLE " & feuilleMaj & " ADD CONSTRAINT PK_" & feuilleMaj & " PRIMARY KEY CLUSTERED (Age)"
  'Set tableOrg = oCat.Tables(feuilleLue)

On Error GoTo 0
  
  'Set newTable = New ADOX.Table
  
  'newTable.Name = feuilleMaj
  
'  'For x = 0 To tableOrg.Columns.Count - 1
'    Set newfield = New ADOX.Column
'
'    newfield.Name = tableOrg.Columns(x).Name
'    newfield.Type = tableOrg.Columns(x).Type
'
'    Set newfield.ParentCatalog = oCat
'
'    For Each oProp In tableOrg.Columns(x).Properties
'      newfield.Properties(oProp.Name) = oProp.Value
'    Next
'
'    newTable.Columns.Append newfield
'  Next x
'
'  On Error GoTo MsgError
'  oCat.Tables.Append newTable
'
'  ' Create a new Primary Key for the table
'  Set oKey = New ADOX.key
'  With oKey
'      .Name = feuilleMaj & "_Age"
'      .Type = adKeyPrimary
'      .RelatedTable = feuilleMaj
'      .Columns.Append "Age"
'  End With
'  newTable.Keys.Append oKey
  
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
  If Trim(feuilleLueIncap) <> "" Then
    m_theRecordsetList.LoadRecordset feuilleLueIncap
  End If
  If Trim(feuilleLueInval) <> "" Then
    m_theRecordsetList.LoadRecordset feuilleLueInval
  End If
    
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
  Set rs = m_dataSource.OpenRecordset(rq, Dynamic)
  
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
        PT = FctCalcProvIncap(x, anc, Duree, franchise, feuilleLue, iIncap, fraisGestionIncap)
        
      ElseIf garantie = "Passage" Then
                       
        PT = FctCalcProvPassage(x, xdepartRetraite, anc, Duree, franchise, feuilleLueIncap, feuilleLue, feuilleLueInval, iIncap, fraisGestionIncap, iInval, fraisGestionInval)
      
      ElseIf garantie = "Invalidité" Then
        PT = FctCalcProvInval(x, anc, Duree, franchise, feuilleLue, iInval, fraisGestionInval)
           
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
          Screen.MousePointer = vbDefault  ' interruption du calcul
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
          m_Logger.EcritTraceDansLog "Erreur in CalcTableauProv " & Err & " : " & Err.Description
        End If
        
        PT = 0
      End If
      
      
      On Error Resume Next  ' si le champs n'existant pas dans la table
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
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
    
    
    If Not autoMode Then
      MsgBox Msg, , "Code erreur dans CalcTableauProv ", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur in CalcTableauProv " & Err & " : " & Err.Description
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul de ax
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' anNaissance : annee de naissance
'
'##ModelId=5C8A75DD03C2
Public Function FctCalc_ax(sTable As String, i As Double, x As Integer, ByVal anNaissance As Integer, fWait As Form) As Double
  Dim k As Integer
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim lx As Double, lxk As Double, v As Double
  
  On Error GoTo err_FctCalc_ax
  
  ' init
  bGeneration = False
  
  ' cherche la ligne contenant les données
  Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
  rq = "NomTable='" & sTable & "'"
  m_dataHelper.Multi_Find rs, rq
  
  If rs.fields("TYPETABLE") = cdTypeTableGeneration Then
    bGeneration = True
    rq = rq & " AND Naissance=" & anNaissance
    m_dataHelper.Multi_Find rs, rq
  End If
  
  If rs.EOF Then
    ' recherche du millesime Max
    anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
    
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
    
      ' erreur dans la table
      
      
      If Not autoMode Then
        fWait.Hide
        Screen.MousePointer = vbDefault
        MsgBox "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "ax est forcé à 0 !", vbCritical
        fWait.Show
        Screen.MousePointer = vbArrow
      Else
        m_Logger.EcritTraceDansLog "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "ax est forcé à 0 !"
      End If
      
  '    rs.Close
      FctCalc_ax = 0
      Exit Function
    End If
  
  End If
  
  ' recupere lx
  If IsNull(rs.fields("l" & x)) Then
    ' erreur dans la table
    
    
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x & " inconnu dans la table " & sTable & vbLf & "ax est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x & " inconnu dans la table " & sTable & vbLf & "ax est forcé à 0 !"
    End If
    
'    rs.Close
    FctCalc_ax = 0
    Exit Function
  Else
    lx = rs.fields("l" & x)
  End If
  
  ' calcul de v
  v = 1# / (1# + i)
  
  '
  ' calcul de la somme de : l(x+k)/lx * v^k
  '
  k = 1
  FctCalc_ax = 0
  Do
    ' recupere l(x+k)
    If IsNull(rs.fields("l" & x + k)) Then
      ' fin de boucle
'      rs.Close
      Exit Function
    Else
      lxk = rs.fields("l" & x + k)
    End If
    
    FctCalc_ax = FctCalc_ax + ((lxk / lx) * (v ^ k))
    
    k = k + 1
  Loop
  
'  rs.Close
  
  Exit Function
  
err_FctCalc_ax:
  ' erreur dans la table
  If Err = 3265 Then
    ' fin de boucle
'    rs.Close
    Exit Function
  End If
  
  If Not autoMode Then
    fWait.Hide
    Screen.MousePointer = vbDefault
    MsgBox "FctCalc_ax() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "ax est forcé à 0 !", vbCritical
    fWait.Show
    Screen.MousePointer = vbArrow
  Else
    m_Logger.EcritTraceDansLog "FctCalc_ax() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "ax est forcé à 0 !"
  End If
  
'  rs.Close
  FctCalc_ax = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul de äx
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' anNaissance : annee de naissance
'
'##ModelId=5C8A75DE0047
Public Function FctCalc_äx(sTable As String, i As Double, x As Integer, ByVal anNaissance As Integer, fWait As Form) As Double
  Dim k As Integer
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim lx As Double, lxk As Double, v As Double
  
  On Error GoTo err_FctCalc_äx
  
  ' init
  bGeneration = False
  
  ' cherche la ligne contenant les données
  Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
  rq = "NomTable='" & sTable & "'"
  m_dataHelper.Multi_Find rs, rq
  
  If rs.fields("TYPETABLE") = cdTypeTableGeneration Then
    bGeneration = True
    rq = rq & " AND Naissance=" & anNaissance
    m_dataHelper.Multi_Find rs, rq
  End If
  
  If rs.EOF Then
    ' recherche du millesime Max
    anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
    
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
    
      ' erreur dans la table
      
      
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "äx est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "äx est forcé à 0 !"
    End If
      
  '    rs.Close
      FctCalc_äx = 0
      Exit Function
    End If
  End If
  
  ' recupere lx
  If IsNull(rs.fields("l" & x)) Then
    ' erreur dans la table
    
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x & " inconnu dans la table " & sTable & vbLf & "äx est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x & " inconnu dans la table " & sTable & vbLf & "äx est forcé à 0 !"
    End If
    
'    rs.Close
    FctCalc_äx = 0
    Exit Function
  Else
    lx = rs.fields("l" & x)
  End If
  
  ' calcul de v
  v = 1# / (1# + i)
  
  '
  ' calcul de la somme de : l(x+k)/lx * v^k
  '
  k = 0
  FctCalc_äx = 0
  Do
    ' recupere l(x+k)
    If IsNull(rs.fields("l" & x + k)) Then
      ' fin de boucle
'      rs.Close
      Exit Function
    Else
      lxk = rs.fields("l" & x + k)
    End If
    
    FctCalc_äx = FctCalc_äx + ((lxk / lx) * (v ^ k))
    
    k = k + 1
  Loop
  
'  rs.Close
  
  Exit Function
  
err_FctCalc_äx:
  ' fin de table
  If Err = 3265 Then
'    rs.Close
    Exit Function
  End If
  
  ' erreur dans la table
  
  If Not autoMode Then
    fWait.Hide
    Screen.MousePointer = vbDefault
    MsgBox "FctCalc_äx() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "ax est forcé à 0 !", vbCritical
    fWait.Show
    Screen.MousePointer = vbArrow
  Else
    m_Logger.EcritTraceDansLog "FctCalc_äx() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "ax est forcé à 0 !"
  End If
  
'  rs.Close
  FctCalc_äx = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul de axn
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' duree       : duree de calcul
' anNaissance : annee de naissance
'
'##ModelId=5C8A75DE00B4
Public Function FctCalc_axn(sTable As String, i As Double, x As Integer, Duree As Integer, ByVal anNaissance As Integer, fWait As Form) As Double
  Dim k As Integer
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim lx As Double, lxk As Double, v As Double
  
  On Error GoTo err_FctCalc_axn
  
  If Duree = 0 Then
    FctCalc_axn = 0
  End If
  
  ' init
  bGeneration = False
  
  ' cherche la ligne contenant les données
  Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
  rq = "NomTable='" & sTable & "'"
  m_dataHelper.Multi_Find rs, rq
  
  If rs.fields("TYPETABLE") = cdTypeTableGeneration Then
    bGeneration = True
    rq = rq & " AND Naissance=" & anNaissance
    m_dataHelper.Multi_Find rs, rq
  End If
  
  If rs.EOF Then
    ' recherche du millesime Max
    anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
    
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
    
      ' erreur dans la table
      
      
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "axn est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "axn est forcé à 0 !"
    End If
      
  '    rs.Close
      FctCalc_axn = 0
      Exit Function
    End If
  End If
  
  ' recupere lx
  If IsNull(rs.fields("l" & x)) Then
    ' erreur dans la table
    
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x & " inconnu dans la table " & sTable & vbLf & "axn est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x & " inconnu dans la table " & sTable & vbLf & "axn est forcé à 0 !"
    End If
    
'    rs.Close
    FctCalc_axn = 0
    Exit Function
  Else
    lx = rs.fields("l" & x)
  End If
  
  ' calcul de v
  v = 1# / (1# + i)
  
  '
  ' calcul de la somme de : l(x+k)/lx * v^k
  '
  FctCalc_axn = 0
  For k = 1 To Duree
    ' recupere l(x+k)
    If IsNull(rs.fields("l" & x + k)) Then
      ' fin de boucle
      Exit For
    Else
      lxk = rs.fields("l" & x + k)
    End If
    
    FctCalc_axn = FctCalc_axn + ((lxk / lx) * (v ^ k))
  Next
  
'  rs.Close
  
  Exit Function
  
err_FctCalc_axn:
  ' fin de table
  If Err = 3265 Then
'    rs.Close
    Exit Function
  End If
  
  ' erreur dans la table
  
  If Not autoMode Then
    fWait.Hide
    Screen.MousePointer = vbDefault
    MsgBox "FctCalc_axn() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "axn est forcé à 0 !", vbCritical
    fWait.Show
    Screen.MousePointer = vbArrow
  Else
    m_Logger.EcritTraceDansLog "FctCalc_axn() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "axn est forcé à 0 !"
  End If
  
'  rs.Close
  FctCalc_axn = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul de äxn
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' duree       : duree de calcul
' anNaissance : annee de naissance
'
'##ModelId=5C8A75DE0151
Public Function FctCalc_äxn(sTable As String, i As Double, x As Integer, Duree As Integer, ByVal anNaissance As Integer, fWait As Form) As Double
  Dim k As Integer
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim lx As Double, lxk As Double, v As Double
  
  On Error GoTo err_FctCalc_äxn
  
  If Duree = 0 Then
    FctCalc_äxn = 0
  End If
  
  ' init
  bGeneration = False
  
  ' cherche la ligne contenant les données
  Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
  rq = "NomTable='" & sTable & "'"
  m_dataHelper.Multi_Find rs, rq
  
  If rs.fields("TYPETABLE") = cdTypeTableGeneration Then
    bGeneration = True
    rq = rq & " AND Naissance=" & anNaissance
    m_dataHelper.Multi_Find rs, rq
  End If
  
  If rs.EOF Then
    ' recherche du millesime Max
    anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
    
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
    
      ' erreur dans la table
      If Not autoMode Then
        fWait.Hide
        Screen.MousePointer = vbDefault
        MsgBox "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "äxn est forcé à 0 !", vbCritical
        fWait.Show
        Screen.MousePointer = vbArrow
      Else
        m_Logger.EcritTraceDansLog "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "äxn est forcé à 0 !"
      End If
      
  '    rs.Close
      FctCalc_äxn = 0
      Exit Function
    End If
  End If
  
  ' recupere lx
  If IsNull(rs.fields("l" & x)) Then
    ' erreur dans la table
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x & " inconnu dans la table " & sTable & vbLf & "äxn est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x & " inconnu dans la table " & sTable & vbLf & "äxn est forcé à 0 !"
    End If
    
    
'    rs.Close
    FctCalc_äxn = 0
    Exit Function
  Else
    lx = rs.fields("l" & x)
  End If
  
  ' calcul de v
  v = 1# / (1# + i)
  
  '
  ' calcul de la somme de : l(x+k)/lx * v^k
  '
  FctCalc_äxn = 0
  For k = 0 To Duree - 1
    ' recupere l(x+k)
    If IsNull(rs.fields("l" & x + k)) Then
      ' fin de boucle
      Exit For
    Else
      lxk = rs.fields("l" & x + k)
    End If
    
    FctCalc_äxn = FctCalc_äxn + ((lxk / lx) * (v ^ k))
  Next
  
'  rs.Close
  
  Exit Function
  
err_FctCalc_äxn:
  ' fin de table
  If Err = 3265 Then
'    rs.Close
    Exit Function
  End If
  
  ' erreur dans la table
  If Not autoMode Then
    fWait.Hide
    Screen.MousePointer = vbDefault
    MsgBox "FctCalc_äxn() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "äxn est forcé à 0 !", vbCritical
    fWait.Show
    Screen.MousePointer = vbArrow
  Else
    m_Logger.EcritTraceDansLog "FctCalc_äxn() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "äxn est forcé à 0 !"
  End If
  
'  rs.Close
  FctCalc_äxn = 0
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul de nEx
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' duree       : duree de calcul
' anNaissance : annee de naissance
'
'##ModelId=5C8A75DE01DD
Public Function FctCalc_nEx(sTable As String, i As Double, x As Integer, Duree As Integer, ByVal anNaissance As Integer, fWait As Form) As Double
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim lx As Double, lxn As Double, v As Double
  
  On Error GoTo err_FctCalc_nEx
  
  If Duree = 0 Then
    FctCalc_nEx = 0
  End If
  
  ' init
  bGeneration = False
  
  ' cherche la ligne contenant les données
  Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
  rq = "NomTable='" & sTable & "'"
  m_dataHelper.Multi_Find rs, rq
  
  If rs.fields("TYPETABLE") = cdTypeTableGeneration Then
    bGeneration = True
    rq = rq & " AND Naissance=" & anNaissance
    m_dataHelper.Multi_Find rs, rq
  End If
  
  If rs.EOF Then
    ' recherche du millesime Max
    anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
    
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
    
      ' erreur dans la table
      If Not autoMode Then
        fWait.Hide
        Screen.MousePointer = vbDefault
        MsgBox "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "nEx est forcé à 0 !", vbCritical
        fWait.Show
        Screen.MousePointer = vbArrow
      Else
        m_Logger.EcritTraceDansLog "Le millesime " & anNaissance & " n'existe pas dans la table " & sTable & vbLf & "nEx est forcé à 0 !"
      End If
      
  '    rs.Close
      FctCalc_nEx = 0
      Exit Function
    End If
  End If
  
  ' recupere lx
  If IsNull(rs.fields("l" & x)) Then
    ' erreur dans la table
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x & " inconnu dans la table " & sTable & vbLf & "nEx est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x & " inconnu dans la table " & sTable & vbLf & "nEx est forcé à 0 !"
    End If
    
'    rs.Close
    FctCalc_nEx = 0
    Exit Function
  Else
    lx = rs.fields("l" & x)
  End If
  
  ' recupere l(x+n)
  If IsNull(rs.fields("l" & x + Duree)) Then
    ' erreur dans la table
    If Not autoMode Then
      fWait.Hide
      Screen.MousePointer = vbDefault
      MsgBox "l" & x + Duree & " inconnu dans la table " & sTable & vbLf & "nEx est forcé à 0 !", vbCritical
      fWait.Show
      Screen.MousePointer = vbArrow
    Else
      m_Logger.EcritTraceDansLog "l" & x + Duree & " inconnu dans la table " & sTable & vbLf & "nEx est forcé à 0 !"
    End If
    
'    rs.Close
    FctCalc_nEx = 0
    Exit Function
  Else
    lxn = rs.fields("l" & x + Duree)
  End If
  
  ' calcul de v
  v = 1# / (1# + i)
  
  '
  ' calcul de la somme de : l(x+duree)/lx * v^duree
  '
  FctCalc_nEx = ((lxn / lx) * (v ^ Duree))
  
'  rs.Close
  
  Exit Function
  
err_FctCalc_nEx:
  ' fin de table
  If Err = 3265 Then
    FctCalc_nEx = 0
'    rs.Close
    Exit Function
  End If
  
  ' erreur dans la table
  If Not autoMode Then
    fWait.Hide
    Screen.MousePointer = vbDefault
    MsgBox "FctCalc_nEx() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "nEx est forcé à 0 !", vbCritical
    fWait.Show
    Screen.MousePointer = vbArrow
  Else
    m_Logger.EcritTraceDansLog "FctCalc_nEx() : erreur " & Err & vbLf & Err.Description & vbLf & vbLf & "nEx est forcé à 0 !"
  End If
  
'  rs.Close
  FctCalc_nEx = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul FctCalc_mafract_xn (am_xn_k nom BFA)
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' n           : duree de calcul
' k           : différé
' fract       : (m-1)/(2*m)
' anNaissance : annee de naissance
'

'##ModelId=5C8A75DE026D
Public Function FctCalc_mafract_xn(sTable As String, i As Double, x As Integer, k As Integer, fract As Double, n As Integer, anNaissance As Integer, fWait As Form) As Double
  
  Dim kEx As Double
  Dim nkEx As Double
  Dim axkn As Double
  Dim xDebut As Integer
  
  If k < 0 Then
    k = 0
  End If
  
  If n < 0 Then
    n = 0
  End If

  If n = 0 Then
    FctCalc_mafract_xn = 0
  End If
  
  xDebut = GetAgeMiniTableGeneration(sTable, x, anNaissance)
  
  kEx = FctCalc_nEx(sTable, i, xDebut, k, anNaissance, fWait)
  nkEx = FctCalc_nEx(sTable, i, xDebut, n + k, anNaissance, fWait)
  axkn = FctCalc_axn(sTable, i, xDebut + k, n, anNaissance, fWait)
  
  FctCalc_mafract_xn = kEx * axkn + fract * (kEx - nkEx)

End Function


'##ModelId=5C8A75DE0333
Public Function GetAgeMiniTableGeneration(sTable As String, x As Integer, anNaissance As Integer) As Integer
  Dim bGeneration As Boolean
  Dim rs As ADODB.Recordset
  Dim rq As String
  Dim i As Integer
  
  GetAgeMiniTableGeneration = x

  ' recherche de l'age mini pour le millesime en cours
  bGeneration = False
  If Left(sTable, 10) = "TABLEGENER" Or sTable = "TG05H" Or sTable = "TG05F" Then
    bGeneration = True
  End If
  
  If bGeneration Then
    
    Set rs = m_theRecordsetList.LoadRecordset("TableMortalite")
    rq = "NomTable='" & sTable & "'"
    If bGeneration Then
      rq = rq & " AND Naissance=" & anNaissance
    End If
    m_dataHelper.Multi_Find rs, rq
    
    If rs.EOF Then
      ' recherche du millesime Max
      anNaissance = m_dataHelper.GetParameterAsDouble("SELECT max(Naissance) From TableMortalite Where NomTable='" & sTable & "'")
      
      rq = "NomTable='" & sTable & "'"
      If bGeneration Then
        rq = rq & " AND Naissance=" & anNaissance
      End If
      m_dataHelper.Multi_Find rs, rq
      
      If rs.EOF Then
        ' erreur dans la table
        GetAgeMiniTableGeneration = x
        Exit Function
      End If
    End If
    
    ' recupere lx
    If IsNull(rs.fields("l" & x)) Then
      
      ' erreur dans la table
      For i = x + 1 To 120
        If Not IsNull(rs.fields("l" & i)) Then
          GetAgeMiniTableGeneration = i
          Exit Function
        End If
      Next
    
    Else
      
      GetAgeMiniTableGeneration = x
      Exit Function
    
    End If
  
  End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Fonction de calcul FctCalc_mäfrac_xn (adotm_xn_k nom BFA)
' sTable      : table de mortalite
' i           : taux Technique
' x           : age de l'assure
' n           : duree de calcul
' k           : différé
' fract       : (m-1)/(2*m)
' anNaissance : annee de naissance
'

'##ModelId=5C8A75DE0391
Public Function FctCalc_mäfract_xn(sTable As String, i As Double, x As Integer, k As Integer, fract As Double, n As Integer, anNaissance As Integer, fWait As Form) As Double
  
  Dim kEx As Double
  Dim nkEx As Double
  Dim äxkn As Double
  Dim xDebut As Integer

  If k < 0 Then
    k = 0
  End If
  
  If n < 0 Then
    n = 0
  End If
  
  If n = 0 Then
    FctCalc_mäfract_xn = 0
  End If
  
  xDebut = GetAgeMiniTableGeneration(sTable, x, anNaissance)

  kEx = FctCalc_nEx(sTable, i, xDebut, k, anNaissance, fWait)
  nkEx = FctCalc_nEx(sTable, i, xDebut, n + k, anNaissance, fWait)
  äxkn = FctCalc_äxn(sTable, i, xDebut + k, n, anNaissance, fWait)
  
  FctCalc_mäfract_xn = kEx * äxkn - fract * (kEx - nkEx)

End Function


'*******************************
'  CALCUL PROVISION MATERNITE, MENSU, Chomage *  FONCTION FctCalcProvMensualiteCertaine calcule UNE SEULE PROVISION pour les paramères passé
'*******************************  'FctCalcProvMensualiteCertaine(x, anc, duree, franchise, feuilleLue, i, fraisGestion)
                                
'##ModelId=5C8A75DF0054
Public Function FctCalcProvMensualiteCertaine(anc As Integer, Duree As Integer, franchise As Integer, i As Double, fraisGestion As Double) As Double
  Dim v As Double, m As Double, PT As Double
  Dim k As Integer, kdeb As Integer

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvMaternite = 0
 
  m = 12       ' fractionnement mensuel
  v = 1 / (1 + i)
  PT = 0
  
  
  If anc = Duree Then
    FctCalcProvMensualiteCertaine = 0
    Exit Function
  End If
  
    
  If anc > franchise Then
    kdeb = anc + 1
  Else
    kdeb = franchise + 1
  End If
  
  
  For k = kdeb To Duree               ' duree = 10, 12 ou 24 mois
    
    PT = PT + (v ^ ((k - anc) / m))
    
  Next k

  
  FctCalcProvMensualiteCertaine = Arrondi(PT * (1 + fraisGestion), NbDecimalePM)  ' avec frais de gestion
  
  
  Exit Function
  
MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, vbCritical, "Code erreur fonction FctCalcProvMaternite", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcProvMaternite " & Err & " : " & Err.Description
    End If
    
    FctCalcProvMensualiteCertaine = 0
    
    Exit Function
    
    Resume Next
  End If
  
End Function


'*******************************
'  CALCUL CORRECTIF INCAPACITE *  FONCTION FctCalcCorrectifIncapacite calcule le correctif pour les paramètres passés
'*******************************
                                
'##ModelId=5C8A75DF00D1
Public Function FctCalcCorrectifIncapacite(ByVal x As Integer, anc_calcul As Integer, anc_reglement As Integer, feuilleLueIncap As String, feuilleLuePassage As String) As Double
  Dim PT As Double, Lanc_calcul As Double, Lanc_regl As Double, Sk As Double
  Dim k As Integer
  Dim rs As ADODB.Recordset
  Dim rs2 As ADODB.Recordset
  Dim wfinIteration As Boolean
  wfinIteration = False

  On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
  
  ' sélection de la ligne pour l'age demandé dans la feuilleIncap
  If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
    x = 20
  End If
  
  Set rs = m_theRecordsetList.LoadRecordset(feuilleLueIncap)
  m_dataHelper.Multi_Find rs, "Age=" & x
  
  If rs.EOF Then
    If Not autoMode Then
      Screen.MousePointer = vbDefault
    End If
    FctCalcCorrectifIncapacite = 0
    Exit Function
  End If
   
  Dim pasDeSk As Boolean
   
  Set rs2 = m_theRecordsetList.LoadRecordset(feuilleLuePassage)
  m_dataHelper.Multi_Find rs2, "Age=" & x
  
  pasDeSk = False
  If rs2.EOF Then
    'Screen.MousePointer = vbDefault
    'FctCalcCorrectifIncapacite = 0
    'Exit Function
    pasDeSk = True
  End If
  
  On Error GoTo Lanc_calculAbsent
  Lanc_calcul = val(rs.fields("Anc" & anc_calcul))            ' effectifs pour anc mois en incap
  If Lanc_calcul = 0 Then
    FctCalcCorrectifIncapacite = 0
    Exit Function
  End If
  
  On Error GoTo Lx_reglAbsent
  Lanc_regl = val(rs.fields("Anc" & anc_reglement))            ' effectifs pour anc mois en incap
  If Lanc_regl = 0 Then
    FctCalcCorrectifIncapacite = 0
    Exit Function
  End If

  PT = 0
  If pasDeSk = False Then
    For k = anc_reglement + 1 To anc_calcul
      
      On Error GoTo SkAbsent
      Sk = val(rs2.fields("Anc" & k))         ' lue dans feuilleLuePassage
      
      PT = PT + Sk
      
      If wfinIteration = True Then
        Exit For
      End If
        
    Next
  End If
  
  PT = (PT + Lanc_calcul) / Lanc_regl
  If PT > 1# Then
    PT = 1#
  End If
  
  FctCalcCorrectifIncapacite = PT
  
  Exit Function

Lanc_calculAbsent:
  Lanc_calcul = 0
  Resume Next

Lx_reglAbsent:
  Lanc_regl = 0
  Resume Next

SkAbsent:
  Sk = 0
  If errCollectionAbsent = Err.Number Then
    wfinIteration = True ' l'élément Sk1 n'existe pas dans la ligne, on arrête l'itération
  End If
  Resume Next
  
MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, vbCritical, "Code erreur fonction FctCalcCorrectifIncapacite", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcCorrectifIncapacite " & Err & " : " & Err.Description
    End If
    
    FctCalcCorrectifIncapacite = 0
    
    Exit Function
    
    Resume Next
  End If
End Function
'******************************************
'  CALCUL des effectifs Lanc *  FONCTION FctCalcLanc recherche des effectifs dans une loi de maintien du type Incap ou dépendance
'******************************************
                                
'##ModelId=5C8A75DF014E
Public Function FctCalcLanc(ByVal x As Integer, anc_calcul As Integer, feuilleLueIncap As String, fraisGestion As Double) As Double
    
    Dim PT As Double, Lanc_calcul As Double, Lanc_regl As Double, Sk As Double
    Dim k As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo MsgError  ' sortie de la fonction avec FctCalcProvPassage = 0
    
    ' sélection de la ligne pour l'age demandé dans la feuilleIncap
    If x < 20 Then ' 1ere ligne de la table = 20ans ou moins
     x = 20
    End If
    
    Set rs = m_theRecordsetList.LoadRecordset(feuilleLueIncap)
    m_dataHelper.Multi_Find rs, "Age=" & x
    
    If rs.EOF Then
        If Not autoMode Then
          Screen.MousePointer = vbDefault
        End If
        FctCalcLanc = 0
        Exit Function
    End If

    On Error GoTo Lanc_calculAbsent
    FctCalcLanc = Arrondi(rs.fields("Anc" & anc_calcul) * (1 + fraisGestion), NbDecimalePM)  ' valeur lue "effectifs" pour anc mois en incap
    If FctCalcLanc = 0 Then
        FctCalcLanc = 0
        Exit Function
    End If
Exit Function
    
Lanc_calculAbsent:
    FctCalcLanc = 0
    Resume Next

MsgError:
  
  Dim Msg As String
  
  If Err.Number <> 0 Then
    Msg = "Erreur " & Err.Number & " générée par " & Err.Source & vbLf & Err.Description
    
    If Not autoMode Then
      MsgBox Msg, vbCritical, "Code erreur fonction FctCalcLanc ", Err.HelpFile, Err.HelpContext
    Else
      m_Logger.EcritTraceDansLog "Erreur dans FctCalcLanc " & Err & " : " & Err.Description
    End If
  
    FctCalcLanc = 0
    Exit Function
    Resume Next
  End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interpolation entre (entier(x), t1) et (entier(x+1), t2)
'
'##ModelId=5C8A75DF01BC
Public Function Interpolation(ByVal t1 As Double, ByVal t2 As Double, ByVal x As Double) As Double
  
  Interpolation = t1 + (x - Int(x)) * (t2 - t1)

End Function


