VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Fonctions d'aide pouir l'acc�s aux donn�es."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F851B7403B9"
Option Explicit

'
' VARIABLES LOCALES
'

Public Enum eSQLMode
  SQLMode_NotSet = 0
  MySQL = 1
  SQLServer
  Oracle
  MSAccess
  MaxDB
End Enum


'##ModelId=3F851B7403DA
Private m_DataSource As DataAccess 'local copy
Private m_SQLMode As eSQLMode '

'
' PROPERTIES
'

'##ModelId=3F851B75002E
Public Property Set DataSource(ByVal vData As DataAccess)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DataSource = Form1
    Set m_DataSource = vData
End Property

Public Property Get DataSource() As DataAccess
    Set DataSource = m_DataSource
End Property


Public Property Let SqlMode(ByVal vData As eSQLMode)
    m_SQLMode = vData
End Property

Public Property Get SqlMode() As eSQLMode
    SqlMode = m_SQLMode
End Property


'
' METHODES
'

'--------------------------------------------------------------------
'--------------------------------------------------------------------
'##ModelId=3F851B75006D
Public Sub Affiche_Liste(oObjList As Object, ChaineSQL As String, MyDatabase As String)

    'ObjLit     : Nom de la listeview
    'Chaine SQL : Requete SQL du type Select * from... ou Select toto, titi .. from ...
    'MyDatabase : Nom de la base de donn�es c:\base de donn�es\toto.mdb
    
    ' Creation d'une liste avec determination des colonnes au prealable
    ' IL EST IMPERATIF DE CREE LES COLONNES AU PREALABLE

    ' Attention : Il faut que le nombre de colonne de la listeview soit inferieur
    '           ou egale au nombre de champs define dans la requete

    Dim MyReq As ADODB.Recordset
    Dim itmX As ListItem
    Dim NbCol As Long
    Dim compteur As Long
    Dim theList As ListView
    
    Set theList = oObjList
    
    Set MyReq = m_DataSource.OpenRecordset(ChaineSQL, Snapshot)
    
    theList.ListItems.Clear
    NbCol = theList.ColumnHeaders.Count  'Nbcol est egale au nombre de colonne
    
    While Not MyReq.EOF
    
        Set itmX = theList.ListItems.Add(, , MyReq.Fields(0))
    
        For compteur = 1 To NbCol - 1 Step 1
          If Not IsNull(MyReq.Fields(compteur).Value) Then
               itmX.SubItems(compteur) = CStr(MyReq.Fields(compteur).Value)
          End If
        Next compteur
        
        MyReq.MoveNext   ' Passe � l'enregistrement suivant.
    Wend
    MyReq.Close
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Creation d'une liste sans determination des colonnes au prealable
' le nom des colonnes correspond au nom des champs de la chaine SQL
'
'##ModelId=3F851B7500BB
Public Sub Affiche_Liste_Table(oObjList As Object, ChaineSQL As String, fErase As Boolean)

    'ObjLit     : Nom de la listeview
    'Chaine SQL : Requete SQL du type Select * from... ou Select toto, titi .. from ...
    
    ' Creation d'une liste sans determination des colonnes au prealable
    ' le nom des colonnes correspond au nom des champs de la chaine SQL
    
    Dim MyReq As ADODB.Recordset
    Dim itmX As ListItem
    Dim NbCol As Long
    Dim compteur As Long
    Dim clmX As ColumnHeader
    Dim i As Long
    Dim j As Long
    Dim ColWidth() As Integer
    Dim theList As ListView
    
    Set theList = oObjList
    
    Set MyReq = m_DataSource.OpenRecordset(ChaineSQL, Snapshot)
    
    If fErase = True Or theList.ColumnHeaders.Count = 0 Then
      theList.View = lvwReport    'Determination de l'affichage sous forme de liste
      theList.ListItems.Clear     'Suppression des elements de la liste
      theList.ColumnHeaders.Clear 'Suppression des colonnes
    End If
    
    If theList.ColumnHeaders.Count = 0 Then
      For i = 0 To MyReq.Fields.Count - 1
        'Creation des colonnes
        Set clmX = theList.ColumnHeaders.Add(, , MyReq.Fields(i).Name, theList.Width / MyReq.Fields.Count)
        If MyReq.Fields(i).Type = adDate _
           Or MyReq.Fields(i).Type = adDBDate _
           Or MyReq.Fields(i).Type = adDBTime _
           Or MyReq.Fields(i).Type = adDBTimeStamp Then
          clmX.Tag = "Date"
        Else
          clmX.Tag = " "
        End If
      Next i
    End If

    NbCol = MyReq.Fields.Count
    
    While Not MyReq.EOF
        'Ajout des champs
        Set itmX = theList.ListItems.Add(, , CStr(IIf(IsNull(MyReq.Fields(0).Value), "", MyReq.Fields(0).Value))) 'Champ principal
        For compteur = 1 To NbCol - 1 Step 1
          If Not IsNull(MyReq.Fields(compteur)) Then
            itmX.SubItems(compteur) = CStr(MyReq.Fields(compteur).Value) ' Champs secondaire
          Else
            itmX.SubItems(compteur) = " "
          End If
        Next compteur
        
        If NbCol <> theList.ColumnHeaders.Count Then
          For compteur = NbCol To theList.ColumnHeaders.Count - 1 Step 1
            itmX.SubItems(compteur) = " "
          Next compteur
        End If
        
        MyReq.MoveNext   ' Passe � l'enregistrement suivant.
    Wend
    MyReq.Close
    
    ' calcule la largeur des colonnes
    'Call LargeurAutomatique(ObjWnd, ObjList)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Creation d'une liste sans determination des colonnes au prealable
' le nom des colonnes correspond au nom des champs de la chaine SQL
'
'##ModelId=3F851B750119
Public Sub Affiche_Liste_Choix(oObjList As Object, ChaineSQL As String)

    'ObjLit     : Nom de la listeview
    'Chaine SQL : Requete SQL du type Select * from... ou Select toto, titi .. from ...
    ' ATTENTION : le champs 0 de la requete doit �tre le code identifiant le choix
    '             ex : CodeAdresse pour une ville
    
    ' Creation d'une liste sans determination des colonnes au prealable
    ' le nom des colonnes correspond au nom des champs de la chaine SQL
    
    Dim MyReq As ADODB.Recordset
    Dim itmX As ListItem
    Dim NbCol As Integer
    Dim compteur As Long
    Dim clmX As ColumnHeader
    Dim i As Long, j As Long
    Dim ColWidth() As Integer
    Dim theList As ListView
    
    Set theList = oObjList
    
    Set MyReq = m_DataSource.OpenRecordset(ChaineSQL, Snapshot)
    
    theList.View = lvwReport   'Determination de l'affichage sous forme de liste
    theList.ListItems.Clear     'Suppression des elements de la liste
    theList.ColumnHeaders.Clear 'Suppression des colonnes
    
    NbCol = MyReq.Fields.Count - 1
    For i = 1 To NbCol
        'Creation des colonnes
        Set clmX = theList.ColumnHeaders.Add(, , MyReq.Fields(i).Name, theList.Width / MyReq.Fields.Count)
        If MyReq.Fields(i).Type = adDate _
           Or MyReq.Fields(i).Type = adDBDate _
           Or MyReq.Fields(i).Type = adDBTime _
           Or MyReq.Fields(i).Type = adDBTimeStamp Then
          clmX.Tag = "Date"
        Else
          clmX.Tag = " "
        End If
    Next i

    NbCol = theList.ColumnHeaders.Count
    
    While Not MyReq.EOF
        'Ajout des champs
        Set itmX = theList.ListItems.Add(, , CStr(MyReq.Fields(1).Value)) 'Champ principal
        itmX.Tag = CStr(MyReq.Fields(0).Value)
        For compteur = 2 To NbCol Step 1
          If Not IsNull(MyReq.Fields(compteur).Value) Then
            itmX.SubItems(compteur - 1) = CStr(MyReq.Fields(compteur).Value) ' Champs secondaire
          End If
        Next compteur
        
        MyReq.MoveNext   ' Passe � l'enregistrement suivant.
    Wend
    MyReq.Close
    
    ' calcule la largeur des colonnes
    'Call LargeurAutomatique(ObjWnd, ObjList)
End Sub

' Rempli le combo avec le r�sultat de la requete
' le premier champs de la requete doit �tre le code
'##ModelId=3F851B750148
Public Sub FillCombo(oCbo As Object, ChaineSQL As String, SelectItem As Long, Optional NonExclusif As Boolean, Optional ClearCombo As Boolean)
    Dim i As Long
    Dim NbCol As Integer
    Dim MyReq As ADODB.Recordset
    Dim text As String
    Dim theCombo As ComboBox
    
    Set theCombo = oCbo
    
    Set MyReq = m_DataSource.OpenRecordset(ChaineSQL, Snapshot)

    
    If Not IsMissing(ClearCombo) Then If ClearCombo = True Then theCombo.Clear
    If IsMissing(ClearCombo) Then theCombo.Clear
    
    
    'cbo.Sorted = False
    While Not MyReq.EOF
        'Ajout des champs
        text = ""
        NbCol = MyReq.Fields.Count - 1
        For i = 1 To NbCol Step 1
          If Not IsNull(MyReq.Fields(i).Value) Then
            If i > 1 Then
              text = text & " "
            End If
            text = text & MyReq.Fields(i).Value
          End If
        Next i
        
        theCombo.AddItem text
        theCombo.ItemData(theCombo.ListCount - 1) = MyReq.Fields(0).Value
        
        If SelectItem <> -1 And MyReq.Fields(0).Value = SelectItem Then theCombo.ListIndex = theCombo.ListCount - 1
        
        MyReq.MoveNext   ' Passe � l'enregistrement suivant.
        
        If Not IsMissing(NonExclusif) Then If NonExclusif = True Then DoEvents
    Wend
    'cbo.Sorted = True
    MyReq.Close
End Sub

' Rempli la liste avec le r�sultat de la requete
' le premier champs de la requete doit �tre le code
'##ModelId=3F851B7501B5
Public Sub FillList(oList As Object, ChaineSQL As String, SelectItem As Long, fClear As Boolean)
    Dim i As Long
    Dim NbCol As Integer
    Dim MyReq As ADODB.Recordset
    Dim text As String
    Dim theList As ListBox
    
    Set theList = oList
    
    Set MyReq = m_DataSource.OpenRecordset(ChaineSQL, Snapshot)

    If fClear = True Then
      theList.Clear
    End If
    
    While Not MyReq.EOF
        'Ajout des champs
        text = ""
        NbCol = MyReq.Fields.Count - 1
        For i = 1 To NbCol Step 1
          If Not IsNull(MyReq.Fields(i).Value) Then
            If i > 1 Then
              text = text & " "
            End If
            text = text & MyReq.Fields(i).Value
          End If
        Next i
        
        If text <> "" Then
          theList.AddItem text
          theList.ItemData(theList.ListCount - 1) = MyReq.Fields(0).Value
        End If
        
        If SelectItem <> -1 And MyReq.Fields(0).Value = SelectItem Then
          theList.ListIndex = theList.ListCount - 1
        End If
        
        MyReq.MoveNext   ' Passe � l'enregistrement suivant.
    Wend
    'cbo.Sorted = True
    MyReq.Close
End Sub

Public Function BuildDateString(dDate As Date) As String
  Select Case m_SQLMode
    Case SQLServer
      BuildDateString = " convert(datetime, '" & Format(dDate, "dd/mm/yyyy") & "', 103) " ' SQL Server : format 103 = 'dd/mm/yyyy'
    
    Case MySQL
      BuildDateString = " convert('" & Format(dDate, "yyyy-mm-dd") & "', DATE) " ' MySQL : les dates sont au format yyyy-mm-dd
    
    Case MaxDB
      BuildDateString = " '" & Format(dDate, "yyyy-mm-dd") & "' " ' MaxDB : les dates sont au format yyyy-mm-dd
    
    Case MSAccess
      BuildDateString = " #" & Format(dDate, "mm/dd/yyyy") & "# " ' MS Access

    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
End Function

Public Function BuildSQLDisplayDate(fieldName As String) As String
  Select Case m_SQLMode
    Case SQLServer
      BuildSQLDisplayDate = " convert(varchar, " & fieldName & ", 103) " ' SQL Server : format 103 = 'dd/mm/yyyy'
    
    Case MySQL
      BuildSQLDisplayDate = " DATE_FORMAT(" & fieldName & ", '%d/%m/%Y') " ' MySQL : les dates sont au format yyyy-mm-dd
    
    Case MSAccess
      BuildSQLDisplayDate = " Format(" & fieldName & ", ""mm/dd/yyyy"") " ' MS Access

    Case MaxDB
      'BuildSQLDisplayDate = " CHAR(" & fieldName & ", 'DD/MM/YYYY') " ' MaxDB
      BuildSQLDisplayDate = fieldName    ' MaxDB
    
    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
End Function

Public Function BuildIsNull(fieldName As String) As String
  Select Case m_SQLMode
    Case SQLServer, MySQL, MaxDB
      BuildIsNull = " (" & fieldName & " IS NULL) "
      'BuildIsNull = fieldName & " IS NULL "
    
    Case MSAccess
      BuildIsNull = "IsNull(" & fieldName & ")"

    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
End Function

'##ModelId=3F851B750213
Public Function BuildDateLimit(fieldName As String, Debut As String, Fin As String) As String
  Select Case m_SQLMode
    Case SQLServer
      BuildDateLimit = " ( " & fieldName & " >= convert(datetime, '" & Debut & "', 103) AND " & fieldName & " <= convert(datetime, '" & Fin & "', 103) ) "
    
    Case MySQL
      BuildDateLimit = " ( " & fieldName & " >= convert('" & Format(Debut, "yyyy-mm-dd") & "', DATE) AND " & fieldName & " <= convert('" & Format(Fin, "yyyy-mm-dd") & "', DATE) ) "
    
    Case MaxDB
      BuildDateLimit = " ( " & fieldName & " >= '" & Format(Debut, "yyyy-mm-dd") & "' AND " & fieldName & " <= '" & Format(Fin, "yyyy-mm-dd") & "' ) "
      
    Case MSAccess
      BuildDateLimit = fieldName & " BETWEEN #" & Format(CDate(Debut), "mm/dd/yyyy") & "# AND #" & Format(CDate(Fin), "mm/dd/yyyy") & "#"
  
    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
End Function

'##ModelId=3F851B750251
'FIXIT: Declare 'GetParameter' with an early-bound data type                               FixIT90210ae-R1672-R1B8ZE
Public Function GetParameter(rq As String) As Variant
  Dim rs As ADODB.Recordset
  
  Set rs = m_DataSource.OpenRecordset(rq, Dynamic)
  
  If Not rs.EOF Then
    GetParameter = rs.Fields(0).Value
  Else
    GetParameter = ""
  End If
  
  rs.Close
End Function

Public Function PeriodeExists(numPeriode As Long, GroupeCle As Integer) As Boolean
  
  Dim rs As ADODB.Recordset
  Set rs = m_DataSource.OpenRecordset("SELECT * FROM Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode, Dynamic)
  
  If Not rs.EOF Then
    PeriodeExists = True
  Else
    PeriodeExists = False
  End If
  
  rs.Close
  
End Function


'##ModelId=3F851B750271
Public Function GetParameterAsDouble(rq As String) As Double
  Dim rs As ADODB.Recordset
  
  Set rs = m_DataSource.OpenRecordset(rq, Dynamic)
  
  If Not rs.EOF Then
    If IsNull(rs.Fields(0).Value) Then
      GetParameterAsDouble = 0
    Else
      GetParameterAsDouble = rs.Fields(0).Value
    End If
  Else
    'Err.Raise -1, GetParameterAsDouble, "Requete invalide"
    GetParameterAsDouble = 0
  End If
  
  rs.Close
End Function


Public Function GetParameterAsDoubleWithParam(rq As String, param1 As String) As Double
  'D�claration des variables
  Dim theCommand As ADODB.Command
  Dim prm1 As ADODB.Parameter
  Dim rs As ADODB.Recordset
  
  'Instanciation des variables
  Set theCommand = New ADODB.Command
  
  'Pr�paration de l'objet Command
  theCommand.CommandText = rq
  
  'Pr�paration du param�tre
  Set prm1 = theCommand.CreateParameter("@param1", adVarChar, adParamInput, IIf(IsEmpty(param1), 1, Len(param1)), param1)
  theCommand.Parameters.Append prm1
  
  'Ex�cution de la requ�te
  Set rs = theCommand.Execute
  
  If Not rs.EOF Then
    If IsNull(rs.Fields(0).Value) Then
      GetParameterAsDoubleWithParam = 0
    Else
      GetParameterAsDoubleWithParam = rs.Fields(0).Value
    End If
  Else
    'Err.Raise -1, GetParameterAsDouble, "Requete invalide"
    GetParameterAsDoubleWithParam = 0
  End If
  
  rs.Close
End Function


Public Function GetParameterAsLongWithParam(rq As String, param1 As String) As Long
  'D�claration des variables
  Dim theCommand As ADODB.Command
  Dim prm1 As ADODB.Parameter
  Dim rs As ADODB.Recordset
  
  'Instanciation des variables
  Set theCommand = New ADODB.Command
  
  'Pr�paration de l'objet Command
  theCommand.ActiveConnection = m_DataSource.Connection
  theCommand.CommandType = adCmdText
  theCommand.CommandText = rq
  
  'Pr�paration du param�tre
  Set prm1 = theCommand.CreateParameter("@param1", adVarChar, adParamInput, IIf(IsEmpty(param1), 1, Len(param1)), param1)
  theCommand.Parameters.Append prm1
  
  'Ex�cution de la requ�te
  Set rs = theCommand.Execute
  
  If Not rs.EOF Then
    If IsNull(rs.Fields(0).Value) Then
      GetParameterAsLongWithParam = 0
    Else
      GetParameterAsLongWithParam = rs.Fields(0).Value
    End If
  Else
    'Err.Raise -1, GetParameterAsDouble, "Requete invalide"
    GetParameterAsLongWithParam = 0
  End If
  
  rs.Close
End Function


'##ModelId=3F851B750271

Public Function GetParameterAsLong(rq As String) As Long
  Dim rs As ADODB.Recordset
  
  Set rs = m_DataSource.OpenRecordset(rq, Dynamic)
  
  If Not rs.EOF Then
    If IsNull(rs.Fields(0).Value) Then
      GetParameterAsLong = 0
    Else
      GetParameterAsLong = rs.Fields(0).Value
    End If
  Else
    'Err.Raise -1, GetParameterAsLong, "Requete invalide"
    GetParameterAsLong = 0
  End If
  
  rs.Close
End Function

'##ModelId=3F851B75029F
Public Function GetParameterAsStringCRW(rq As String) As String
  Dim rs As ADODB.Recordset
  
  Set rs = m_DataSource.OpenRecordset(rq, Dynamic)
  
  If Not rs.EOF Then
    If IsNull(rs.Fields(0).Value) Then
      GetParameterAsStringCRW = 0
    Else
      GetParameterAsStringCRW = VirguleVersPoint(rs.Fields(0).Value)
    End If
  Else
    'Err.Raise -1, GetParameterAsStringCRW, "Requete invalide"
    GetParameterAsStringCRW = 0
  End If
  
  rs.Close
End Function


Public Function VirguleVersPoint(texte As String) As String
  Dim pos As Integer
    
  pos = InStr(1, texte, ",")
  If pos <> 0 Then
    VirguleVersPoint = Left$(texte, pos - 1) & "." & Right$(texte, Len(texte) - pos)
  Else
    VirguleVersPoint = texte
  End If
End Function

Public Function PointVersVirgule(texte As String) As String
  Dim pos As Integer
    
  pos = InStr(1, texte, ".")
  If pos <> 0 Then
    PointVersVirgule = Left$(texte, pos - 1) & "," & Right$(texte, Len(texte) - pos)
  Else
    PointVersVirgule = texte
  End If
End Function

Public Sub GetDouble(ByRef f As ADODB.Field, t As String)
  Dim e As Boolean
  Dim tt As String
  
  On Error GoTo errGetDouble
  
  ' essaye
  e = False
  f.Value = CDbl(t)
  If e = False Then Exit Sub
  
  ' essaye sans '.'
  e = False
  tt = PointVersVirgule(t)
  f.Value = CDbl(tt)
  If e = False Then Exit Sub
  
  ' essaye sans ','
  e = False
  tt = VirguleVersPoint(t)
  f.Value = CDbl(tt)
  If e = False Then Exit Sub
  
  f.Value = 0
    
  On Error GoTo 0
  
  Exit Sub
  
errGetDouble:
  e = True
  Resume Next
End Sub

Public Function GetDouble2(t As String) As Double

  Dim e As Boolean
  Dim tt As String
  
  On Error GoTo errGetDouble2
  
  ' essaye
  e = False
  GetDouble2 = CDbl(t)
  If e = False Then Exit Function
  
  ' essaye sans '.'
  e = False
  tt = PointVersVirgule(t)
  GetDouble2 = CDbl(tt)
  If e = False Then Exit Function
  
  ' essaye sans ','
  e = False
  tt = VirguleVersPoint(t)
  GetDouble2 = CDbl(tt)
  If e = False Then Exit Function
  
  GetDouble2 = 0
    
  On Error GoTo 0
  
  Exit Function
  
errGetDouble2:
  e = True
  Resume Next
End Function

Public Sub Multi_Find(ByRef oRs As ADODB.Recordset, sCriteria As String)
  If Left$(Trim$(UCase$(sCriteria)), 4) = "AND " Then
    sCriteria = Mid$(sCriteria, 5)
  End If
  
  oRs.Filter = adFilterNone
  oRs.Filter = sCriteria
End Sub

Public Function ValidateSQL(sSQL As String) As String
  ValidateSQL = m_DataSource.ValidateSQL(sSQL)
End Function
