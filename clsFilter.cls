VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75D902BB"
Option Explicit

'
' Definitions
'

' type de donnée
Public Enum Filter_Elem_Type
  Number = 1
  Char
End Enum

' type de chaine de sélection à créer
Public Enum Selection_Type
  sql = 1
  CRW
  Ecran
End Enum

' élément de filtre
' SQLToDisplayName : la requete sera passée à GetParameter en remplacant :
'                    <Value> par la valeur du filtre
'                    <ClePeriode> par m_ClePeriode
'                    <CleGroupe> par m_CleGroupe
'                    ex : "SELECT SONOM FROM Societe WHERE SOCLE=<Value> AND SOGROUPE=<CleGroupe>"
Private Type Filter_Elem
  DataType As Filter_Elem_Type
  
  Name As String
  FieldName As String
  CRWFieldName As String
  
  SQLToDisplayName As String
  
  CanBeNull As Boolean ' si true, IsNull est generé comme filtre si value="(null)"
  ErrorIfNull As Boolean ' si true, génére une erreur si value=0, "" ou "(null)"
  
  UseLike As Boolean ' si true, genere une commande SQL : <fieldname> like "<value>%"
  ShowDetail As Boolean ' si true, apparait dans la chaine affiché dans lblFilter
  
  Value As String
End Type

'
' members list
'
'##ModelId=5C8A75D903BF
Private m_NbFilterElem As Integer
'##ModelId=5C8A75D903E0
Private m_FilterElem() As Filter_Elem

'##ModelId=5C8A75DA0016
Private m_ClePeriode As Long
'##ModelId=5C8A75DA0025
Private m_CleGroupe As Long
  
' pour GetFilterValue
'##ModelId=5C8A75DA0054
Private last_name As String
'##ModelId=5C8A75DA0073
Private last_value As String

'##ModelId=5C8A75DA0093
Public autoMode As Boolean
'

'##ModelId=5C8A67E90254


'##ModelId=5C8A75DA00B2
Private Sub Class_Initialize()
  Clear
End Sub

'##ModelId=5C8A75DA00C1
Private Sub Class_Terminate()
  Clear
End Sub

'##ModelId=5C8A75DA00D1
Public Sub Clear()
  last_name = ""
  last_value = ""
  
  m_NbFilterElem = 0
  Erase m_FilterElem
End Sub

'##ModelId=5C8A75DA00E1
Public Sub ClearValue()
  Dim i As Integer
  
  last_name = ""
  last_value = ""
  
  For i = 0 To m_NbFilterElem - 1
    m_FilterElem(i).Value = ""
  Next
End Sub

'##ModelId=5C8A75DA00F0
Public Sub AddFilterElem(DataType As Filter_Elem_Type, Name As String, FieldName As String, CRWFieldName As String, SQLToDisplayName As String, CanBeNull As Boolean, ErrorIfNull As Boolean, UseLike As Boolean, ShowDetail As Boolean)
  Dim i As Integer
  
  ' vérifie que le nom est unique
  For i = 0 To m_NbFilterElem - 1
    If Name = m_FilterElem(i).Name Then
      Err.Raise -1, "clsFilter", "Filtre " & Name & " déjà existant !"
      Exit Sub
    End If
  Next
  
  ' alloue de la mémoire pour l'élément
  ReDim Preserve m_FilterElem(m_NbFilterElem + 1)
  
  ' sauve les données de l'élément
  With m_FilterElem(m_NbFilterElem)
    .DataType = DataType
    .Name = Name
    .FieldName = FieldName
    .CRWFieldName = CRWFieldName
    .SQLToDisplayName = SQLToDisplayName
    .CanBeNull = CanBeNull
    .ErrorIfNull = ErrorIfNull
    .UseLike = UseLike
    .ShowDetail = ShowDetail
    .Value = ""
  End With
  
  m_NbFilterElem = m_NbFilterElem + 1
End Sub

'##ModelId=5C8A75DA01CB
Public Sub SetFilterElemValue(Name As String, Value As String)
  Dim i As Integer, bFound As Boolean
  
  last_name = ""
  last_value = ""
  
  If m_NbFilterElem = 0 Then Exit Sub
  
  bFound = False
  For i = 0 To m_NbFilterElem - 1
    If Name = m_FilterElem(i).Name Then
      If Value = "" Or Value = FILTER_VALUE_NULL Then
        ' erreur si null ?
        If m_FilterElem(i).ErrorIfNull = True Then
          Err.Raise -1, "clsFilter", "Filtre " & m_FilterElem(i).Name & " ne peut être '(null)' !"
          Exit Sub
        End If
        
        ' can be null ?
        If m_FilterElem(i).CanBeNull = False Then
          Value = ""
        End If
      End If
      
      ' cas particulier pour le NCA :
      '   le format est "0 00 000000 00 00". les " " et "0" en tete doivent etre supprimés.
      '   la base de données ne les stocke pas.
      If m_FilterElem(i).Name = "NCA" Then
        Value = ReplaceInString(Value, " ", "")
        
        Do While Left(Value, 1) = "0"
          Value = mID(Value, 2)
        Loop
      End If
      
      m_FilterElem(i).Value = Value
      bFound = True
      
      Exit For
    End If
  Next

  If bFound = False Then
    Err.Raise -1, "clsFilter", "Filtre " & Name & " introuvable !"
  End If
End Sub

'##ModelId=5C8A75DA0210
Public Function GetFilterElemValue(Name As String) As String
  Dim i As Integer, bFound As Boolean
  
  ' speed up loop
  If last_name = Name Then
    GetFilterElemValue = last_value
    Exit Function
  End If
  
  GetFilterElemValue = ""
  
  If m_NbFilterElem = 0 Then Exit Function
  
  bFound = False
  For i = 0 To m_NbFilterElem - 1
    If Name = m_FilterElem(i).Name Then
      GetFilterElemValue = m_FilterElem(i).Value
      
      If GetFilterElemValue = "" And m_FilterElem(i).DataType = Number Then
        GetFilterElemValue = "0"
      End If
      
      bFound = True
      Exit For
    End If
  Next

  If bFound = False Then
    last_name = ""
    last_value = ""
    
    Err.Raise -1, "clsFilter", "Filtre " & Name & " introuvable !"
    Exit Function
  End If
  
  ' speed up loop
  last_name = Name
  last_value = GetFilterElemValue
End Function

'##ModelId=5C8A75DA0229
Public Property Let numPeriode(n As Long)
  m_ClePeriode = n
  If n = 0 Then
    Err.Raise Number:=-1, Description:="N° de période invalide : " & n, Source:="clsFilter"
  End If
End Property

'##ModelId=5C8A75DA0259
Public Property Let NumGroupe(n As Long)
  m_CleGroupe = n
  If n = 0 Then
    Err.Raise Number:=-1, Description:="N° de groupe invalide : " & n, Source:="clsFilter"
  End If
End Property

'
' sauve le filtre en cours dans le .INI
'
'##ModelId=5C8A75DA0278
Public Sub Save(ClePeriode As Long)
  Dim dom As New DOMDocument, xmlRoot As IXMLDOMElement
  Dim nodeGroupe As IXMLDOMElement, nodeSousGroupe As IXMLDOMElement
  Dim nomSection As String, i As Integer
  
  On Error GoTo err_Save
  
  nomSection = "Filtre_Periode_" & ClePeriode
  Set xmlRoot = dom.createElement(nomSection)
  dom.appendChild xmlRoot
  
  Set nodeGroupe = dom.createElement("Filtre")
  
  For i = 0 To m_NbFilterElem - 1
    nodeGroupe.setAttribute Replace(m_FilterElem(i).Name, " ", "_"), m_FilterElem(i).Value
  Next
  
  xmlRoot.appendChild nodeGroupe
  
  ' sauve le XML dans le .INI
  Call SaveSettingIni(CompanyName, nomSection, "Xml", Trim(dom.xml))
  
  Exit Sub
  
err_Save:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub

'
' recharge le filtre depuis le .INI
'
'##ModelId=5C8A75DA0298
Public Sub Load(ClePeriode As Long)
  Dim dom As New DOMDocument, xmlRoot As IXMLDOMElement
  Dim nodeGroupe As IXMLDOMElement, nodeSousGroupe As IXMLDOMElement
  Dim attr As IXMLDOMAttribute, node As IXMLDOMNode
  Dim nomSection As String, i As Integer
  
  On Error GoTo err_Load
  
  nomSection = "Filtre_Periode_" & ClePeriode
  
  ' relie le XML depuis le .INI
  dom.loadXML GetSettingIni(CompanyName, nomSection, "Xml", "")
  
  ' decode le XML
  Set xmlRoot = dom.documentElement
  
  If xmlRoot Is Nothing Then Exit Sub
  
  Set nodeGroupe = xmlRoot.childNodes(0)
  
  For i = 0 To m_NbFilterElem - 1
    Set node = nodeGroupe.Attributes.getNamedItem(Replace(m_FilterElem(i).Name, " ", "_"))
    If Not (node Is Nothing) Then
      If m_FilterElem(i).DataType = Number Then
        If node.text <> "" Then
          If node.text = FILTER_VALUE_NULL Then
            m_FilterElem(i).Value = FILTER_VALUE_NULL
          Else
            m_FilterElem(i).Value = CLng(node.text)
          End If
        Else
          m_FilterElem(i).Value = ""
        End If
      Else
        m_FilterElem(i).Value = node.text
      End If
    End If
  Next
  
  Exit Sub
  
err_Load:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub

'##ModelId=5C8A75DA02D2
Public Function BuildSelectionString(SelType As Selection_Type) As String
  Dim i As Integer, szSQL As String
  Dim FieldName As String
    
  BuildSelectionString = ""
  
  For i = 0 To m_NbFilterElem - 1
    If m_FilterElem(i).Value <> "" Then
      If SelType <> Ecran Or m_FilterElem(i).ShowDetail = True Then
        Select Case SelType
          Case sql
            FieldName = m_FilterElem(i).FieldName
          
          Case CRW
            FieldName = m_FilterElem(i).CRWFieldName
          
          Case Ecran
            FieldName = m_FilterElem(i).Name
        End Select
        
        ' (null) ?
        If m_FilterElem(i).Value = FILTER_VALUE_NULL Then
          If m_FilterElem(i).CanBeNull Then
            If BuildSelectionString <> "" Then
              BuildSelectionString = BuildSelectionString & IIf(SelType = Ecran, ", ", " AND ")
            End If
            BuildSelectionString = BuildSelectionString & IIf(SelType = Ecran, FieldName & " (aucun)", "IsNull(" & FieldName & ")")
          End If
        Else
          ' non null
          If BuildSelectionString <> "" Then
            BuildSelectionString = BuildSelectionString & IIf(SelType = Ecran, ", ", " AND ")
          End If
          
          ' like "...%" ?
          If m_FilterElem(i).UseLike Then
            BuildSelectionString = BuildSelectionString & IIf(SelType = CRW, "UpperCase(" & FieldName & ")", FieldName)
            BuildSelectionString = BuildSelectionString & IIf(SelType = Ecran, "=", " LIKE ")
          Else
            BuildSelectionString = BuildSelectionString & FieldName & "="
          End If
          
          ' type=char ?
          If m_FilterElem(i).DataType = Char Then
            BuildSelectionString = BuildSelectionString & "'"
          End If
          
          If SelType = Ecran And m_FilterElem(i).SQLToDisplayName <> "" Then
            ' SQLToDisplayName : la requete sera passée à GetParameter en remplacant :
            '                    <Value> par la valeur du filtre
            '                    <ClePeriode> par m_ClePeriode
            '                    <CleGroupe> par m_CleGroupe
            '                    ex : "SELECT SONOM FROM Societe WHERE SOCLE=<Value> AND SOGROUPE=<CleGroupe>"
            szSQL = m_FilterElem(i).SQLToDisplayName
            
            szSQL = Replace(szSQL, "<Value>", m_FilterElem(i).Value)
            szSQL = Replace(szSQL, "<ClePeriode>", m_ClePeriode)
            szSQL = Replace(szSQL, "<CleGroupe>", m_CleGroupe)
            
            BuildSelectionString = BuildSelectionString & m_dataHelper.GetParameter(szSQL)
          Else
            If SelType = Ecran And m_FilterElem(i).Name = "NCA" Then
              BuildSelectionString = BuildSelectionString & Format(m_FilterElem(i).Value, m_FormatNCA)
            Else
              If m_FilterElem(i).UseLike Then
                BuildSelectionString = BuildSelectionString & Replace(m_FilterElem(i).Value, "'", "''")
              Else
                BuildSelectionString = BuildSelectionString & m_FilterElem(i).Value
              End If
            End If
          End If
        
          ' like "...%" ?
          If m_FilterElem(i).UseLike Then
            BuildSelectionString = BuildSelectionString & IIf(SelType = Ecran, "*", "%")
          End If
          
          ' type=char ?
          If m_FilterElem(i).DataType = Char Then
            BuildSelectionString = BuildSelectionString & "'"
          End If
        End If
      End If
    End If
  Next
  
  If SelType <> Ecran And BuildSelectionString <> "" Then
    BuildSelectionString = " AND (" & BuildSelectionString & ") "
  End If
End Function

'##ModelId=5C8A75DA02F1
Public Function GetSelectionSQLString() As String
  GetSelectionSQLString = BuildSelectionString(sql)
End Function

'##ModelId=5C8A75DA0301
Public Function GetCRWFilterSQLString() As String
  GetCRWFilterSQLString = BuildSelectionString(CRW)
End Function

'##ModelId=5C8A75DA0310
Public Function SelectionString() As String
  SelectionString = BuildSelectionString(Ecran)
End Function

'##ModelId=5C8A75DA0320
Public Function BuildFilterString() As String

  Dim i As Integer, szSQL As String
  Dim FieldName As String
    
  BuildFilterString = ""
    
  For i = 0 To m_NbFilterElem - 1
    If m_FilterElem(i).Value <> "" Then
              
      FieldName = m_FilterElem(i).FieldName
      
      If AssureDisplays.CurrentlySelectedDisplay.ContainsFieldName(FieldName) Then
      
        ' (null) ?
        If m_FilterElem(i).Value = FILTER_VALUE_NULL Then
          If m_FilterElem(i).CanBeNull Then
            If BuildFilterString <> "" Then
              BuildFilterString = BuildFilterString & " AND "
            End If
            BuildFilterString = BuildFilterString & "IsNull(" & FieldName & ")"
          End If
        Else
          ' non null
          If BuildFilterString <> "" Then
            BuildFilterString = BuildFilterString & " AND "
          End If
          
          ' like "...%" ?
          If m_FilterElem(i).UseLike Then
            BuildFilterString = BuildFilterString & FieldName
            BuildFilterString = BuildFilterString & " LIKE "
          Else
            BuildFilterString = BuildFilterString & FieldName & "="
          End If
          
          ' type=char ?
          If m_FilterElem(i).DataType = Char Then
            BuildFilterString = BuildFilterString & "'"
          End If
          
          If m_FilterElem(i).UseLike Then
            BuildFilterString = BuildFilterString & Replace(m_FilterElem(i).Value, "'", "''")
          Else
            BuildFilterString = BuildFilterString & m_FilterElem(i).Value
          End If
        
          ' like "...%" ?
          If m_FilterElem(i).UseLike Then
            BuildFilterString = BuildFilterString & "%"
          End If
          
          ' type=char ?
          If m_FilterElem(i).DataType = Char Then
            BuildFilterString = BuildFilterString & "'"
          End If
        End If
      
      End If
      
    End If
  Next
  
  If BuildFilterString <> "" Then
    BuildFilterString = " AND (" & BuildFilterString & ") "
  End If
  
End Function


