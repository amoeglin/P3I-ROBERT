VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListCategorieIntrouvable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75DB0339"
Option Explicit

Private Type defCategorie
'  Regime As Long
'  Categorie As Long
  Regime As String
  Categorie As String
End Type

'##ModelId=5C8A75DC007A
Private m_nbCatNotFound As Long
'##ModelId=5C8A75DC009B
Private m_lstCatNotFound() As defCategorie

'##ModelId=5C8A75DC00AB
Private m_Logger As clsLogger
'##ModelId=5C8A75DC00AC
Private m_TableName As String
'##ModelId=5C8A75DC00C8
Private m_NomCategorie As String
'##ModelId=5C8A75DC00E8
Private m_NomRegime As String

'##ModelId=5C8A75DC0107
Public autoMode As Boolean



'##ModelId=5C8A75DC0126
Private Sub Class_Initialize()
  m_nbCatNotFound = 0
  m_NomCategorie = "Catégorie"
  m_NomRegime = "Régime"
End Sub


'##ModelId=5C8A75DC0136
Private Sub Class_Terminate()
  m_nbCatNotFound = 0
  Erase m_lstCatNotFound
End Sub


'##ModelId=5C8A75DC0145
Public Property Set Logger(l As clsLogger)
  Set m_Logger = l
End Property


'##ModelId=5C8A75DC0174
Public Property Let tableName(t As String)
  m_TableName = t
End Property


'##ModelId=5C8A75DC01A3
Public Property Let NomCategorie(t As String)
  m_NomCategorie = t
End Property


'##ModelId=5C8A75DC01C2
Public Property Let NomRegime(t As String)
  m_NomRegime = t
End Property


'Public Sub AddCategorie(Regime As Long, Categorie As Long, Optional bLog As
'Boolean = False)
'##ModelId=5C8A75DC01F1
Public Sub AddCategorie(ByVal Regime As String, ByVal Categorie As String, Optional bLog As Boolean = False)
  Dim iCat As Integer, bFound As Boolean
  
  ' recherche si la categorie a déjà été rencontré
  bFound = False
  For iCat = 0 To m_nbCatNotFound - 1
    If m_lstCatNotFound(iCat).Regime = Regime And m_lstCatNotFound(iCat).Categorie = Categorie Then
      bFound = True
      Exit For
    End If
  Next iCat
  
  If Not bFound Then
    ReDim Preserve m_lstCatNotFound(m_nbCatNotFound + 1) As defCategorie
    
    m_lstCatNotFound(m_nbCatNotFound).Regime = Regime
    m_lstCatNotFound(m_nbCatNotFound).Categorie = Categorie
    
    m_nbCatNotFound = m_nbCatNotFound + 1
    
    If bLog Then
      m_Logger.EcritTraceDansLog IIf(Regime <> 0, "Regime/", "") & "Catégorie à ajouter dans la table '" & m_TableName & "' : " & IIf(Regime <> 0, Regime & "/", "") & Categorie
    End If
  End If
End Sub

'##ModelId=5C8A75DC0238
Public Sub LogCategorie()
  Dim iCat As Integer, bFound As Boolean
  
  On Error GoTo err_log
  
  ' recherche si la categorie a déjà été rencontré
  bFound = False
  If m_nbCatNotFound <> 0 Then
    m_Logger.EcritTraceDansLog IIf(m_lstCatNotFound(iCat).Regime <> "0", m_NomRegime & "/", "") & m_NomCategorie & " à ajouter dans la table '" & m_TableName & "'"
    For iCat = 0 To m_nbCatNotFound - 1
      m_Logger.EcritTraceDansLog "     " & IIf(m_lstCatNotFound(iCat).Regime <> "0", m_NomRegime & " " & m_lstCatNotFound(iCat).Regime & " / ", "") & m_NomCategorie & " " & m_lstCatNotFound(iCat).Categorie
    Next iCat
    m_Logger.EcritTraceDansLog "     "
  End If
  
  Exit Sub
  
err_log:
  Resume Next
End Sub

