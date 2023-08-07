VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListeMessageUnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A75F50246"
Option Explicit

'##ModelId=5C8A75F50384
Private m_nb As Long
'##ModelId=5C8A75F503A3
Private m_lst() As String

'##ModelId=5C8A75F503C4
Private m_Logger As clsLogger

'##ModelId=5C8A75F503C5
Public autoMode As Boolean
'


'##ModelId=5C8A75F503E1
Private Sub Class_Initialize()
  m_nb = 0
End Sub


'##ModelId=5C8A75F60009
Private Sub Class_Terminate()
  m_nb = 0
  Erase m_lst
End Sub


'##ModelId=5C8A75F60019
Public Property Set Logger(l As clsLogger)
  Set m_Logger = l
End Property


'##ModelId=5C8A75F60048
Public Sub AddMessage(ByVal Message As String)
  Dim i As Integer, bFound As Boolean
  
  ' recherche si la categorie a déjà été rencontré
  bFound = False
  For i = 0 To m_nb - 1
    If m_lst(i) = Message Then
      bFound = True
      Exit For
    End If
  Next i
  
  If Not bFound Then
    ReDim Preserve m_lst(m_nb + 1) As String
    
    m_lst(m_nb) = Message
   
    m_nb = m_nb + 1
  End If
End Sub

'##ModelId=5C8A75F60076
Public Sub LogMessage()
  Dim i As Integer, bFound As Boolean
  
  On Error GoTo err_log
  
  ' recherche si la categorie a déjà été rencontré
  bFound = False
  For i = 0 To m_nb - 1
    m_Logger.EcritTraceDansLog "     " & m_lst(i)
  Next i
  m_Logger.EcritTraceDansLog "     "
  
  Exit Sub
  
err_log:
  Resume Next
End Sub

