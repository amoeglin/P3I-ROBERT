VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProvisionBCAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A7602039F"

Option Explicit

'
' Members
'

'##ModelId=5C8A760300FF
Public m_CleTable As Variant
'##ModelId=5C8A7603011E
Public m_Anciennete As Variant
'##ModelId=5C8A7603012E
Public m_Age As Variant
'##ModelId=5C8A7603014D
Public m_Provision As Variant

'##ModelId=5C8A7603016C
Public autoMode As Boolean
'##ModelId=5C8A7603018E
Public m_Logger As clsLogger



'
' Constructors
'


'##ModelId=5C8A7603019B
Private Sub Class_Initialize()
    
   m_CleTable = Null
   m_Anciennete = Null
   m_Age = Null
   m_Provision = Null

End Sub



'##ModelId=5C8A760301AB
Public Sub Copy(ByRef src As clsProvisionBCAC)

   m_CleTable = src.m_CleTable
   m_Anciennete = src.m_Anciennete
   m_Age = src.m_Age
   m_Provision = src.m_Provision

End Sub



'---------------------------------------------------
'##ModelId=5C8A760301DA
Public Function Load(DB As DataAccess, CleTable As Long, Anciennete As Long, Age As Long) As Boolean

   Dim theData As ADODB.Recordset
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Load

   Set theData = DB.OpenRecordset("SELECT * FROM ProvisionBCAC WHERE  CleTable=" & CleTable & " AND  Anciennete=" & Anciennete & " AND  Age=" & Age, Snapshot)
   If theData.EOF = False Then

      If IsNull(theData.fields("CleTable")) Then
         m_CleTable = 0
      Else
         m_CleTable = theData.fields("CleTable")
      End If

      If IsNull(theData.fields("Anciennete")) Then
         m_Anciennete = 0
      Else
         m_Anciennete = theData.fields("Anciennete")
      End If

      If IsNull(theData.fields("Age")) Then
         m_Age = 0
      Else
         m_Age = theData.fields("Age")
      End If

      If IsNull(theData.fields("Provision")) Then
         m_Provision = 0
      Else
         m_Provision = theData.fields("Provision")
      End If


   Else
      ret = False
   End If

   Load = ret

   Exit Function

err_Load:
   'MsgBox "Erreur dans clsProvisionBCAC::Load()", vbCritical
   If Not autoMode Then
      MsgBox "Erreur dans clsProvisionBCAC::Load() " & Err & vbLf & Err.Description, vbCritical
   Else
      m_Logger.EcritTraceDansLog "Erreur dans clsProvisionBCAC::Load() " & Err & " : " & Err.Description
   End If
   
   Load = False

End Function


'---------------------------------------------------
'##ModelId=5C8A76030228
Public Function Save(DB As DataAccess) As Boolean
   Dim sql As String, ret As Boolean
   
   ret = True

   On Error GoTo err_Save
   
   sql = "INSERT INTO ProvisionBCAC(CleTable, Anciennete, Age, Provision) VALUES(?, ?, ?, ?)"

   Dim theCommand As New ADODB.Command, prm As ADODB.Parameter
   theCommand.ActiveConnection = DB.Connection
   theCommand.CommandType = adCmdText
   theCommand.CommandText = sql

   Set prm = theCommand.CreateParameter("@CleTable", adInteger, adParamInput, 0, m_CleTable)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@Anciennete", adInteger, adParamInput, 0, m_Anciennete)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@Age", adInteger, adParamInput, 0, m_Age)
   theCommand.Parameters.Append prm

   Set prm = theCommand.CreateParameter("@Provision", adDouble, adParamInput, 0, m_Provision)
   theCommand.Parameters.Append prm

   theCommand.Execute

   Save = ret

   Exit Function

err_Save:
   MsgBox "Erreur dans clsProvisionBCAC::Save()" & vbLf & Err.Description, vbCritical
   Save = False

End Function


'---------------------------------------------------
'##ModelId=5C8A76030257
Public Function Delete(DB As DataAccess) As Boolean
   Dim ret As Boolean
   ret = True

   On Error GoTo err_Delete

   Dim theCommand As New ADODB.Command
   theCommand.ActiveConnection = DB.Connection
   theCommand.CommandType = adCmdText
   theCommand.CommandText = "DELETE FROM ProvisionBCAC WHERE  CleTable= " & m_CleTable & " AND  Anciennete= " & m_Anciennete & " AND  Age= " & m_Age

   theCommand.Execute

   m_CleTable = 0
m_Anciennete = 0
m_Age = 0


   Delete = ret

   Exit Function

err_Delete:
   MsgBox "Erreur dans clsProvisionBCAC::Delete()", vbCritical
   Delete = False

End Function




