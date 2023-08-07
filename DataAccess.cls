VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F851B74036B"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'              Classe d'accès aux données de la base                '
'    Permet d'avoir un framework standard qq soit l'application     '
' frmMain, frmPeriode et frmEditPériode deviennent donc indépendants'
' de l'appli (P3I, P2I, IFC, CAPISICA, ...                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Base 0

Option Explicit

Public Enum eRecordsetMode
  Dynamic
  Snapshot
  Table
End Enum

'##ModelId=3F8549100043
Private m_theAdoDB As ADODB.Connection

'##ModelId=3F854910004E
Private m_DatabaseConnectionString As String

'##ModelId=3F8530140261
Private m_Connected As Boolean

Private m_SQLMode As eSQLMode '

'##ModelId=3F855445000F
Public Sub Execute(SQLQuery As String, Optional additionnalOption As Long)
  If IsMissing(additionnalOption) Then
    m_theAdoDB.Execute ValidateSQL(SQLQuery), , adExecuteNoRecords
  Else
    m_theAdoDB.Execute ValidateSQL(SQLQuery), , additionnalOption
  End If
End Sub

'##ModelId=3F8553800222
Public Sub BeginTrans()
  On Error GoTo BeginTransErr

  m_theAdoDB.BeginTrans

  Exit Sub
BeginTransErr:
  MsgBox "Erreur dans DataAccess:BeginTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=3F85538902EE
Public Sub CommitTrans()
  On Error GoTo CommitTransErr

  m_theAdoDB.CommitTrans

  Exit Sub
CommitTransErr:
  MsgBox "Erreur dans DataAccess:CommitTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=3F85539200CB
Public Sub RollbackTrans()
  On Error GoTo RollbackTransErr

  m_theAdoDB.RollbackTrans

  Exit Sub
RollbackTransErr:
  MsgBox "Erreur dans DataAccess:RollbackTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'return the resulting recordset of a SQL query.
'##ModelId=3F854D9801B5
Public Function OpenRecordset(SQLQuery As String, recorsetType As eRecordsetMode) As ADODB.Recordset
  On Error GoTo OpenRecordsetErr
    
  ' la requete ne doit pas etre vide
  If Trim(SQLQuery) = "" Then
    Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset : requète invalide")
    Exit Function
  End If
  
  ' convertie les " en '
  Dim n As Integer
  Do
    n = InStr(SQLQuery, """")
    If n <> 0 Then
      SQLQuery = Left(SQLQuery, n - 1) & "'" & Mid(SQLQuery, n + 1)
    Else
      Exit Do
    End If
  Loop
  
  ' enleve les P3IUser. si MySQL
  SQLQuery = ValidateSQL(SQLQuery)
  
  ' ouvre le recordset
  Set OpenRecordset = New ADODB.Recordset
  
  ' conversion du type de recordset
  Dim typeRecordset As Long
  
  Select Case recorsetType
    Case Dynamic
      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenKeyset, adLockOptimistic
  
    Case Snapshot
      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenStatic, adLockOptimistic
      
    Case Table
      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenStatic, adLockOptimistic, adCmdTableDirect
      
    Case Else
      Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset : type de recordset invalide")
      Exit Function
  End Select

        
  Exit Function
OpenRecordsetErr:
  Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset Method")
End Function

'Create a DataHelper instance for this DataAccess instance
'##ModelId=3F854948003E
Public Function CreateHelper() As DataHelper
  On Error GoTo CreateHelperErr

  Set CreateHelper = New DataHelper
  Set CreateHelper.DataSource = Me
  
  If InStr(UCase(m_DatabaseConnectionString), "SQLOLEDB") <> 0 Then
    CreateHelper.SqlMode = SQLServer
  ElseIf InStr(UCase(m_DatabaseConnectionString), "MYSQL") <> 0 Then
    CreateHelper.SqlMode = MySQL
  ElseIf InStr(UCase(m_DatabaseConnectionString), ".MDB") <> 0 Then
    CreateHelper.SqlMode = MSAccess
  ElseIf InStr(UCase(m_DatabaseConnectionString), "SAP DB") <> 0 Then
    CreateHelper.SqlMode = MaxDB
  End If
  
  Exit Function
CreateHelperErr:
  Call RaiseError(MyUnhandledError, "DataAccess::CreateHelper Method")
End Function


'##ModelId=3F8530D401B5
Public Property Get Connected() As Boolean
  Let Connected = m_Connected
End Property

'##ModelId=3F854910007D
Public Property Get CurrentConnection() As String
  Let CurrentConnection = m_DatabaseConnectionString
End Property

'##ModelId=3F854910009C
Public Property Get Connection() As ADODB.Connection
  Set Connection = m_theAdoDB
End Property

'##ModelId=3F852FD900AB
Public Sub Disconnect()
  On Error GoTo DisconnectErr

    If Not m_theAdoDB Is Nothing Then
      If m_theAdoDB.State = adStateOpen Then m_theAdoDB.Close
    End If
    Set m_theAdoDB = Nothing
    
    m_DatabaseConnectionString = ""
    
    m_Connected = False
    
  Exit Sub
DisconnectErr:
  Call RaiseError(MyUnhandledError, "DataAccess:Disconnect Method")
End Sub

'##ModelId=3F851D8A0128
Private Sub Class_Terminate()
  #If DebugMode Then
    'the class is being destroyed
  Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
  #End If
  
  If m_Connected = True Then
    Disconnect
  End If
End Sub

'##ModelId=3F851D8A0119
Private Sub Class_Initialize()
  #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
    mlClassDebugID = GetNextClassDebugID()
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
  #End If
  
  Set m_theAdoDB = Nothing
  m_DatabaseConnectionString = ""
  m_Connected = False
End Sub

'Ouvre la connexion à la base.
'ConnectionString donne la chaine de connexion (ADO)
'AdditionnalOption donne les options nécessaire (Password pour DAO, etc ...)
'##ModelId=3F851B8C0138
Public Function Connect(ConnectionString As String) As Boolean
  On Error GoTo ConnectErr

    ' ouverture de la base en mode partage, lecture/ecriture
    'm_DatabaseConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\P3I_ADO\P3I.mdb;Persist Security Info=False;Jet OLEDB:Database Password=P3I32"
    'm_DatabaseConnectionString = "Provider=SQLOLEDB;Password=P3IPwd;Persist Security Info=True;User ID=P3IUser;Initial Catalog=P3I;Data Source=PORTABLE"
    m_DatabaseConnectionString = ConnectionString
    
    Set m_theAdoDB = New ADODB.Connection
    
    If InStr(UCase(m_DatabaseConnectionString), "SQLOLEDB") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = SQLServer
    ElseIf InStr(UCase(m_DatabaseConnectionString), "MYSQL") <> 0 Then
      m_theAdoDB.CursorLocation = adUseClient
      m_SQLMode = MySQL
    ElseIf InStr(UCase(m_DatabaseConnectionString), "SAP DB") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = MaxDB
    ElseIf InStr(UCase(m_DatabaseConnectionString), ".MDB") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = MSAccess
    End If
    m_theAdoDB.Open m_DatabaseConnectionString
    
    If m_SQLMode = MaxDB Then
      MsgBox "MaxDB n'est pas supporté à ce jour !", vbCritical
    End If
    
    m_Connected = True
    Connect = True
  Exit Function

ConnectErr:
  Call RaiseError(MyUnhandledError, "DataAccess:Connect Method")
End Function

Public Sub SetDatabase(dta As Object)
  ' pour DAO
  'dta.DatabaseName = DatabaseFileName
  'dta.Connect = "Access 2000;"
  
  If UCase(TypeName(dta)) <> "ADODC" Then
    MsgBox "Type de datacontrol invalid " & dta.Name, vbCritical
    Exit Sub
  End If
  
  Select Case m_SQLMode
    Case SQLServer
      dta.RecordSource = "SELECT 1"
     
    Case MySQL
      dta.RecordSource = "SELECT 1"
      
    Case MaxDB
      dta.RecordSource = "SELECT 1 FROM DUAL"
    
    Case MSAccess
      dta.RecordSource = "SELECT * FROM CATR9"
  
    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
  dta.ConnectionString = m_DatabaseConnectionString
End Sub

Public Function ValidateSQL(sSQL As String) As String
  Select Case m_SQLMode
    Case SQLServer
      ValidateSQL = sSQL
     
    Case MySQL
      ValidateSQL = Replace(sSQL, "[", "'")
      ValidateSQL = Replace(ValidateSQL, "]", "'")
      ValidateSQL = Replace(ValidateSQL, "P3IUser.", "")
      
    Case MaxDB
      ValidateSQL = Replace(sSQL, "[", """")
      ValidateSQL = Replace(ValidateSQL, "]", """")
      ValidateSQL = Replace(ValidateSQL, "P3IUser.", "")
      'If InStr(sSQL, " JOIN ") <> 0 Then
      '  Err.Raise -1, , "ValidateSQL() : Synthaxe JOIN non supporté par MaxDB"
      'End If
    
    Case MSAccess
      ValidateSQL = Replace(sSQL, "P3IUser.", "")
  
    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
End Function

