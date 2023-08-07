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
  Disconnected
  StoredProcedure
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
Public Sub BeginTrans(Optional tName As String = "")
  On Error GoTo BeginTransErr

  m_theAdoDB.BeginTrans
'  Execute "BEGIN TRANSACTION " '& tName
  
  Exit Sub
BeginTransErr:
  MsgBox "Erreur dans DataAccess:BeginTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=3F85538902EE
Public Sub CommitTrans(Optional tName As String = "")
  On Error GoTo CommitTransErr

  m_theAdoDB.CommitTrans
'  Execute "COMMIT TRANSACTION " '& tName
  
  Exit Sub
CommitTransErr:
  MsgBox "Erreur dans DataAccess:CommitTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=3F85539200CB
Public Sub RollbackTrans(Optional tName As String = "")
  On Error GoTo RollbackTransErr

  m_theAdoDB.RollbackTrans
'  Execute "ROLLBACK TRANSACTION " ' & tName

  Exit Sub
RollbackTransErr:
  MsgBox "Erreur dans DataAccess:RollbackTrans n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'return the resulting recordset of a SQL query.
'##ModelId=3F854D9801B5
Public Function OpenRecordset(SQLQuery As String, recorsetType As eRecordsetMode) As ADODB.Recordset
'  On Error GoTo OpenRecordsetErr
    
  ' la requete ne doit pas etre vide
  If Trim$(SQLQuery) = "" Then
    Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset : requète invalide")
    Exit Function
  End If
  
  ' convertie les " en '
  Dim n As Long
  Do
    n = InStr(SQLQuery, """")
    If n <> 0 Then
      SQLQuery = Left$(SQLQuery, n - 1) & "'" & Mid$(SQLQuery, n + 1)
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
      
    Case Disconnected
      OpenRecordset.CursorLocation = adUseClient
      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenStatic, adLockBatchOptimistic
      Set OpenRecordset.ActiveConnection = Nothing
      
    Case Table
'      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenStatic, adLockOptimistic, adCmdTableDirect
      OpenRecordset.Open SQLQuery, m_theAdoDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
    Case StoredProcedure
      OpenRecordset.Open SQLQuery, m_theAdoDB ' doit contenir SET NOCOUNT ON
    
    Case Else
      Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset : type de recordset invalide")
      Exit Function
  End Select

        
  Exit Function
'OpenRecordsetErr:
'  Call RaiseError(MyUnhandledError, "DataAccess::OpenRecordset Method")
End Function

'Create a DataHelper instance for this DataAccess instance
'##ModelId=3F854948003E
Public Function CreateHelper() As DataHelper
  On Error GoTo CreateHelperErr

  Set CreateHelper = New DataHelper
  Set CreateHelper.DataSource = Me
  
  If InStr(UCase$(m_DatabaseConnectionString), "SQLOLEDB") <> 0 _
     Or InStr(UCase$(m_DatabaseConnectionString), "SQLNCLI") <> 0 Then
    CreateHelper.SqlMode = SQLServer
  ElseIf InStr(UCase$(m_DatabaseConnectionString), "MYSQL") <> 0 Then
    CreateHelper.SqlMode = MySQL
  ElseIf InStr(UCase$(m_DatabaseConnectionString), ".MDB") <> 0 Then
    CreateHelper.SqlMode = MSAccess
  ElseIf InStr(UCase$(m_DatabaseConnectionString), "SAP DB") <> 0 Then
    CreateHelper.SqlMode = MaxDB
  ElseIf InStr(UCase$(m_DatabaseConnectionString), "ORACLE") <> 0 _
        Or InStr(UCase$(m_DatabaseConnectionString), "MSDAORA") <> 0 Then
    CreateHelper.SqlMode = Oracle
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
  MsgBox "Erreur dans DataAccess:Disconnect n°" & Err & vbLf & Err.Description, vbCritical
End Sub

'##ModelId=3F851D8A0128
Private Sub Class_Terminate()
  If m_Connected = True Then
    Disconnect
  End If
End Sub

'##ModelId=3F851D8A0119
Private Sub Class_Initialize()
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
    
    If InStr(UCase$(m_DatabaseConnectionString), "SQLOLEDB") <> 0 _
       Or InStr(UCase$(m_DatabaseConnectionString), "SQLNCLI") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_theAdoDB.IsolationLevel = adXactIsolated
      m_SQLMode = SQLServer
    ElseIf InStr(UCase$(m_DatabaseConnectionString), "MYSQL") <> 0 Then
      m_theAdoDB.CursorLocation = adUseClient
      m_SQLMode = MySQL
    ElseIf InStr(UCase$(m_DatabaseConnectionString), "SAP DB") <> 0 Then
      m_theAdoDB.CursorLocation = adUseClient
      m_SQLMode = MaxDB
    ElseIf InStr(UCase$(m_DatabaseConnectionString), ".MDB") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = MSAccess
    ElseIf InStr(UCase$(m_DatabaseConnectionString), ".XLS") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = MSAccess
    ElseIf InStr(UCase$(m_DatabaseConnectionString), "ORACLE") <> 0 _
        Or InStr(UCase$(m_DatabaseConnectionString), "MSDAORA") <> 0 Then
      m_theAdoDB.CursorLocation = adUseServer
      m_SQLMode = Oracle
    End If
    
    m_theAdoDB.CommandTimeout = 900
    
    m_theAdoDB.Open m_DatabaseConnectionString
    
    If m_SQLMode = MaxDB Then
      MsgBox "MaxDB n'est pas supporté à ce jour !", vbCritical
    End If
    
    m_Connected = True
    Connect = True
  Exit Function

ConnectErr:
  MsgBox "Erreur dans DataAccess:Connect n°" & Err & vbLf & Err.Description, vbCritical
End Function

Public Sub SetDatabase(dta As Object)
  ' pour DAO
  'dta.DatabaseName = DatabaseFileName
  'dta.Connect = "Access 2000;"
  
  If UCase$(TypeName(dta)) <> "ADODC" Then
    MsgBox "Type de datacontrol invalid " & dta.Name, vbCritical
    Exit Sub
  End If
  
  Dim theData As MSAdodcLib.Adodc
  
  Set theData = dta
  
  Select Case m_SQLMode
    Case SQLServer
      theData.RecordSource = "SELECT 1"
     
    Case MySQL
      theData.RecordSource = "SELECT 1"
      
    Case MaxDB
      theData.RecordSource = "SELECT 1 FROM DUAL"
    
    Case Oracle
      theData.RecordSource = "SELECT 1 FROM DUAL"
    
    Case MSAccess
      theData.RecordSource = "SELECT 1" '"SELECT * FROM MSysObjects"
  
    Case Else
      Err.Raise -1, , "SQLMode Inconnu"
  End Select
  
  theData.ConnectionString = m_DatabaseConnectionString
  theData.ConnectionTimeout = 60
  theData.CommandTimeout = 900
End Sub

Public Function ValidateSQL(sSQL As String) As String
  Select Case m_SQLMode
    Case SQLServer
      ValidateSQL = Replace(sSQL, "P3IUser.", "")
     
    Case MySQL
      ValidateSQL = Replace(sSQL, "[", "'")
      ValidateSQL = Replace(ValidateSQL, "]", "'")
      ValidateSQL = Replace(ValidateSQL, "P3IUser.", "")
      ValidateSQL = Replace(ValidateSQL, "GetDate()", "Now()")
      
    Case MaxDB, Oracle
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

Public Function CreateView(Nom As String, SQLQuery As String) As Boolean
  On Error Resume Next

  CreateView = True
  
  If m_SQLMode <> MSAccess Then
    Err = 0
    Me.Execute "DROP VIEW " & Nom
    If Err <> 0 And Err <> -2147217865 Then
      MsgBox "Erreur lors de la suppression de la vue " & Nom & vbLf & Err.Description, vbCritical
    End If

    Err = 0
    Me.Execute "CREATE VIEW " & Nom & " AS " & SQLQuery
    If Err <> 0 Then
      MsgBox "Erreur lors de la création de la vue " & Nom & vbLf & Err.Description, vbCritical
      CreateView = False
    End If
  Else
    Dim oCat As New ADOX.Catalog
    Dim oView As ADOX.View
    Dim cmd As New ADODB.Command
    
    oCat.ActiveConnection = Me.Connection
    
    ' delete existing view
    oCat.Views.Delete ValidateSQL(Nom)
    
    ' Create the command representing the view.
    Err = 0
    cmd.CommandText = Me.ValidateSQL(SQLQuery)
    If Err <> 0 Then
      MsgBox "Erreur lors de la création de la vue " & Nom & vbLf & Err.Description, vbCritical
      CreateView = False
    End If
    
    ' Create the new View
    Err = 0
    oCat.Views.Append ValidateSQL(Nom), cmd
    If Err <> 0 Then
      MsgBox "Erreur lors de la création de la vue " & Nom & vbLf & Err.Description, vbCritical
      CreateView = False
    End If
    
    'Clean up
    Set oCat.ActiveConnection = Nothing
    Set oCat = Nothing
    Set cmd = Nothing
  End If
End Function
