VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArchiveRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A7603035E"
'##ModelId=5C8A760400ED
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Declare Function LookupAccountName Lib "advapi32.dll" Alias "LookupAccountNameA" (lpSystemName As String, ByVal lpAccountName As String, sid As Any, cbSid As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long

'##ModelId=5C8A76040070
Dim oldNewAssureRecno() As String

'##ModelId=5C8A7604012B
Public Function GetWinUser() As String

  Dim sDomainName As String * 255
  Dim lDomainNameLength As Long
  Dim sUserName As String
  Dim bUserSid(255) As Byte
  Dim lSIDType As Long
  
  sUserName = String(100, Chr$(0))
  GetUserName sUserName, 100
  GetWinUser = Left$(sUserName, InStr(sUserName, Chr$(0)) - 1)

End Function

'Complete Archive Procedure
'##ModelId=5C8A7604012C
Public Function ArchivePeriode(connStrProd As String, connStrArch As String, numPeriode As Integer, statusMessage As String) As Boolean
    
    ArchivePeriode = False
    
    Dim Start As Single
    Dim operationRes As OperationStatus
    Dim rsCount As Recordset
    Dim recCount As Long
    Dim sqlStrCount As String
        
    On Error GoTo DBConnectionError

    Dim cnxnProd As ADODB.Connection
    Dim cnxnArch As ADODB.Connection

    Set cnxnProd = New ADODB.Connection
    Set cnxnArch = New ADODB.Connection
    
    cnxnProd.CommandTimeout = 900
    cnxnArch.CommandTimeout = 900

    cnxnProd.CursorLocation = adUseClient
    cnxnArch.CursorLocation = adUseClient
    
    cnxnProd.Open connStrProd
    cnxnArch.Open connStrArch
    
    If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then
        
        On Error GoTo TransError
        
        cnxnProd.BeginTrans
        cnxnArch.BeginTrans
        
'### testing
'GoTo test
        
        '1: Update the Status and StatusDate fields in the Prod DB
        UpdatePeriodeStatus cnxnProd, numPeriode, "Archivée"
        'UpdatePeriodeStatus numPeriode, "Archivée"
        DoEvents
        
        '2: Insert or update Periode in Archives DB
        InsertOrUpdatePeriode cnxnProd, cnxnArch, numPeriode
    
        DoEvents
        
        '3: Synchronize all reference tables in Archives DB - except Assure Tables
        'Delays: Sync_TBQREGA: 1.461s -- Sync_Correspondance_CatOption: 5.633s -- Sync_CoeffAmortissement: 0.969s
        
        'Start = Timer
        Sync_AgeDepartRetraite cnxnProd, cnxnArch, numPeriode
        'Debug.Print "Sync_AgeDepartRetraite: " & Format(Timer - Start, "0.000")
        DoEvents
        
        Sync_AgeDepartRetraiteInval1 cnxnProd, cnxnArch, numPeriode
        DoEvents
        'Start = Timer
        Sync_Capitaux_Moyens cnxnProd, cnxnArch, numPeriode
        'Debug.Print "Sync_Capitaux_Moyens: " & Format(Timer - Start, "0.000")
        DoEvents
        Sync_CatR9 cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_CATR9INVAL cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_CDSITUAT cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_CODECATINV cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_CodesCat cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_CoeffAmortissement cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_Correspondance_CatOption cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_DonneesSociales cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_ParamCalcul cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_ParamRentes cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_PassageNCA cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_PM_Retenue cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_ProvisionsOuverture cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_Reassurance cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_TBQREGA cnxnProd, cnxnArch, numPeriode
        DoEvents
        
        'Start = Timer
        Sync_Assure_P3ILOGTRAIT cnxnProd, cnxnArch, numPeriode
        'Debug.Print "Sync_Assure_P3ILOGTRAIT: " & Format(Timer - Start, "0.000")
        DoEvents
        
        Sync_TypeCalcul cnxnProd, cnxnArch, numPeriode
        DoEvents
        Sync_TypePeriode cnxnProd, cnxnArch, numPeriode
        DoEvents
        
        
        '4: Bulk Insert all Assure data from Prod DB into Archives DB
        operationRes = CreateCSVFile(cnxnProd, numPeriode, "P3IUser.Assure", GetWinUser & "_Assure.csv", "POPERCLE")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnArch, numPeriode, "P3IUser.Assure", GetWinUser & "_Assure.csv", "POPERCLE"
            DoEvents
        End If
        
        'Get the new RECNO values from the Assure table in the Archive DB
        UpdateOldNewRecnoArray cnxnArch, numPeriode
        
        operationRes = CreateCSVFile(cnxnProd, numPeriode, "P3IUser.Assure_Retraite", GetWinUser & "_Assure_Retraite.csv", "POPERCLE")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnArch, numPeriode, "P3IUser.Assure_Retraite", GetWinUser & "_Assure_Retraite.csv", "POPERCLE"
            DoEvents
        End If
        
        operationRes = CreateCSVFile(cnxnProd, numPeriode, "P3IUser.Assure_P3IPROVCOLL", GetWinUser & "_Assure_P3IPROVCOLL.csv", "NumPeriode")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnArch, numPeriode, "P3IUser.Assure_P3IPROVCOLL", GetWinUser & "_Assure_P3IPROVCOLL.csv", "NumPeriode"
            DoEvents
        End If
        
        
        '5: delete assure for specific periode from Prod DB after Bulk Insert into Archive DB
        'sqlStr = "DELETE FROM P3IUser.Assure WHERE POPERCLE = " & NumPeriode
        'cnxnProd.Execute sqlStr, , adExecuteNoRecords
        
        
'### testing
'test:

If True Then
        
Delete_Assure:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure WHERE POPERCLE = " & numPeriode
        cnxnProd.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnProd.Execute("SELECT COUNT(*)FROM P3IUser.Assure where popercle = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure
            
        
        'sqlStr = "DELETE FROM P3IUser.Assure_Retraite WHERE POPERCLE = " & NumPeriode
        'cnxnProd.Execute sqlStr, , adExecuteNoRecords
Delete_Assure_Retraite:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure_Retraite WHERE POPERCLE = " & numPeriode
        cnxnProd.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnProd.Execute("SELECT COUNT(*)FROM P3IUser.Assure_Retraite where popercle = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées Retraité à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure_Retraite
        
        
        'sqlStr = "DELETE FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & NumPeriode
        'cnxnProd.Execute sqlStr, , adExecuteNoRecords
Delete_Assure_P3IPROVCOLL:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & numPeriode
        cnxnProd.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnProd.Execute("SELECT COUNT(*)FROM P3IUser.Assure_P3IPROVCOLL where NumPeriode = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées PROVCOLL à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure_P3IPROVCOLL
        
        
End If
        
        
        'everything worked - commit all updates to both DB's
        cnxnProd.CommitTrans
        cnxnArch.CommitTrans
        
    Else
        'handle connection problem
        MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: ArchiveRestore.ArchivePeriode", vbCritical
    End If
    
    ArchivePeriode = True
    
    
Cleanup:
    If Not cnxnProd Is Nothing Then
        If cnxnProd.State = adStateOpen Then
            cnxnProd.Close
        End If
    End If

    If Not cnxnArch Is Nothing Then
        If cnxnArch.State = adStateOpen Then
            cnxnArch.Close
        End If
    End If

    Set cnxnProd = Nothing
    Set cnxnArch = Nothing
    
    Exit Function
    
TransError:
    
    cnxnProd.RollbackTrans
    cnxnArch.RollbackTrans
    
    'Handle all ADO Errors - Select Case...
        
    'Raise Error: Periode could not be archived
    If Err <> 0 Then
      MsgBox "Erreur pendant l'archivage du période : " & numPeriode & vbLf & _
      "Erreur : " & Err.Number & vbLf & _
      "Source: " & Err.Source & vbLf & _
      "Description: " & Err.Description
    End If

    GoTo Cleanup


DBConnectionError:

    'Raise Error: DB connection cannot be established
    If Err <> 0 Then
      MsgBox "Impossible d'ouvrir la base de données !" & vbLf & _
      "Erreur : " & Err.Number & vbLf & _
      "Source: ArchiveRestore.ArchivePeriode" & vbLf & _
      "Description: " & Err.Description
    End If
        
    GoTo Cleanup


End Function

'Complete Restore Procedure
'##ModelId=5C8A76040179
Public Function RestorePeriode(connStrProd As String, connStrArch As String, numPeriode As Integer, statusMessage As String) As Boolean

    Dim Start As Single
    Dim cnxnProd As ADODB.Connection
    Dim cnxnArch As ADODB.Connection
    Dim operationRes As OperationStatus
    Dim rsCount As Recordset
    Dim recCount As Long
    Dim sqlStrCount As String
        
    Set cnxnProd = New ADODB.Connection
    Set cnxnArch = New ADODB.Connection
    
    cnxnProd.CommandTimeout = 900
    cnxnArch.CommandTimeout = 900
    
    RestorePeriode = False
    
    On Error GoTo DBConnectionError
    
    cnxnProd.CursorLocation = adUseClient
    cnxnArch.CursorLocation = adUseClient
    
    cnxnProd.Open connStrProd
    cnxnArch.Open connStrArch
    
    If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then
        
        On Error GoTo TransError
        
        cnxnProd.BeginTrans
        cnxnArch.BeginTrans
        
        '1: Verify if the periode has been deleted from the Production DB
        If (PeriodeHasBeenDeleted(cnxnProd, numPeriode)) Then
          
          'insert the missing periode
          InsertOrUpdatePeriode cnxnArch, cnxnProd, numPeriode
          
          'restore all records for related tables
          Sync_AgeDepartRetraite cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          'restore all records for related tables
          Sync_AgeDepartRetraiteInval1 cnxnArch, cnxnProd, numPeriode
          DoEvents
                    
          Sync_Capitaux_Moyens cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CatR9 cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CATR9INVAL cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CDSITUAT cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CODECATINV cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CodesCat cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_CoeffAmortissement cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_Correspondance_CatOption cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_DonneesSociales cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_ParamCalcul cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_ParamRentes cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_PassageNCA cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_PM_Retenue cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_ProvisionsOuverture cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_Reassurance cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_TBQREGA cnxnArch, cnxnProd, numPeriode
          DoEvents
          
          Sync_Assure_P3IPROVCOLL cnxnArch, cnxnProd, numPeriode
          DoEvents
          Sync_Assure_P3ILOGTRAIT cnxnArch, cnxnProd, numPeriode
          DoEvents
          
        Else
        
          'Update Periode if it already exists
          InsertOrUpdatePeriode cnxnArch, cnxnProd, numPeriode
          DoEvents
                  
        End If
        
        '2: Update the Status and StatusDate fields in both DB's
        UpdatePeriodeStatus cnxnProd, numPeriode, "Restaurée"
        'not required - periode will be deleted in Archive
        'UpdatePeriodeStatus cnxnArch, numPeriode, "Restaurée"
        DoEvents
        
        
        '4: Bulk Insert all Assure data
        operationRes = CreateCSVFile(cnxnArch, numPeriode, "P3IUser.Assure", GetWinUser & "_Assure.csv", "POPERCLE")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnProd, numPeriode, "P3IUser.Assure", GetWinUser & "_Assure.csv", "POPERCLE"
            DoEvents
        End If
        
        'Get the new RECNO values from the Assure table in the Archive DB
        UpdateOldNewRecnoArray cnxnProd, numPeriode
        
        operationRes = CreateCSVFile(cnxnArch, numPeriode, "P3IUser.Assure_Retraite", GetWinUser & "_Assure_Retraite.csv", "POPERCLE")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnProd, numPeriode, "P3IUser.Assure_Retraite", GetWinUser & "_Assure_Retraite.csv", "POPERCLE"
            DoEvents
        End If
        
        operationRes = CreateCSVFile(cnxnArch, numPeriode, "P3IUser.Assure_P3IPROVCOLL", GetWinUser & "_Assure_P3IPROVCOLL.csv", "NumPeriode")
        DoEvents
        If operationRes = eSuccess Then
            BulkInsert cnxnProd, numPeriode, "P3IUser.Assure_P3IPROVCOLL", GetWinUser & "_Assure_P3IPROVCOLL.csv", "NumPeriode"
            DoEvents
        End If
        
        
        '5: delete assure for specific periode from Archive DB after Bulk Insert into Archive DB
        'sqlStr = "DELETE FROM P3IUser.Assure WHERE POPERCLE = " & NumPeriode
        'cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
        'sqlStr = "DELETE FROM P3IUser.Assure_Retraite WHERE POPERCLE = " & NumPeriode
        'cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
        'sqlStr = "DELETE FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & NumPeriode
        'cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
Delete_Assure:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure WHERE POPERCLE = " & numPeriode
        cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnArch.Execute("SELECT COUNT(*)FROM P3IUser.Assure where popercle = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure
        
        
Delete_Assure_Retraite:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure_Retraite WHERE POPERCLE = " & numPeriode
        cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnArch.Execute("SELECT COUNT(*)FROM P3IUser.Assure_Retraite where popercle = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure_Retraite
        
        
Delete_Assure_P3IPROVCOLL:
        sqlStr = "DELETE TOP (500) FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & numPeriode
        cnxnArch.Execute sqlStr, , adExecuteNoRecords
        
        Set rsCount = cnxnArch.Execute("SELECT COUNT(*)FROM P3IUser.Assure_P3IPROVCOLL where NumPeriode = " & numPeriode)
        res = rsCount(0).Value
        
        DoEvents
        frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        DoEvents
        
        If res > 0 Then GoTo Delete_Assure_P3IPROVCOLL
        
        
        '6: delete the periode and all data in linked reference tables from Archive
        DeletePeriodeAndReferenceData cnxnArch, numPeriode
        
        
        'everything worked - commit all updates in both DB's
        cnxnProd.CommitTrans
        cnxnArch.CommitTrans
        
        RestorePeriode = True
        
    Else
        'handle connection problem
        MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: ArchiveRestore.RestorePeriode"
    End If
    
    
Cleanup:
    If Not cnxnProd Is Nothing Then
        If cnxnProd.State = adStateOpen Then
            cnxnProd.Close
        End If
    End If
    
    If Not cnxnArch Is Nothing Then
        If cnxnArch.State = adStateOpen Then
            cnxnArch.Close
        End If
    End If
    
    Set cnxnProd = Nothing
    Set cnxnArch = Nothing
    
    Exit Function
    
TransError:
    
    cnxnProd.RollbackTrans
    cnxnArch.RollbackTrans
    
    'Handle all ADO Errors - Select Case...
        
    'Raise Error: Periode could not be restored
    If Err <> 0 Then
      MsgBox "Erreur pendant la restauration du période : " & numPeriode & vbLf & _
      "Erreur : " & Err.Number & vbLf & _
      "Source: " & Err.Source & vbLf & _
      "Description: " & Err.Description
    End If
    
    GoTo Cleanup


DBConnectionError:

    'Raise Error: DB connection cannot be established
    If Err <> 0 Then
      MsgBox "Impossible d'ouvrir la base de données !" & vbLf & _
      "Erreur : " & Err.Number & vbLf & _
      "Source: ArchiveRestore.RestorePeriode" & vbLf & _
      "Description: " & Err.Description
    End If
        
    GoTo Cleanup


End Function

'delete periode and all corresponding data from linked reference tables
'##ModelId=5C8A76040277
Public Sub DeletePeriodeAndReferenceData(cnxn As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
        
    'all ADO errors are handled by the calling procedure
    If (cnxn.State = adStateOpen) Then
               
        cnxn.Execute "DELETE FROM P3IUser.TBQREGA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CodesCat WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CodeCatInv WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.Correspondance_CatOption WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.Capitaux_Moyens WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CATR9 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CATR9INVAL WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.PassageNCA WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.PM_Retenue WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.Reassurance WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CDSITUAT WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.ParamRentes WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.AgeDepartRetraite WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.AgeDepartRetraiteInval1 WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.CoeffAmortissement WHERE GroupeCle = " & GroupeCle & " AND NumPeriode = " & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.Donneessociales WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.ProvisionsOuverture WHERE GPECLE=" & GroupeCle & " AND NUMCLE=" & numPeriode, , adExecuteNoRecords

        cnxn.Execute "DELETE FROM P3IUser.Assure_P3ILOGTRAIT WHERE CleGroupe=" & GroupeCle & " AND NumPeriode=" & numPeriode, , adExecuteNoRecords
        cnxn.Execute "DELETE FROM P3IUser.ParamCalcul WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode, , adExecuteNoRecords
        
        cnxn.Execute "DELETE FROM P3IUser.Periode WHERE PEGPECLE = " & GroupeCle & " AND PENUMCLE = " & numPeriode, , adExecuteNoRecords
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise -1, "ArchiveRestore.DeletePeriodeAndReferenceData", "Impossible d'ouvrir la base de données"
    End If
  
End Sub

'copy records from one DB to the same table in another DB
'##ModelId=5C8A76040324
Public Sub CopyRecordsBetweenDatabases(rsCopyFrom As ADODB.Recordset, rsCopyTo As ADODB.Recordset)
     
    Dim objFields As ADODB.fields
    Dim i As Integer
    
    Set objFields = rsCopyFrom.fields
    
    Do While Not rsCopyFrom.EOF
            
        With rsCopyTo
            .AddNew
            
            For i = 0 To (objFields.Count - 1)
                .fields(i).Value = rsCopyFrom.fields(i).Value
            Next
            
            .Update
            rsCopyFrom.MoveNext
        End With
    Loop

End Sub

'copy all fields from a record in one DB to the same record in another DB
'##ModelId=5C8A76040371
Public Sub CopyAllFieldsOfRecord(rsCopyFrom As ADODB.Recordset, rsCopyTo As ADODB.Recordset)
     
    Dim objFields As ADODB.fields
    Dim i As Integer
    
    Set objFields = rsCopyFrom.fields
                
    With rsCopyTo
        
        For i = 0 To (objFields.Count - 1)
            .fields(i).Value = rsCopyFrom.fields(i).Value
        Next
        
    End With

End Sub

'Verify if the periode with the specified number has been deleted from the DB
'##ModelId=5C8A760403A0
Public Function PeriodeHasBeenDeleted(cnxn As ADODB.Connection, numPeriode As Integer) As Boolean

  Dim ret As Boolean
  Dim sqlStr As String
  Dim rs As ADODB.Recordset
    
  ret = False
  
  Set rs = New ADODB.Recordset
  sqlStr = "SELECT * FROM P3IUser.Periode WHERE PENUMCLE = " & numPeriode
  Set rs = cnxn.Execute(sqlStr)
  
  If rs.RecordCount < 1 Then
    ret = True
  End If

  PeriodeHasBeenDeleted = ret
  
End Function

'Update the Status and Archiving Date in the Periode table
'##ModelId=5C8A760403DE
Public Sub UpdatePeriodeStatus(cnxnSource As ADODB.Connection, numPeriode As Integer, status As String)
'Public Sub UpdatePeriodeStatus(numPeriode As Integer, status As String)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
        
    'all ADO errors are handled by the calling procedure
    If (cnxnSource.State = adStateOpen) Then
    'If (m_dataSource.Connection.State = adStateOpen) Then
        
        'get data from Production DB
        Set rsSource = New ADODB.Recordset
        
        sqlStr = "SELECT * FROM P3IUser.Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        'Set rsSource = m_dataSource.OpenRecordset(sqlStr, Snapshot)
        
        If rsSource.RecordCount > 0 Then
            rsSource.MoveFirst
            
            rsSource("StatutArchRest") = status ' "Archivée"
            rsSource("DateArchRest") = Format(Now, "DD/MM/YYYY")
                
            rsSource.Update
        Else
            'no record was found in the Prod DB - raise an error to the calling function
            Err.Raise -1, "ArchiveRestore.UpdatePeriodeStatus", "La période " & numPeriode & " n'existe pas dans la base des données"
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise -1, "ArchiveRestore.UpdatePeriodeStatus", "Impossible d'ouvrir la base de données"
    End If
    
End Sub

'If the requested Periode already exists, update all fields, otherwise create
'the Periode
'##ModelId=5C8A76050035
Public Sub InsertOrUpdatePeriode(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    'all ADO errors are handled by the calling procedure
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        'get data from Production DB
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        sqlStr = "SELECT * FROM P3IUser.Periode WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            rsSource.MoveFirst
            
            If rsDest.RecordCount > 0 Then
                'the periode already exists - do an Update
                rsDest.MoveFirst
                
                rsDest("PEDATEDEB") = rsSource("PEDATEDEB")
                rsDest("PEDATEFIN") = rsSource("PEDATEFIN")
                rsDest("PEDATEEXT") = rsSource("PEDATEEXT")
                rsDest("PECOMMENTAIRE") = rsSource("PECOMMENTAIRE")
                rsDest("PELOCKED") = rsSource("PELOCKED")
                rsDest("PENBJOURMAX") = rsSource("PENBJOURMAX")
                rsDest("PENBJOURDC") = rsSource("PENBJOURDC")
                rsDest("NUTRAITP3I") = rsSource("NUTRAITP3I")
                rsDest("DateDernierControle") = rsSource("DateDernierControle")
                rsDest("DateDernierImport") = rsSource("DateDernierImport")
                rsDest("IdTypeCalcul") = rsSource("IdTypeCalcul")
                rsDest("PETYPEPERIODE") = rsSource("PETYPEPERIODE")
                rsDest("PEP3I_INDIVIDUEL") = rsSource("PEP3I_INDIVIDUEL")
                rsDest("PECALCULER") = rsSource("PECALCULER")
                rsDest("StatutArchRest") = rsSource("StatutArchRest")
                rsDest("DateArchRest") = rsSource("DateArchRest")
                
                rsDest.Update
            Else
                'the periode does not exist - do an Insert
                rsDest.AddNew
                
                rsDest("PEGPECLE") = rsSource("PEGPECLE")
                rsDest("PENUMCLE") = rsSource("PENUMCLE")
                rsDest("PEDATEDEB") = rsSource("PEDATEDEB")
                rsDest("PEDATEFIN") = rsSource("PEDATEFIN")
                rsDest("PEDATEEXT") = rsSource("PEDATEEXT")
                rsDest("PECOMMENTAIRE") = rsSource("PECOMMENTAIRE")
                rsDest("PELOCKED") = rsSource("PELOCKED")
                rsDest("PENBJOURMAX") = rsSource("PENBJOURMAX")
                rsDest("PENBJOURDC") = rsSource("PENBJOURDC")
                rsDest("NUTRAITP3I") = rsSource("NUTRAITP3I")
                rsDest("DateDernierControle") = rsSource("DateDernierControle")
                rsDest("DateDernierImport") = rsSource("DateDernierImport")
                rsDest("IdTypeCalcul") = rsSource("IdTypeCalcul")
                rsDest("PETYPEPERIODE") = rsSource("PETYPEPERIODE")
                rsDest("PEP3I_INDIVIDUEL") = rsSource("PEP3I_INDIVIDUEL")
                rsDest("PECALCULER") = rsSource("PECALCULER")
                rsDest("StatutArchRest") = rsSource("StatutArchRest")
                rsDest("DateArchRest") = rsSource("DateArchRest")
        
                rsDest.Update
            End If
        
        Else
            'no record was found in the Prod DB - raise an error to the calling function
            Err.Raise -1, "ArchiveRestore.InsertOrUpdatePeriode", "La période " & numPeriode & " n'existe pas dans la base des données"
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise -1, "ArchiveRestore.InsertOrUpdatePeriode", "Impossible d'ouvrir la base de données"
    End If
    
End Sub

'****************************************************************
'********* Functions: Synchronize Reference Tables **************
'****************************************************************

'##ModelId=5C8A76050083
Public Sub Sync_CatR9(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CATR9 WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CATR9 WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
          Do While Not rsSource.EOF
              rsDest.AddNew
              
              CopyAllFieldsOfRecord rsSource, rsDest

              rsDest.Update
              rsSource.MoveNext
          Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CatR9", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A7605011F
Public Sub Sync_AgeDepartRetraite(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.AgeDepartRetraite WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecord
        
        sqlStr = "SELECT * FROM P3IUser.AgeDepartRetraite WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_AgeDepartRetraite", "Impossible d'ouvrir la base de données"
    End If

End Sub
Public Sub Sync_AgeDepartRetraiteInval1(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.AgeDepartRetraiteInval1 WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecord
        
        sqlStr = "SELECT * FROM P3IUser.AgeDepartRetraiteInval1 WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_AgeDepartRetraiteInval1", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A7605015E
Public Sub Sync_Capitaux_Moyens(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.Capitaux_Moyens WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.Capitaux_Moyens WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Capitaux_Moyens", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760501AC
Public Sub Sync_CATR9INVAL(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CATR9INVAL WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CATR9INVAL WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
                        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CATR9INVAL", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760501FA
Public Sub Sync_CDSITUAT(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CDSITUAT WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CDSITUAT WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CDSITUAT", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76050238
Public Sub Sync_CODECATINV(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CODECATINV WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CODECATINV WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CODECATINV", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76050279
Public Sub Sync_CodesCat(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CodesCat WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CodesCat WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CodesCat", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760502D4
Public Sub Sync_CoeffAmortissement(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.CoeffAmortissement WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.CoeffAmortissement WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_CoeffAmortissement", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A7605038F
Public Sub Sync_Correspondance_CatOption(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.Correspondance_CatOption WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.Correspondance_CatOption WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        'rsDest.Open "P3IUser.Correspondance_CatOption", cnxnDest, adOpenKeyset, adLockOptimistic, adCmdTable
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
          
          Do While Not rsSource.EOF
              rsDest.AddNew
              
              CopyAllFieldsOfRecord rsSource, rsDest
      
              rsDest.Update
              rsSource.MoveNext
          Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Correspondance_CatOption", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76060005
Public Sub Sync_DonneesSociales(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.DonneesSociales WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.DonneesSociales WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
                
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_DonneesSociales", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76060053
Public Sub Sync_ParamCalcul(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.ParamCalcul WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.ParamCalcul WHERE PENUMCLE = " & numPeriode & " AND PEGPECLE = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                rsDest("PEGPECLE") = rsSource("PEGPECLE")
                rsDest("PENUMCLE") = rsSource("PENUMCLE")
                
                rsDest("PENUMPARAMCALCUL") = rsSource("PENUMPARAMCALCUL")
                rsDest("PENOMPARAM") = rsSource("PENOMPARAM")
                rsDest("PECOMMENTAIREPARAM") = rsSource("PECOMMENTAIREPARAM")
                rsDest("PELMINCAP") = rsSource("PELMINCAP")
                rsDest("PELMPASSAGE") = rsSource("PELMPASSAGE")
                rsDest("PETXINCAP") = rsSource("PETXINCAP")
                rsDest("PEGESTIONINCAP") = rsSource("PEGESTIONINCAP")
                rsDest("PELMINVAL") = rsSource("PELMINVAL")
                rsDest("PETXINVAL") = rsSource("PETXINVAL")
                rsDest("PEGESTIONINVAL") = rsSource("PEGESTIONINVAL")
                
                rsDest("PEGESTIONC7") = rsSource("PEGESTIONC7")
                rsDest("PETXREVENUC7") = rsSource("PETXREVENUC7")
                rsDest("PEAGERETRAITE") = rsSource("PEAGERETRAITE")
                rsDest("PEDUREEINCAP") = rsSource("PEDUREEINCAP")
                rsDest("PEFRANCHISE") = rsSource("PEFRANCHISE")
                rsDest("PETXREVALO") = rsSource("PETXREVALO")
                rsDest("PEDUREEREVALO") = rsSource("PEDUREEREVALO")
                rsDest("PETMO") = rsSource("PETMO")
                rsDest("PEDUREELISSAGE") = rsSource("PEDUREELISSAGE")
                rsDest("PETRENTECONJOINT") = rsSource("PETRENTECONJOINT")
                
                rsDest("PETXRTCONJOINT") = rsSource("PETXRTCONJOINT")
                rsDest("PEGESTIONCONJOINT") = rsSource("PEGESTIONCONJOINT")
                rsDest("PEPAIEMENTCONJOINT") = rsSource("PEPAIEMENTCONJOINT")
                rsDest("PEFRACTIONCONJOINT") = rsSource("PEFRACTIONCONJOINT")
                rsDest("PETRENTEEDUC") = rsSource("PETRENTEEDUC")
                rsDest("PETXRTEDUC") = rsSource("PETXRTEDUC")
                rsDest("PEGESTIONEDUC") = rsSource("PEGESTIONEDUC")
                rsDest("PEPAIEMENTEDUC") = rsSource("PEPAIEMENTEDUC")
                rsDest("PEFRACTIONEDUC") = rsSource("PEFRACTIONEDUC")
                rsDest("PERecalculBCAC") = rsSource("PERecalculBCAC")
                
                rsDest("PETauxTechRecalculBCAC_MDC") = rsSource("PETauxTechRecalculBCAC_MDC")
                rsDest("PEGestionRecalculBCAC_MDC") = rsSource("PEGestionRecalculBCAC_MDC")
                rsDest("PECleTableMortBCACInval_MDC") = rsSource("PECleTableMortBCACInval_MDC")
                rsDest("PECleTableMortBCACIncap_MDC") = rsSource("PECleTableMortBCACIncap_MDC")
                rsDest("PECleTableBCACInval_MDC") = rsSource("PECleTableBCACInval_MDC")
                rsDest("PECleTableBCACIncap_MDC") = rsSource("PECleTableBCACIncap_MDC")
                rsDest("PEGestionDeces_MDC") = rsSource("PEGestionDeces_MDC")
                rsDest("PEGestionEduc_MDC") = rsSource("PEGestionEduc_MDC")
                rsDest("PEGestionConjoint_MDC") = rsSource("PEGestionConjoint_MDC")
                rsDest("PEAgeConjoint_MDC") = rsSource("PEAgeConjoint_MDC")
                
                rsDest("PECapConstitMoyenTempo_MDC") = rsSource("PECapConstitMoyenTempo_MDC")
                rsDest("PECapConstitMoyenViagere_MDC") = rsSource("PECapConstitMoyenViagere_MDC")
                rsDest("PEForcerCapConstitMoyen_MDC") = rsSource("PEForcerCapConstitMoyen_MDC")
                rsDest("PELissageProvision") = rsSource("PELissageProvision")
                rsDest("PEAnnualisationDernierFlux") = rsSource("PEAnnualisationDernierFlux")
                rsDest("PEPMGDForcerInval") = rsSource("PEPMGDForcerInval")
                rsDest("PEMethodeCalcul") = rsSource("PEMethodeCalcul")
                rsDest("PEPCTACCIDENT") = rsSource("PEPCTACCIDENT")
                rsDest("PEAgeLimiteCalul_MDC") = rsSource("PEAgeLimiteCalul_MDC")
                rsDest("PEAGE_ANNIVERSAIRE") = rsSource("PEAGE_ANNIVERSAIRE")
                
                rsDest("PEINTERPOLATION_INVAL") = rsSource("PEINTERPOLATION_INVAL")
                rsDest("PESANSLISSAGE_PASSAGE") = rsSource("PESANSLISSAGE_PASSAGE")
                rsDest("PEBRIDAGE_AGE") = rsSource("PEBRIDAGE_AGE")
                rsDest("PEPctPMCalculeeDC") = rsSource("PEPctPMCalculeeDC")
                rsDest("PEPortefeuilleSalaries") = rsSource("PEPortefeuilleSalaries")
                rsDest("PEAgeLimiteInvalCat1") = rsSource("PEAgeLimiteInvalCat1")
                rsDest("PEAgeLimiteInvalCat1_Retraite") = rsSource("PEAgeLimiteInvalCat1_Retraite")
                rsDest("PEAgeLimiteCalul_MDC_Retraite") = rsSource("PEAgeLimiteCalul_MDC_Retraite")
                rsDest("PECleTableBCACIncap_MDC_Retraite") = rsSource("PECleTableBCACIncap_MDC_Retraite")
                rsDest("PECleTableBCACInval_MDC_Retraite") = rsSource("PECleTableBCACInval_MDC_Retraite")
                
                rsDest("PECleTableMortBCACIncap_MDC_Retraite") = rsSource("PECleTableMortBCACIncap_MDC_Retraite")
                rsDest("PECleTableMortBCACInval_MDC_Retraite") = rsSource("PECleTableMortBCACInval_MDC_Retraite")
                rsDest("PELMINVAL_Retraite") = rsSource("PELMINVAL_Retraite")
                rsDest("PELMPASSAGE_Retraite") = rsSource("PELMPASSAGE_Retraite")
                rsDest("PELMINCAP_Retraite") = rsSource("PELMINCAP_Retraite")
                rsDest("PEANNUALISATION_PASSAGE") = rsSource("PEANNUALISATION_PASSAGE")
                rsDest("PELMDEPENDANCE") = rsSource("PELMDEPENDANCE")
                rsDest("PEINVAL1CAT") = rsSource("PEINVAL1CAT")
                rsDest("PEDUREECHOMAGE") = rsSource("PEDUREECHOMAGE")
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_ParamCalcul", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760600B1
Public Sub Sync_ParamRentes(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.ParamRentes WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.ParamRentes WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.UpdateBatch
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_ParamRentes", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760600EF
Public Sub Sync_PassageNCA(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.PassageNCA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.PassageNCA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
      
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_PassageNCA", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A7606014D
Public Sub Sync_PM_Retenue(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.PM_Retenue WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.PM_Retenue WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
      
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_PM_Retenue", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A7606019B
Public Sub Sync_ProvisionsOuverture(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.ProvisionsOuverture WHERE NUMCLE = " & numPeriode & " AND GPECLE = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.ProvisionsOuverture WHERE NUMCLE = " & numPeriode & " AND GPECLE = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
              
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                rsDest("GPECLE") = rsSource("GPECLE")
                rsDest("NUMCLE") = rsSource("NUMCLE")
                
                rsDest("POSTECLE") = rsSource("POSTECLE")
                rsDest("PROV_ANn") = rsSource("PROV_ANn")
                rsDest("PROV_ANn1") = rsSource("PROV_ANn1")
                rsDest("PROV_ANn2") = rsSource("PROV_ANn2")
                rsDest("PROV_ANn3") = rsSource("PROV_ANn3")
                rsDest("PROV_ANn4") = rsSource("PROV_ANn4")
                rsDest("PROV_ANn5") = rsSource("PROV_ANn5")
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_ProvisionsOuverture", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760601E9
Public Sub Sync_Reassurance(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.Reassurance WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.Reassurance WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
                
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
        
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Reassurance", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76060246
Public Sub Sync_TBQREGA(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.TBQREGA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.TBQREGA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
                
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_TBQREGA", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76060299
Public Sub Sync_Assure_P3ILOGTRAIT(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
        sqlStr = "DELETE FROM P3IUser.Assure_P3ILOGTRAIT WHERE NumPeriode = " & numPeriode & " AND CleGroupe = " & GroupeCle
        cnxnDest.Execute sqlStr, , adExecuteNoRecords
        
        sqlStr = "SELECT * FROM P3IUser.Assure_P3ILOGTRAIT WHERE NumPeriode = " & numPeriode & " AND CleGroupe = " & GroupeCle
        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
        
        'insert data into DB
        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
        
        If rsSource.RecordCount > 0 Then
            Do While Not rsSource.EOF
                rsDest.AddNew
                
                CopyAllFieldsOfRecord rsSource, rsDest
                
                rsDest.Update
                rsSource.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Assure_P3ILOGTRAIT", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A760602F1
Public Sub Sync_Assure_P3IPROVCOLL(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
'        Set rsSource = New ADODB.Recordset
'        Set rsDest = New ADODB.Recordset
'
'        'delete old data from DB
'        sqlStr = "DELETE FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & numPeriode & " AND CleGroupe = " & GroupeCle
'        cnxnDest.Execute sqlStr, , adExecuteNoRecords
'
'        sqlStr = "SELECT * FROM P3IUser.Assure_P3IPROVCOLL WHERE NumPeriode = " & numPeriode & " AND CleGroupe = " & GroupeCle
'        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
'
'        'insert data into DB
'        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
'
'        If rsSource.RecordCount > 0 Then
'            Do While Not rsSource.EOF
'                rsDest.AddNew
'
'                CopyAllFieldsOfRecord rsSource, rsDest
'
'                rsDest.Update
'                rsSource.MoveNext
'            Loop
'        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Assure_P3IPROVCOLL", "Impossible d'ouvrir la base de données"
    End If

End Sub


'****************************************************************
'******************** Editions Tables ***************************
'****************************************************************

'##ModelId=5C8A7606036E
Public Sub Sync_Editions(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsSource As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    
    If (cnxnSource.State = adStateOpen) And (cnxnDest.State = adStateOpen) Then
        
        Set rsSource = New ADODB.Recordset
        Set rsDest = New ADODB.Recordset
        
        'delete old data from DB
'        sqlStr = "DELETE FROM P3IUser.TBQREGA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
'        cnxnDest.Execute sqlStr, , adExecuteNoRecords
'
'        sqlStr = "SELECT * FROM P3IUser.TBQREGA WHERE NumPeriode = " & numPeriode & " AND GroupeCle = " & GroupeCle
'        rsSource.Open sqlStr, cnxnSource, adOpenKeyset, adLockOptimistic
'
'        'insert data into DB
'        rsDest.Open sqlStr, cnxnDest, adOpenKeyset, adLockOptimistic
'
'        If rsSource.RecordCount > 0 Then
'            Do While Not rsSource.EOF
'                rsDest.AddNew
'
'                CopyAllFieldsOfRecord rsSource, rsDest
'
'                rsDest.Update
'                rsSource.MoveNext
'            Loop
'        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_Editions", "Impossible d'ouvrir la base de données"
    End If

End Sub


'****************************************************************
'************* Update Parent Tables of Periode ******************
'****************************************************************

'##ModelId=5C8A760603CB
Public Sub Sync_TypeCalcul(cnxnProd As ADODB.Connection, cnxnArch As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsProd As ADODB.Recordset
    Dim rsArch As ADODB.Recordset
    
    If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then
        
        'get data from Production DB
        Set rsProd = New ADODB.Recordset
        Set rsArch = New ADODB.Recordset
        
        sqlStr = "SELECT * FROM P3IUser.TypeCalcul"
        rsProd.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
                
        'insert data into Archive DB
        rsArch.Open "P3IUser.TypeCalcul", cnxnArch, adOpenKeyset, adLockOptimistic, adCmdTable
                
        If rsProd.RecordCount > 0 Then
            Do While Not rsProd.EOF
                rsArch.filter = "IdTypeCalcul = " & rsProd("IdTypeCalcul")
                
                If rsArch.RecordCount >= 1 Then
                    'update the record
                    rsArch("IdTypeCalcul") = rsProd("IdTypeCalcul")
                    rsArch("Libelle") = rsProd("Libelle")
                        
                    rsArch.UpdateBatch
                Else
                    'add a new record
                    rsArch.AddNew
                    rsArch("IdTypeCalcul") = rsProd("IdTypeCalcul")
                    rsArch("Libelle") = rsProd("Libelle")
                        
                    rsArch.UpdateBatch
                End If
                                
                rsProd.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_TypeCalcul", "Impossible d'ouvrir la base de données"
    End If

End Sub

'##ModelId=5C8A76070032
Public Sub Sync_TypePeriode(cnxnProd As ADODB.Connection, cnxnArch As ADODB.Connection, numPeriode As Integer)

    Dim sqlStr As String
    Dim rsProd As ADODB.Recordset
    Dim rsArch As ADODB.Recordset
    
    If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then
        
        'get data from Production DB
        Set rsProd = New ADODB.Recordset
        Set rsArch = New ADODB.Recordset
        
        sqlStr = "SELECT * FROM P3IUser.TypePeriode"
        rsProd.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
                
        'insert data into Archive DB
        rsArch.Open "P3IUser.TypePeriode", cnxnArch, adOpenKeyset, adLockOptimistic, adCmdTable
                
        If rsProd.RecordCount > 0 Then
            Do While Not rsProd.EOF
                rsArch.filter = "IdTypePeriode = " & rsProd("IdTypePeriode")
                
                If rsArch.RecordCount >= 1 Then
                    'update the record
                    rsArch("IdTypePeriode") = rsProd("IdTypePeriode")
                    rsArch("Libelle") = rsProd("Libelle")
                        
                    rsArch.UpdateBatch
                Else
                    'add a new record
                    rsArch.AddNew
                    rsArch("IdTypePeriode") = rsProd("IdTypePeriode")
                    rsArch("Libelle") = rsProd("Libelle")
                        
                    rsArch.UpdateBatch
                End If
                                
                rsProd.MoveNext
            Loop
        End If
        
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.Sync_TypePeriode", "Impossible d'ouvrir la base de données"
    End If

End Sub

'****************************************************************
'****************** Bulk Insert & CSV Creation ******************
'****************************************************************

'##ModelId=5C8A76070080
Public Function AddNewRecnoToAssureTxtFile(cnxnProd As ADODB.Connection, numPeriode As Integer) As OperationStatus

  Dim rsProd As ADODB.Recordset
    Dim sqlStr As String
    
    Dim i As Integer
    Dim nbFields As Integer
    Dim nbRecords As Long
    Dim lineRecno As String
    Dim fieldEndchar As String
    Dim fieldValue As String
    Dim FieldName As String
    Dim fieldType As String
    
    AddNewRecnoToAssureTxtFile = efailure
    
    If (cnxnProd.State = adStateOpen) Then
    
        On Error Resume Next
        Close #1
        
OpenTxt:
        
        Open CSVUNCPath & GetWinUser & "_AssureOldNewRecno.txt" For Output As #2
        
        If Err.Number = 70 Then
             
           MsgBox "Il semble que le fichier " & CSVUNCPath & GetWinUser & "_AssureOldNewRecno.txt" & _
            " est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok.", vbCritical
            
          Err.Clear
          GoTo OpenTxt
         End If
        
        On Error GoTo CSVError
        
        Set rsProd = New ADODB.Recordset
        
        sqlStr = "SELECT RECNO FROM P3IUser.Assure WHERE POPERCLE = " & numPeriode & " Order By RECNO"
                
        rsProd.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
        
        nbRecords = rsProd.RecordCount
        nbFields = rsProd.fields.Count
        
        If nbRecords = 0 Then
            AddNewRecnoToAssureTxtFile = eNoData
            Close #1
            Exit Function
        End If
           
        rsProd.MoveFirst
        
        With rsProd.fields
            Do While Not rsProd.EOF
                
                lineRecno = .Item(0).Value + "&&" + vbCrLf
                
                Print #1, lineRecno
                
                rsProd.MoveNext
            Loop
        End With
    
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise -1, "ArchiveRestore.AddNewRecnoToAssureTxtFile", "Impossible d'ouvrir la base de données"
    End If

    Close #1
    
    AddNewRecnoToAssureTxtFile = eSuccess
    
    Exit Function
    
CSVError:
    
    Close #1
    
    AddNewRecnoToAssureTxtFile = efailure
    
    Err.Raise Err.Number, "ArchiveRestore.AddNewRecnoToAssureTxtFile", Err.Description
    
End Function

'##ModelId=5C8A760700BE
Public Sub UpdateOldNewRecnoArray(cnxnProd As ADODB.Connection, numPeriode As Integer)

    Dim rsProd As ADODB.Recordset
    Dim sqlStr As String
    
    Dim i As Integer
    Dim j As Long
    Dim nbFields As Integer
    Dim nbRecords As Long
    Dim arrSize As Long
    Dim lineRecno As String
    Dim fieldEndchar As String
    Dim fieldValue As String
    Dim FieldName As String
    Dim fieldType As String
        
    If (cnxnProd.State = adStateOpen) Then
        
        On Error GoTo CSVError
        
        Set rsProd = New ADODB.Recordset
        
        sqlStr = "SELECT RECNO FROM P3IUser.Assure WHERE POPERCLE = " & numPeriode & " Order By RECNO"
                
        rsProd.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
        
        nbRecords = rsProd.RecordCount
        nbFields = rsProd.fields.Count
        
        'test if the array is empty
        On Error Resume Next
        arrSize = UBound(oldNewAssureRecno)
        If Err.Number = 9 Then
          'the array is empty
          Err.Clear
          On Error GoTo CSVError
          
          Exit Sub
        End If
        
        If (UBound(oldNewAssureRecno) <> nbRecords) Then
          MsgBox ("Error: Number of records in archived database do not correspond with number of records in production database!")
          Exit Sub
        End If
        
        j = 0
        If nbRecords > 0 Then
          rsProd.MoveFirst
          
          With rsProd.fields
              Do While Not rsProd.EOF
                j = j + 1
                oldNewAssureRecno(j) = oldNewAssureRecno(j) + "&&" + CStr(.Item(0).Value)
                  
                rsProd.MoveNext
              Loop
          End With
        End If
    
    Else
        Err.Raise -1, "ArchiveRestore.UpdateOldNewRecnoArray", "Impossible d'ouvrir la base de données"
    End If
    
    Exit Sub
    
CSVError:
    
    Err.Raise Err.Number, "ArchiveRestore.UpdateOldNewRecnoArray", Err.Description
    
End Sub

'##ModelId=5C8A7607010D
Public Function CreateCSVFile(cnxnProd As ADODB.Connection, numPeriode As Integer, table As String, csvFile As String, KeyName As String) As OperationStatus

    Dim rsProd As ADODB.Recordset
    Dim sqlStr As String
    
    Dim i As Integer
    Dim j As Long
    Dim nbFields As Integer
    Dim nbRecords As Long
    Dim line As String
    Dim lineRecno As String
    Dim newRecno As String
    Dim fieldEndchar As String
    Dim fieldValue As String
    Dim FieldName As String
    Dim fieldType As String
    
    CreateCSVFile = efailure
    
    If (cnxnProd.State = adStateOpen) Then
    
        On Error Resume Next
        Close #1
        
OpenCSV:
        
        Open CSVUNCPath & csvFile For Output As #1
        'Open "T:\P3I_Generali\Import\" & csvFile For Output As #1
        
        If Err.Number = 70 Then
          'File is already open and needs to be closed
          'Err.Raise -1, "ArchiveRestore.CreateCSVFile", "Il semble que le fichier " & CSVUNCPath & csvFile & _
          '" est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok."
               
          MsgBox "Il semble que le fichier " & CSVUNCPath & csvFile & _
          " est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok.", vbCritical
          
          Err.Clear
          GoTo OpenCSV
         End If
                
        'revert back to standard error handling
        'On Error GoTo 0
        On Error GoTo CSVError
        
        'create a dictionary table with old and new RECNO field
        If table = "P3IUser.Assure" Then
          Erase oldNewAssureRecno
          
'          On Error Resume Next
'          Close #2
          
'OpenCSV2:
'          Open CSVUNCPath & GetWinUser & "_AssureOldNewRecno.txt" For Output As #2
'
'          If Err.Number = 70 Then
'
'            MsgBox "Il semble que le fichier " & CSVUNCPath & GetWinUser & "_AssureOldNewRecno.txt" & _
'            " est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok.", vbCritical
'
'            Err.Clear
'            GoTo OpenCSV2
'           End If
'
'           On Error GoTo CSVError
'
       End If
        
        
        Set rsProd = New ADODB.Recordset
        
        If table = "P3IUser.Assure" Or table = "P3IUser.Assure_Retraite" Then
          sqlStr = "SELECT * FROM " & table & " WHERE " & KeyName & " = " & numPeriode & " Order By RECNO"
        Else
          sqlStr = "SELECT * FROM " & table & " WHERE " & KeyName & " = " & numPeriode
        End If
        
        rsProd.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
        
        nbRecords = rsProd.RecordCount
        nbFields = rsProd.fields.Count
        
        If table = "P3IUser.Assure" And nbRecords > 0 Then
          ReDim oldNewAssureRecno(1 To nbRecords)
        End If
        
        If nbRecords = 0 Then
            CreateCSVFile = eNoData
            Close #1
            
'            If table = "P3IUser.Assure" Then
'              Close #2
'            End If
          
            Exit Function
        End If
           
        rsProd.MoveFirst
        j = 0
        
        With rsProd.fields
            Do While Not rsProd.EOF
                
                line = ""
                lineRecno = ""
                j = j + 1
                
                For i = 0 To nbFields - 1
                
                    'manage the character that is added after each field
                    If i = nbFields - 1 Then
                        fieldEndchar = "" ' vbCrLf
                    Else
                        fieldEndchar = ";"
                    End If
                    
                    FieldName = .Item(i).Name
                    
                    'test for null value
                    If IsNull(.Item(i).Value) Then
                        fieldValue = ""
                    Else
                        fieldValue = .Item(i).Value
                    End If
                    
                    If table = "P3IUser.Assure" And FieldName = "RECNO" Then
                      oldNewAssureRecno(j) = fieldValue
                      'lineRecno = .Item(i).Value + "&&"
                    End If
                    
                    'change the value of the POIdAssure field in the Assure_Retraite table
                    If table = "P3IUser.Assure_Retraite" And FieldName = "POIdAssure" Then
                      If fieldValue <> "" Then
                        'search the array for this value
                        'c = Join(Filter(a, "20000&&"), "#")
                        newRecno = Join(filter(oldNewAssureRecno, fieldValue + "&&"), "#")
                        If newRecno <> "" Then
                          If (InStr(newRecno, "#")) > 0 Then
                            'take the first part of the string
                            newRecno = Left(newRecno, InStr(newRecno, "#") - 1)
                          End If
                          
                          newRecno = Replace(newRecno, fieldValue + "&&", "")
                          fieldValue = newRecno
                        
                        Else
                          'the element was not found
                          'MsgBox ("Element not found!")
                        End If
                      Else
                        'PoldAssure does not contain anything
                        MsgBox ("PoldAssure does not contain anything!")
                      End If
                    End If
                    
                    
                    'convert DateTime
                    If .Item(i).Type = 135 And fieldValue <> "" Then
                      fieldValue = Format(fieldValue, "mm/dd/yyyy")
                    End If
                    
                    'convert float data types : , -> .
                    If .Item(i).Type = 5 And fieldValue <> "" Then
                      fieldValue = Replace$(fieldValue, ",", ".")
                    End If
                    
                    If .Item(i).Type = 131 And fieldValue <> "" Then
                      fieldValue = Replace$(fieldValue, ",", ".")
                    End If
                    
                    'delete the character ; from text fields
                    If .Item(i).Type = 200 And InStr(fieldValue, ";") > 0 Then ' fieldValue <> "" Then
                      fieldValue = Replace$(fieldValue, ";", " ")
                    End If
    
                    'manage all the bit fields
                    If table = "P3IUser.Assure" Or table = "P3IUser.Assure_Retraite" Or table = "P3IUser.Assure_P3IPROVCOLL" Then
                      If FieldName = "PODATEPAIEMENTESTIMEE" Or FieldName = "POPSAPCAPMOYEN" Or FieldName = "PODOSSIERCLOS" _
                      Or FieldName = "POIsCadre" Or FieldName = "POPMReassAvecCorrectif" Or FieldName = "POPMAvecCorrectif" _
                      Or FieldName = "POCaptive" Or FieldName = "POTopAmortissable" Then
                          If UCase$(fieldValue) = "FALSE" Or UCase$(fieldValue) = "FAUX" Then
                              fieldValue = "0"
                          ElseIf UCase$(fieldValue) = "TRUE" Or UCase$(fieldValue) = "VRAI" Then
                              fieldValue = "1"
                          End If
                      End If
                    End If
                    
                    
'                    If table = "P3IUser.Assure_P3IPROVCOLL" Then
'                      If FieldName = "LBASSURE" Then
'                          If InStr(fieldValue, ";") > 0 Then
'                              fieldValue = Replace$(fieldValue, ";", " ")
'                          End If
'                      End If
'                    End If

                    'treat all delimiters and replace them with a space
                    If InStr(fieldValue, ";") > 0 Then
                        fieldValue = Replace$(fieldValue, ";", " ")
                    End If
                    
                    line = line + fieldValue + fieldEndchar
                Next i
                
                Print #1, line
                
'                If table = "P3IUser.Assure" Then
'                  Print #2, lineRecno
'                End If
                
                rsProd.MoveNext
            Loop
        End With
    
    Else
        'no DB connection available - raise an error to the calling function
        Err.Raise -1, "ArchiveRestore.CreateCSVFile - File: " & csvFile, "Impossible d'ouvrir la base de données"
    End If

    Close #1
    
'    If table = "P3IUser.Assure" Then
'      Close #2
'    End If
    
    CreateCSVFile = eSuccess
    
    Exit Function
    
CSVError:
    
    Close #1
    
'    If table = "P3IUser.Assure" Then
'      Close #2
'    End If
    
    CreateCSVFile = efailure
    
    'File is already open and needs to be closed
    Err.Raise Err.Number, "ArchiveRestore.CreateCSVFile", Err.Description
    '"Erreur pendant la création du fichier CSV : " & CSVUNCPath & csvFile
    
        
End Function

'Bulk Insert into Archives DB
'##ModelId=5C8A76070189
Public Sub BulkInsert(cnxnSource As ADODB.Connection, numPeriode As Integer, table As String, csvFile As String, KeyName As String)
    
    Dim rsSource As ADODB.Recordset
    Dim rsTestRows As ADODB.Recordset
    Dim sqlStr As String
    Dim compName As String
    Dim ret As Long
    Dim numbRecords As Long
    
    'delete assure for specific periode before doing a Bulk Insert for that periode
    sqlStr = "DELETE FROM " & table & " WHERE " & KeyName & " = " & numPeriode
    
    cnxnSource.Execute sqlStr, , adExecuteNoRecords
        
    'Dim compRes As String * 255
    'GetComputerNameA compRes, 255
    'compName = Left$(compRes, InStr(compRes, Chr$(0)) - 1)
    
    sqlStr = "BULK INSERT " & table & " FROM "
    sqlStr = sqlStr & "'" & CSVUNCPath & csvFile & "' "
    
    '###Test keepidentity
    sqlStr = sqlStr & "WITH (FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
    'sqlStr = sqlStr & "WITH (KEEPIDENTITY, FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
        
    On Error GoTo BulkError
    cnxnSource.Execute sqlStr, , adExecuteNoRecords
    On Error GoTo 0
    
    'raise error if no records have been inserted
    Set rsTestRows = New ADODB.Recordset
    sqlStr = "SELECT TOP 1 * FROM " & table & " WHERE " & KeyName & " = " & numPeriode  ' SELECT TOP 1 *
    Set rsTestRows = cnxnSource.Execute(sqlStr)
    
    If rsTestRows.RecordCount < 1 Then
      Err.Raise ArchiveRestoreErrors.errBulkInsertNoRecordsInserted, "ArchiveRestore.BulkInsert", _
      "Pendant l'opération Bulk Insert, aucune donnée n'a été inséré dans la table " & table & " pour la période " & numPeriode
    End If
    
    Exit Sub
    
BulkError:

  Err.Raise Err.Number, "ArchiveRestore.BulkInsert - " & Err.Source, _
      "Table : " & table & " - Période : " & numPeriode & " - " & Err.Description
    
    
End Sub




'*******************************************************************************************
'*********************** Archivation of General Tables *************************************
'*******************************************************************************************


'##ModelId=5C8A76070212
Public Sub SyncGeneralTablesFromProdToArchive()

  If MsgBox("Voulez-vous vraiment ARCHIVER toutes les TABLES GENERALES dans la base ARCHIVE ?", vbYesNo) = vbYes Then
  
  Dim table As String
  Dim tableCount As Integer
  Dim cnxnProd As ADODB.Connection
  Dim cnxnArch As ADODB.Connection
  Dim rsTables As ADODB.Recordset
  Dim colBadTables As Collection
  Dim cnt As Integer
  
  frmMain.StatusBar1.Visible = True
  Set colBadTables = New Collection
  
  Set cnxnProd = New ADODB.Connection
  Set cnxnArch = New ADODB.Connection
  
  cnxnProd.CommandTimeout = 900
  cnxnArch.CommandTimeout = 900

  cnxnProd.CursorLocation = adUseClient
  cnxnArch.CursorLocation = adUseClient
  
  cnxnProd.Open DatabaseFileName
  cnxnArch.Open DatabaseFileNameArchive
  
  If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then

      'iterate all tables  &synchronize entire table content to archive
      'On Error GoTo DBError
      Set rsTables = cnxnProd.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
      tableCount = rsTables.RecordCount
      cnt = 0
      
      rsTables.MoveFirst
      Do While Not rsTables.EOF
        table = LCase(rsTables!TABLE_NAME)
        SyncPKTables cnxnProd, cnxnArch, table, colBadTables, tableCount, cnt
        rsTables.MoveNext
      Loop
      
      SyncListTableLoi cnxnProd, cnxnArch, "listetableloi", colBadTables, tableCount, cnt
      
      rsTables.MoveFirst
      Do While Not rsTables.EOF
        table = LCase(rsTables!TABLE_NAME)
        SyncGeneralTables cnxnProd, cnxnArch, table, colBadTables, tableCount, cnt
        rsTables.MoveNext
      Loop
  
  
      frmMain.StatusBarProgress frmMain.StatusBar1, 2, 100, vbBlue
      frmMain.StatusBar1.Panels(1).text = "Nombre des tables déjà archivées : " & tableCount & "/" & tableCount
       
  
  Else
    'no DB connection available - raise an error to the calling function
    Err.Raise errDBConnectionProblem, "ArchiveRestore.SyncGeneralTablesFromProdToArchive", "Impossible d'ouvrir la base de données"
  End If
  
  
Cleanup:
    If Not cnxnProd Is Nothing Then
        If cnxnProd.State = adStateOpen Then
            cnxnProd.Close
        End If
    End If

    If Not cnxnArch Is Nothing Then
        If cnxnArch.State = adStateOpen Then
            cnxnArch.Close
        End If
    End If

    Set cnxnProd = Nothing
    Set cnxnArch = Nothing
    Set rsTables = Nothing
       
    'Finished
    
    If colBadTables.Count > 0 Then
      Dim strErr As String
      Dim colCount As Integer
      
      strErr = "La mise à jour des tables générales a échoué - il faudra relancer l'opération ! " & vbLf & vbLf
      
      For colCount = 1 To colBadTables.Count
        strErr = strErr + colBadTables.Item(colCount) & vbLf
      Next
      
      MsgBox strErr
    Else
      MsgBox "L'archivage des données générales est fini !"
    End If
    
    frmMain.StatusBar1.Visible = False
    
    
  End If 'from messagebox

End Sub

'##ModelId=5C8A76070217
Public Sub SyncGeneralTablesFromArchiveToProd()

  Dim table As String
  Dim tableCount As Integer
  Dim cnxnProd As ADODB.Connection
  Dim cnxnArch As ADODB.Connection
  Dim rsTables As ADODB.Recordset
  Dim colBadTables As Collection
  Dim cnt As Integer
  
  strMess = "ATTENTION -> Cette opération ECRASE TOUTES LES TABLES GENERALES dans votre base de PRODUCTION ! " & vbLf & vbLf
  strMess = strMess & "Avant de cliquer sur le bouton Ok, demander par sécurité à votre administrateur de faire une sauvegarde complète de votre base de Production dans SqlServer (full backup) " & vbLf & vbLf
  strMess = strMess & "En effet, l'application P3I peut ne plus fonctionner correctement si, par exemple, vous avez mis à jour des Tables Générales en Production et vous ne les avez pas Archivées  !" & vbLf & vbLf


'  strMess = "Est-ce que vous est sur de vouloir restaurer toutes les données générales ?" & vbLf & vbLf
'  strMess = strMess & "Cette opération va écraser tous les données dans les tables générales de la base production et peut potentiellement empêcher le bon fonctionnement de l'application P3I !" & vbLf & vbLf
'  strMess = strMess & "Avant de cliquer sur le bouton Ok il est fortement conseille de créer un sauvegarde de la base production !" & vbLf & vbLf
'  strMess = strMess & "" & vbLf & vbLf

  If MsgBox(strMess, vbYesNo) = vbYes Then
  
    If MsgBox("Est-ce que vous est sûr d'avoir créé une sauvegarde de la base Production?", vbYesNo) = vbYes Then
  
  
      frmMain.StatusBar1.Visible = True
      Set colBadTables = New Collection
      
      Set cnxnProd = New ADODB.Connection
      Set cnxnArch = New ADODB.Connection
      
      cnxnProd.CommandTimeout = 900
      cnxnArch.CommandTimeout = 900
      
      cnxnProd.CursorLocation = adUseClient
      cnxnArch.CursorLocation = adUseClient
      
      'DB's are swapped in order to allow a transfer of data from Archive to Prod
      cnxnProd.Open DatabaseFileNameArchive
      cnxnArch.Open DatabaseFileName
      
      If (cnxnProd.State = adStateOpen) And (cnxnArch.State = adStateOpen) Then
      
          'iterate all tables  &synchronize entire table content to archive
          'On Error GoTo DBError
          Set rsTables = cnxnProd.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
          tableCount = rsTables.RecordCount
          cnt = 0
          
          rsTables.MoveFirst
          Do While Not rsTables.EOF
            table = LCase(rsTables!TABLE_NAME)
            SyncPKTables cnxnProd, cnxnArch, table, colBadTables, tableCount, cnt
            rsTables.MoveNext
          Loop
          
          SyncListTableLoi cnxnProd, cnxnArch, "listetableloi", colBadTables, tableCount, cnt
          
          rsTables.MoveFirst
          Do While Not rsTables.EOF
            table = LCase(rsTables!TABLE_NAME)
            SyncGeneralTables cnxnProd, cnxnArch, table, colBadTables, tableCount, cnt
            rsTables.MoveNext
          Loop
      
      
          frmMain.StatusBarProgress frmMain.StatusBar1, 2, 100, vbBlue
          frmMain.StatusBar1.Panels(1).text = "Nombre des tables déjà archivées : " & tableCount & "/" & tableCount
           
      
      Else
        'no DB connection available - raise an error to the calling function
        Err.Raise errDBConnectionProblem, "ArchiveRestore.SyncGeneralTablesFromArchiveToProd", "Impossible d'ouvrir la base de données"
      End If
  
  
Cleanup:
      If Not cnxnProd Is Nothing Then
          If cnxnProd.State = adStateOpen Then
              cnxnProd.Close
          End If
      End If
  
      If Not cnxnArch Is Nothing Then
          If cnxnArch.State = adStateOpen Then
              cnxnArch.Close
          End If
      End If
  
      Set cnxnProd = Nothing
      Set cnxnArch = Nothing
      Set rsTables = Nothing
         
      'Finished
      
      If colBadTables.Count > 0 Then
        Dim strErr As String
        Dim colCount As Integer
        
        strErr = "La mise à jour des tables générales a échoué - il faudra relancer l'opération ! " & vbLf & vbLf
        
        For colCount = 1 To colBadTables.Count
          strErr = strErr + colBadTables.Item(colCount) & vbLf
        Next
        
        MsgBox strErr
      Else
        MsgBox "La restauration des données générales est fini !"
      End If
      
      frmMain.StatusBar1.Visible = False
      
      
    End If 'from messagebox
  
  End If

End Sub

'##ModelId=5C8A76070236
Private Sub SyncGeneralTables(cnxnProd As ADODB.Connection, cnxnArch As ADODB.Connection, table As String, colBadTables As Collection, tableCount As Integer, cnt As Integer)

  Dim sqlStr As String
  Dim sqlIdentity As String
  Dim rsSource As ADODB.Recordset
  Dim rsIdentityCheck As ADODB.Recordset
  Dim rsDest As ADODB.Recordset
  Dim tableHasIdentityColumn As Boolean
  Dim tableAvailable As Boolean
          
  If IsGeneralTable(table) Then
             
    Set rsSource = New ADODB.Recordset
    Set rsDest = New ADODB.Recordset
    cnt = cnt + 1
      
    'verify if that table exists at the destination DB  - if not, create it - if table creation is not successful, goto cleanup
    tableAvailable = False
    
    If Not TableExists(cnxnArch, table) Then
      If CreateTable(cnxnProd, cnxnArch, table) Then
        tableAvailable = True
      Else
        tableAvailable = False
      End If
    Else
      tableAvailable = True
    End If
      
    If tableAvailable = True Then
    
      'verify if the table has an identity column
      tableHasIdentityColumn = False
      Set rsIdentityCheck = New ADODB.Recordset
      
      sqlIdentity = "SELECT OBJECTPROPERTY(OBJECT_ID('" & table & "'), 'TableHasIdentity')"
      rsIdentityCheck.Open sqlIdentity, cnxnArch, adOpenKeyset, adLockOptimistic
      
      
      If rsIdentityCheck.fields.Count > 0 Then
        tableHasIdentityColumn = IIf(rsIdentityCheck.fields(0) = 0, False, True)
      Else
        'something is wrong
        GoTo TableError
      End If
      Set rsIdentityCheck = Nothing
      
      'delete old data from Archive DB
      sqlStr = "Truncate Table P3IUser. " & table
      Err.Clear
      
      On Error GoTo TableError
      cnxnArch.Execute sqlStr, , adExecuteNoRecords
      On Error GoTo 0
      
      'cannot truncate because of PK violation
      If Err.Number <> 0 Then  ' = -2147217865 Then
        Err.Clear
        GoTo DBError
      End If
      
      
      On Error GoTo DBError
      
      sqlStr = "SELECT * FROM P3IUser." & table
      rsSource.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
      
      If tableHasIdentityColumn Then
        sqlIdentity = "set identity_insert " & table & " on"
        cnxnArch.Execute sqlIdentity
      End If
                  
      'insert data into DB
      rsDest.Open sqlStr, cnxnArch, adOpenKeyset, adLockOptimistic
      
      If rsSource.RecordCount > 0 Then
        Do While Not rsSource.EOF
            rsDest.AddNew
            
            CopyAllFieldsOfRecord rsSource, rsDest
            DoEvents
  
            rsDest.Update
            rsSource.MoveNext
        Loop
      End If
      
      If tableHasIdentityColumn Then
        sqlIdentity = "set identity_insert " & table & " off"
        cnxnArch.Execute sqlIdentity
      End If
      
    End If
      
    frmMain.StatusBarProgress frmMain.StatusBar1, 2, CLng(cnt * 100 / tableCount), vbBlue
    frmMain.StatusBar1.Panels(1).text = "Nombre des tables déjà archivées : " & cnt & "/" & tableCount & " - Table: " & UCase(table) & " en cours d'archivage."
  
    
 End If


Cleanup:
    
  Set rsSource = Nothing
  Set rsDest = Nothing
  
  Exit Sub
  
DBError:

  'tables were not updated - start again
  colBadTables.Add ("Table: " & table & " -- Error: " & Err.Number & " -- " & Err.Description)
  
  GoTo Cleanup
  
  
TableError:

  GoTo Cleanup

End Sub

'##ModelId=5C8A760702CF
Private Function CreateTable(cnxnSource As ADODB.Connection, cnxnDest As ADODB.Connection, tableName As String) As Boolean

  Dim i As Integer
  Dim szSQL As String
  Dim res As Boolean
  Dim datType As String
  
  Dim rsSchemaCol As ADODB.Recordset
  Dim rsSchemaPK As ADODB.Recordset
  
  res = False

  'verify if table exists on destination
  If Not TableExists(cnxnDest, tableName) Then
  
    'test if primary key = age - otherwise do nothing
    'Debug.Print cnxnSource.DefaultDatabase
    Set rsSchemaPK = cnxnSource.OpenSchema(adSchemaPrimaryKeys, Array(cnxnSource.DefaultDatabase, "P3IUser", tableName))
    
    If rsSchemaPK.RecordCount > 0 Then
      If LCase(rsSchemaPK!COLUMN_NAME) = "age" Then
      
        'get all table columns
        Set rsSchemaCol = cnxnSource.OpenSchema(adSchemaColumns, Array(cnxnSource.DefaultDatabase, "P3IUser", tableName))
    
        szSQL = "CREATE TABLE " & tableName & "( "
        
        While Not rsSchemaCol.EOF
           datType = GetDataType(rsSchemaCol!DATA_TYPE)
           If datType = "Err" Then
            'quit procedure - stop creating this table
            GoTo TableCreationError
           End If
           
           szSQL = szSQL & rsSchemaCol!COLUMN_NAME & " " & datType & " " & GetNullable(rsSchemaCol!IS_NULLABLE) & ", "
           
           rsSchemaCol.MoveNext
        Wend
        
        szSQL = szSQL & " CONSTRAINT PK_" & tableName & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
        
        '### handle error
        On Error Resume Next
        cnxnDest.Execute szSQL
        On Error GoTo 0
        res = True
      
      End If
    End If
      
  End If
  
  
Cleanup:

  Set rsSchemaCol = Nothing
  Set rsSchemaPK = Nothing
  
  CreateTable = res
  
  Exit Function
  
TableCreationError:

  res = False
  GoTo Cleanup

End Function

'##ModelId=5C8A7607031D
Private Function TableExists(cnxnDest As ADODB.Connection, table As String) As Boolean

  Dim rsSchemaTables As ADODB.Recordset
  Dim res As Boolean
  
  res = False
  Set rsSchemaTables = cnxnDest.OpenSchema(adSchemaTables)
  
  While Not rsSchemaTables.EOF
    If LCase(rsSchemaTables!TABLE_NAME) = LCase(table) Then
      res = True
    End If
    
    rsSchemaTables.MoveNext
  Wend
  
  TableExists = res

End Function

'##ModelId=5C8A7607036B
Private Function GetDataType(dat As String) As String

  If dat = "2" Then
    GetDataType = "smallint"
  ElseIf dat = "5" Then
    GetDataType = "float"
  ElseIf dat = "129" Then
    GetDataType = "varchar(255)"
  Else
    'for the type of tables we are treating, there can't be any other data types - abondon the table creation
    GetDataType = "Err"
  End If
  
End Function

'##ModelId=5C8A760703AA
Private Function GetNullable(nullable As Boolean) As String

  If nullable = True Then
    GetNullable = "NULL"
  Else
    GetNullable = "NOT NULL"
  End If
  
End Function

'##ModelId=5C8A760703C9
Private Sub SyncPKTables(cnxnProd As ADODB.Connection, cnxnArch As ADODB.Connection, table As String, colBadTables As Collection, tableCount As Integer, cnt As Integer)

  Dim sqlStr As String
  Dim sqlIdentity As String
  Dim rsSource As ADODB.Recordset
  Dim rsDest As ADODB.Recordset
  Dim rsIdentityCheck As ADODB.Recordset
  Dim tableHasIdentityColumn As Boolean
          
  If IsPKTable(table) Then
  
    'do an insert & update
    Set rsSource = New ADODB.Recordset
    Dim calcProv As Variant
    cnt = cnt + 1
    
    'verify if the table has an identity column
    tableHasIdentityColumn = False
    Set rsIdentityCheck = New ADODB.Recordset
    sqlIdentity = "SELECT OBJECTPROPERTY(OBJECT_ID('" & table & "'), 'TableHasIdentity')"
    rsIdentityCheck.Open sqlIdentity, cnxnArch, adOpenKeyset, adLockOptimistic
    
    If rsIdentityCheck.fields.Count > 0 Then
      tableHasIdentityColumn = IIf(rsIdentityCheck.fields(0) = 0, False, True)
    Else
      'something is wrong
      GoTo TableError
    End If

    Set rsIdentityCheck = Nothing
    
    On Error GoTo DBError
    
    sqlStr = "SELECT * FROM P3IUser." & table
    rsSource.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
    
    If rsSource.RecordCount > 0 Then
      Do While Not rsSource.EOF
          
          'test if this record already exists in destination
          Select Case table
            Case "groupe"
              sqlStr = "SELECT * FROM P3IUser.Groupe Where GROUPECLE = " & rsSource!GroupeCle
            Case "typecalcul"
              sqlStr = "SELECT * FROM P3IUser.typecalcul Where IdTypeCalcul = " & rsSource!IdTypeCalcul
            Case "typedetable"
              sqlStr = "SELECT * FROM P3IUser.typedetable Where CleTypeDeTable = " & rsSource!CleTypeDeTable
            Case "typegarantie"
              sqlStr = "SELECT * FROM P3IUser.typegarantie Where TYPEGARCLE = " & rsSource!TYPEGARCLE
            Case "typeperiode"
              sqlStr = "SELECT * FROM P3IUser.typeperiode Where IdTypePeriode = " & rsSource!IdTypePeriode
              
            Case Else
              GoTo DBError
          End Select
          
          Set rsDest = New ADODB.Recordset
          rsDest.Open sqlStr, cnxnArch, adOpenKeyset, adLockOptimistic
          
          If rsDest.RecordCount > 0 Then
            'Record exists, do an Update : UPDATE Customers SET ContactName='test' WHERE CustomerID=12;
            Select Case table
              Case "groupe"
                sqlStr = "Update P3IUser.Groupe Set NOM = '" & Replace(rsSource!nom, "'", "''") & "' Where GROUPECLE = " & rsSource!GroupeCle
              Case "typecalcul"
                sqlStr = "Update P3IUser.typecalcul Set Libelle = '" & Replace(rsSource!Libelle, "'", "''") & "' Where IdTypeCalcul = " & rsSource!IdTypeCalcul
              Case "typedetable"
                sqlStr = "Update P3IUser.typedetable Set Libelle = '" & Replace(rsSource!Libelle, "'", "''") & "' Where CleTypeDeTable = " & rsSource!CleTypeDeTable
              Case "typegarantie"
                
                If rsSource!CalculProvMaintien = True Then
                  calcProv = 1
                ElseIf rsSource!CalculProvMaintien = False Then
                  calcProv = 0
                Else
                  calcProv = "null"
                End If
                sqlStr = "Update P3IUser.typegarantie Set NOM = '" & Replace(rsSource!nom, "'", "''") & "', CalculProvMaintien = " & calcProv & " Where TYPEGARCLE = " & rsSource!TYPEGARCLE
              Case "typeperiode"
                sqlStr = "Update P3IUser.typeperiode Set Libelle = '" & Replace(rsSource!Libelle, "'", "''") & "' Where IdTypePeriode = " & rsSource!IdTypePeriode
                
              Case Else
                GoTo DBError
            End Select
            
            cnxnArch.Execute sqlStr
          
          Else
            'Record does not exist, do an Insert
            
            If tableHasIdentityColumn Then
              sqlIdentity = "set identity_insert " & table & " on"
              cnxnArch.Execute sqlIdentity
            End If
            
            Set rsDest = New ADODB.Recordset
            sqlStr = "SELECT * FROM P3IUser." & table
            rsDest.Open sqlStr, cnxnArch, adOpenKeyset, adLockOptimistic
            
            rsDest.AddNew
          
            CopyAllFieldsOfRecord rsSource, rsDest
            DoEvents
            rsDest.Update
            
            If tableHasIdentityColumn Then
              sqlIdentity = "set identity_insert " & table & " off"
              cnxnArch.Execute sqlIdentity
            End If
            
          End If
          
          rsSource.MoveNext
      Loop
    End If
    
    frmMain.StatusBarProgress frmMain.StatusBar1, 2, CLng(cnt * 100 / tableCount), vbBlue
    frmMain.StatusBar1.Panels(1).text = "Nombre des tables déjà archivées : " & cnt & "/" & tableCount & " - Table: " & UCase(table) & " en cours d'archivage."
   
  
  End If
  
    

Cleanup:
    
  Set rsSource = Nothing
  Set rsDest = Nothing
  
  Exit Sub
  
DBError:

  'tables were not updated - start again
  colBadTables.Add ("Table: " & table & " -- Error: " & Err.Number & " -- " & Err.Description)
  
  GoTo Cleanup
  
TableError:
 
  GoTo Cleanup


End Sub

'##ModelId=5C8A7608007D
Private Sub SyncListTableLoi(cnxnProd As ADODB.Connection, cnxnArch As ADODB.Connection, table As String, colBadTables As Collection, tableCount As Integer, cnt As Integer)

  Dim sqlStr As String
  Dim sqlIdentity As String
  Dim rsSource As ADODB.Recordset
  Dim rsDest As ADODB.Recordset
  Dim rsIdentityCheck As ADODB.Recordset
      
  '************************************************************************
  'treat: listetableloi after other PK Tables because of its dependency on TypedeTable
  '************************************************************************
  
  Set rsSource = New ADODB.Recordset
  cnt = cnt + 1
        
  On Error GoTo DBError
  Dim tabUtilisat As Variant
  
  sqlStr = "SELECT * FROM P3IUser.ListeTableLoi"
  rsSource.Open sqlStr, cnxnProd, adOpenKeyset, adLockOptimistic
  
  If rsSource.RecordCount > 0 Then
    Do While Not rsSource.EOF
        
        'test if this record already exists in destination
        sqlStr = "SELECT * FROM P3IUser.ListeTableLoi Where TABLECLE = " & rsSource!TABLECLE
        
        Set rsDest = New ADODB.Recordset
        rsDest.Open sqlStr, cnxnArch, adOpenKeyset, adLockOptimistic
        
        If rsDest.RecordCount > 0 Then
          
          If rsSource!TableUtilisateur = True Then
            tabUtilisat = 1
          ElseIf rsSource!TableUtilisateur = False Then
            tabUtilisat = 0
          Else
            tabUtilisat = "null"
          End If
          
          'Record exists, do an Update
          sqlStr = "Update P3IUser.listetableloi Set LIBTABLE = '" & Replace(rsSource!libTable, "'", "''") & "', NOMTABLE = '" & rsSource!nomTable & _
          "', TYPETABLE = " & rsSource!typeTable & ", TableUtilisateur = " & tabUtilisat & _
          " Where TABLECLE = " & rsSource!TABLECLE
          
          cnxnArch.Execute sqlStr
        
        Else
          'Record does not exist, do an Insert
          
          Set rsDest = New ADODB.Recordset
          sqlStr = "SELECT * FROM P3IUser.listetableloi"
          rsDest.Open sqlStr, cnxnArch, adOpenKeyset, adLockOptimistic
          
          rsDest.AddNew
        
          CopyAllFieldsOfRecord rsSource, rsDest
          DoEvents
          rsDest.Update
          
        End If
        
        rsSource.MoveNext
    Loop
  End If
      
  
Cleanup:
    
  Set rsSource = Nothing
  Set rsDest = Nothing
  
  Exit Sub
 
DBError:

  colBadTables.Add ("Table: " & table & " -- Error: " & Err.Number & " -- " & Err.Description)
  
  GoTo Cleanup
   

End Sub

'##ModelId=5C8A7608011A
Private Function IsGeneralTable(table As String) As Boolean

  Dim notGeneral As String
  
  notGeneral = "&&agedepartretraiteinval1&&agedepartretraite&&capitaux_moyens&&catr9&&catr9inval&&cdsituat&&codecatinv&&codescat&&coeffamortissement&&"
  notGeneral = notGeneral & "correspondance_catoption&&donneessociales&&paramcalcul&&paramrentes&&passagenca&&pm_retenue&&"
  notGeneral = notGeneral & "provisionsouverture&&reassurance&&tbqrega&&assure_retraite&&assure_retraitetemp&&assuretemp&&"
  notGeneral = notGeneral & "assure&&editionii&&editioniii&&editioniitemp&&editioniiitemp&&editionrevalo&&editions&&"
  notGeneral = notGeneral & "editionstemp&&sysdiagrams&&periode&&lockedperiods&&"
  
  'to be decided
  notGeneral = notGeneral & "assure_p3ilogtrait&&assure_p3iprovcoll&&p3ilogtrait&&p3iprovcoll&&"
  notGeneral = notGeneral & "p3itrace&&ttlogtrait&&ttprovcoll&&"
  
  'PK-Tables
  notGeneral = notGeneral & "groupe&&listetableloi&&typecalcul&&typedetable&&"
  notGeneral = notGeneral & "typegarantie&&typeperiode&&"
  
  '### New Tables - delete those lines later, because they are also part of the general tables
  notGeneral = notGeneral & "assure_display_field&&assuredisplay&&assurefield&&"
  
  '### include this just for testing - delete later
  'notGeneral = notGeneral & "provisionbcac&&"
  
  
  If InStr(notGeneral, "&&" & LCase(table) & "&&") > 0 Then
    IsGeneralTable = False
  Else
    IsGeneralTable = True
  End If

End Function

'##ModelId=5C8A76080149
Private Function IsPKTable(table As String) As Boolean

  Dim pkTables As String
  
  'PK-Tables -- listetableloi is not included because it is treated seperately at the very end
  'pkTables = "&&groupe&&"
  pkTables = "&&groupe&&typecalcul&&typedetable&&"
  pkTables = pkTables & "typegarantie&&typeperiode&&"
  
  If InStr(pkTables, "&&" & LCase(table) & "&&") > 0 Then
    IsPKTable = True
  Else
    IsPKTable = False
  End If

End Function










