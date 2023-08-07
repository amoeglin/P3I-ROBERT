VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssureDisplays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A760D017B"

'##ModelId=5C8A760D0289
Private mCol As Collection
'##ModelId=5C8A760D02B8
Private mCurrentlySelectedDisplay As AssureDisplay

'##ModelId=5C8A760D031E
Public Function Add(ID As Integer, Name As String, UserName As String, Description As String, IsDefault As Boolean) As AssureDisplay

    Dim objAssureDisplay As AssureDisplay
    Set objAssureDisplay = New AssureDisplay
    
    objAssureDisplay.ID = ID
    objAssureDisplay.Name = Name
    objAssureDisplay.UserName = UserName
    objAssureDisplay.Description = Description
    objAssureDisplay.IsDefault = IsDefault
    
    mCol.Add objAssureDisplay, ID

    Set Add = objAssureDisplay
    Set objAssureDisplay = Nothing

End Function

'##ModelId=5C8A760D038B
Public Sub AddDispObject(display As AssureDisplay)
    
    mCol.Add display, CStr(display.ID)

End Sub

'##ModelId=5C8A760D03AA
Public Property Get Item(vntIndexKey As String) As AssureDisplay
  Set Item = mCol(vntIndexKey)
End Property

'##ModelId=5C8A760D03D9
Public Property Get GetDisplayAllFields() As AssureDisplay
  
  Dim disp As AssureDisplay
  For Each disp In mCol
    If disp.ID = cDispAllFieldsID Then
      Set GetDisplayAllFields = disp
      Exit For
    End If
  Next
  
End Property

'##ModelId=5C8A760E002F
Public Property Get CurrentlySelectedDisplay() As AssureDisplay
  Set CurrentlySelectedDisplay = mCurrentlySelectedDisplay
End Property

'##ModelId=5C8A760E0001
Public Property Set CurrentlySelectedDisplay(ByVal display As AssureDisplay)
    Set mCurrentlySelectedDisplay = display
End Property

'##ModelId=5C8A760E003F
Public Property Get Count() As Long
    Count = mCol.Count
End Property

'##ModelId=5C8A760E004F
Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

'##ModelId=5C8A760E006E
Public Sub Clear()
    Set mCol = New Collection
End Sub


'##ModelId=5C8A760E007E
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Function

'##ModelId=5C8A760E009D
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

'##ModelId=5C8A760E00AC
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'##ModelId=5C8A760E00CC
Public Sub InitDisplayObjects()
  
  Dim disp As AssureDisplay
  Dim colAssureFields As AssureFields
  Dim AssureField As AssureField

  Dim cnxnProd As ADODB.Connection
  Dim rsDisplays As ADODB.Recordset
  Dim rsDisp As ADODB.Recordset
  Dim rsFields As ADODB.Recordset
  Dim sql As String
  Dim noBracket As String
  Dim defaultDescription As String
  Dim dispID As String
  
  On Error GoTo DBConnectionError
  
  Set cnxnProd = New ADODB.Connection
  cnxnProd.Open DatabaseFileName
  cnxnProd.CursorLocation = adUseClient
  Set rsDisplays = New ADODB.Recordset
  
  If (cnxnProd.State = adStateOpen) Then
    
    On Error GoTo TransError
      
    'get all displays for current user
    'sql = "SELECT * FROM P3IUser.AssureDisplay WHERE UserName LIKE '" & user_name & "' Or UserName LIKE '" & cAllUsersUserName & "' Order By ID"
    sql = "SELECT * FROM P3IUser.AssureDisplay WHERE UserName LIKE '" & user_name & "' Order By ID"
    rsDisplays.Open sql, cnxnProd, adOpenStatic, adLockOptimistic    ' adOpenKeyset, adLockOptimistic
    
    If rsDisplays.RecordCount = 0 Then
      'There are no displays for this user - create a default display
      defaultDescription = "Cette affichage utilise tous les champs disponibles."
      
      sql = "Insert into AssureDisplay (UserName, Name, Description, IsDefault) Values ('" & user_name & "', '" & cDefaultDisplayName & "', '" _
      & defaultDescription & "', 1)"
    
      cnxnProd.Execute sql, True
      
      'get id of new display
      Set rsDisp = cnxnProd.Execute("Select @@Identity")
      dispID = rsDisp.fields(0).Value '  rsDisp![ID]
      
      'add display fields into table: Assure_Display_Field
      sql = "Select ID From AssureField"
      Set rsDisp = cnxnProd.Execute(sql)
      
      If rsDisp.RecordCount > 0 Then
        Do While Not rsDisp.EOF
          sql = "Insert into Assure_Display_Field (DisplayID, FieldID) Values (" & dispID & ", " & rsDisp.fields(0).Value & ")"
          cnxnProd.Execute sql, True
          rsDisp.MoveNext
        Loop
      End If
      
      sql = "SELECT * FROM P3IUser.AssureDisplay WHERE UserName LIKE '" & user_name & "' Order By ID"
      If rsDisplays.State = 1 Then rsDisplays.Close
      rsDisplays.Open sql, cnxnProd, adOpenStatic, adLockOptimistic
    End If
    
    If rsDisplays.RecordCount > 0 Then
      rsDisplays.MoveFirst
      
      Do While Not rsDisplays.EOF
        Set disp = New AssureDisplay
        disp.ID = rsDisplays![ID]
        disp.Name = rsDisplays![Name]
        disp.Description = rsDisplays![Description]
        disp.UserName = rsDisplays![UserName]
        disp.IsDefault = rsDisplays![IsDefault]
        
        'get all fields for the current display
        sql = "SELECT P3IUser.AssureField.DisplayField, P3IUser.AssureField.DBQuery, P3IUser.AssureField.SpreadColor, " _
          & "P3IUser.AssureField.DataFillRequired, P3IUser.AssureField.IsDateColumn, P3IUser.AssureField.ID " _
          & "FROM P3IUser.AssureDisplay INNER JOIN P3IUser.Assure_Display_Field ON " _
          & "P3IUser.AssureDisplay.ID = P3IUser.Assure_Display_Field.DisplayID INNER JOIN " _
          & "P3IUser.AssureField ON P3IUser.Assure_Display_Field.FieldID = P3IUser.AssureField.ID " _
          & "WHERE P3IUser.AssureDisplay.ID = " & disp.ID
          
        Set rsFields = New ADODB.Recordset
        rsFields.Open sql, cnxnProd, adOpenStatic, adLockOptimistic
        
        If rsFields.RecordCount > 0 Then
          rsFields.MoveFirst
          
          Set colAssureFields = New AssureFields
          
          Do While Not rsFields.EOF
            noBracket = Replace(rsFields![DisplayField], "[", "")
            noBracket = Replace(noBracket, "]", "")
            
            Set AssureField = New AssureField
            AssureField.ID = rsFields![ID]
            AssureField.DispalyField = rsFields![DisplayField]
            AssureField.DispalyFieldNoBrackets = noBracket
            AssureField.DBQuery = rsFields![DBQuery]
            AssureField.SpreadColor = rsFields![SpreadColor]
            AssureField.DataFillRequired = rsFields![DataFillRequired]
            AssureField.IsDateColumn = rsFields![ID]
            
            'add each field to the fields collection
            colAssureFields.AddFieldObject AssureField
            
            rsFields.MoveNext
          Loop
        
          'add the fields collection to the display
          Set disp.AssureFields = colAssureFields
          
        End If  ' of: If rsFields.RecordCount > 0
        
        If disp.IsDefault Then
          Set mCurrentlySelectedDisplay = disp
        End If
        
        'add Display to collection
        Me.AddDispObject disp
        
        rsDisplays.MoveNext
      Loop
      'End With
    
    End If ' of: If rsDisplays.RecordCount > 0
    
    'get all display fields - this is no longer required
    Set disp = New AssureDisplay
    disp.ID = cDispAllFieldsID
    disp.Name = cDispAllFieldsName

    sql = "Select * From AssureField"
    Set rsFields = New ADODB.Recordset
    rsFields.Open sql, cnxnProd, adOpenStatic, adLockOptimistic

    If rsFields.RecordCount > 0 Then
      rsFields.MoveFirst

      Set colAssureFields = New AssureFields

      Do While Not rsFields.EOF
        noBracket = Replace(rsFields![DisplayField], "[", "")
        noBracket = Replace(noBracket, "]", "")

        Set AssureField = New AssureField
        AssureField.ID = rsFields![ID]
        AssureField.DispalyField = rsFields![DisplayField]
        AssureField.DispalyFieldNoBrackets = noBracket
        AssureField.DBQuery = rsFields![DBQuery]
        AssureField.SpreadColor = rsFields![SpreadColor]
        AssureField.DataFillRequired = rsFields![DataFillRequired]
        AssureField.IsDateColumn = rsFields![ID]

        'add each field to the fields collection
        colAssureFields.AddFieldObject AssureField

        rsFields.MoveNext
      Loop

      Set disp.AssureFields = colAssureFields

      'add Display to collection
      Me.AddDispObject disp

    End If
    
  Else
      'handle connection problem
      MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: AssureDisplays.InitDisplayObjects", vbCritical
  End If
  
  
Cleanup:
  If Not cnxnProd Is Nothing Then
      If cnxnProd.State = adStateOpen Then
          cnxnProd.Close
      End If
  End If

  Set cnxnProd = Nothing
  Set rsFields = Nothing
  Set rsDisplays = Nothing
  
  Exit Sub
    
TransError:
    
  If Err <> 0 Then
    MsgBox "Erreur : " & Err.Number & vbLf & _
    "Source: " & "AssureDisplays.InitDisplayObjects" & vbLf & _
    "Description: " & Err.Description
  End If

  GoTo Cleanup

    
DBConnectionError:

  'Raise Error: DB connection cannot be established
  If Err <> 0 Then
    MsgBox "Impossible d'ouvrir la base de données !" & vbLf & _
    "Erreur : " & Err.Number & vbLf & _
    "Source: AssureDisplays.InitDisplayObjects" & vbLf & _
    "Description: " & Err.Description
  End If
      
  GoTo Cleanup
  
  
End Sub



