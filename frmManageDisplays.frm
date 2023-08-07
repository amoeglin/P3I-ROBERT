VERSION 5.00
Begin VB.Form frmManageDisplays 
   Caption         =   "Gestion des Affichages"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLeft 
      Height          =   3375
      ItemData        =   "frmManageDisplays.frx":0000
      Left            =   240
      List            =   "frmManageDisplays.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   28
      Top             =   1440
      Width           =   2895
   End
   Begin VB.ListBox lstRight 
      Height          =   3375
      ItemData        =   "frmManageDisplays.frx":0004
      Left            =   4200
      List            =   "frmManageDisplays.frx":0006
      MultiSelect     =   2  'Extended
      TabIndex        =   27
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdDown 
      Height          =   500
      Left            =   7320
      Picture         =   "frmManageDisplays.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3360
      Width           =   300
   End
   Begin VB.CommandButton cmdUp 
      Height          =   500
      Left            =   7320
      Picture         =   "frmManageDisplays.frx":034A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   300
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Créer"
      Height          =   375
      Left            =   10320
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Mise à jour"
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Make Default"
      Height          =   300
      Left            =   12840
      TabIndex        =   23
      Top             =   8520
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   780
      Left            =   13920
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox txtDescription 
      Height          =   1620
      Left            =   9240
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   13920
      TabIndex        =   18
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   12600
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command7 
         Caption         =   "&Fermer"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdUnselectAll 
      Caption         =   "Désélectionner tous"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10680
      TabIndex        =   9
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdUse 
      Caption         =   "Utiliser l'affichage sélectionné"
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Supprimer l'affichage sélectionné"
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdselectAll 
      Caption         =   "Sélectionner tous"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   280
      Left            =   3360
      Picture         =   "frmManageDisplays.frx":068C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton cmdRight 
      Height          =   280
      Left            =   3360
      Picture         =   "frmManageDisplays.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   600
   End
   Begin VB.ComboBox cmbDisplay 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   250
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   7920
      TabIndex        =   11
      Top             =   960
      Width           =   4695
      Begin VB.CheckBox chkDefault 
         Caption         =   "Utiliser comme affichage par défaut"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblDescription 
         Caption         =   "Description "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Nom du modèle d'affichage "
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Champs déjà sélectionnes"
      Height          =   225
      Left            =   4200
      TabIndex        =   29
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Champs à sélectionner"
      Height          =   225
      Left            =   240
      TabIndex        =   26
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label lblSelectDisplay 
      Caption         =   "Choisissez un modèle d'affichage "
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   315
      Width           =   2655
   End
End
Attribute VB_Name = "frmManageDisplays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A68320254"


'##ModelId=5C8A6832034E
Private Sub cmdUp_Click()

  Dim currIndex As Integer
  Dim currItem  As String
  Dim currItemData  As String
  Dim bSel As Boolean

  If lstRight.ListIndex >= 1 Then
      currIndex = lstRight.ListIndex
      currItem = lstRight.List(currIndex)
      currItemData = lstRight.ItemData(currIndex)
      
      bSel = lstRight.Selected(currIndex)
      lstRight.RemoveItem currIndex
      lstRight.AddItem currItem, currIndex - 1
      lstRight.ItemData(currIndex - 1) = currItemData
      
      lstRight.ListIndex = currIndex - 1
      lstRight.Selected(currIndex - 1) = bSel
  End If
    
End Sub
     
'##ModelId=5C8A6832036D
Private Sub cmdDown_Click()

  Dim currIndex As Integer
  Dim currItem  As String
  Dim currItemData  As String
  Dim bSel As Boolean
  
  If lstRight.ListIndex <> -1 And lstRight.ListIndex < lstRight.ListCount - 1 Then
      currIndex = lstRight.ListIndex
      currItem = lstRight.List(currIndex)
      currItemData = lstRight.ItemData(currIndex)
      
      bSel = lstRight.Selected(currIndex)
      lstRight.RemoveItem currIndex
      lstRight.AddItem currItem, currIndex + 1
      lstRight.ItemData(currIndex + 1) = currItemData
      
      lstRight.ListIndex = currIndex + 1
      lstRight.Selected(currIndex + 1) = bSel
  End If
    
End Sub

'##ModelId=5C8A6832037D
Private Sub lstRight_Click()

  ' enable/disable buttons when the current item changes
  cmdUp.Enabled = (lstRight.ListIndex > 0)
  cmdDown.Enabled = (lstRight.ListIndex <> -1 And lstRight.ListIndex < lstRight.ListCount - 1)
  
End Sub




'##ModelId=5C8A6832038C
Private Sub cmdAdd_Click()

  Dim cnxnProd As ADODB.Connection
  Dim sql As String
  Dim rsDisp As Recordset
  Dim dispID As String
  Dim i As Integer
  
  'make sure the name of this display is different from the name of the default display
  If txtName.text = cDefaultDisplayName Then
    MsgBox "S'il vous plait, choisissez un nom d'affichage diffèrent parce que le nom que vous avez choisi est réservé pour l'affichage par défaut.", vbOKOnly
    Exit Sub
  End If
  
  Set cnxnProd = New ADODB.Connection
  cnxnProd.CursorLocation = adUseClient
    
  On Error GoTo DBConnectionError
    
  cnxnProd.Open DatabaseFileName
    
  If (cnxnProd.State = adStateOpen) Then
  
    On Error GoTo TransError
    
    cnxnProd.BeginTrans
  
    'if IsDefault was set to True for this display, set the IsDefault value to false for the current default display
    If chkDefault.Value = 1 Then
      'sql = "Select ID From AssureDisplay Where IsDefault = 1 And UserName Not Like '" & cAllUsersUserName & "'"
      sql = "Select ID From AssureDisplay Where IsDefault = 1 AND UserName = '" & user_name & "'"
      Set rsDisp = cnxnProd.Execute(sql)
      
      If rsDisp.RecordCount > 0 Then
        Do While Not rsDisp.EOF
          'sql = "Update AssureDisplay Set IsDefault = 0 Where ID = " & rsDisp![ID] & " And UserName Not Like '" & cAllUsersUserName & "'"
          sql = "Update AssureDisplay Set IsDefault = 0 Where ID = " & rsDisp![ID]
          cnxnProd.Execute (sql)
          rsDisp.MoveNext
        Loop
      End If
    End If
    
    'insert new display
    sql = "Insert into AssureDisplay (UserName, Name, Description, IsDefault) Values ('" & user_name & "', '" & txtName.text & "', '" _
    & txtDescription.text & "', " & chkDefault.Value & ")"
  
    cnxnProd.Execute sql, True
    
    'get id of new display
    Set rsDisp = cnxnProd.Execute("Select @@Identity")
    dispID = rsDisp.fields(0).Value '  rsDisp![ID]
    
    'add display fields into table: Assure_Display_Field
    For i = 0 To lstRight.ListCount - 1
      sql = "Insert into Assure_Display_Field (DisplayID, FieldID) Values (" & dispID & ", " & lstRight.ItemData(i) & ")"
      cnxnProd.Execute sql, True
    Next
    
    cnxnProd.CommitTrans
    
    'Update AssureDisplays collection
    Set AssureDisplays = New AssureDisplays
    AssureDisplays.InitDisplayObjects
    FillFormElements True
    
    For i = 0 To cmbDisplay.ListCount - 1
      If cmbDisplay.ItemData(i) = dispID Then
        cmbDisplay.ListIndex = i
        Exit For
      End If
    Next
  
  Else
      'handle connection problem
      MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: frmManageDisplays.cmdAdd_Click", vbCritical
  End If
  
  
Cleanup:
  If Not cnxnProd Is Nothing Then
      If cnxnProd.State = adStateOpen Then
          cnxnProd.Close
      End If
  End If

  Set cnxnProd = Nothing
    
      
  Exit Sub
  
  
TransError:
    
  cnxnProd.RollbackTrans
  
  If Err <> 0 Then
    MsgBox "Erreur : " & Err.Number & vbLf & _
    "Source: " & "frmManageDisplays.cmdAdd_Click" & vbLf & _
    "Description: " & Err.Description
  End If

  GoTo Cleanup
  
    
DBConnectionError:
  
  'Raise Error: DB connection cannot be established
  If Err <> 0 Then
    MsgBox "Erreur base de données !" & vbLf & _
    "Erreur : " & Err.Number & vbLf & _
    "Source: frmManageDisplays.cmdAdd_Click" & vbLf & _
    "Description: " & Err.Description
  End If
      
  GoTo Cleanup

End Sub

'##ModelId=5C8A683203BB
Private Sub cmdDelete_Click()

  If MsgBox("Voulez-vous vraiment supprimer l'affichage sélectionné ?", vbYesNo) = vbYes Then
  
    Dim cnxnProd As ADODB.Connection
    Dim sql As String
    
    Set cnxnProd = New ADODB.Connection
    cnxnProd.CursorLocation = adUseClient
      
    On Error GoTo DBConnectionError
      
    cnxnProd.Open DatabaseFileName
      
    If (cnxnProd.State = adStateOpen) Then
    
      On Error GoTo TransError
      
      'cnxnProd.BeginTrans
      
      'if IsDefault was set to True for this display, set the IsDefault value to True on te Default Display
'      If chkDefault.Value = 1 Then
'        sql = "Update AssureDisplay Set IsDefault = 1 Where Name = '" & cDefaultDisplayName & "'"
'        cnxnProd.Execute (sql)
'      End If
      
      'selete display
      sql = "Delete From Assure_Display_Field Where DisplayID = " & cmbDisplay.ItemData(cmbDisplay.ListIndex)
      cnxnProd.Execute sql, True
      
      sql = "Delete From AssureDisplay Where ID = " & cmbDisplay.ItemData(cmbDisplay.ListIndex)
      cnxnProd.Execute sql, True
      
      'cnxnProd.CommitTrans
      
      'Update AssureDisplays collection
      Set AssureDisplays = New AssureDisplays
      AssureDisplays.InitDisplayObjects
      FillFormElements True
    
    Else
        'handle connection problem
        MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: frmManageDisplays.cmdAdd_Click", vbCritical
    End If
    
    
Cleanup:
    If Not cnxnProd Is Nothing Then
        If cnxnProd.State = adStateOpen Then
            cnxnProd.Close
        End If
    End If
  
    Set cnxnProd = Nothing
        
    Exit Sub
    
    
TransError:
      
    'cnxnProd.RollbackTrans
    
    If Err <> 0 Then
      MsgBox "Erreur : " & Err.Number & vbLf & _
      "Source: " & "frmManageDisplays.cmdDelete_Click" & vbLf & _
      "Description: " & Err.Description
    End If
  
    GoTo Cleanup
    
      
DBConnectionError:
    
    'Raise Error: DB connection cannot be established
    If Err <> 0 Then
      MsgBox "Erreur base de données !" & vbLf & _
      "Erreur : " & Err.Number & vbLf & _
      "Source: frmManageDisplays.cmdDelete_Click" & vbLf & _
      "Description: " & Err.Description
    End If
        
    GoTo Cleanup
  
  End If

End Sub

'##ModelId=5C8A683203DB
Private Sub cmdUpdate_Click()

  Dim cnxnProd As ADODB.Connection
  Dim sql As String
  Dim rsDisp As Recordset
  Dim dispID As String
  Dim i As Integer
  
  'make sure the name of this display is different from the name of the default display
'  If txtName.text = cDefaultDisplayName Then
'    MsgBox "S'il vous plait, choisissez un nom d'affichage diffèrent parce que le nom que vous avez choisi est réservé pour l'affichage par défaut.", vbOKOnly
'    Exit Sub
'  End If
  
  Set cnxnProd = New ADODB.Connection
  cnxnProd.CursorLocation = adUseClient
    
  On Error GoTo DBConnectionError
    
  cnxnProd.Open DatabaseFileName
    
  If (cnxnProd.State = adStateOpen) Then
  
    On Error GoTo TransError
    
    cnxnProd.BeginTrans
  
    'if IsDefault was set to True for this display, set the IsDefault value to false for the current default display
    If chkDefault.Value = 1 Then
      sql = "Select ID From AssureDisplay Where IsDefault = 1 AND UserName = '" & user_name & "'"
      Set rsDisp = cnxnProd.Execute(sql)
      
      If rsDisp.RecordCount > 0 Then
        Do While Not rsDisp.EOF
          'sql = "Update AssureDisplay Set IsDefault = 0 Where ID = " & rsDisp![ID] & " And UserName Not Like '" & cAllUsersUserName & "'"
          sql = "Update AssureDisplay Set IsDefault = 0 Where ID = " & rsDisp![ID]
          cnxnProd.Execute (sql)
          rsDisp.MoveNext
        Loop
      End If
      
    End If
    
    'Update new display
    sql = "Update AssureDisplay Set Name = '" & txtName.text & "', Description = '" & txtDescription.text & "', IsDefault = " _
    & chkDefault.Value & " Where ID = " & cmbDisplay.ItemData(cmbDisplay.ListIndex)
  
    cnxnProd.Execute sql, True
    
    sql = "Delete From Assure_Display_Field Where DisplayID = " & cmbDisplay.ItemData(cmbDisplay.ListIndex)
    cnxnProd.Execute sql, True
    
    'add display fields into table: Assure_Display_Field
    For i = 0 To lstRight.ListCount - 1
      sql = "Insert into Assure_Display_Field (DisplayID, FieldID) Values (" & cmbDisplay.ItemData(cmbDisplay.ListIndex) & ", " & lstRight.ItemData(i) & ")"
      cnxnProd.Execute sql, True
    Next
    
    cnxnProd.CommitTrans
    
    'Update AssureDisplays collection
    Set AssureDisplays = New AssureDisplays
    AssureDisplays.InitDisplayObjects
    FillFormElements False, cmbDisplay.ItemData(cmbDisplay.ListIndex)
    
'    For i = 0 To cmbDisplay.ListCount - 1
'      If cmbDisplay.ItemData(i) = dispID Then
'        cmbDisplay.ListIndex = i
'        Exit For
'      End If
'    Next
  
  Else
      'handle connection problem
      MsgBox "Impossible d'ouvrir la base de données !" & vbLf & "Source: frmManageDisplays.cmdAdd_Click", vbCritical
  End If
  
  
Cleanup:
  If Not cnxnProd Is Nothing Then
      If cnxnProd.State = adStateOpen Then
          cnxnProd.Close
      End If
  End If

  Set cnxnProd = Nothing
    
      
  Exit Sub
  
  
TransError:
    
  cnxnProd.RollbackTrans
  
  If Err <> 0 Then
    MsgBox "Erreur : " & Err.Number & vbLf & _
    "Source: " & "frmManageDisplays.cmdAdd_Click" & vbLf & _
    "Description: " & Err.Description
  End If

  GoTo Cleanup
  
    
DBConnectionError:
  
  'Raise Error: DB connection cannot be established
  If Err <> 0 Then
    MsgBox "Erreur base de données !" & vbLf & _
    "Erreur : " & Err.Number & vbLf & _
    "Source: frmManageDisplays.cmdAdd_Click" & vbLf & _
    "Description: " & Err.Description
  End If
      
  GoTo Cleanup
  
End Sub

'##ModelId=5C8A68330012
Private Sub cmdUse_Click()

  Dim disp As AssureDisplay
  
  For Each disp In AssureDisplays
    If disp.ID = cmbDisplay.ItemData(cmbDisplay.ListIndex) Then
      Set AssureDisplays.CurrentlySelectedDisplay = disp
      Exit For
    End If
  Next
  
  Unload Me
        
End Sub

'##ModelId=5C8A68330031
Private Sub Form_Load()

  'Get all the latest display objects
  Set AssureDisplays = New AssureDisplays
  AssureDisplays.InitDisplayObjects
  
  'Dim disp As AssureDisplay
  'Set disp = AssureDisplays.GetDisplayAllFields
  
  FillFormElements True
    
End Sub

'##ModelId=5C8A68330041
Private Sub cmbDisplay_Click()

  Dim dispIndex As String
  Dim disp As AssureDisplay

  dispIndex = CStr(cmbDisplay.ItemData(cmbDisplay.ListIndex))

  FillFormElements False, cmbDisplay.ItemData(cmbDisplay.ListIndex)

  Set disp = AssureDisplays.Item(dispIndex)

  txtName.text = disp.Name
  txtDescription.text = disp.Description
  
  If disp.Name = cDefaultDisplayName Then
    'cmdDelete.Enabled = False
    'cmdUpdate.Enabled = False
    'cmdAdd.Enabled = False
  Else
    'cmdDelete.Enabled = True
    'cmdUpdate.Enabled = True
    'cmdAdd.Enabled = True
  End If
  
End Sub

'##ModelId=5C8A68330050
Private Sub FillFormElements(useDefaultDisplay As Boolean, Optional dispID As Integer)

  Dim disp As AssureDisplay
  Dim selectedDisplay As AssureDisplay
  Dim field As AssureField
  Dim defaultID As Integer
  Dim i As Integer
  Dim j As Integer
  Dim defaultDisplayExists As Boolean
  Dim DefaultDisplaySet As Boolean
   
  If useDefaultDisplay Then
    cmbDisplay.Clear
  End If
  
  lstRight.Clear
  lstLeft.Clear
  
  defaultDisplayExists = False
  DefaultDisplaySet = False
  
  For Each disp In AssureDisplays
    If disp.IsDefault Then
      defaultDisplayExists = True
      Exit For
    End If
  Next
  
  For Each disp In AssureDisplays
    
    If useDefaultDisplay Then
      If disp.ID <> cDispAllFieldsID Then
        cmbDisplay.AddItem disp.Name
        cmbDisplay.ItemData(cmbDisplay.NewIndex) = disp.ID
      
        If disp.IsDefault Or (Not defaultDisplayExists And Not DefaultDisplaySet) Then
          Set selectedDisplay = disp
          DefaultDisplaySet = True
          defaultID = disp.ID
          txtName.text = disp.Name
          txtDescription.text = disp.Description
          chkDefault.Value = IIf(disp.IsDefault, 1, 0)
          
          'fill the fields box
          For Each field In disp.AssureFields
            lstRight.AddItem field.DispalyFieldNoBrackets
            lstRight.ItemData(lstRight.NewIndex) = field.ID
          Next
        
        End If
      End If
    Else
      If disp.ID = dispID Then
        Set selectedDisplay = disp
        defaultID = disp.ID
        txtName.text = disp.Name
        txtDescription.text = disp.Description
        chkDefault.Value = IIf(disp.IsDefault, 1, 0)
        
        'fill the fields box
        For Each field In disp.AssureFields
          lstRight.AddItem field.DispalyFieldNoBrackets
          lstRight.ItemData(lstRight.NewIndex) = field.ID
        Next
      End If
    End If
  Next
  
  'fill the left listbox
  Dim dispAllFields As AssureDisplay
  Set dispAllFields = AssureDisplays.GetDisplayAllFields
    
  If Not selectedDisplay Is Nothing Then
    For Each field In dispAllFields.AssureFields
      If Not selectedDisplay.ContainsFieldID(field.ID) Then
        lstLeft.AddItem field.DispalyFieldNoBrackets
        lstLeft.ItemData(lstLeft.NewIndex) = field.ID
      End If
    Next
  End If
  
  
  'select the default item
  If useDefaultDisplay Then
  For i = 0 To cmbDisplay.ListCount - 1
    If cmbDisplay.ItemData(i) = defaultID Then
      cmbDisplay.ListIndex = i
      Exit For
    End If
  Next
  End If
  
End Sub

'##ModelId=5C8A6833009E
Private Sub cmdRight_Click()

  Dim i As Integer

  For i = lstLeft.ListCount To 1 Step -1
  
    If lstLeft.Selected(i - 1) Then
      lstRight.AddItem lstLeft.List(i - 1)
      lstRight.ItemData(lstRight.NewIndex) = lstLeft.ItemData(i - 1)
      lstLeft.RemoveItem i - 1
    End If
  
  Next

End Sub

'##ModelId=5C8A683300AE
Private Sub cmdLeft_Click()

  Dim i As Integer

  For i = lstRight.ListCount To 1 Step -1
  
    If lstRight.Selected(i - 1) Then
      lstLeft.AddItem lstRight.List(i - 1)
      lstLeft.ItemData(lstLeft.NewIndex) = lstRight.ItemData(i - 1)
      lstRight.RemoveItem i - 1
    End If
  
  Next

End Sub

'##ModelId=5C8A683300BE
Private Sub cmdselectAll_Click()
 
  lstLeft.Clear
  lstRight.Clear
  
  Dim dispAllFields As AssureDisplay
  Set dispAllFields = AssureDisplays.GetDisplayAllFields
    
  For Each field In dispAllFields.AssureFields
    lstRight.AddItem field.DispalyField
    lstRight.ItemData(lstRight.NewIndex) = field.ID
  Next
  
End Sub

'##ModelId=5C8A683300CD
Private Sub cmdUnselectAll_Click()

  lstLeft.Clear
  lstRight.Clear
  
  Dim dispAllFields As AssureDisplay
  Set dispAllFields = AssureDisplays.GetDisplayAllFields
    
  For Each field In dispAllFields.AssureFields
    lstLeft.AddItem field.DispalyField
    lstLeft.ItemData(lstLeft.NewIndex) = field.ID
  Next
  
End Sub



'##ModelId=5C8A683300DD
Private Sub cmdClose_Click()
  Unload Me
End Sub
