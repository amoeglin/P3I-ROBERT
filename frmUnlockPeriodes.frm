VERSION 5.00
Begin VB.Form frmUnlockPeriodes 
   Caption         =   "Périodes Bloquées"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Débloquer les périodes sélectionnées "
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   4695
   End
   Begin VB.ListBox lstPeriodes 
      Height          =   2205
      ItemData        =   "frmUnlockPeriodes.frx":0000
      Left            =   240
      List            =   "frmUnlockPeriodes.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmUnlockPeriodes.frx":0004
      Height          =   775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frmUnlockPeriodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6834038C"
'##ModelId=5C8A6835009E
Private Sub Command1_Click()

  'm_dataSource.Execute "Delete From LockedPeriods"
  'lstPeriodes.Clear
  
  Dim per As String
  
  For i = lstPeriodes.ListCount To 1 Step -1
  
    If lstPeriodes.Selected(i - 1) Then
      per = Replace(lstPeriodes.List(i - 1), "Période : ", "")
      per = Trim(Left(per, InStr(per, "--") - 1))
      
      lstPeriodes.RemoveItem i - 1
      m_dataSource.Execute "Delete From LockedPeriods where Periode = " & per
    End If
  
  Next
  
  Unload Me
  
  MsgBox "Les périodes sélectionnées étaient débloquées !"
  
End Sub

'##ModelId=5C8A683500BE
Private Sub Form_Load()

  Dim rs As ADODB.Recordset
  
  lstPeriodes.Clear
  
  'lock this periode for the current user
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM LockedPeriods", Snapshot)
  
  If rs.RecordCount > 0 Then
    Do While Not rs.EOF
      lstPeriodes.AddItem "Période : " & rs.fields("Periode") & " -- Bloquée  par : " & rs.fields("UserName")
      rs.MoveNext
    Loop
  End If
  
  rs.Close
  
End Sub
