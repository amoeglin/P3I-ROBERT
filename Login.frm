VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login..."
   ClientHeight    =   1770
   ClientLeft      =   3300
   ClientTop       =   2895
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1770
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboGroupe 
      Height          =   315
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   3210
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1395
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   450
      Width           =   3210
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -15
      TabIndex        =   8
      Top             =   1245
      Width           =   4650
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2685
      TabIndex        =   7
      Top             =   1395
      Width           =   1335
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   6
      Top             =   1395
      Width           =   1335
   End
   Begin VB.TextBox Password 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1380
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   810
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Groupe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Utilisateur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   495
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Mot de passe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   855
      Width           =   1275
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67A60398"
Option Explicit

'##ModelId=5C8A67A700AA
Private Sub Cancel_Click()
    ret_code = -1
    
    Unload Me
End Sub

'##ModelId=5C8A67A700BA
Private Sub cboGroupe_Click()
  If cboGroupe.ListIndex <> -1 Then
    ' rempli la liste des utilisateurs
    m_dataHelper.FillCombo cboUser, "SELECT USERCLE, TANOM FROM Utilisateur WHERE TAGPECLE = " & cboGroupe.ItemData(cboGroupe.ListIndex) & " ORDER BY TANOM", 1
  End If
End Sub

'##ModelId=5C8A67A700D9
Private Sub Form_Activate()
  Password.SetFocus
End Sub

'##ModelId=5C8A67A700E8
Private Sub Form_Load()
    ' Centre la fenetre
    Left = (Screen.Width - Width) / 2
    top = (Screen.Height - Height) / 2
    
    ' rempli la liste des groupes
    m_dataHelper.FillCombo cboGroupe, "SELECT GROUPECLE, NOM FROM GROUPE ORDER BY NOM", GetSettingIni(CompanyName, SectionName, "LastGroupe", "1")
    m_dataHelper.FillCombo cboUser, "SELECT USERCLE, TANOM FROM Utilisateur WHERE TAGPECLE = " & cboGroupe.ItemData(cboGroupe.ListIndex) & " ORDER BY TANOM", GetSettingIni(CompanyName, SectionName, "LastUser", "-1")
    
    
    '###
    'Password.text = "PS"
        
    
End Sub

'##ModelId=5C8A67A700F8
Private Sub Ok_Click()
    Dim UserCle As Integer
    
    ' utilisateur et mot de passe doivent etre rempli
    If cboUser.text = "" Then
        Beep
        cboUser.SetFocus
        Exit Sub
    End If
    If Password.text = "" Then
        Beep
        Password.SetFocus
        Exit Sub
    End If
    
    ' set global data
    GroupeCle = cboGroupe.ItemData(cboGroupe.ListIndex)
    UserCle = cboUser.ItemData(cboUser.ListIndex)
    user_name = cboUser.text
    user_pwd = Password.text
    
    Call SaveSettingIni(CompanyName, SectionName, "LastGroupe", GroupeCle)
    Call SaveSettingIni(CompanyName, SectionName, "LastUser", UserCle)
    
    ret_code = 0
    Unload Me
End Sub
