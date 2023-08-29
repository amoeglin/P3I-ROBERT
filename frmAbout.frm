VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   6630
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6345
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4576.149
   ScaleMode       =   0  'User
   ScaleWidth      =   5958.283
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   90
      Picture         =   "frmAbout.frx":1BB2
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   3735
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Default         =   -1  'True
      Height          =   345
      Left            =   2550
      TabIndex        =   0
      Top             =   6255
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   5
      Top             =   6120
      Width           =   6540
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   0
      Picture         =   "frmAbout.frx":3764
      ScaleHeight     =   3600
      ScaleWidth      =   6300
      TabIndex        =   8
      Top             =   0
      Width           =   6360
   End
   Begin VB.Label Label2 
      Caption         =   "Conception, réalisation et maintenance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   90
      TabIndex        =   10
      Top             =   4320
      Width           =   3705
   End
   Begin VB.Label Label1 
      Caption         =   "http://www.actuariesservices.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3825
      MouseIcon       =   "frmAbout.frx":1C566
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Site WEB du Cabinet MOEGLIN"
      Top             =   4320
      Width           =   2490
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calculs des provisions techniques..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Left            =   4005
      TabIndex        =   2
      Top             =   5535
      Width           =   2265
   End
   Begin VB.Label lblLicense 
      Caption         =   "lblLicense"
      Height          =   225
      Left            =   90
      TabIndex        =   7
      Top             =   5850
      Width           =   6180
   End
   Begin VB.Label lblDB 
      Caption         =   "DB..."
      Height          =   1215
      Left            =   90
      TabIndex        =   6
      Top             =   4545
      Width           =   6180
   End
   Begin VB.Label lblTitle 
      Caption         =   "Provisions Techniques Incapacités Invalidités (P3I)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   675
      TabIndex        =   3
      Top             =   3735
      Width           =   5595
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   675
      TabIndex        =   4
      Top             =   4005
      Width           =   5640
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67A001C5"
Option Explicit

' Reg Key Security Options...
'##ModelId=5C8A67A002AD
Const READ_CONTROL = &H20000
'##ModelId=5C8A67A002CD
Const KEY_QUERY_VALUE = &H1
'##ModelId=5C8A67A002DC
Const KEY_SET_VALUE = &H2
'##ModelId=5C8A67A002FC
Const KEY_CREATE_SUB_KEY = &H4
'##ModelId=5C8A67A0030B
Const KEY_ENUMERATE_SUB_KEYS = &H8
'##ModelId=5C8A67A0032A
Const KEY_NOTIFY = &H10
'##ModelId=5C8A67A0034A
Const KEY_CREATE_LINK = &H20
'##ModelId=5C8A67A00359
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
'##ModelId=5C8A67A00379
Const HKEY_LOCAL_MACHINE = &H80000002
'##ModelId=5C8A67A00398
Const ERROR_SUCCESS = 0
'##ModelId=5C8A67A003B7
Const REG_SZ = 1                         ' Unicode nul terminated string
'##ModelId=5C8A67A003D6
Const REG_DWORD = 4                      ' 32-bit number

'##ModelId=5C8A67A1001D
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'##ModelId=5C8A67A1002D
Const gREGVALSYSINFOLOC = "MSINFO"
'##ModelId=5C8A67A1004C
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'##ModelId=5C8A67A1006B
Const gREGVALSYSINFO = "PATH"

'##ModelId=5C8A67A1008B
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'##ModelId=5C8A67A100F8
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'##ModelId=5C8A67A10175
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'##ModelId=5C8A67A101A4
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'

'##ModelId=5C8A67A10230
Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

'##ModelId=5C8A67A10240
Private Sub cmdOk_Click()
  Unload Me
End Sub

'##ModelId=5C8A67A10250
Private Sub Form_Load()
    Me.Caption = "A propos de " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & ".0." & App.Revision & " - © 1997-2023 Actuaries Services"
    
    lblTitle.Caption = "Provisions Techniques Incapacités Invalidités (P3I)"
    'lblDescription.Caption = vbLf & "Cabinet Moeglin " & vbLf & "Tel : 01.45.92.31.41" & vbLf & "Fax : 01.43.03.77.76" & vbLf & "www.moeglin.com"
    lblDescription.Visible = False
    
    'lblDB = "Base de données : " & DatabaseFileName
    lblDB = "       Actuaries Services" & vbLf _
            & "        Service commercial: sales@actuariesservices.com" & vbLf _
            & "        Support technique: support@actuariesservices.com"
    
    lblLicense = GetAboutString
End Sub

'##ModelId=5C8A67A1026F
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

'##ModelId=5C8A67A1027F
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(mID(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(mID(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

'##ModelId=5C8A67A102CD
Public Sub HyperJump(ByVal URL As String)
  Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub

'##ModelId=5C8A67A102EC
Private Sub Label1_Click()
  HyperJump Label1.Caption
End Sub

'##ModelId=5C8A67A1030B
Private Sub picIcon_DblClick()
  Dim Msg As String, SASP3I As String
  
  SASP3I = GetSettingIni(SectionName, "DB", "SASP3IConnectionString", "#")

  
  Msg = "-Environnement de P3I:" & vbLf
  Msg = Msg & "    Database: " & DatabaseFileName & vbLf
  Msg = Msg & "    SASP3I: " & SASP3I & vbLf
  Msg = Msg & "    INI: " & sFichierIni & vbLf
  
  Dim wsh As Object
  
  Set wsh = CreateObject("WScript.Network")
  
  Msg = Msg & vbLf & "-Environnement réseau:" & vbLf
  Msg = Msg & "    Ordinateur: " & wsh.ComputerName & vbLf
  Msg = Msg & "    Domaine: " & wsh.UserDomain & vbLf
  Msg = Msg & "    Utilisateur: " & wsh.UserName & vbLf
  
  Set wsh = Nothing
  
  MsgBox Msg, vbInformation
End Sub


