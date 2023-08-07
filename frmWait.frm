VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   1755
      TabIndex        =   3
      Top             =   1170
      Width           =   1410
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Veuillez Patienter..."
      Height          =   240
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   4830
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Opération en cours"
      Height          =   240
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4830
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fTravailAnnule As Boolean
Public hideStop As Boolean

' Variable d'affichage
Public sText As String
Public nValue As Integer

' constante pour hWndInsertAfter
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1

' constante pour wFlags
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
                 ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                 ByVal wFlags As Long) As Long
                 
Private Sub Command1_Click()
  fTravailAnnule = True
End Sub

Private Sub Form_Load()
  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2 - Height

  fTravailAnnule = False
  
  If hideStop = True Then
    Command1.Visible = False
  End If
  
  SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

