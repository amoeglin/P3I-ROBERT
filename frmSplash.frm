VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   -90
      Picture         =   "frmSplash.frx":00FA
      ScaleHeight     =   3600
      ScaleWidth      =   6300
      TabIndex        =   0
      Top             =   -90
      Width           =   6360
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E20189"
Option Explicit

'##ModelId=5C8A67E20263
Private Sub Form_Load()
  Picture1.Left = 0
  Picture1.top = 0
  Me.Width = Picture1.Width - 30
  Me.Height = Picture1.Height - 30
  
  'On Error Resume Next
  'Picture1.Picture = LoadPicture("Splash.bmp")
  'On Error GoTo 0
End Sub

