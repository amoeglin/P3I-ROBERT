VERSION 5.00
Begin VB.Form frmSelectPurge 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purge"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "selectPurge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPurger 
      Caption         =   "&Purger"
      Height          =   375
      Left            =   180
      TabIndex        =   4
      Top             =   2790
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2790
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   0
      TabIndex        =   2
      Top             =   2655
      Width           =   4695
   End
   Begin VB.ListBox lstDate 
      Height          =   2205
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   4605
   End
   Begin VB.Label Label1 
      Caption         =   "Sélectionnez la date d'import des enregistrements à purger :"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4335
   End
End
Attribute VB_Name = "frmSelectPurge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E40190"
Option Explicit

'##ModelId=5C8A67E4028A
Private Sub btnCancel_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A67E402AA
Private Sub btnPurger_Click()
  If lstDate.ListIndex = -1 Then Exit Sub
  
  If MsgBox("Les enregistrements importés le " & lstDate.List(lstDate.ListIndex) & " vont être effacés définitivement." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) = vbYes Then
    If MsgBox("2ème COMFIRMATION : ETES-VOUS SUR DE VOULOIR SUPPRIMER LES ENREGISTREMENTS IMPORTES LE " & lstDate.List(lstDate.ListIndex) & " ?", vbQuestion + vbYesNo) = vbYes Then
      ret_code = 0
      Me.Hide
      Exit Sub
    End If
  End If
  
  ret_code = -1
End Sub

'##ModelId=5C8A67E402B9
Private Sub Form_Activate()
  ret_code = -1
  
  If lstDate.ListCount = 0 Then
    btnPurger.Enabled = False
  End If
End Sub

