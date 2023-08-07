VERSION 5.00
Begin VB.Form frmAutoLog 
   Caption         =   "Logs"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   5550
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   6255
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Afficher Log"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Supprimer Log"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
End
Attribute VB_Name = "frmAutoLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6843012B"

'##ModelId=5C8A68430225
Private Sub cmdClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A68430244
Private Sub cmdDisplay_Click()

  Dim frm As New frmDisplayLog
  frm.FichierLog = m_logPathAuto & File1.filename
  frm.Show vbModal
  Set frm = Nothing
  
End Sub

'##ModelId=5C8A68430254
Private Sub Form_Load()

  If Dir(m_logPathAuto) = "" Then
    MkDir m_logPathAuto
  End If

  File1.Path = m_logPathAuto
End Sub

'##ModelId=5C8A68430273
Private Sub cmdDelete_Click()

  If MsgBox("Est-ce que vous êtes sur de vouloir supprimer le fichier sélectionné ?", vbYesNo) = vbYes Then
    Kill m_logPathAuto & File1.filename
  End If

End Sub

