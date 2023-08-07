VERSION 5.00
Begin VB.Form frmDetailAffichagePeriode 
   Caption         =   "Détails à afficher"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSignalisation 
      Caption         =   "Lignes de détail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   135
      TabIndex        =   2
      Top             =   105
      Width           =   2445
      Begin VB.CheckBox chkResteAAmortir 
         Caption         =   "Montant restant à amortir"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   1575
         Width           =   2220
      End
      Begin VB.CheckBox chkDejaAmorti 
         Caption         =   "Montant déjà amorti"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   1260
         Width           =   1860
      End
      Begin VB.CheckBox chkEcart 
         Caption         =   "Ecart"
         Height          =   240
         Left            =   135
         TabIndex        =   7
         Top             =   945
         Width           =   1860
      End
      Begin VB.CheckBox chkApres 
         Caption         =   "Calcul après réforme"
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   630
         Width           =   1860
      End
      Begin VB.CheckBox chkAvant 
         Caption         =   "Calcul avant réforme"
         Height          =   240
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type d'affichage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   2445
      Begin VB.CheckBox chkDonneesBrutes 
         Caption         =   "Données brutes"
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Visible         =   0   'False
         Width           =   2085
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   225
      TabIndex        =   1
      Top             =   2265
      Width           =   1035
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Top             =   2265
      Width           =   1035
   End
End
Attribute VB_Name = "frmDetailAffichagePeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A682201E6"
Option Explicit

'##ModelId=5C8A68220300
Public m_TypePeriode As Integer
'##ModelId=5C8A68220321
Public m_DetailAffichagePeriode As clsDetailAffichagePeriode
'##ModelId=5C8A6822037D
Public ret_code As Integer

'##ModelId=5C8A6822039C
Private Sub CancelButton_Click()
  ret_code = 0
  Me.Hide
End Sub

'##ModelId=5C8A682203BB
Private Sub Form_Load()
  ret_code = 0
  
  chkDonneesBrutes.Value = IIf(m_DetailAffichagePeriode.DonneesBrutes, vbChecked, vbUnchecked)
  
  chkAvant.Value = IIf(m_DetailAffichagePeriode.Avant, vbChecked, vbUnchecked)
  chkApres.Value = IIf(m_DetailAffichagePeriode.Apres, vbChecked, vbUnchecked)
  chkEcart.Value = IIf(m_DetailAffichagePeriode.Ecart, vbChecked, vbUnchecked)
  chkDejaAmorti.Value = IIf(m_DetailAffichagePeriode.DejaAmorti, vbChecked, vbUnchecked)
  chkResteAAmortir.Value = IIf(m_DetailAffichagePeriode.ResteAAmortir, vbChecked, vbUnchecked)
  
  If (eProvisionRetraite <> m_TypePeriode And eProvisionRetraiteRevalo <> m_TypePeriode) Then
    chkAvant.Value = vbChecked
    chkApres.Value = vbUnchecked
    chkEcart.Value = vbUnchecked
    chkDejaAmorti.Value = vbUnchecked
    chkResteAAmortir.Value = vbUnchecked
    
    chkAvant.Enabled = False
    chkApres.Enabled = False
    chkEcart.Enabled = False
    chkDejaAmorti.Enabled = False
    chkResteAAmortir.Enabled = False
  End If
    
  chkAvant.Value = vbChecked
  chkAvant.Enabled = False
End Sub

'##ModelId=5C8A682203CB
Private Sub OKButton_Click()
  
  If (chkAvant.Value Or chkApres.Value Or chkEcart.Value Or chkDejaAmorti.Value Or chkResteAAmortir.Value) = False Then
    MsgBox "Vous devez sélectionner un type de lignes de détail !", vbCritical
    Exit Sub
  End If
  
'  If (chkAvant.Value Or chkApres.Value) = False Then
'    MsgBox "Vous devez sélectionner un type de lignes de données (Avant ou Après) !", vbCritical
'    Exit Sub
'  End If
  
  m_DetailAffichagePeriode.DonneesBrutes = chkDonneesBrutes.Value
  
  m_DetailAffichagePeriode.Avant = (chkAvant.Value = vbChecked)
  m_DetailAffichagePeriode.Apres = (chkApres.Value = vbChecked)
  m_DetailAffichagePeriode.Ecart = (chkEcart.Value = vbChecked)
  m_DetailAffichagePeriode.DejaAmorti = (chkDejaAmorti.Value = vbChecked)
  m_DetailAffichagePeriode.ResteAAmortir = (chkResteAAmortir.Value = vbChecked)

  ret_code = 1
  Me.Hide
  
End Sub
