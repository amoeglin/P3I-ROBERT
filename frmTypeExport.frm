VERSION 5.00
Begin VB.Form frmTypeExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export des données vers l'infocentre"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmSignalisation 
      Caption         =   "Fichier de signalisation"
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
      Left            =   90
      TabIndex        =   9
      Top             =   2430
      Width           =   4110
      Begin VB.CommandButton btnForceTop 
         Caption         =   "Générer"
         Default         =   -1  'True
         Height          =   330
         Left            =   2970
         TabIndex        =   11
         Top             =   270
         Width           =   945
      End
      Begin VB.CheckBox chkCreateSignalisation 
         Caption         =   "Créer le fichier de signalisation"
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   2535
      End
   End
   Begin VB.Frame frmModification 
      Caption         =   "Modification de l'existant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   4110
      Begin VB.OptionButton rdoAddExistant 
         Caption         =   "Ajouter aux données présentes dans les tables TTLOGTRAIT et TTPROVCOLL"
         Height          =   375
         Left            =   135
         TabIndex        =   1
         Top             =   810
         Width           =   3885
      End
      Begin VB.OptionButton rdoDelExistant 
         Caption         =   "Ecraser les données présentes dans les tables TTLOGTRAIT et TTPROVCOLL"
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   3885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type de Provision"
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
      Left            =   90
      TabIndex        =   2
      Top             =   1575
      Width           =   4110
      Begin VB.OptionButton rdoClient 
         Caption         =   "Client"
         Height          =   285
         Left            =   1485
         TabIndex        =   4
         Top             =   315
         Width           =   1005
      End
      Begin VB.OptionButton rdoSimul 
         Caption         =   "Simulation"
         Height          =   285
         Left            =   2835
         TabIndex        =   5
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton rdoBilan 
         Caption         =   "Bilan"
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   1140
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2205
      TabIndex        =   8
      Top             =   3375
      Width           =   1170
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   765
      TabIndex        =   7
      Top             =   3375
      Width           =   1215
   End
End
Attribute VB_Name = "frmTypeExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A681902A5"

Option Explicit

'##ModelId=5C8A681903CE
Public frmTypeProvision As Integer
'##ModelId=5C8A681A0005
Public frmDelExistant As Boolean
'##ModelId=5C8A681A0015
Public frmCreateSignalisation As Boolean
'

'##ModelId=5C8A681A0034
Private Sub btnForceTop_Click()
  CreationFichierSignalisation
End Sub

'##ModelId=5C8A681A0054
Private Sub CancelButton_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A681A0063
Private Sub Form_Load()
  rdoDelExistant.Value = True
  rdoBilan.Value = True
  
  Select Case frmTypeProvision
    Case 1
      rdoBilan.Value = True
    Case 2
      rdoClient.Value = True
    Case 3
      rdoSimul.Value = True
    'Case Else
    '  MsgBox "Vous devez sélectionner un type de provision valide !", vbCritical
  End Select
  
  frmDelExistant = True
  frmCreateSignalisation = True
  
End Sub

'##ModelId=5C8A681A0073
Private Sub OKButton_Click()
  Dim typeChoisi As Integer
  
  If (rdoBilan.Value = True) Then
    typeChoisi = 1
  ElseIf (rdoClient.Value = True) Then
    typeChoisi = 2
  ElseIf (rdoSimul.Value = True) Then
    typeChoisi = 3
  Else
    MsgBox "Vous devez sélectionner un type de provision !", vbCritical
    Exit Sub
  End If
  
  If frmTypeProvision <> typeChoisi Then
    'If MsgBox("Vous avez sélectionné un type de provision différent de celui de la période !" & vbLf & "Voulez-vous continuer ?", vbYesNo + vbQuestion) = vbNo Then
    '  Exit Sub
    'End If
  End If
  
  frmTypeProvision = typeChoisi
  
  If rdoDelExistant.Value = True Then
    frmDelExistant = True
  ElseIf rdoAddExistant.Value = True Then
    frmDelExistant = False
  Else
    MsgBox "Vous devez sélectionner un type de export !", vbCritical
    Exit Sub
  End If
    
      
  If chkCreateSignalisation.Value = vbChecked Then
    frmCreateSignalisation = True
  Else
    frmCreateSignalisation = False
  End If
    
    
  ret_code = 0
  Me.Hide
End Sub
