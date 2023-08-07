VERSION 5.00
Begin VB.Form frmChoixControle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contrôle du lôt de données n°"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRenteEduc 
      Caption         =   "Age <=26 ans pour Garantie Rente Education"
      Height          =   195
      Left            =   4770
      TabIndex        =   19
      Top             =   2250
      Width           =   4110
   End
   Begin VB.CheckBox chkDateInvalSurv 
      Caption         =   "Date de mise en invalidité < Date de Survenance"
      Height          =   195
      Left            =   4770
      TabIndex        =   18
      Top             =   1890
      Width           =   4110
   End
   Begin VB.CheckBox chkChoixPrest 
      Caption         =   "CDCHOIXPREST inconnu dans CODECATINV"
      Height          =   195
      Left            =   4770
      TabIndex        =   17
      Top             =   1530
      Width           =   4110
   End
   Begin VB.CheckBox chkCATR9 
      Caption         =   "Catégories dans CATR9 ou CATR9INVAL"
      Height          =   195
      Left            =   4770
      TabIndex        =   16
      Top             =   1170
      Width           =   4110
   End
   Begin VB.CheckBox chkNom 
      Caption         =   "Nom d'assuré non renseigné"
      Height          =   195
      Left            =   4770
      TabIndex        =   15
      Top             =   810
      Width           =   4110
   End
   Begin VB.CheckBox chkCDPRODUIT 
      Caption         =   "CDPRODUIT ou NumParamCalcul inconnu dans CODESCAT"
      Height          =   195
      Left            =   270
      TabIndex        =   14
      Top             =   2970
      Width           =   4740
   End
   Begin VB.CheckBox chkCapitauxConstitutif 
      Caption         =   "Liste des Capitaux Constitutif GA < 2006"
      Height          =   195
      Left            =   270
      TabIndex        =   13
      Top             =   3330
      Width           =   4110
   End
   Begin VB.CommandButton btnTous 
      Caption         =   "Tous"
      Height          =   285
      Left            =   135
      TabIndex        =   12
      Top             =   3645
      Width           =   555
   End
   Begin VB.CheckBox chkDoublon 
      Caption         =   "Recherche des doublons"
      Height          =   195
      Left            =   4770
      TabIndex        =   11
      Top             =   450
      Width           =   4110
   End
   Begin VB.CheckBox chkAgeRente 
      Caption         =   "Age >=18 ans et Age <=65 ans pour Garantie Hors Rente Education"
      Height          =   195
      Left            =   270
      TabIndex        =   10
      Top             =   2610
      Width           =   5235
   End
   Begin VB.CheckBox chkNaiss1900 
      Caption         =   "Date de Naissance < 01/01/1900"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   2250
      Width           =   4110
   End
   Begin VB.CheckBox chkDateNaissSurv 
      Caption         =   "Date de Naissance < Date de Survenance"
      Height          =   195
      Left            =   270
      TabIndex        =   8
      Top             =   1890
      Width           =   4110
   End
   Begin VB.CheckBox chkMontant 
      Caption         =   "Montant Prestation >  sz"
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   1530
      Width           =   4110
   End
   Begin VB.CheckBox chkDate 
      Caption         =   "Dates incorrectes"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1170
      Width           =   4110
   End
   Begin VB.CheckBox chkCodeGE 
      Caption         =   "Code GE absent de TBQREGA"
      Height          =   195
      Left            =   270
      TabIndex        =   4
      Top             =   810
      Width           =   4110
   End
   Begin VB.CheckBox chkCodeProv 
      Caption         =   "Code_Prov invalide dans TBQREGA"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   450
      Width           =   4110
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   9210
      TabIndex        =   0
      Top             =   3945
      Width           =   9210
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Annuler"
         Height          =   345
         Left            =   2393
         TabIndex        =   2
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Contrôler"
         Default         =   -1  'True
         Height          =   345
         Left            =   1313
         TabIndex        =   1
         Top             =   45
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Veuillez sélectionner les contrôles à effectuer :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   4380
   End
End
Attribute VB_Name = "frmChoixControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public sFichierIni As String


Private Sub btnTous_Click()
  Dim v As Integer
  
  v = IIf(chkCodeProv.Value = vbChecked, vbUnchecked, vbChecked)
  
  chkAgeRente.Value = v
  chkCapitauxConstitutif.Value = v
  chkCodeGE.Value = v
  chkCodeProv.Value = v
  chkDate.Value = v
  chkDateNaissSurv.Value = v
  chkDateInvalSurv.Value = v
  chkDoublon.Value = v
  chkMontant.Value = v
  chkNaiss1900.Value = v
  chkCDPRODUIT.Value = v
  chkNom.Value = v
  chkCATR9.Value = v
  chkRenteEduc.Value = v
  chkChoixPrest.Value = v
End Sub

Private Sub cmdClose_Click()
  ret_code = -1
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
  
  If chkAgeRente.Value = vbUnchecked _
     And chkCapitauxConstitutif.Value = vbUnchecked _
     And chkCodeGE.Value = vbUnchecked _
     And chkCodeProv.Value = vbUnchecked _
     And chkDate.Value = vbUnchecked _
     And chkDateNaissSurv.Value = vbUnchecked _
     And chkDateInvalSurv.Value = vbUnchecked _
     And chkDoublon.Value = vbUnchecked _
     And chkMontant.Value = vbUnchecked _
     And chkNaiss1900.Value = vbUnchecked _
     And chkCDPRODUIT.Value = vbUnchecked _
     And chkCATR9.Value = vbUnchecked _
     And chkRenteEduc.Value = vbUnchecked _
     And chkChoixPrest.Value = vbUnchecked _
     And chkNom.Value = vbUnchecked Then
    MsgBox "Vous devez sélectionner un contrôle !", vbCritical + vbOKOnly
    Exit Sub
  End If
  
  ret_code = 0
  Me.Hide
End Sub

Private Sub Form_Load()
  Dim sz As String
  
  sz = sReadIniFile("P3I", "ControleMontantMax", "100000", 20, sFichierIni)
  chkMontant.Caption = "Montant Prestation <= 0 ou > " & sz
  
  chkAgeRente.Value = vbChecked
  chkCapitauxConstitutif.Value = vbChecked
  chkCodeGE.Value = vbChecked
  chkCodeProv.Value = vbChecked
  chkDate.Value = vbChecked
  chkDateNaissSurv.Value = vbChecked
  chkDateInvalSurv.Value = vbChecked
  chkDoublon.Value = vbChecked
  chkMontant.Value = vbChecked
  chkNaiss1900.Value = vbChecked
  chkCDPRODUIT.Value = vbChecked
  chkCATR9.Value = vbChecked
  chkChoixPrest.Value = vbChecked
  chkNom.Value = vbChecked
  chkRenteEduc.Value = vbChecked
End Sub
