VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelectDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import des donn�es..."
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   Icon            =   "SelectDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   -45
      TabIndex        =   9
      Top             =   3555
      Width           =   5685
      Begin VB.OptionButton rdoEnsemblePaiement 
         Caption         =   "Ensemble des paiements"
         Height          =   285
         Left            =   945
         TabIndex        =   18
         Top             =   450
         Width           =   2085
      End
      Begin VB.OptionButton rdoDernierPaiement 
         Caption         =   "Dernier paiement"
         Height          =   285
         Left            =   3105
         TabIndex        =   17
         Top             =   450
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "- Calcul de l'ANNUALISATION du montant r�gl� :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   180
         Width           =   5325
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   -45
      TabIndex        =   8
      Top             =   1350
      Width           =   5685
      Begin VB.OptionButton rdoImportComplet 
         Caption         =   "Import Complet"
         Height          =   285
         Left            =   285
         TabIndex        =   12
         Top             =   990
         Width           =   1455
      End
      Begin VB.OptionButton rdoImportDonneesSeules 
         Caption         =   "Donn�es seules"
         Height          =   285
         Left            =   1815
         TabIndex        =   11
         Top             =   990
         Width           =   1500
      End
      Begin VB.OptionButton rdoImportTableParametre 
         Caption         =   "Tables de param�trage"
         Height          =   285
         Left            =   3390
         TabIndex        =   10
         Top             =   990
         Width           =   1950
      End
      Begin VB.Label Label5 
         Caption         =   $"SelectDate.frx":1BB2
         Height          =   555
         Left            =   225
         TabIndex        =   21
         Top             =   405
         Width           =   5190
      End
      Begin VB.Label Label2 
         Caption         =   "- Importer les TABLES DE PARAMETRAGE avec les donn�es ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   180
         Width           =   5415
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5580
      TabIndex        =   4
      Top             =   4485
      Width           =   5580
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Importer"
         Default         =   -1  'True
         Height          =   345
         Left            =   3420
         TabIndex        =   2
         Top             =   45
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Annuler"
         Height          =   345
         Left            =   4500
         TabIndex        =   3
         Top             =   45
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   -45
      TabIndex        =   0
      Top             =   2520
      Width           =   5730
      Begin VB.OptionButton rdoDateFinPeriode 
         Caption         =   "Date de fin de p�riode (PeriodeAu)"
         Height          =   285
         Left            =   2700
         TabIndex        =   15
         Top             =   720
         Width           =   2760
      End
      Begin VB.OptionButton rdoDatePaiement 
         Caption         =   "Date de paiement (Paiement)"
         Height          =   285
         Left            =   225
         TabIndex        =   14
         Top             =   720
         Width           =   2400
      End
      Begin VB.Label lblDate2 
         Caption         =   "(PM=0 si nb jours entre cette date et la date d'arret� > 888 jours)"
         Height          =   240
         Left            =   225
         TabIndex        =   22
         Top             =   450
         Width           =   5100
      End
      Begin VB.Label lblDate 
         Caption         =   "- DATE � utiliser pour le CALCUL du DELAI D'INACTIVITE ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   5370
      End
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   2670
      TabIndex        =   1
      Top             =   945
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   367788035
      CurrentDate     =   36114
   End
   Begin VB.Label Label4 
      Caption         =   "Cette date doit �tre comprise entre le d�but et la fin de la p�riode en cours :"
      Height          =   240
      Left            =   90
      TabIndex        =   20
      Top             =   675
      Width           =   5370
   End
   Begin VB.Label Label1 
      Caption         =   "- Veuillez s�lectionner la DATE D'ARRETE des comptes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   450
      Width           =   5370
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date d'arret�"
      Height          =   255
      Index           =   4
      Left            =   1635
      TabIndex        =   6
      Top             =   990
      Width           =   1005
   End
   Begin VB.Label lblGroupe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Import des donn�es : Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   5370
   End
End
Attribute VB_Name = "frmSelectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gDateDebut As Date
Public gDateFin As Date

Private Sub cmdClose_Click()
  ret_code = -1
  Unload Me
End Sub

Private Sub cmdUpdate_Click()
  If DTPicker2 > CDate(gDateFin) Or DTPicker2 < CDate(gDateDebut) Then
    MsgBox "Vous devez choisir une date d'arret� incluse dans la p�riode !", vbExclamation
    DTPicker2.SetFocus
    Exit Sub
  End If
  
  If rdoImportComplet.Value = False _
     And rdoImportDonneesSeules.Value = False _
     And rdoImportTableParametre.Value = False Then
    MsgBox "Vous devez choisir le type de donn�es � importer !", vbExclamation
    rdoImportComplet.SetFocus
    Exit Sub
  End If
  
  If rdoDatePaiement.Value = False _
     And rdoDateFinPeriode.Value = False Then
    MsgBox "Vous devez choisir le type de dates de r�f�rence !", vbExclamation
    rdoDatePaiement.SetFocus
    Exit Sub
  End If
  
  If rdoEnsemblePaiement.Value = False _
     And rdoDernierPaiement.Value = False Then
    MsgBox "Vous devez choisir le type de calcul de l'annualisation !", vbExclamation
    rdoEnsemblePaiement.SetFocus
    Exit Sub
  End If
  
  ret_code = 0
  Me.Hide
End Sub

Private Sub Form_Activate()
  ' rempli le nom du groupe
  DTPicker2.MinDate = CDate(gDateDebut)
  DTPicker2.MaxDate = CDate(gDateFin)
  
  'rdoImportComplet.Value = True
  
  rdoDateFinPeriode.Value = True
  
  rdoDernierPaiement.Value = True
End Sub
