VERSION 5.00
Begin VB.Form frmProvOuverture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Provisions à l'ouverture"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmProvOuverture.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Enregistrer"
      Height          =   345
      Left            =   1575
      TabIndex        =   15
      Top             =   3285
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Fermer"
      Height          =   345
      Left            =   2700
      TabIndex        =   14
      Top             =   3285
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   13
      Top             =   3150
      Width           =   5460
   End
   Begin VB.TextBox txtAnN5 
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Top             =   2790
      Width           =   1300
   End
   Begin VB.TextBox txtAnN4 
      Height          =   285
      Left            =   2520
      TabIndex        =   10
      Top             =   2430
      Width           =   1300
   End
   Begin VB.TextBox txtAnN3 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   2070
      Width           =   1300
   End
   Begin VB.TextBox txtAnN2 
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Top             =   1710
      Width           =   1300
   End
   Begin VB.TextBox txtAnN1 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   1350
      Width           =   1300
   End
   Begin VB.TextBox txtAnN 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   990
      Width           =   1300
   End
   Begin VB.Label Label8 
      Caption         =   "et antérieure"
      Height          =   240
      Left            =   3915
      TabIndex        =   17
      Top             =   2835
      Width           =   1050
   End
   Begin VB.Label Label7 
      Caption         =   "Veuillez entrer les montants des provisions à l'ouverture pour les années :"
      Height          =   285
      Left            =   45
      TabIndex        =   16
      Top             =   630
      Width           =   5370
   End
   Begin VB.Label lblAnN5 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   11
      Top             =   2835
      Width           =   1050
   End
   Begin VB.Label lblAnN4 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   9
      Top             =   2475
      Width           =   1050
   End
   Begin VB.Label lblAnN3 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   7
      Top             =   2115
      Width           =   1050
   End
   Begin VB.Label lblAnN2 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   5
      Top             =   1755
      Width           =   1050
   End
   Begin VB.Label lblAnN1 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   3
      Top             =   1395
      Width           =   1050
   End
   Begin VB.Label lblAnN 
      Alignment       =   1  'Right Justify
      Caption         =   "Année N"
      Height          =   240
      Left            =   1395
      TabIndex        =   1
      Top             =   1035
      Width           =   1050
   End
   Begin VB.Label lblGroupe 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Groupe sdfd dfg dfg  dfg d fg d fg dfgdfg df g df g dfg d g df gdfg d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5370
   End
End
Attribute VB_Name = "frmProvOuverture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67DE037D"
Option Explicit

'##ModelId=5C8A67DF006F
Public NumSte As Long

'##ModelId=5C8A67DF008F
Private Sub cmdClose_Click()
  Unload Me
End Sub

'##ModelId=5C8A67DF009E
Private Sub cmdUpdate_Click()
  ' charge les valeurs
  Dim rs As ADODB.Recordset
  
  On Error GoTo errUpdateProv
  
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM ProvisionsOuverture WHERE GPECLE = " & GroupeCle & " AND NUMCLE = " & numPeriode & " AND POSTECLE=" & NumSte, Dynamic)
  If rs.EOF Then
    rs.AddNew
  End If
  
  rs.fields("GPECLE").Value = GroupeCle
  rs.fields("NUMCLE").Value = numPeriode
  rs.fields("POSTECLE").Value = NumSte
  
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn"), txtAnN)
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn1"), txtAnN1)
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn2"), txtAnN2)
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn3"), txtAnN3)
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn4"), txtAnN4)
  Call m_dataHelper.GetDouble(rs.fields("PROV_ANn5"), txtAnN5)
  
  rs.Update
  
  rs.Close
  
  Unload Me
  
  On Error GoTo 0
    
  Exit Sub
  
errUpdateProv:
  On Error GoTo 0
  
  MsgBox "Erreur lors de la sauvegarde : " & vbLf & Err.Description
End Sub

'##ModelId=5C8A67DF00AE
Private Sub Form_Load()
  Dim df As Integer, ns As String
    
  df = Year(m_dataHelper.GetParameter("SELECT PEDATEFIN FROM Periode WHERE PEGPECLE=" & GroupeCle & " AND PENUMCLE=" & numPeriode))

  lblAnN = "Année " & df
  lblAnN1 = "Année " & df - 1
  lblAnN2 = "Année " & df - 2
  lblAnN3 = "Année " & df - 3
  lblAnN4 = "Année " & df - 4
  lblAnN5 = "Année " & df - 5
  
  ns = m_dataHelper.GetParameter("SELECT SONOM FROM Societe WHERE SOGROUPE = " & GroupeCle & " AND SOCLE = " & NumSte)
  
  lblGroupe = "Période n° " & numPeriode & " du Groupe '" & NomGroupe & "'" & vbLf & "pour la Société '" & ns & "'"
  
  ' charge les valeurs
  Dim rs As ADODB.Recordset
  
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM ProvisionsOuverture WHERE GPECLE = " & GroupeCle & " AND NUMCLE = " & numPeriode & " AND POSTECLE=" & NumSte, Snapshot)
  If Not rs.EOF Then
    txtAnN = rs.fields("PROV_ANn").Value
    txtAnN1 = rs.fields("PROV_ANn1").Value
    txtAnN2 = rs.fields("PROV_ANn2").Value
    txtAnN3 = rs.fields("PROV_ANn3").Value
    txtAnN4 = rs.fields("PROV_ANn4").Value
    txtAnN5 = rs.fields("PROV_ANn5").Value
  End If
  rs.Close
End Sub
