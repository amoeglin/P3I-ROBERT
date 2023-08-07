VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportTable 
   Caption         =   "Import de table"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFilename 
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1035
      Width           =   4110
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "..."
      Height          =   330
      Left            =   4275
      TabIndex        =   4
      Top             =   1035
      Width           =   330
   End
   Begin VB.CommandButton btnLoadData 
      Caption         =   ">>"
      Height          =   330
      Left            =   4275
      TabIndex        =   7
      Top             =   1710
      Width           =   330
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4050
      Top             =   2925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Fichier Excel à importer"
      FileName        =   "*.xls"
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -90
      TabIndex        =   11
      Top             =   2790
      Width           =   4875
   End
   Begin VB.TextBox txtNomTable 
      Height          =   315
      Left            =   90
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1710
      Width           =   4110
   End
   Begin VB.TextBox txtLibelleTable 
      Height          =   315
      Left            =   90
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2385
      Width           =   4515
   End
   Begin VB.ComboBox cboTypeTable 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4515
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Importer"
      Height          =   375
      Left            =   1170
      TabIndex        =   10
      Top             =   2970
      Width           =   1125
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2430
      TabIndex        =   12
      Top             =   2970
      Width           =   1080
   End
   Begin VB.Label Label4 
      Caption         =   "Fichier Excel"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   765
      Width           =   2310
   End
   Begin VB.Label Label3 
      Caption         =   "Nom de la table et Nom de la zone de données dans Excel"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1440
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "Description de la table "
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   2115
      Width           =   2310
   End
   Begin VB.Label Label1 
      Caption         =   "Type de table"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1095
   End
End
Attribute VB_Name = "frmImportTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6827012B"
Option Explicit

'##ModelId=5C8A68270235
Public ret_code As Integer

'##ModelId=5C8A68270244
Private Sub btnBrowse_Click()
    
  ' demande le nom du fichier xls
  CommonDialog1.filename = IIf(txtFilename <> "", "*.xls", txtFilename)
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Import d'une table..."
  CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  
  If cboTypeTable.ListIndex <> -1 Then
    If cboTypeTable.ItemData(cboTypeTable.ListIndex) = cdTypeTable_BaremeAnneeStatutaire Then
      CommonDialog1.filename = IIf(txtFilename <> "", "*.csv", txtFilename)
      CommonDialog1.DefaultExt = ".csv"
      CommonDialog1.DialogTitle = "Import d'une table..."
      CommonDialog1.filter = "Fichiers CSV|*.csv"
      CommonDialog1.InitDir = CSVUNCPath
    End If
  End If
  
  CommonDialog1.ShowOpen
  
  If CommonDialog1.filename = "" _
     Or CommonDialog1.filename = "*.xls" _
     Or CommonDialog1.filename = "*.csv" _
     Or CommonDialog1.filename = "*.xlsx" _
     Or CommonDialog1.filename = "*.*" Then
    Exit Sub
  End If

  txtFilename = CommonDialog1.filename

End Sub

'##ModelId=5C8A68270254
Private Sub btnLoadData_Click()
  Dim idTable As Long
  
  idTable = m_dataHelper.GetParameterAsLong("SELECT TABLECLE FROM ListeTableLoi WHERE NOMTABLE='" & Trim(txtNomTable) & "'")
  If idTable <> 0 Then
    Dim cListe As clsListeTableLoi
    
    Set cListe = New clsListeTableLoi
    
    cListe.Load m_dataSource, idTable
    
    If cListe.m_TableUtilisateur = False Then
      MsgBox "La table '" & txtNomTable & "' est une table système et ne pas être modifiée !", vbExclamation, "Import table..."
      Exit Sub
    End If
        
    txtLibelleTable = cListe.m_LIBTABLE
    
    Dim i As Integer
    
    For i = 0 To cboTypeTable.ListCount - 1
      If cboTypeTable.ItemData(i) = cListe.m_TYPETABLE Then
        cboTypeTable.ListIndex = i
        Exit For
      End If
    Next
    
  End If
End Sub


'##ModelId=5C8A68270273
Private Sub CancelButton_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A68270283
Private Sub add_type(nom As String, code As Integer)
  cboTypeTable.AddItem nom
  cboTypeTable.ItemData(cboTypeTable.ListCount - 1) = code
End Sub

'##ModelId=5C8A682702B2
Private Sub cboTypeTable_Click()

  If cboTypeTable.ListIndex <> -1 Then
    If cboTypeTable.ItemData(cboTypeTable.ListIndex) = cdTypeTable_BaremeAnneeStatutaire Then
      Label3 = "Nom de la table"
      Label4 = "Fichier CSV"
    Else
      Label3 = "Nom de la table et Nom de la zone de données dans Excel"
      Label4 = "Fichier Excel"
    End If
  End If
  
End Sub

'##ModelId=5C8A682702D1
Private Sub Form_Load()
  cboTypeTable.Clear
  
  add_type "Loi de maintien en incapacité", cdTypeTable_LoiMaintienIncapacite
  add_type "Loi de passage d'incapacité en invalidité", cdTypeTable_LoiPassage
  add_type "Loi de maintien en invalidité", cdTypeTable_LoiMaintienInvalidite
  add_type "Loi de maintien en DEPENDANCE", cdTypeTable_LoiDependance
  add_type "Table de mortalité", cdTypeTableMortalite
  add_type "Table de génération", cdTypeTableGeneration
  add_type "Mortalité des personnes en incapacité", cdTypeTable_MortaliteIncap
  add_type "Mortalité des personnes en invalidité", cdTypeTable_MortaliteInval
  
  add_type "Risque STATUTAIRE", cdTypeTable_BaremeAnneeStatutaire
  

  
  txtFilename = ""
  txtLibelleTable = ""
  txtNomTable = ""
End Sub

'##ModelId=5C8A682702E0
Private Sub OKButton_Click()
  
  If cboTypeTable.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner un type de table", vbCritical
    cboTypeTable.SetFocus
    Exit Sub
  End If
  
  If Trim(txtNomTable) = "" Or ValidateNomTable(txtNomTable) = False Then
    MsgBox "Vous devez saisir un nom de table. Les caractères autorisés sont : '-', '_', 'A-Z', 'a-z', '0-9'.", vbCritical
    txtNomTable.SetFocus
    Exit Sub
  End If
  
  If Trim(txtLibelleTable) = "" Then
    MsgBox "Vous devez saisir une description pour cette table.", vbCritical
    txtLibelleTable.SetFocus
    Exit Sub
  End If
  
  If Trim(txtFilename) = "" Then
    MsgBox "Vous devez sélectionner un fichier.", vbCritical
    txtFilename.SetFocus
    Exit Sub
  End If
  
  txtNomTable = Trim(txtNomTable)
  txtLibelleTable = Trim(txtLibelleTable)
  txtFilename = Trim(txtFilename)
  
  ret_code = 0
  Me.Hide
  
End Sub

'##ModelId=5C8A68270300
Private Function ValidateNomTable(nom As String) As Boolean
  Dim Position As Integer
  Dim Character As String
  Dim AsciiCharacter As Integer
  
  For Position = 1 To Len(nom)

    Character = UCase(mID(nom, Position, 1))
    AsciiCharacter = Asc(Character)

    If Character = "_" Or Character = "." _
       Or (AsciiCharacter >= Asc("A") And AsciiCharacter <= Asc("Z")) _
       Or (AsciiCharacter >= Asc("0") And AsciiCharacter <= Asc("9")) Then

    Else 'otherwise, invalid character
      ValidateNomTable = False
      Exit Function
    End If

  Next

  ValidateNomTable = True
End Function
