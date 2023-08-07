VERSION 5.00
Begin VB.Form frmMenu0 
   BackColor       =   &H00FFFF00&
   Caption         =   "Menu0    Provisions Incapacité Invalidité"
   ClientHeight    =   6120
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox edtTauxTechnique 
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton cmdPurges 
      Caption         =   "&Purges"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   3680
      Width           =   2895
   End
   Begin VB.CommandButton cmdEdition 
      Caption         =   "&Editions"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2200
      Width           =   2895
   End
   Begin VB.CommandButton cmdExtraction 
      Caption         =   "&Extraction des données et calcul des provisions"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   5175
   End
   Begin VB.CommandButton cmdQuitter 
      Caption         =   "&Quitter"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Menu mnuExtraction 
      Caption         =   "&Extraction"
   End
   Begin VB.Menu mnuEdition 
      Caption         =   "&Editions"
   End
   Begin VB.Menu mnuPurges 
      Caption         =   "&Purges"
   End
   Begin VB.Menu mnuBaremes 
      Caption         =   "&Barèmes"
   End
   Begin VB.Menu mnuQuitter 
      Caption         =   "&Quitter"
   End
End
Attribute VB_Name = "frmMenu0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExtraction_Click()
Dim db As Database
Dim rs As Recordset

Set db = OpenDatabase
Set rsEffectif = db.OpenRecordset("", dbOpenSnapshot)
Set rsProvision = db.OpenRecordset("", dbOpenDynaset)

While Not rs.EOF
    ' calcul

    ' next
    rs.MoveNext
Wend

' k est le n° du champs
nom = "Champs" & k

rs.AddNew

toto = rs.Fields(nom)

rs.Fields(nom) = toto

rs.Update
''''''''''''''''''''''''''''''''
    i = edtTauxTechnique ' Worksheets("paramètres").Cells(11, "C").Value            ' taux technique
    wlig = 4
    wcol = 8
    xdeb = 30
    xfin = 30
    x = 0
    anc = 0
    franchise = 0
    k = 0
    
    PT = 0
    v = 1 / (1 + i)
    

For x = xdeb To xfin
    
    
    ''''''''''''''''''''
    For anc = 0 To xfin - xdeb
    ''''''''''''''''''''
    
    PT = 0
    
    For k = anc To 60 - anc - 1
     
    If x + k = 60 Then GoTo FIN_CALCUL_X

    Lanc = Worksheets("invalidité").Cells(x, anc + 2).Value
    If Lanc = 0 Then GoTo FIN_CALCUL_X
    
    Lk = Worksheets("invalidité").Cells(x, k + 2).Value  '
    Lk1 = Worksheets("invalidité").Cells(x, k + 3).Value
        
    
    PT = PT + (1 / (2 * Lanc)) * (Lk * (v ^ (k - anc)) + Lk1 * (v ^ (k + 1 - anc)))
    
    Next k

FIN_CALCUL_X:
    
    Worksheets("provisions invalidité").Cells(x, anc + 2).Value = PT

    '''''''''''''''''''
    Next anc
    '''''''''''''''''''
    

Next x

End Sub

Private Sub cmdQuitter_Click()
Beep
End
End Sub

Private Sub edtTauxTechnique_Change()

End Sub

Private Sub Form_Load()

End Sub

Private Sub mnuBaremes_Click()

End Sub

Private Sub mnuQuitter_Click()

End Sub

Private Sub Text1_Change()

End Sub
