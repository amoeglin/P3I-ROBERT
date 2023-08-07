VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAnnexe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Annexes"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "frmAnnexe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNewGroupe 
      Caption         =   "&Nouveau"
      Height          =   330
      Left            =   105
      TabIndex        =   11
      Top             =   4680
      Width           =   1380
   End
   Begin VB.CommandButton btnDelGroupe 
      Caption         =   "&Supprimer"
      Height          =   330
      Left            =   1575
      TabIndex        =   10
      Top             =   4680
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   315
      Left            =   5670
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   8123
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Groupes"
      TabPicture(0)   =   "frmAnnexe.frx":1BB2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpreadGroupe"
      Tab(0).Control(1)=   "dtaGroupe"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Sociétés"
      TabPicture(1)   =   "frmAnnexe.frx":1BCE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtaSte"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboGroupeSte"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SpreadSte"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Utilisateurs"
      TabPicture(2)   =   "frmAnnexe.frx":1BEA
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cboGroupeUser"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "spreadUser"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "dtaUser"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Garanties"
      TabPicture(3)   =   "frmAnnexe.frx":1C06
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dtaGarantie"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "spreadGarantie"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cboGroupeGar"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboSocieteGar"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label3"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin MSAdodcLib.Adodc dtaGroupe 
         Height          =   330
         Left            =   -70455
         Top             =   4140
         Visible         =   0   'False
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dtaGroupe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc dtaSte 
         Height          =   330
         Left            =   -70455
         Top             =   4140
         Visible         =   0   'False
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT 1"
         Caption         =   "dtaSte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread vaSpreadGroupe 
         Height          =   4020
         Left            =   -74865
         TabIndex        =   14
         Top             =   450
         Width           =   6495
         _Version        =   524288
         _ExtentX        =   11456
         _ExtentY        =   7091
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   1000000
         SpreadDesigner  =   "frmAnnexe.frx":1C22
         VirtualMode     =   -1  'True
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaUser 
         Height          =   330
         Left            =   4275
         Top             =   4095
         Visible         =   0   'False
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM P3IUser.Utilisateur"
         Caption         =   "dtaUser"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread spreadUser 
         Height          =   3615
         Left            =   90
         TabIndex        =   13
         Top             =   855
         Width           =   6540
         _Version        =   524288
         _ExtentX        =   11536
         _ExtentY        =   6376
         _StockProps     =   64
         DAutoSizeCols   =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   1000000
         SpreadDesigner  =   "frmAnnexe.frx":20F3
         VirtualMode     =   -1  'True
         AppearanceStyle =   0
      End
      Begin MSAdodcLib.Adodc dtaGarantie 
         Height          =   330
         Left            =   -70905
         Top             =   4095
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SELECT * FROM P3IUser.Garantie"
         Caption         =   "dtaGarantie"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin FPSpreadADO.fpSpread spreadGarantie 
         Height          =   3255
         Left            =   -74865
         TabIndex        =   12
         Top             =   1215
         Width           =   6495
         _Version        =   524288
         _ExtentX        =   11456
         _ExtentY        =   5741
         _StockProps     =   64
         DAutoSizeCols   =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   1000000
         SpreadDesigner  =   "frmAnnexe.frx":265D
         VirtualMode     =   -1  'True
         AppearanceStyle =   0
      End
      Begin VB.ComboBox cboGroupeGar 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   480
         Width           =   3930
      End
      Begin VB.ComboBox cboSocieteGar 
         Height          =   315
         Left            =   -73800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   3930
      End
      Begin VB.ComboBox cboGroupeUser 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox cboGroupeSte 
         Height          =   315
         ItemData        =   "frmAnnexe.frx":2B6F
         Left            =   -73800
         List            =   "frmAnnexe.frx":2B71
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin FPSpreadADO.fpSpread SpreadSte 
         Height          =   3570
         Left            =   -74910
         TabIndex        =   15
         Top             =   900
         Width           =   6585
         _Version        =   524288
         _ExtentX        =   11615
         _ExtentY        =   6297
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   1000000
         SpreadDesigner  =   "frmAnnexe.frx":2B73
         VirtualMode     =   -1  'True
         AppearanceStyle =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Société :"
         Height          =   255
         Left            =   -74865
         TabIndex        =   7
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Groupe :"
         Height          =   255
         Left            =   -74865
         TabIndex        =   6
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Groupe :"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Groupe :"
         Height          =   255
         Left            =   -74820
         TabIndex        =   3
         Top             =   495
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmAnnexe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67AC01A5"
Option Explicit

'##ModelId=5C8A67AC0271
Private Sub btnDelGroupe_Click()
  Select Case SSTab1.Tab
    Case 0:  ' groupes
      If MsgBox("Le groupe '" & dtaGroupe.Recordset.fields("NOM") & "' va être supprimé définitivement." & vbLf & "Voulez vous vraiment le supprimer ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
      
      dtaGroupe.Recordset.Delete
      dtaGroupe.Refresh
      
      vaSpreadGroupe.VirtualMaxRows = -1
      vaSpreadGroupe.Refresh
      
    Case 1:  ' societes
      If cboGroupeSte.ListIndex = -1 Then Exit Sub
      
      If MsgBox("La société '" & dtaSte.Recordset.fields("SONOM") & "' va être supprimée définitivement." & vbLf & "Voulez vous vraiment la supprimer ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
      
      dtaSte.Recordset.Delete
      dtaSte.Recordset.Requery
      dtaSte.Refresh
      
      SpreadSte.VirtualMaxRows = -1
      SpreadSte.Refresh
      
    Case 2:  ' utilisateurs
      If cboGroupeUser.ListIndex = -1 Then Exit Sub
      
      If MsgBox("L'utilisateur '" & dtaUser.Recordset.fields("Nom") & "' va être supprimé définitivement." & vbLf & "Voulez vous vraiment le supprimer ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
      
      dtaUser.Recordset.Delete
      dtaUser.Refresh
      
      spreadUser.VirtualMaxRows = -1
      spreadUser.Refresh
          
    Case 3: ' Garanties
      If cboSocieteGar.ListIndex = -1 Or cboGroupeGar.ListIndex = -1 Then Exit Sub

      If MsgBox("La garantie '" & dtaGarantie.Recordset.fields("Nom") & "' va être supprimée définitivement." & vbLf & "Voulez vous vraiment la supprimer ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
      
      dtaGarantie.Recordset.Delete
      dtaGarantie.Refresh
      
      spreadGarantie.VirtualMaxRows = -1
      spreadGarantie.Refresh
  End Select
End Sub

'##ModelId=5C8A67AC0290
Private Sub btnNewGroupe_Click()
  Select Case SSTab1.Tab
    Case 0:  ' groupes
      dtaGroupe.Recordset.AddNew
      dtaGroupe.Recordset.fields("NOM") = "Nouveau Groupe"
      dtaGroupe.Recordset.Update
      dtaGroupe.Refresh
      
      vaSpreadGroupe.VirtualMaxRows = -1
      vaSpreadGroupe.Refresh
      
    Case 1:  ' societes
      If cboGroupeSte.ListIndex = -1 Then Exit Sub
      
      dtaSte.Recordset.AddNew
      'dtaSte.Recordset.fields("SOCLE") = m_dataHelper.GetParameterAsDouble("SELECT MAX(SOCLE) FROM Societe") + 1
      dtaSte.Recordset.fields("SOGROUPE") = cboGroupeSte.ItemData(cboGroupeSte.ListIndex)
      dtaSte.Recordset.fields("SONOM") = "Nouvelle Societe"
      dtaSte.Recordset.Update
      dtaSte.Refresh
      
      SpreadSte.VirtualMaxRows = -1
      SpreadSte.Refresh
      
    Case 2:  ' utilisateurs
      If cboGroupeUser.ListIndex = -1 Then Exit Sub
      
      dtaUser.Recordset.AddNew
      dtaUser.Recordset.fields("TAGPECLE") = cboGroupeUser.ItemData(cboGroupeUser.ListIndex)
      dtaUser.Recordset.fields("NOM") = "Nouvel Utilisateur"
      dtaUser.Recordset.fields("Mot de passe") = ""
      dtaUser.Recordset.fields("ID") = "NU"
      dtaUser.Recordset.fields("ADMIN") = False
      dtaUser.Recordset.Update
      dtaUser.Refresh
      
      spreadUser.VirtualMaxRows = -1
      spreadUser.Refresh
    
    Case 3:  ' garanties
      If cboSocieteGar.ListIndex = -1 Or cboGroupeGar.ListIndex = -1 Then Exit Sub
      
      dtaGarantie.Recordset.AddNew
      dtaGarantie.Recordset.fields("Groupe") = cboGroupeGar.ItemData(cboGroupeGar.ListIndex)
      dtaGarantie.Recordset.fields("Société") = cboSocieteGar.ItemData(cboSocieteGar.ListIndex)
      dtaGarantie.Recordset.fields("Nom") = "Nouvelle Garantie"
      'dtaGarantie.Recordset.Fields("GAGROUPCLE") = cboGroupeGar.ItemData(cboGroupeGar.ListIndex)
      'dtaGarantie.Recordset.Fields("GASTECLE") = cboSocieteGar.ItemData(cboSocieteGar.ListIndex)
      'dtaGarantie.Recordset.Fields("GALIB") = "Nouvelle Garantie"
      dtaGarantie.Recordset.Update
      dtaGarantie.Refresh
      
      spreadGarantie.VirtualMaxRows = -1
      spreadGarantie.Refresh
  End Select
End Sub

'##ModelId=5C8A67AC029F
Private Sub cboSocieteGar_Click()
  
  If cboSocieteGar.ListIndex <> -1 And cboGroupeGar.ListIndex <> -1 Then
    dtaGarantie.RecordSource = m_dataHelper.ValidateSQL("SELECT GAGARCLE as [Id], GAGROUPCLE as [Groupe], GASTECLE as [Société], " _
                             & " GATYPEGAR as [Type], GALIB as [Nom], GACALPASS as [Passage=0]" _
                             & " FROM Garantie WHERE GAGROUPCLE = " & cboGroupeGar.ItemData(cboGroupeGar.ListIndex) & " AND GASTECLE = " & cboSocieteGar.ItemData(cboSocieteGar.ListIndex))
    dtaGarantie.Refresh
    spreadGarantie.VirtualMaxRows = -1
    spreadGarantie.Refresh
    
    Dim i As Integer
    For i = 1 To spreadGarantie.MaxCols
      spreadGarantie.ColWidth(i) = spreadGarantie.MaxTextColWidth(i) + 1
    Next i
  End If
End Sub

'##ModelId=5C8A67AC02AF
Private Sub cboGroupeGar_Click()
    If cboGroupeGar.ListIndex = -1 Then Exit Sub
    m_dataHelper.FillCombo cboSocieteGar, "Select SOCLE, SONOM FROM Societe WHERE SOGROUPE = " & cboGroupeGar.ItemData(cboGroupeGar.ListIndex), -1
    
    dtaGarantie.RecordSource = ""
    'dtaGarantie.Refresh
    
    spreadGarantie.VirtualMaxRows = 0
    spreadGarantie.Refresh
End Sub

'##ModelId=5C8A67AC02BF
Private Sub cboGroupeSte_Click()
  If cboGroupeSte.ListIndex <> -1 Then
    dtaSte.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Societe WHERE SOGROUPE = " & cboGroupeSte.ItemData(cboGroupeSte.ListIndex))
    dtaSte.Refresh
    SpreadSte.VirtualMaxRows = -1
    SpreadSte.Refresh
    
    'SpreadSte.Col = 1
    'SpreadSte.ColHidden = False
    
    Dim i As Integer
    For i = 1 To SpreadSte.MaxCols
      SpreadSte.ColWidth(i) = SpreadSte.MaxTextColWidth(i) + 1
    Next i
  End If
End Sub

'##ModelId=5C8A67AC02CE
Private Sub cboGroupeUser_Click()
  If cboGroupeUser.ListIndex <> -1 Then
    dtaUser.RecordSource = m_dataHelper.ValidateSQL("SELECT USERCLE, TAGPECLE, TANUMCLE as ID, TANOM as Nom, TAPASS as [Mot de Passe], TAADMIN as Admin FROM Utilisateur WHERE TAGPECLE = " & cboGroupeUser.ItemData(cboGroupeUser.ListIndex))
    
    spreadUser.VirtualMaxRows = -1
    spreadUser.Refresh
  End If
End Sub

'##ModelId=5C8A67AC02DE
Private Sub Command2_Click()
  Unload Me
End Sub


'##ModelId=5C8A67AC02EE
Private Sub Form_Load()
  m_dataSource.SetDatabase dtaUser
  m_dataSource.SetDatabase dtaGarantie
  m_dataSource.SetDatabase dtaSte
  m_dataSource.SetDatabase dtaGroupe
  
  dtaUser.Enabled = True
  dtaGarantie.Enabled = True
  dtaSte.Enabled = True
  dtaGroupe.Enabled = True
  
  dtaGroupe.RecordSource = m_dataHelper.ValidateSQL("Groupe")
  
  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  
  ' active le premier tab
  SSTab1.Tab = 0
End Sub


'##ModelId=5C8A67AC02FD
Private Sub spreadUser_Change(ByVal Col As Long, ByVal Row As Long)
  spreadUser.Col = Col
  spreadUser.Row = Row
  dtaUser.Recordset.fields(spreadUser.DataField).Value = spreadUser.Value
  dtaUser.Recordset.Update
End Sub

'##ModelId=5C8A67AC031D
Private Sub SSTab1_Click(PreviousTab As Integer)
  Select Case SSTab1.Tab
    Case 0:  ' groupes
      dtaSte.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Goupe")
      dtaGroupe.Refresh
      Set vaSpreadGroupe.DataSource = dtaGroupe
      vaSpreadGroupe.VirtualMaxRows = -1
      vaSpreadGroupe.Refresh
   
    Case 1:  ' societes
      m_dataHelper.FillCombo cboGroupeSte, "Select GROUPECLE, NOM FROM Groupe", -1
      If cboGroupeSte.ListCount > 0 Then
        cboGroupeSte.ListIndex = 0
      End If
      'dtaSte.RecordSource = m_dataHelper.ValidateSQL("SELECT * FROM Societe WHERE SOCLE=0")
      'dtaSte.Refresh
      Set SpreadSte.DataSource = dtaSte
      'SpreadSte.VirtualMaxRows = 0
      SpreadSte.Refresh
      
    Case 2:  ' utilisateurs
      m_dataHelper.FillCombo cboGroupeUser, "Select GROUPECLE, NOM FROM Groupe", -1
      If cboGroupeUser.ListCount > 0 Then
        cboGroupeUser.ListIndex = 0
      End If
      'dtaUser.RecordSource = ""
      'dtaUser.Refresh
      Set spreadUser.DataSource = dtaUser
      'spreadUser.VirtualMaxRows = 0
      spreadUser.Refresh
      
    Case 3:  ' garanties
      m_dataHelper.FillCombo cboGroupeGar, "Select GROUPECLE, NOM FROM Groupe", -1
      cboSocieteGar.ListIndex = -1
      'dtaGarantie.RecordSource = ""
      'dtaGarantie.Refresh
      Set spreadGarantie.DataSource = dtaGarantie
      spreadGarantie.VirtualMaxRows = 0
      spreadGarantie.Refresh
  End Select
End Sub
