VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "P3I Generali"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7965
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   6135
      Visible         =   0   'False
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   14631
            MinWidth        =   14640
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerProtect 
      Interval        =   60000
      Left            =   1035
      Top             =   720
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "openCahier"
            Object.ToolTipText     =   "Liste des périodes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "about"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin FPSpreadADO.fpSpread sprLargeur 
         Height          =   285
         Left            =   1890
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   825
         _Version        =   524288
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "main.frx":1BB2
         AppearanceStyle =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1F9C
            Key             =   "openCahier"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":20A6
            Key             =   "openPeriode"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":21B0
            Key             =   "About"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":22BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFichiers 
      Caption         =   "&Fichiers"
      Begin VB.Menu mnuPeriodes 
         Caption         =   "&Périodes"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuAnnexes 
         Caption         =   "&Annexes"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Afficher les Logs"
         Begin VB.Menu mnuLogCopy 
            Caption         =   "&Copie des lots depuis SASP3I"
         End
         Begin VB.Menu mnuControlImport 
            Caption         =   "&Controle avant l'import"
         End
         Begin VB.Menu mnuErreurImport 
            Caption         =   "&Erreurs durant l'import"
         End
         Begin VB.Menu mnuErreurCalcul 
            Caption         =   "&Erreurs de calcul et Contrôles"
         End
         Begin VB.Menu mnuErreurExport 
            Caption         =   "&Erreurs durant l'export"
         End
      End
      Begin VB.Menu munSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuParametres 
      Caption         =   "&Paramètres"
      Begin VB.Menu mnuLoiMaintien 
         Caption         =   "&Lois de maintien, taux ..."
      End
      Begin VB.Menu mnuGestionAffichages 
         Caption         =   "&Gestion d'affichages"
      End
      Begin VB.Menu mnuPeriodsLocked 
         Caption         =   "&Périodes Bloquées"
      End
   End
   Begin VB.Menu mnuGenericData 
      Caption         =   "&Données Générales"
      Begin VB.Menu mnuArchGeneral 
         Caption         =   "&Archivage des données générales"
      End
      Begin VB.Menu mnuRestGeneral 
         Caption         =   "&Restauration des données générales"
      End
   End
   Begin VB.Menu mnuAuto1 
      Caption         =   "&Automatisation"
      Begin VB.Menu mnuImportTables 
         Caption         =   "Import des tables de paramétrages"
      End
      Begin VB.Menu mnuLancerProc 
         Caption         =   "Lancer une procédure"
      End
      Begin VB.Menu mnuLogAuto 
         Caption         =   "Afficher les logs"
      End
      Begin VB.Menu mnuTest 
         Caption         =   "Test"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAide 
      Caption         =   "&?"
      Begin VB.Menu mnuAbout 
         Caption         =   "&A propos..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67A7026F"
Option Explicit

'Déclaration de l'API d'attente
'##ModelId=5C8A67A7035B
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'##ModelId=5C8A67A7037A
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'##ModelId=5C8A67A703A9
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'##ModelId=5C8A67A8001F
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type


Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

'##ModelId=5C8A67A8005D
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

'##ModelId=5C8A67A7033C
Private Const ICC_USEREX_CLASSES = &H200

'##ModelId=5C8A67A8007D
Private Sub Command1_Click()

  'm_dataSource.Execute "Delete from P3IUser.Assure Where POPERCLE = 268"

End Sub

'Public Function InitCommonControlsVB() As Boolean
'   On Error Resume Next
'   Dim iccex As tagInitCommonControlsEx
'   ' Ensure CC available:
'   With iccex
'       .lngSize = LenB(iccex)
'       .lngICC = ICC_USEREX_CLASSES
'   End With
'   InitCommonControlsEx iccex
'   InitCommonControlsVB = (Err.Number = 0)
'   On Error GoTo 0
'End Function

'##ModelId=5C8A67A8008C
Private Sub MDIForm_Load()
    ' Verification piratage
    Dim dataSize As Double
    Dim dataLen As Double, l As Double
    
    
    
'    Dim oApp As New Excel.Application
'    MsgBox oApp.version
        
    
    InitCommonControls
    'InitCommonControlsVB
    
    Me.Show
    Me.WindowState = vbMaximized
    
    frmSplash.Show
    frmSplash.Refresh
    
    '*** init des variables globales
    '*** lecture de la ligne de commande
    Dim cmdLine As String
    
    DatabaseFileName = ""
    sFichierIni = ""
    
    ' decodage de la ligne de commande
    cmdLine = Trim(Command)
    
    Dim direc As String
    direc = Dir("T:\VB6\P3I_Generali\User\Cigogne\p3i.ini")
    
    ' INI
    sFichierIni = GetParameterFromCmdLine(cmdLine, "/INI=")
    If sFichierIni <> "" And Dir(sFichierIni) = "" Then
      MsgBox "Le fichier " & sFichierIni & " n'existe pas!", vbCritical
      sFichierIni = ""
    End If
      
    ' init du chemin du fichier de configuration
    If sFichierIni = "" Then
      ' get ini file path into sFichierINI
      Call BuildINIFilePath("P3I.INI")
    End If
    
    ' DB
    DatabaseFileName = GetParameterFromCmdLine(cmdLine, "/DB=")
    CRWDatabaseConnexion = GetParameterFromCmdLine(cmdLine, "/CRW=")
    
    DatabaseFileNameArchive = GetParameterFromCmdLine(cmdLine, "/DBARCH=")
    CRWDatabaseConnexionArchive = GetParameterFromCmdLine(cmdLine, "/CRWArch=")
    'If DatabaseFileName <> "" And Dir(DatabaseFileName) = "" Then
    '  MsgBox "Le fichier " & DatabaseFileName & " n'existe pas!", vbCritical
    '  DatabaseFileName = ""
    'End If
    
    ' default path for db and INI file
    If DatabaseFileName = "" Then
      DatabaseFileName = GetSettingIni(SectionName, "DB", "ConnectionString", "#")
      'If Right(DatabaseFileName, 1) <> "\" Then
      '  DatabaseFileName = DatabaseFileName & "\"
      'End If
      'DatabaseFileName = DatabaseFileName & "P3I.mdb"
      'DatabaseFileName = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseFileName & "P3I.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=P3I32"
    End If
    
    If DatabaseFileNameArchive = "" Then
      DatabaseFileNameArchive = GetSettingIni(SectionName, "DB", "ConnectionStringArchive", "")
    End If
    
    If CRWDatabaseConnexion = "" Then
      CRWDatabaseConnexion = GetSettingIni(SectionName, "DB", "CRWConnectionString", "#")
    End If
    
    If CRWDatabaseConnexionArchive = "" Then
      CRWDatabaseConnexionArchive = GetSettingIni(SectionName, "DB", "CRWConnectionStringArchive", "#")
    End If
    
    
    'CSV Files for exported tables
    If CSVUNCPath = "" Then
      CSVUNCPath = GetSettingIni(SectionName, "BULKINSERT", "CSVUNCPath", "#")
    End If
    
    
    '***
    '*** DEBUT PROTECTION
    '***
    
    ' Verification piratage
    licenceDatPath = GetSettingIni(SectionName, "Dir", "LicencesPath", App.Path)
    
    If CheckProtection(licenceDatPath) = False Then
      LeaveProtection
      Unload Me
      End
    End If
    
    '***
    '*** FIN PROTECTION
    '***

    ' chemin pour la création des fichiers log
    m_logPath = GetSettingIni(SectionName, "Dir", "LogPath", App.Path)
    If Right(m_logPath, 1) <> "\" Then
      m_logPath = m_logPath & "\"
    End If
    
    m_logPathAuto = GetSettingIni(SectionName, "Dir", "LogPathAuto", App.Path)
    If Right(m_logPathAuto, 1) <> "\" Then
      m_logPathAuto = m_logPathAuto & "\"
    End If
    
    ' init du chemin de la base
    If GetSettingIni(SectionName, "DB", "DBPath", "#") = "#" Then
      Call SaveSettingIni(SectionName, "DB", "DBPath", App.Path)
    End If
    
    If GetSettingIni(SectionName, "DB", "ConnectionString", "#") = "#" Then
      Call SaveSettingIni(SectionName, "DB", "ConnectionString", DatabaseFileName)
    End If

    
    ' ouverture de la base en mode partage, lecture/ecriture
On Error GoTo errLoadDB
    DatabasePassword = ";PWD=P3I32"
    
    Set m_dataSource = New P3IGeneraliDataAccess.DataAccess
    If m_dataSource.Connect(DatabaseFileName) = False Then
      MsgBox "Impossible d'ouvrir la base de donnée!" & vbLf & "Connection : " & DatabaseFileName, vbCritical
      LeaveProtection
      Unload Me
      End
    End If
    
    'Set theDB = m_dataSource.Connection
    Set m_dataHelper = m_dataSource.CreateHelper
    
    '###
    Call Sleep(2000)
    
    frmSplash.Hide
    Unload frmSplash
      
    ' login ...
    '#### UNCOMMENT
    If UserLogin(True) = False Then
      Unload Me
      End
      Exit Sub
    End If
    
On Error GoTo 0
    
    ' nb décimales calculs taux de provisions chargés par défaut
    NbDecimalePM = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "NbDecimalPM", "2"))
    NbDecimaleCalcul = m_dataHelper.GetDouble2(GetSettingIni(CompanyName, SectionName, "NbDecimalCalcul", "6"))
    
    
    ' fabrique le titre de la fenetre en fonction du groupe en cours
    NomGroupe = m_dataHelper.GetParameter("SELECT NOM FROM Groupe WHERE GroupeCle = " & GroupeCle)
        
    ' format des NCA
    m_FormatNCA = GetSettingIni(CompanyName, SectionName, "FormatNCA", "0 00 000000 00 00")
    
    
    '###
    'create the collection class that contains all the Assure Displays (used in frmEditPeriode) for the current user
    'Set AssureDisplays = New AssureDisplays
    'Command1_Click

    
    '### TEST
'    Dim frm As New frmStatImport
'    frm.PeriodeType = "STAT"
'    frm.Show vbModal
    
    'CalculatePSAPGeneric ("MO")
    
        
    Exit Sub
    
errLoadDB:
  MsgBox Err.Description
  End
End Sub

'##ModelId=5C8A67A8009C
Private Sub MDIForm_Unload(Cancel As Integer)
  If Not m_dataSource Is Nothing Then
    If m_dataSource.Connected Then
      m_dataSource.Disconnect
      'Set theDB = Nothing
      Set m_dataSource = Nothing
      Set m_dataHelper = Nothing
    End If
  End If
  
  LeaveProtection
End Sub

'##ModelId=5C8A67A800BB
Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

'##ModelId=5C8A67A800DA
Private Sub mnuAnnexes_Click()
  If DroitAdmin Then
    frmAnnexe.Show vbModal
  End If
End Sub

'##ModelId=5C8A67A800EA
Private Sub ShowLog(filename As String, titre As String)
  
  Dim frm As frmDisplayLog
  
  Set frm = New frmDisplayLog
  frm.FichierLog = m_logPath & filename
  frm.Caption = titre
  frm.lblLog.text = m_logPath & filename
  
  frm.Show vbModal
  
  Set frm = Nothing
End Sub


'##ModelId=5C8A67A80128
Private Sub mnuAuto_Click()

End Sub

'##ModelId=5C8A67A80138
Private Sub mnuGestionAffichages_Click()

  Dim FM As frmManageDisplays
  Set FM = New frmManageDisplays
  FM.Show vbModal
  Set FM = Nothing
    
End Sub

'##ModelId=5C8A67A80148
Private Sub mnuImportTables_Click()
  Dim FM As frmAutoImportTables
  Set FM = New frmAutoImportTables
  FM.Show vbModal
  Set FM = Nothing
End Sub

'##ModelId=5C8A67A80177
Private Sub mnuLancerProc_Click()
  Dim FM As frmAutomatisation
  Set FM = New frmAutomatisation
  FM.Show vbModal
  Set FM = Nothing
End Sub

'##ModelId=5C8A67A80196
Private Sub mnuLogAuto_Click()
  frmAutoLog.Show vbModal
End Sub

'##ModelId=5C8A67A801A5
Private Sub mnuPeriodsLocked_Click()
  frmUnlockPeriodes.Show vbModal
End Sub

'##ModelId=5C8A67A801B5
Private Sub mnuTest_Click()
  
  Dim selectedPeriods As New Collection
  selectedPeriods.Add (920)
  selectedPeriods.Add (921)
  selectedPeriods.Add (922)
  selectedPeriods.Add (923)
  selectedPeriods.Add (924)
  
  Dim frm As frmAutoStat
  Set frm = New frmAutoStat
  Set frm.selectedPeriods = selectedPeriods
  frm.Show vbModal
  Set frm = Nothing
  
  
End Sub

'##ModelId=5C8A67A801D4
Private Sub TimerProtect_Timer()
  Static delai As Integer
  
  TimerProtect.Interval = 60000 ' 1 mn
  TimerProtect.Enabled = False
  
  #If No_Protection = 1 Then
    Exit Sub
  #End If
  
  delai = delai + 1
  If delai = 4 Then
    ' Verification piratage
    If CheckProtection(licenceDatPath) = False Then
      Unload Me
      End
    End If
    delai = 0
  End If
End Sub

'##ModelId=5C8A67A801E4
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' Utilise la propriété Key avec l'instruction SelectCase pour   spécifier une action.
    Select Case Button.key
        Case Is = "openCahier"
          frmPeriode.Show
        Case Is = "about"
          frmAbout.Show vbModal
    End Select

End Sub

'##ModelId=5C8A67A80203
Private Sub mnuControlImport_Click()
  ShowLog GetWinUser & "_ControleImport.log", "Contrôles avant l'import"
End Sub

'##ModelId=5C8A67A80222
Private Sub mnuErreurCalcul_Click()
  ShowLog GetWinUser & "_ErreurCalcul.log", "Erreurs de calcul"
End Sub

'##ModelId=5C8A67A80232
Private Sub mnuErreurExport_Click()
  ShowLog GetWinUser & "_ErreurExport.log", "Erreurs durant l'export"
End Sub

'##ModelId=5C8A67A80251
Private Sub mnuErreurImport_Click()
  ShowLog GetWinUser & "_ErreurImport.log", "Erreurs durant l'import"
End Sub

'##ModelId=5C8A67A80261
Private Sub mnuLogCopy_Click()
  ShowLog GetWinUser & "_CopyLot.log", "Erreurs durant la copie d'un lot depuis SASP3I"
End Sub

'##ModelId=5C8A67A80271
Private Sub mnuLoiMaintien_Click()
  frmParametre.Show vbModal
End Sub

'##ModelId=5C8A67A80280
Private Sub mnuPeriodes_Click()
    frmPeriode.Show
End Sub

'##ModelId=5C8A67A80290
Private Sub mnuQuitter_Click()
  Unload Me
End Sub

'##ModelId=5C8A67A8029F
Private Sub mnuArchGeneral_Click()

  Dim archRest As New ArchiveRestore
  archRest.SyncGeneralTablesFromProdToArchive
  
  'SyncArchive

End Sub

'##ModelId=5C8A67A802BF
Private Sub mnuRestGeneral_Click()

  Dim archRest As New ArchiveRestore
  archRest.SyncGeneralTablesFromArchiveToProd
  
End Sub

'##ModelId=5C8A67A802CE
Public Function StatusBarProgress(sbStatus As StatusBar, vPannel As Variant, ByVal lPercentComplete As Long, Optional lColor As OLE_COLOR = &H9D9793) As Boolean
    
    Static soProgressBar As ProgressBar, slLastColor As Long
    Const clBorder As Long = 25
    Const WM_USER = &H400, CCM_FIRST As Long = &H2000&
    Const CCM_SETBKCOLOR As Long = (CCM_FIRST + 1), PBM_SETBKCOLOR As Long = CCM_SETBKCOLOR, PBM_SETBARCOLOR As Long = (WM_USER + 9)
    
    On Error GoTo ErrFailed
    If soProgressBar Is Nothing Then
        Set soProgressBar = Controls.Add("MSComctlLib.ProgCtrl.2", "ProgressBar1")
        Call SetParent(soProgressBar.hwnd, sbStatus.hwnd)
        soProgressBar.Visible = True
        soProgressBar.BorderStyle = ccNone
        soProgressBar.Appearance = ccFlat
        With sbStatus.Panels(vPannel)
            soProgressBar.Move .Left + clBorder, Screen.TwipsPerPixelY * 2 + clBorder, .Width - (clBorder * 2), sbStatus.Height - (Screen.TwipsPerPixelY * 3) - (clBorder * 2)
        End With
        soProgressBar.Min = 0
        soProgressBar.Max = 100
    End If
    
    soProgressBar.Value = lPercentComplete
    StatusBarProgress = True
    
    If lColor <> slLastColor Then
        'change the bar colour
        slLastColor = lColor
        Call SendMessage(soProgressBar.hwnd, PBM_SETBARCOLOR, 0&, ByVal lColor)
    End If
    
    Exit Function

ErrFailed:
    Debug.Print "Error in StatusBarProgress2: " & Err.Description
    Debug.Assert False
    StatusBarProgress = False
    
End Function





















