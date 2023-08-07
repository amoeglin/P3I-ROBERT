Attribute VB_Name = "Protection"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67E202D1"
Option Explicit

'##ModelId=5C8A67E30189
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'##ModelId=5C8A67E301C7
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Schreiben sie den folgenden Code in ein Öffengtliches Modul
'##ModelId=5C8A67E30206
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'##ModelId=5C8A67E30263
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'##ModelId=5C8A67E203BB
Const WM_TIMER = &H113 'Timer Ereigniss Trifft ein

'##ModelId=5C8A67E30002
Private hEvent As Long

' parametre contre le piratage
'##ModelId=5C8A67E30012
Public Const VeuillezContacterMoeglin As String = "Veuillez contacter Actuaries Services à support@actuariesservices.com !"

' {680BE2F0-73E9-11d2-8847-0020AFDD5298} : P2I
' {68FFFFFF-73E9-11d2-8847-0020AFDD5298} : P2I GENERALI
' {680BE2F1-73E9-11d2-8847-0020AFDD5298} : TPFM
' {680BE2F2-73E9-11d2-8847-0020AFDD5298} : IFC
' {680BE2F3-73E9-11d2-8847-0020AFDD5298} : ECP
' {680BE2F4-73E9-11d2-8847-0020AFDD5298} : CAPISICA
' {680BE2F5-73E9-11d2-8847-0020AFDD5298} : P3I
'##ModelId=5C8A67E30041
Private Const RacinePirate As String = "Classes\CLSID"
'##ModelId=5C8A67E30050
Private Const SectionPirate As String = "{680BE2F5-73E9-11d2-8847-0020AFDD5298}\MiscStatus"
'##ModelId=5C8A67E3006F
Private Const NombrePirate As Long = 1130
'##ModelId=5C8A67E3008F
Private Const DomainePirate As String = "GROUPEISICA"

' protection guilleray
'##ModelId=5C8A67E300AE
Private Const NomProgramme As String = "P3I"

'##ModelId=5C8A67E300DF
Private mLicenceServer As New CLicenceServer
'##ModelId=5C8A67E300E0
Private mUserId As String
'##ModelId=5C8A67E300FC
Private databasePath As String

'##ModelId=5C8A67E3011B
Public licenceDatPath As String

'##ModelId=5C8A67E3013B
Private Const CompKey = 546852  ' la valeur est arbitraire

' timer
'##ModelId=5C8A67E3015A
Private Const Delai_Timer = 60000
'

'##ModelId=5C8A67E302B2
Public Function GetAboutString() As String
  Dim CurrentLicence As CLicence
  Dim CurrentLicenceFile As New CLicenceFile
  
  #If No_Protection = 1 Then
    GetAboutString = "PAS DE PROTECTION"
    Exit Function
  #End If

  ' arrete le timer
  DisableTimer
    
  ' inutile de tester la validité du fichier
  ' si il est altéré/corrompu ou autre les données ne seront pas visibles
  ' une cascade d'erreurs (trappées) informera l'utilisateur
  If CurrentLicenceFile.Load(databasePath + "\Licences.dat", DatFile, sFichierIni) = True Then
      Set CurrentLicence = CurrentLicenceFile.Licence(NomProgramme)
      If CurrentLicence Is Nothing Then Exit Function
      GetAboutString = "Nb Licences : " & CurrentLicence.inUse & " / " & CurrentLicence.LicencesCount & " - Expirant le " & CurrentLicence.EndDate
  End If
  
  ' relance le timer avec un battement de 1mn
  EnableTimer Delai_Timer
End Function

'Die Timer Prozedur die jedesmal zu den Eigegebenen Millisekunden ein Ereigniss
'sendet
'##ModelId=5C8A67E302C1
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Static nbDelai As Integer
  Static inMessage As Integer
  
  If uMsg = WM_TIMER Then
    nbDelai = nbDelai + 1
    
    If inMessage = 1 Then
      ' arrete le timer
      DisableTimer
      Unload frmMain
      End
    End If
    
    If nbDelai > 2 Then
      ' Verification piratage
      If CheckProtection(licenceDatPath) = False Then
        EnableTimer 10000
        inMessage = 1
        MsgBox "Il n'y a plus de licence disponible, le programme va être arreté." & vbLf & VeuillezContacterMoeglin, vbCritical
        LeaveProtection
        Unload frmMain
        End
      End If
      nbDelai = 0
    End If
    
    ' arrete le timer
    DisableTimer
    
    ' relance le timer avec un battement de 1mn
    EnableTimer Delai_Timer
  End If
End Sub

'Startet den Timer
'##ModelId=5C8A67E3031F
Public Sub EnableTimer(ByVal msInterval As Long)
  If hEvent <> 0 Then Exit Sub
  hEvent = SetTimer(0&, 0&, msInterval, AddressOf TimerProc)
End Sub

'Beendet den Timer
'##ModelId=5C8A67E3034E
Public Sub DisableTimer()
  If hEvent = 0 Then Exit Sub
  KillTimer 0&, hEvent
  hEvent = 0
End Sub

'##ModelId=5C8A67E3035D
Public Sub LeaveProtection()
  DisableTimer

#If No_Protection = 1 Then
  Exit Sub
#End If

  mLicenceServer.UserLeave mUserId, databasePath, NomProgramme, sFichierIni
  
  If Not mLicenceServer Is Nothing Then
    Set mLicenceServer = Nothing
  End If
  
End Sub

'##ModelId=5C8A67E3036D
Public Function CheckProtection(licenceFilePath As String) As Boolean
  On Error GoTo errProtection
    
  #If No_Protection = 1 Then
    MsgBox "PAS DE PROTECTION !", vbCritical
    CheckProtection = True
    Exit Function
  #End If
    
  CheckProtection = False
  
  ' protection guilleray
  databasePath = Trim(licenceFilePath)
  If Right(databasePath, 1) = "\" Then
    databasePath = Left(databasePath, Len(databasePath) - 1)
  End If
      
  If mLicenceServer.CheckLicence(mUserId, databasePath, NomProgramme, sFichierIni) Then
    CheckProtection = True
    
    ' demarre la boucle de verification avec un delai aléatoire de max 45s
    Randomize
    EnableTimer Int(45000 * Rnd + 15000) ' first timer between 15s and 60s
  End If
  
  Exit Function
    
errProtection:
  MsgBox "Erreur : " & Err.Description & vbLf & "Le programme a été mal installé !"
  'Set pObj = Nothing
  
  CheckProtection = False
End Function

'##ModelId=5C8A67E3038C
Public Sub BuildINIFilePath(sRoot As String)
  Dim sWork As String, sCompName As String
  Dim l As Long
  
  ' get computer name
  sCompName = Space(300)
  l = 300
  Call GetComputerName(sCompName, l)
  sCompName = Left(sCompName, l)
  

  ' seach for ini file under .\User
  sWork = App.Path & "\User\" & sCompName & "\" & sRoot
  If Dir(sWork) <> "" Then
    sFichierIni = sWork
    Exit Sub
  End If
  
  ' seach for ini file under .\..\User
  sWork = App.Path & "\..\User\" & sCompName & "\" & sRoot
  If Dir(sWork) <> "" Then
    sFichierIni = sWork
    Exit Sub
  End If

  sFichierIni = sRoot
End Sub

