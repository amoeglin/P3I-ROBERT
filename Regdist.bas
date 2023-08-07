Attribute VB_Name = "Regdist"
'----------------------------------------------------------------
' DESCRIPTION
'----------------------------------------------------------------
'   Module des fonctions liées  à la registry
'
'----------------------------------------------------------------
' FONCTIONS ET PROCEDURES PUBLIQUES
'----------------------------------------------------------------
'
'   Procédure de fonctionnement identique à l'équivalent VB. Seul l'emplacement dans la Registry
'   est différent
'      Sub SaveSetting(appname As String, section As String, key As String, setting As String)
'
'   Procédure de fonctionnement identique à l'équivalent VB. Seul l'emplacement dans la Registry
'   est différent
'       Function GetSetting(appname As String, section As String, key As String, default As String) As Variant
'
'   Fonction retournant les paramètres nécessaires à une connexion depuis un poste client sur un
'   serveur NT Oracle
'       Function GetParamClientServeur(ByVal NomServeur As String, cle As String, Passwd As String) As Boolean
'
'----------------------------------------------------------------

Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" _
    (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, lpData As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
     lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal _
     lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
     lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
     lpdwDisposition As Long) As Long
     
Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
     ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
     ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003
Private Const SYNCHRONIZE = &H100000

Private Const READ_CONTROL = &H20000

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&

Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4
Private Const REG_NOTIFY_CHANGE_NAME = &H1
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_BINARY = 3
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_OPENED_EXISTING_KEY = &H2
Private Const REG_OPTION_BACKUP_RESTORE = 4
Private Const REG_OPTION_CREATE_LINK = 2
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_RESERVED = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_REFRESH_HIVE = &H2
Private Const REG_RESOURCE_LIST = 8
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Private Const REG_SZ = 1
Private Const REG_WHOLE_HIVE_VOLATILE = &H1
Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Private Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)

' endroit où vont lire/écrire les fonctions d'accès à la registry
Private Const REG_NOM_HANDLE = HKEY_CURRENT_USER

'Déclaration des fonction de l'API Windows permettant la lecture et l'écriture dans des fichiers .ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Fonction de liste des sections d'un fichier ini
'
Public Function EnumSections(ByRef aSections() As String, sIniPath As String) As Boolean
    On Error GoTo err_EnumSections

'    on récupère toutes les clés de la section
    Dim sRet As String, sTempSection As String
    
    sRet = String$(458752, vbNullChar)
    sTempSection = Left$(sRet, GetPrivateProfileSectionNames(sRet, Len(sRet), sIniPath) - 1)

'   en principe c'est pas vide, mais dans le  doute...
    If LenB(sTempSection) = 0 Then
        EnumSections = False
    Else
'       on renvoie toutes les sections
        aSections() = Split(sTempSection, vbNullChar)
        EnumSections = True
    End If
    
    sRet = vbNullString
    sTempSection = vbNullString
    
    Exit Function
    
err_EnumSections:
  MsgBox "Error " & Err & " dans EnumSections() : " & vbLf & Err.Description, vbCritical
  Resume Next
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Fonction de Lecture dans un fichier .ini
'sSection : nom de la section entre [], à rechercher
'sEntree : nom de l'entrée de la section sSection, à rechercher
'sDefaut : valeur par défaut à retourner, si l'entrée, la section, ou le fichier ne sont pas trouvés
'lLgMax : nombre maximum de caractères devant être lus
'sFichier : nom du fichier .ini
'
'La fonction retourne directement la chaine trouvée
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function sReadIniFile(sSection As String, sEntree As String, sDefaut, lLgMax As Long, sFichier As String) As String
  Dim lRes As Long
  Dim sValeur As String
  
  'Prépare la chaine de retour (longueur lLgMax +1, car caractère de fin de chaine)
  sValeur = Space(lLgMax + 1)
  
  'Lecture dans fichier en utilisant fonction API Win32 (lRes contient le nombre de caractères lus)
  lRes = GetPrivateProfileString(sSection, sEntree, sDefaut, sValeur, Len(sValeur), sFichier)
  
  'Retour de la valeur trouvée (ou de la valeur par défaut, sinon), uniquement les caractères "utiles"
  sReadIniFile = Left(sValeur, lRes)
End Function

#If P3IImportGenerali <> 1 Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' wrapper pour GetSetting et SaveSetting mais dans un .INI au lieu de dans la registry
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetSettingIni(appname As String, Section As String, key As String, default As String, Optional fLocalMachine) As Variant
  GetSettingIni = sReadIniFile(Section, key, default, 1024, sFichierIni)
End Function

Sub SaveSettingIni(appname As String, Section As String, key As String, setting As Variant, Optional fLocalMachine)
  Call WritePrivateProfileString(Section, key, CStr(setting), sFichierIni)
End Sub

Sub DeleteSection(appname As String, Section As String)
  Call WritePrivateProfileString(Section, vbNullString, vbNullString, sFichierIni)
End Sub

Function SafeGetSettingIni(appname As String, Section As String, key As String, default As String, Optional fLocalMachine) As Variant
  SafeGetSettingIni = sReadIniFile(Section, key, default, 1024, "WIN.INI")
End Function

Sub SafeSaveSettingIni(appname As String, Section As String, key As String, setting As Variant, Optional fLocalMachine)
  Call WritePrivateProfileString(Section, key, CStr(setting), "WIN.INI")
End Sub


'------------------------------------------------------------------------------------------------
'   Procédure de fonctionnement identique à l'équivalent VB. Seul l'emplacement dans la Registry
'   est différent
'   En entrée : appname : Racine de l'ensemble des paramètres propres à l'environnemment
'                         d'exploitation. Exemple :"MOEGLIN"
'               section : Nom de l'application. Exemple : "P2I"
'               key : Zone renseignée. Exemple : "DateInstall"
'               setting : Valeur affectée
'------------------------------------------------------------------------------------------------
'
Sub SaveSetting(appname As String, Section As String, key As String, setting As Variant, Optional fLocalMachine)

    Dim lret As Long
    
    On Error GoTo ErrSaveSetting
    
    Dim sKey As String
    Dim lLen As Long

    sKey = "SOFTWARE\" & appname & "\" & Section
    
    Dim hKey As Long
    Dim lRetVal As Long
    Dim sa As SECURITY_ATTRIBUTES
    
    If IsMissing(fLocalMachine) Then
      ' sauvegarde pour le user courant
      lRetVal = RegCreateKeyEx(HKEY_CURRENT_USER, sKey, 0&, _
                vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                sa, hKey, lRetVal)
    Else
      ' sauvegarde dans LOCAL_MACHINE
      lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, _
                vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                sa, hKey, lRetVal)
    End If
        
    If lret = ERROR_SUCCESS Then
        Dim Ssetting As String
        Ssetting = CStr(setting)
        lLen = Len(Ssetting)
        lret = SetValueEx(hKey, key, REG_SZ, Ssetting)
        lret = RegCloseKey(hKey)
    End If

ExitSaveSetting:
    Exit Sub

ErrSaveSetting:
    Resume ExitSaveSetting
    
End Sub
'------------------------------------------------------------------------------------------------
'   Procédure de fonctionnement identique à l'équivalent VB. Seul l'emplacement dans la Registry
'   est différent
'   En entrée : appname : Racine de l'ensemble des paramètres propres à l'environnemment
'                         d'exploitation. Exemple :"MOEGLIN"
'               section : Nom de l'application. Exemple : "P2I"
'               key : Zone contenant la valeur cherchée. Exemple : DateInstall
'               default : Valeur retournée si la regsitry en contient pas l'information cherchée
'------------------------------------------------------------------------------------------------
'
Function GetSetting(appname As String, Section As String, key As String, default As String, Optional fLocalMachine) As Variant

    Dim lret As Long
    Dim hRegKey As Long
    
    On Error GoTo ErrGetSetting
    
    If IsMissing(fLocalMachine) Then
      ' sauvegarde pour le user courant
      lret = RegConnectRegistry("", HKEY_CURRENT_USER, hRegKey)
    Else
      ' sauvegarde dans LOCAL_MACHINE
      lret = RegConnectRegistry("", HKEY_LOCAL_MACHINE, hRegKey)
    End If
    
    If lret <> ERROR_SUCCESS Then
        Exit Function
    Else
        Dim sBuf As String
        Dim hKey As Long
        Dim lSam As Long
        Dim sKey As String
        Dim lLen As Long

        sBuf = String(1024, vbNullChar)
        sKey = "SOFTWARE\" & appname & "\" & Section
        lSam = KEY_READ
        
        lret = RegOpenKeyEx(hRegKey, sKey, 0, lSam, hKey)
        
        If lret = ERROR_SUCCESS Then
            sBuf = String(1024, vbNullChar)
            lLen = Len(sBuf)
            lret = QueryValueEx(hKey, key, sBuf)
            If lret = ERROR_SUCCESS Then
                GetSetting = RTrim(sBuf)
                GetSetting = Left(GetSetting, Len(GetSetting) - 1)
            Else
                GetSetting = default
                Exit Function
            End If
            
        Else
            GetSetting = default
            Exit Function
        End If
        lret = RegCloseKey(hRegKey)
    End If

ExitGetSetting:
    Exit Function

ErrGetSetting:
    GetSetting = default
    Resume ExitGetSetting
    
End Function

'------------------------------------------------------------------------------------------------
'   Fonction encapsulant les requêtes sur le contenu d'une clé de la registry
'   Les valeurs de type binaire ne sont pas remontées
'
'------------------------------------------------------------------------------------------------
Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine la taille et le type des données à lire
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_SUCCESS Then Error 5

    Select Case lType
        ' string
        Case REG_SZ:
                sValue = String(cch, 0)
                lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_SUCCESS Then
                vValue = Left$(sValue, cch)
            Else
                vValue = Empty
            End If
        '  DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_SUCCESS Then vValue = lValue
        ' Binaire
        Case Else
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function


'------------------------------------------------------------------------------------------------
'   Fonction encapsulant les mises à jour du contenu d'une clé de la registry
'   Les valeurs de type binaire ne sont pas utilisées
'
'------------------------------------------------------------------------------------------------
Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Dim bvalue() As Byte
    Dim lenValue As Long
    
    
    Select Case lType
        ' string
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        '  DWORDS
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        '   Binary
        Case REG_BINARY
            bvalue() = vValue
            lenValue = UBound(bvalue) - LBound(bvalue) + 1
            SetValueEx = RegSetValueExBinary(hKey, sValueName, 0&, lType, bvalue(0), lenValue)
        Case Else
            SetValueEx = -1
    End Select
End Function


'------------------------------------------------------------------------------------------------
'   Lecture de la base de la base de registre
'getkey (nom complet de la section sans \ à la fin, nom de la clé, valeur par défaut)
'------------------------------------------------------------------------------------------------
Public Function GetKey(Section As String, key As String, default As String) As Variant
    Dim lret As Long
    Dim hRegKey As Long
    Dim sRootKey As String
    Dim hRootKEY As Long
    
    Dim posSection As Long
        
    Dim sBuf As String
    Dim hKey As Long
    Dim lSam As Long
    Dim sKey As String
    
    On Error GoTo ErrGetKey
    
    'Extraction du nom de la branche de base
    posSection = InStr(1, Section, "\")
    If posSection = 0 Then
        Err.Raise Number:=-1
    End If
    sRootKey = Left(Section, posSection - 1)
    
    'Transformation du nom de la branche de base en clé de connexion
    Select Case UCase(sRootKey)
        Case "HKEY_CLASSES_ROOT"
            hRootKEY = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            hRootKEY = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            hRootKEY = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            hRootKEY = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            hRootKEY = HKEY_USERS
        Case "HKEY_CURRENT_USER"
            hRootKEY = HKEY_CURRENT_USER
        Case Else
            Err.Raise Number:=-1
    End Select
    
    'Connection à la base de registre : on obtient le handle hRegKey
    lret = RegConnectRegistry("", hRootKEY, hRegKey)
    If lret <> ERROR_SUCCESS Then
        Err.Raise Number:=-1, Description:="Connexion à la base de registre impossible", Source:="RegDist.GetKey"
    End If

    
    'Obtention d'un handle sur une clé de la base de registre
    sKey = Mid(Section, posSection + 1)
    lSam = KEY_READ
    lret = RegOpenKeyEx(hRegKey, sKey, 0, lSam, hKey)
    If lret <> ERROR_SUCCESS Then
        Err.Raise Number:=-1
    End If
    
    'Lecture de la valeur de la clé
    sBuf = String(1024, vbNullChar)
    lret = QueryValueEx(hKey, key, sBuf)
    If lret <> ERROR_SUCCESS Then
        Err.Raise Number:=-1
    End If
    
    'On retourne la valeur lue dans la base de registre
    GetKey = RTrim(sBuf)
    GetKey = Left(GetKey, Len(GetKey) - 1)
        
    'On rend le handle sur la clé
    lret = RegCloseKey(hRegKey)
    
    
    '??? Fermer la connexion à la base de registre

ExitGetKey:
    Exit Function

ErrGetKey:
    GetKey = default
    Resume ExitGetKey

End Function


'------------------------------------------------------------------------------------------------
'sauvegarde d'une valeur dans la base de registre
'------------------------------------------------------------------------------------------------
'
Sub PutKey(Section As String, key As String, setting As Variant)
    
    
    Dim sRootKey As String
    Dim hRootKEY As Long
    Dim posSection As Long
    Dim lret As Long
    Dim sKey As String
    Dim lLen As Long
    Dim hKey As Long
    Dim lRetVal As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim Ssetting As String
    
    On Error GoTo ErrSaveSetting
    
    'recupération de la branche de base et du chemin
    posSection = InStr(1, Section, "\")
    
    If posSection = 0 Then
        Err.Raise Number:=-1
    End If
    
    sRootKey = Left(Section, posSection - 1)
    sKey = Mid(Section, posSection + 1)
    
    'Transformation du nom de la branche de base en clé de connexion
    Select Case UCase(sRootKey)
        Case "HKEY_CLASSES_ROOT"
            hRootKEY = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_CONFIG"
            hRootKEY = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            hRootKEY = HKEY_DYN_DATA
        Case "HKEY_LOCAL_MACHINE"
            hRootKEY = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            hRootKEY = HKEY_USERS
        Case "HKEY_CURRENT_USER"
            hRootKEY = HKEY_CURRENT_USER
        Case Else
            Err.Raise Number:=-1
    End Select
    
    'création de la section de la clé
    lret = RegCreateKeyEx(hRootKEY, sKey, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              sa, hKey, lRetVal)
        
    If lret <> ERROR_SUCCESS Then
        Err.Raise Number:=-1
    End If
    
    'attribution de la valeur à la clé
    
    
    Select Case VarType(setting)
        Case vbArray + vbByte
            lret = SetValueEx(hKey, key, REG_BINARY, setting)
        Case Else
            Ssetting = CStr(setting)
            lLen = Len(Ssetting)
            lret = SetValueEx(hKey, key, REG_SZ, Ssetting)
    End Select
    lret = RegCloseKey(hKey)


ExitSaveSetting:
    Exit Sub

ErrSaveSetting:
    Resume ExitSaveSetting
    
End Sub


#End If
