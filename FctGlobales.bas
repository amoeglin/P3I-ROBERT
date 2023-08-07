Attribute VB_Name = "FctGlobales"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6823013B"
Option Explicit

'##ModelId=5C8A68240235
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'container for Displays and AssureFields
'##ModelId=5C8A68230237
Public AssureDisplays As AssureDisplays

' variable code retour
'##ModelId=5C8A68230244
Public GroupeCle As Integer
'##ModelId=5C8A68230254
Public NomGroupe As String
'##ModelId=5C8A68230263
Public user_name As String
'##ModelId=5C8A68230283
Public user_pwd As String
'##ModelId=5C8A682302A2
Public DroitAdmin As Boolean
'##ModelId=5C8A682302C1
Public RECNO As Long

'##ModelId=5C8A682302E0
Public archiveMode As Boolean

' base de donnee
'##ModelId=5C8A682302F2
Public m_dataSource As P3IGeneraliDataAccess.DataAccess
'##ModelId=5C8A68230302
Public m_dataHelper As P3IGeneraliDataAccess.DataHelper

'DB Archive
'##ModelId=5C8A6823030F
Public m_dataSourceArchive As P3IGeneraliDataAccess.DataAccess
'##ModelId=5C8A68230312
Public m_dataHelperArchive As P3IGeneraliDataAccess.DataHelper

'Public theDB As dao.Database
'##ModelId=5C8A6823031F
Public DatabaseFileName As String
'##ModelId=5C8A6823032F
Public CRWDatabaseConnexion As String
'##ModelId=5C8A6823034E
Public DatabasePassword As String

'##ModelId=5C8A6823036D
Public DatabaseFileNameArchive As String
'##ModelId=5C8A6823037D
Public CRWDatabaseConnexionArchive As String

'CSV Files For Exported Tables
'##ModelId=5C8A682303AC
Public CSVUNCPath As String


' fichier log pour les calculs
'##ModelId=5C8A682303BB
Public m_logPath As String
'##ModelId=5C8A682303DA
Public m_logPathAuto As String


' taille des bouton
'##ModelId=5C8A68240002
Public Const btnHeight As Integer = 370
'##ModelId=5C8A68240021
Public Const btnWidth As Integer = 2200

' passage des parametres
'##ModelId=5C8A68240041
Public numPeriode As Long
'##ModelId=5C8A68240050
Public NumParamCalcul As Long
'##ModelId=5C8A6824006F
Public DescriptionPeriode As String
'##ModelId=5C8A6824008F
Public SoCle As Double                 ' Pour édition revalo : Société sélectionnée 0 si toutes

' valeur par défaut nb de décimales utilisées dans les calculs des taux de
'provisions
'##ModelId=5C8A6824009E
Public NbDecimalePM As Integer
'##ModelId=5C8A682400BE
Public NbDecimaleCalcul As Integer

' Format d'affichage des NCA (POCONVENTION)
'##ModelId=5C8A682400DD
Public m_FormatNCA As String

' parametre pour la sauvegarde des valeurs par defaut dans la registry
'##ModelId=5C8A682400FC
Public Const CompanyName As String = "Moeglin"
'##ModelId=5C8A6824011B
Public Const SectionName As String = "P3I"
'##ModelId=5C8A6824013B
Public sFichierIni As String

'##ModelId=5C8A6824015A
Public Const DEFAULT_PARAM_SECTION As String = "DefaultParameters_"
'##ModelId=5C8A68240179
Public Const DEFAULT_SECTION As String = SectionName

'Import Statutaire
'##ModelId=5C8A68240189
Public NumPeriodeStat As Long
'##ModelId=5C8A682401A8
Public NumPeriodeNonStat As Long
'##ModelId=5C8A682401C7
Public PathSexFileExcel As String
'##ModelId=5C8A682401D7
Public CategoryCodeSTAT As String
'##ModelId=5C8A682401F6
Public SexAllMale As Boolean
'##ModelId=5C8A68240216
Public TwoLotImport As Boolean

'##ModelId=5C8A68240283
Public Function FieldExistsInRS(ByRef rs As ADODB.Recordset, ByVal FieldName As String)

   Dim fld As ADODB.field
    
   FieldName = UCase(FieldName)
    
   For Each fld In rs.fields
      If UCase(fld.Name) = FieldName Then
         FieldExistsInRS = True
         Exit Function
      End If
   Next
    
   FieldExistsInRS = False
End Function

'##ModelId=5C8A682402C1
Public Function SetCategoryCodeStatVariable()

  'get the STAT category code
  Dim rsCat As New ADODB.Recordset
  Dim cnt As Integer
  Dim res As OperationStatus
  
  res = efailure
  cnt = 1
  CategoryCodeSTAT = ""
  
  'Set rsCat = m_dataSource.OpenRecordset("Select Categorie From Statutaire_Categorie Where Description = 'CodeStatutaire'", Snapshot)
  Set rsCat = m_dataSource.OpenRecordset("Select Categorie From Statutaire_Categorie", Snapshot)
  If Not rsCat.EOF() Then
    Do Until rsCat.EOF
      If cnt = 1 Then
        CategoryCodeSTAT = "'" & Trim$(rsCat.fields("Categorie")) & "'"
      Else
        CategoryCodeSTAT = CategoryCodeSTAT & "," & "'" & Trim$(rsCat.fields("Categorie")) & "'"
      End If
      
      cnt = cnt + 1
      rsCat.MoveNext
    Loop
    
    rsCat.Close
  End If
  
  res = eSuccess
  SetCategoryCodeStatVariable = res
    
  Exit Function
    
Error:

  Err.Raise Err.Number, "FctGlobales.SetCategoryCodeStatVariable - " & Err.Source, _
      " - " & Err.Description
   
  SetCategoryCodeStatVariable = res
        
End Function

'##ModelId=5C8A682402D1
Public Function BulkInsert(cnxn As ADODB.Connection, table As String, csvFile As String)

Dim sqlStr As String
Dim res As OperationStatus

  res = efailure

  'table = "TTPROVCOLL"
  sqlStr = "BULK INSERT " & table & " FROM "
  sqlStr = sqlStr & "'" & csvFile & "' "
   
    sqlStr = sqlStr & "WITH (FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
    'sqlStr = sqlStr & "WITH (KEEPIDENTITY, FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
        
    On Error GoTo BulkError
    cnxn.Execute sqlStr, , adExecuteNoRecords
    On Error GoTo 0
    
    'raise error if no records have been inserted
'    Set rsTestRows = New ADODB.Recordset
'    sqlStr = "SELECT TOP 1 * FROM " & table & " WHERE " & KeyName & " = " & NumPeriode  ' SELECT TOP 1 *
'    Set rsTestRows = cnxnSource.Execute(sqlStr)
'
'    If rsTestRows.RecordCount < 1 Then
'      Err.Raise ArchiveRestoreErrors.errBulkInsertNoRecordsInserted, "ArchiveRestore.BulkInsert", _
'      "Pendant l'opération Bulk Insert, aucune donnée n'a été inséré dans la table " & table & " pour la période " & NumPeriode
'    End If

    res = eSuccess
    BulkInsert = res
    
    Exit Function
    
BulkError:

  Err.Raise Err.Number, "FctGlobales.BulkInsert - " & Err.Source, _
      "Table : " & table & " - " & Err.Description
   
  BulkInsert = res

End Function


'##ModelId=5C8A6824031F
Public Function GetWinUser() As String

  Dim sDomainName As String * 255
  Dim lDomainNameLength As Long
  Dim sUserName As String
  Dim bUserSid(255) As Byte
  Dim lSIDType As Long
  
  sUserName = String(100, Chr$(0))
  GetUserName sUserName, 100
  GetWinUser = Left$(sUserName, InStr(sUserName, Chr$(0)) - 1)

End Function

'Establish connection to Archive DB
'##ModelId=5C8A6824032F
Public Function CreateArchiveConnection() As Boolean

  If DatabaseFileNameArchive = "" Then
    CreateArchiveConnection = False
    Exit Function
  End If
  
  If Not m_dataSourceArchive Is Nothing Then
    If m_dataSourceArchive.Connected Then
      m_dataSourceArchive.Disconnect
    End If
  End If
  
  Set m_dataSourceArchive = New P3IGeneraliDataAccess.DataAccess
  
  If m_dataSourceArchive.Connect(DatabaseFileNameArchive) = False Then
      MsgBox "Impossible d'ouvrir la base de données Archive!" & vbLf & "Source: FctGlobales.CreateArchiveConnection" _
      & vbLf & "Connection : " & DatabaseFileNameArchive, vbCritical
      CreateArchiveConnection = False
      Exit Function
  End If
  
  Set m_dataHelperArchive = m_dataSourceArchive.CreateHelper
  
  CreateArchiveConnection = True
  
End Function

'Close connection to Archive DB
'##ModelId=5C8A6824033E
Public Sub CloseArchiveConnection()

  If Not m_dataSourceArchive Is Nothing Then
    If m_dataSourceArchive.Connected Then
      m_dataSourceArchive.Disconnect
      Set m_dataSourceArchive = Nothing
      Set m_dataHelperArchive = Nothing
    End If
  End If
  
End Sub


'##ModelId=5C8A6824034E
Public Function Arrondi(Valeur As Double, Nbdecimale As Integer) As Double ' fonction arrondi
  Nbdecimale = Abs(Nbdecimale)
  Arrondi = Fix((Valeur * (10 ^ Nbdecimale) + 0.5)) / (10 ^ Nbdecimale)
End Function

'##ModelId=5C8A6824039C
Public Function Maximum(a As Integer, b As Integer) As Integer
  If a > b Then
    Maximum = a
  Else
    Maximum = b
  End If
End Function

'##ModelId=5C8A682403CB
Public Function DateMax(a As Date, b As Date) As Date
  If a > b Then
    DateMax = a
  Else
    DateMax = b
  End If
End Function

'##ModelId=5C8A68250021
Public Function Minimum(a As Integer, b As Integer) As Integer
  If a > b Then
    Minimum = b
  Else
    Minimum = a
  End If
End Function

'##ModelId=5C8A6825006F
Public Function EnableMenu(fEnabled As Boolean)
    With frmMain
        ' invalide les menus
   End With
End Function


' permet de changer d'utilisateur
'##ModelId=5C8A6825008F
Public Function UserLogin(fShowSplash As Boolean) As Boolean
    Dim rs As ADODB.Recordset
    
    ' invalide les menus
    EnableMenu (False)
    
    With frmMain
        On Error GoTo GestionErreurLogin
        
        Do
            ' demande le nom de l'utilisateur
            ret_code = 0
            Login.Show vbModal
            
            If ret_code = -1 Then
                ' valide les menus
                EnableMenu (True)
                UserLogin = False
                Exit Function
            End If
            
            ' initialisation
            ret_code = -1
            
            Set rs = m_dataSource.OpenRecordset("SELECT * FROM Utilisateur", Snapshot)
            rs.MoveFirst
            rs.Find " TANOM = '" & user_name & "'"
            
            If Not rs.EOF Then
                If rs.fields("TAPASS").Value <> user_pwd Then
                    MsgBox "Mot de passe non valide !", vbCritical + vbOKOnly, "Login"
                    rs.Close
                Else
                    ret_code = 0
                End If
            Else
                MsgBox "Utilisateur " & user_name & " inconnu !", vbCritical + vbOKOnly, "Login"
                rs.Close
            End If
           
            ' sortie de boucle si ok
            If ret_code = 0 Then
                Exit Do
            End If
        Loop
        
        ' invalide les menus
        EnableMenu (True)
        
        ' gestion des droits
        DroitAdmin = rs.fields("TAADMIN").Value
        
        frmMain.mnuAnnexes.Enabled = DroitAdmin
        
        rs.Close
    End With
      
    UserLogin = True
    
    Exit Function
    
GestionErreurLogin:
    MsgBox Err.Description, vbRetryCancel + vbCritical + vbApplicationModal, "Login"
    
    ret_code = -1
    
    Resume Next
End Function

'##ModelId=5C8A682500BE
Public Sub PlacePremierBoutton(theBoutton As CommandButton, top As Integer)
  theBoutton.Left = 0
  theBoutton.top = top
  theBoutton.Width = btnWidth
  theBoutton.Height = btnHeight
End Sub


'##ModelId=5C8A6825010C
Public Sub PlaceBoutton(theBoutton As CommandButton, BouttonGauche As CommandButton, top As Integer)
  theBoutton.Left = BouttonGauche.Left + BouttonGauche.Width + 50
  theBoutton.top = top
  theBoutton.Width = btnWidth
  theBoutton.Height = btnHeight
End Sub


' calcule la largeur des colonnes
' ObjWnd  : form contenant l'objet listview
' ObjList : objet listview
'##ModelId=5C8A6825015A
Sub LargeurAutomatique(ObjWnd As Form, ObjList As Object)
  Dim ColWidth() As Integer
  Dim i As Long, j As Long
    
  ' calcule la largeur des colonnes
  ReDim ColWidth(ObjList.ColumnHeaders.Count + 5) As Integer
  
  For i = 0 To ObjList.ColumnHeaders.Count - 1 Step 1
    ColWidth(i) = ObjWnd.TextWidth(ObjList.ColumnHeaders(i + 1))
  Next i
  
  For j = 1 To ObjList.ListItems.Count Step 1
    ColWidth(0) = Maximum(ColWidth(0), ObjWnd.TextWidth(ObjList.ListItems(j)))
  
    For i = 1 To ObjList.ColumnHeaders.Count - 1 Step 1
      ColWidth(i) = Maximum(ColWidth(i), ObjWnd.TextWidth(ObjList.ListItems(j).SubItems(i)))
    Next i
  Next j
  
  For i = 0 To ObjList.ColumnHeaders.Count - 1 Step 1
    ObjList.ColumnHeaders(i + 1).Width = ColWidth(i) + 100
  Next i
End Sub


'##ModelId=5C8A68250199
Public Function GetParameterFromCmdLine(cmdLine As String, param As String) As String
  Dim d As Integer, f As Integer
  
  GetParameterFromCmdLine = ""
  
  d = InStr(cmdLine, param)
  If d <> 0 Then
    f = InStr(d + Len(param), cmdLine, "/")
    If f <> 0 Then
      ' /<param>=path /<Other>
      GetParameterFromCmdLine = mID(cmdLine, d + Len(param), f - d - Len(param))
    Else
      ' /<param>=path
      GetParameterFromCmdLine = mID(cmdLine, d + Len(param))
    End If
    
    GetParameterFromCmdLine = Trim(GetParameterFromCmdLine)
  End If
End Function


'##ModelId=5C8A682501D7
Public Function ReplaceInString(texte As String, quoi As String, par As String) As String
  Dim pos As Integer
    
  ReplaceInString = texte
  Do
    pos = InStr(1, ReplaceInString, quoi)
    If pos <> 0 Then
      ReplaceInString = Left(ReplaceInString, pos - 1) & par & mID(ReplaceInString, pos + Len(quoi))
    End If
  Loop Until pos = 0
End Function

    
'##ModelId=5C8A68250235
Public Sub LargeurMaxColonneSpread(spr As fpSpread)
'  ' largeur des colonnes
'  Dim i As Integer, sprLargeur As fpSpread
'
'  Set sprLargeur = frmMain.sprLargeur
'
'  sprLargeur.maxRows = 1
'  sprLargeur.MaxCols = spr.MaxCols
'
'  spr.Row = 1
'  spr.Col = 1
'
'  sprLargeur.BlockMode = True
'  sprLargeur.Row = 1
'  sprLargeur.Col = 1
'  sprLargeur.Row2 = -1
'  sprLargeur.Col2 = sprLargeur.MaxCols
'  Set sprLargeur.Font = spr.Font
'  sprLargeur.BlockMode = False
'
'  spr.Row = SpreadHeader
'  sprLargeur.Row = 1
'  For i = 1 To spr.MaxCols
'    spr.Col = i
'    sprLargeur.Col = i
'
'    sprLargeur.text = spr.text
'
'    spr.ColWidth(i) = spr.MaxTextColWidth(i) + 1
'    sprLargeur.ColWidth(i) = sprLargeur.MaxTextColWidth(i) + 1
'    If spr.ColWidth(i) < sprLargeur.ColWidth(i) Then
'      spr.ColWidth(i) = sprLargeur.ColWidth(i)
'    End If
'  Next i
  
  spr.MaxRows = spr.MaxRows + 1
  
  ' largeur des colonnes
  Dim i As Integer, w As Double, txt As String
  
  For i = 1 To spr.MaxCols
    spr.Row = 0
    spr.Col = i
    txt = spr.text
    
    spr.Row = spr.MaxRows
    spr.CellType = CellTypeStaticText
    spr.text = txt
    
    spr.ColWidth(i) = spr.MaxTextColWidth(i) + 1
  Next i
  
  spr.MaxRows = spr.MaxRows - 1

End Sub



'##ModelId=5C8A68250263
Public Function Interpolation(ByVal x1 As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal x As Double) As Double
  
  If (x2 - x1) = 0 Then
    
    Interpolation = y1
  
  Else
  
    Interpolation = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
  
  End If

End Function


'**********************************************************
'PURPOSE:    Puts all lines of file into a string array
'PARAMETERS: FileName = FullPath of File
'            TheArray = StringArray to which contents
'                       Of File will be added.
'Example
'  Dim sArray() as String
'  FileToArray "C:\MyTextFile.txt", sArray
'  For lCtr = 0 to Ubound(sArray)
'  Debug.Print sArray(lCtr)
'  Next

'NOTES:
'  --  Requires a reference to Microsoft Scripting Runtime
'      Library
'  --  You can write this method in a number of different ways
'      For instance, you can take advantage of VB 6's ability to
'      return an array.
' --   You can also read all the contents of the file and use the
'      Split function with vbCrlf as the delimiter, but I
'      wanted to illustrate use of the ReadLine
'      and AtEndOfStream methods.
'##ModelId=5C8A682502E0
Public Function FileToString(ByVal filename As String) As String

  Dim oFSO As New FileSystemObject
  Dim oFSTR As Scripting.TextStream
  Dim ret As Long, lCtr As Long
  Dim sz As String

  FileToString = ""

  If Dir(filename) = "" Then Exit Function

  On Error GoTo ErrorHandler
  
  Set oFSTR = oFSO.OpenTextFile(filename)
   
  Do While Not oFSTR.AtEndOfStream
    
    sz = Trim(oFSTR.ReadLine)
    
    If Left(sz, 2) <> "--" Then
      FileToString = FileToString & sz & vbLf
    End If
      
  Loop
  
  oFSTR.Close
     

ErrorHandler:
  
  Set oFSTR = Nothing

End Function


