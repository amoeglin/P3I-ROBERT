Attribute VB_Name = "ImportExport"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67EB0245"
Option Explicit


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' enleve les caractères interdits du nom de fichier
'
'##ModelId=5C8A67EB032F
Public Function NormalizeName(Name As String) As String
  Dim n As Integer
  
  NormalizeName = Name
  
  n = InStr(NormalizeName, "  ")
  Do Until n = 0
    NormalizeName = Replace(NormalizeName, "  ", " ")
    n = InStr(NormalizeName, "  ")
  Loop
  
  NormalizeName = Replace(NormalizeName, "/", "-")
  NormalizeName = Replace(NormalizeName, "\", "-")
  NormalizeName = Replace(NormalizeName, ":", "-")
  NormalizeName = Replace(NormalizeName, ">", "-")
  NormalizeName = Replace(NormalizeName, "<", "-")
  NormalizeName = Replace(NormalizeName, "*", "-")
  NormalizeName = Replace(NormalizeName, "?", "-")
  NormalizeName = Replace(NormalizeName, """", "-")
  NormalizeName = Replace(NormalizeName, "|", "-")
  
  NormalizeName = Left(NormalizeName, 215) ' maximum 215 caracteres
End Function


'##ModelId=5C8A67EB036E
Private Function ExcelColumnName(colIndex As Integer) As String
  ' ExcelColumnName = IIf(colIndex > 26, Chr$(Asc("A") - 1 + Int(colIndex / 26)), "") & IIf(colIndex > 26, Chr$(Asc("A") + colIndex Mod 26), Chr$(Asc("A") - 1 + colIndex))

  If colIndex > 26 Then

    ' 1st character:  Subtract 1 to map the characters to 0-25,
    '                 but you don't have to remap back to 1-26
    '                 after the 'Int' operation since columns
    '                 1-26 have no prefix letter

    ' 2nd character:  Subtract 1 to map the characters to 0-25,
    '                 but then must remap back to 1-26 after
    '                 the 'Mod' operation by adding 1 back in
    '                 (included in the '65')

    ExcelColumnName = Chr(Int((colIndex - 1) / 26) + 64) & Chr(((colIndex - 1) Mod 26) + 65)
  Else
    ' Columns A-Z
    ExcelColumnName = Chr(colIndex + 64)
  End If
  
End Function


'##ModelId=5C8A67EB038D
Public Sub RenameSheetInExcelFile(filename As String, SheetName As String, tableName As String, MaxCol As Integer, nbHeaderRow As Integer)
  Dim xl As Excel.Application
  Dim books As Excel.Workbooks
  Dim srcBook As Excel.Workbook
  Dim sheet As Excel.Worksheet
  
  On Error GoTo err_ExcelFile
  
  'xl.Visible = True
  
  Set xl = New Excel.Application
  
  Set books = xl.Workbooks
  
  Set srcBook = books.Open(filename, , False)
  
  Set sheet = srcBook.Worksheets(SheetName)
  
  ' déverrouille la feuille
  sheet.Unprotect
  
  ' nomme la zone de données
  srcBook.Names.Add Name:=tableName, RefersToR1C1:="='" & SheetName & "'!C1:C" & MaxCol
  
  ' largeur des colonnes
  sheet.Columns("A:" & ExcelColumnName(MaxCol)).EntireColumn.AutoFit
  
  ' centre les noms de colonnes
  sheet.Rows("1:1").Select
  With xl.Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
  End With
  
  ' cache les colonnes inutiles
  sheet.Columns(ExcelColumnName(MaxCol + 1) & ":" & ExcelColumnName(MaxCol + 1)).Select
  sheet.Range(xl.Selection, xl.Selection.End(xlToRight)).Select
  xl.Selection.EntireColumn.Hidden = True

  ' fige les volets
  sheet.Range("A" & CStr(1 + nbHeaderRow)).Select
  xl.ActiveWindow.FreezePanes = True
  
  ' sauve le fichier
  srcBook.Close True
  
  xl.Quit

  Exit Sub

err_ExcelFile:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Export le contenu d'un spread vers un fichier Excel
' FileName      : nom de départ du fichier
' SheetName     : nom de l'onglet
' TableName     : nom de la zone de données créée
' spr           : le spread à exporter
' CommonDialog1 : onject CommonDialog pour le choix de l'emplacement du
'fichier
' strFilter     : texte du filtre si besoin (affichage en tete de fichier
'excel)
' bExportBCAC   : si true, les entetes de colonnes sont générés pour la
'syntaxe
'                 des tables du BCAC
'
'##ModelId=5C8A67EC0003
Public Sub ExportTableToExcelFile(filename As String, SheetName As String, tableName As String, ByRef spr As fpSpread, ByRef CommonDialog1 As CommonDialog, strFilter As String, bExportBCAC As Boolean, Optional bAskForFilename As Boolean = True, Optional progbar As ProgressBar)
  
  Dim nbRow As Integer, theFileName As String
  
  On Error GoTo err_export
  
  theFileName = filename
  
  ' enleve les caractères interdits du nom de fichier
  filename = NormalizeName(filename)
  SheetName = NormalizeName(SheetName)
  tableName = NormalizeName(tableName)
  
  Dim bVirtualMode As Boolean
  
  bVirtualMode = spr.VirtualMode
  
  
  If bAskForFilename = True Then
    
    
    If InStr(1, tableName, "statutaire", 1) > 0 And tableName <> "Statutaire_Garantie_Code" _
      And tableName <> "Statutaire_Garantie" And tableName <> "Statutaire_Categorie" Then ' spr.MaxRows > 65000 Then
    
'      If MsgBox("Cette opération peut être très long." & vbLf & "Voulez-vous continuer l'export ?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'      End If
      
      filename = Replace$(filename, ".xls", ".csv")
      
      Screen.MousePointer = vbHourglass
      
      CreateCSVFileStat tableName, filename, progbar  ' GetWinUser & "_" & filename
      
      
      'CommonDialog1.filename = Replace$(filename, ".xls", ".csv")
      'CommonDialog1.filter = "Fichier CSV|*.csv|All Files|*.*"
    Else
      ' demande le nom de la base (fichier xls)
      CommonDialog1.InitDir = GetSettingIni(CompanyName, "Dir", "ExportPath", App.Path)
      CommonDialog1.filename = filename
    
      CommonDialog1.filter = "Fichier Excel|*.xls|All Files|*.*"
      
      CommonDialog1.CancelError = True
      CommonDialog1.Flags = cdlOFNNoChangeDir + cdlOFNOverwritePrompt
      CommonDialog1.ShowSave
      
      If CommonDialog1.filename = "" Or CommonDialog1.filename = "*.xls" Then
        Exit Sub
      End If
      
      theFileName = CommonDialog1.filename
  
      Screen.MousePointer = vbHourglass
        
      If bVirtualMode And spr.VirtualMaxRows > 4000 Then
        If MsgBox("L'export vers Excel va d'abord charger la totalité des données (" & spr.VirtualMaxRows & " lignes)." & vbLf _
            & "Ceci peut être très long." & vbLf & "Voulez-vous continuer l'export ?", vbQuestion + vbYesNo) = vbNo Then
          Exit Sub
        End If
      End If
      
      spr.VirtualMode = False
      spr.DataRefresh
      
      ' si un filtre existe on ajoute une ligne de plus pour le filtre
      nbRow = IIf(strFilter <> "", 2, 1)
    
      ' etends le spread
      spr.MaxRows = spr.MaxRows + nbRow
      
      ' insere le nom de colonnes
      spr.BlockMode = True
      
      Dim txt As String, i As Integer
      
      spr.InsertRows 1, nbRow
 
      If bExportBCAC Then
       spr.MaxCols = spr.MaxCols + 1
       
       spr.InsertCols 1, 1
  
      End If
    
      spr.RowHeight(nbRow) = spr.RowHeight(nbRow) * 2
        
      spr.Row = nbRow
      spr.Col = 1
      spr.Row2 = nbRow
      spr.Col2 = spr.MaxCols
      spr.CellType = CellTypeStaticText
      spr.TypeHAlign = TypeHAlignCenter
      
      ' couleur des noms de colonnes
      spr.Row = nbRow
      spr.Col = 1
      spr.Row2 = nbRow
      spr.Col2 = spr.MaxCols
      spr.BackColor = &HC0C0C0   ' LTGRAY
      spr.TypeHAlign = TypeHAlignCenter
  
      If bExportBCAC Then
        spr.ColWidth(1) = spr.ColWidth(1) * 2
        
        spr.Row = -1
        spr.Col = 1
        spr.Row2 = -1
        spr.Col2 = 1
        spr.CellType = CellTypeStaticText
        spr.TypeHAlign = TypeHAlignCenter
        
        ' couleur des noms de colonnes
        spr.Row = -1
        spr.Col = 1
        spr.Row2 = -1
        spr.Col2 = 1
        spr.BackColor = &HC0C0C0   ' LTGRAY
        spr.TypeHAlign = TypeHAlignCenter
      End If
      
      ' filtre dans la ligne d'entète
      If strFilter <> "" Then
        ' couleur
        spr.Row = 1
        spr.Col = 1
        spr.Row2 = 1
        spr.Col2 = spr.MaxCols
        spr.BackColor = &HFFFF00   ' CYAN
        spr.TypeHAlign = TypeHAlignLeft
        
        spr.BlockMode = False
        
        ' copie le filtre dans la ligne d'entète
        spr.Row = 1
        spr.Col = 2
        spr.text = strFilter
      End If
      
      spr.BlockMode = False
    
      ' copie le titre des colonnes dans la ligne d'entète
      For i = 1 To spr.MaxCols
        spr.Row = 0
        spr.Col = i
        txt = spr.text
        
        If bExportBCAC Then
          If i = 1 Then
            txt = "Age"
          Else
            txt = Replace(txt, "Anc=", "")
            txt = "Anc" & txt
          End If
        End If
        
        spr.Row = nbRow
        spr.Col = i
        spr.text = txt
      Next i
      
      ' copie le titre des colonnes dans la ligne d'entète
      If bExportBCAC Then
        For i = nbRow + 1 To spr.MaxRows
          spr.Row = i
          spr.Col = 0
          txt = spr.text
          
          spr.Row = i
          spr.Col = 1
          spr.text = txt
        Next i
      End If
      
      ' do the export
      Dim nomSheet As String
      
      nomSheet = SheetName
    
    'If InStr(tableName, "statutaire") > -1 Then
      'CreateCSVFileStat tableName, Replace$(theFileName, ".xlsx", ".csv")
      'ExportSpreadContentToExcel theFileName, nomSheet, spr, tableName
    'Else
      spr.ExportToExcel theFileName, nomSheet, ""
    'End If
    

      If tableName <> "" Then
        ' renomme la zone
        RenameSheetInExcelFile theFileName, nomSheet, tableName, spr.MaxCols, nbRow
      End If
      
      ' supprime le nom de colonnes
      spr.BlockMode = True
    
      spr.DeleteRows 1, nbRow
  
    
      If bExportBCAC Then
        spr.DeleteCols 1, 1
        
        spr.MaxCols = spr.MaxCols - 1
      End If
      
      spr.BlockMode = False
      
      spr.MaxRows = spr.MaxRows - nbRow
      
      spr.VirtualMode = bVirtualMode
  
    End If
  
  End If
    
  Screen.MousePointer = vbDefault
    
  Exit Sub
  
err_export:
End Sub

'##ModelId=5C8A67EC009F
Public Function BulkInsertStat(table As String, csvFile As String) As Boolean
    
    Dim rsSource As ADODB.Recordset
    Dim rsTestRows As ADODB.Recordset
    Dim sqlStr As String
    Dim compName As String
    Dim ret As Long
    Dim numbRecords As Long
   
    sqlStr = "TRUNCATE TABLE " & table
    m_dataSource.Execute sqlStr
    
    sqlStr = "BULK INSERT " & table & " FROM "
    sqlStr = sqlStr & "'" & CSVUNCPath & csvFile & "' "
    
    '###Test keepidentity
    sqlStr = sqlStr & "WITH (FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
    'sqlStr = sqlStr & "WITH (KEEPIDENTITY, FIELDTERMINATOR = ';',ROWTERMINATOR = '\n')"   ', FIRSTROW = 1,KEEPNULLS,CODEPAGE='RAW')"
        
    On Error GoTo BulkError
    m_dataSource.Execute sqlStr ', , adExecuteNoRecords
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
    
    BulkInsertStat = True
    
    Exit Function
    
BulkError:

  BulkInsertStat = False
  
  MsgBox "La fonction 'Bulk Insert' a généré un erreur ! Source : BulkInsertStat. Code erreur :" & Err.Number & " Description erreur : " & Err.Description
    
'  Err.Raise Err.Number, "ImportExport.BulkInsert - " & Err.Source, _
'      "Table : " & table & " - " & Err.Description
    
    
End Function
'##ModelId=5C8A67EC00DD
Public Sub CreateCSVFileStat(table As String, csvFile As String, progbar As ProgressBar)

    Dim rs As ADODB.Recordset
    Dim sqlStr As String

    Dim f As ADODB.field
    Dim fieldValue As String
    Dim FieldName As String
    Dim fieldType As String
    Dim line As String
    Dim cnt As Long
    
    On Error Resume Next
    Close #1
        
OpenCSV:
  
        Open CSVUNCPath & csvFile For Output As #1
        'Open "T:\P3I_Generali\Import\" & csvFile For Output As #1
        
        If Err.Number = 70 Then
          'File is already open and needs to be closed
             
          MsgBox "Il semble que le fichier " & CSVUNCPath & csvFile & _
          " est ouvert. S'il vous plait fermez le fichier et cliquez sur le bouton Ok.", vbCritical
          
          Err.Clear
          GoTo OpenCSV
         End If
                
        'revert back to standard error handling
        'On Error GoTo 0
        On Error GoTo CSVError
                
        Set rs = m_dataSource.OpenRecordset("SELECT * FROM " & table, Dynamic)
        
         
        If rs.EOF Then
           'm_Logger.EcritTraceDansLog " La table : " & table & " est vide !"
           MsgBox "La table sélectionné est vide !"
        Else
          rs.MoveLast
          rs.MoveFirst
          
          If Not progbar Is Nothing Then
            progbar.Visible = True
            progbar.Max = rs.RecordCount
            progbar.Value = 0
          End If
        
          
          Do Until rs.EOF
            line = ""
            For Each f In rs.fields
              
              If Not IsNull(f.Value) Then
                fieldValue = f.Value
                
                'convert DateTime
                If f.Type = 135 And fieldValue <> "" Then
                  fieldValue = Format(fieldValue, "mm/dd/yyyy")
                End If
                
                'convert float data types : , -> .
                If f.Type = 5 And fieldValue <> "" Then
                  fieldValue = Replace$(fieldValue, ",", ".")
                End If

                If f.Type = 131 And fieldValue <> "" Then
                  fieldValue = Replace$(fieldValue, ",", ".")
                End If

                'delete the character ; from text fields
                If f.Type = 200 And InStr(fieldValue, ";") > 0 Then ' fieldValue <> "" Then
                  fieldValue = Replace$(fieldValue, ";", " ")
                End If

                'manage all the bit fields
'                If table = "P3IUser.Assure" Then
'                  If FieldName = "PODATEPAIEMENTESTIMEE" Or FieldName = "POPSAPCAPMOYEN" Or FieldName = "PODOSSIERCLOS" _
'                  Or FieldName = "POIsCadre" Or FieldName = "POPMReassAvecCorrectif" Or FieldName = "POPMAvecCorrectif" _
'                  Or FieldName = "POCaptive" Or FieldName = "POTopAmortissable" Then
'                      If UCase$(fieldValue) = "FALSE" Or UCase$(fieldValue) = "FAUX" Then
'                          fieldValue = "0"
'                      ElseIf UCase$(fieldValue) = "TRUE" Or UCase$(fieldValue) = "VRAI" Then
'                          fieldValue = "1"
'                      End If
'                  End If
'                End If

                'treat all delimiters and replace them with a space
                If InStr(fieldValue, ";") > 0 Then
                    fieldValue = Replace$(fieldValue, ";", " ")
                End If
                
                
                'write to csv file
                If line = "" Then
                  'treat the first element
                  If Not IsNull(f.Value) Then
                    line = fieldValue
                  End If
                Else
                  If Not IsNull(f.Value) Then
                    line = line & ";" & fieldValue
                  Else
                    line = line & ";"
                  End If
                End If
              
              End If
              
            Next
            
            Print #1, line
            DoEvents
            
            cnt = cnt + 1
            progbar.Value = cnt
          
            rs.MoveNext
          Loop
        
        End If
      

    Close #1
    
    MsgBox "Le fichier a été exporté sur : " & CSVUNCPath & csvFile
    
    Exit Sub
    
CSVError:
    
    Close #1
    
    'File is already open and needs to be closed
    Err.Raise Err.Number, "ArchiveRestore.CreateCSVFileStat", Err.Description
    '"Erreur pendant la création du fichier CSV : " & CSVUNCPath & csvFile
    
        
End Sub

'##ModelId=5C8A67EC011C
Public Sub ExportSpreadContentToExcel(sFileName As String, sSheetName As String, spr As fpSpread, nom_zone)
  
  Dim cs As String
  
  Screen.MousePointer = vbHourglass
  
  On Error Resume Next
  
  Kill sFileName
  
  On Error GoTo err_ExportSpreadContentToExcel
  
  'Start a new workbook in Excel
  Dim oApp As New Excel.Application
  Dim oBook As Excel.Workbook
  Dim oSheet As Excel.Worksheet
  
  Set oBook = oApp.Workbooks.Add
  Set oSheet = oBook.Worksheets(1)
  
  oSheet.Name = sSheetName
    
    
  'Add the field names in row 1
  Dim i As Integer, iNumCols As Integer, Row As Long, Col As Integer
  Dim cellVal As String
  
  iNumCols = spr.MaxCols ' rs.fields.Count
  
  'write column headers
  spr.Row = 1
  For i = 1 To iNumCols
    spr.Col = i
    oSheet.Cells(1, i).Value = spr.text
  Next
  
  'Add the data starting at cell A2
  'oSheet.Range("A2").CopyFromRecordset rs
  
  'write all other lines
  For Row = 2 To spr.MaxRows
    DoEvents
    For Col = 1 To spr.MaxCols
      spr.Row = Row
      spr.Col = Col
      
      oSheet.Cells(Row, Col).Value = spr.text
     Next Col
  Next Row
  
  
  'Format the header row as bold and autofit the columns
  With oSheet.Range("A1").Resize(1, iNumCols)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .EntireColumn.AutoFit
  End With
  
  ' Fige les volets
  oSheet.Range("A2").Select
  oApp.ActiveWindow.FreezePanes = True
  
  
  ' Entete sur fond gris
  oSheet.Rows("1:1").Select
  With oApp.Selection.Interior
      .ColorIndex = 15
      .Pattern = xlSolid
  End With
  oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1)).Select
  
  ' nom de la zone de données
  If nom_zone <> "" Then
    ' nomme la zone de données
    'oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(rs.RecordCount + 1, iNumCols)).Select
    oBook.Names.Add Name:=nom_zone, RefersToR1C1:="='" & sSheetName & "'!C1:C" & iNumCols, Visible:=True
  End If
  
  ' largeur des colonnes
  oSheet.Columns("A:" & ExcelColumnName(iNumCols)).EntireColumn.AutoFit
  
  ' Sauvegarde
  oSheet.SaveAs sFileName    ', 51    ', XlFileFormat.xlOpenXMLWorkbook   ', FileFormat:=56  ', XlFileFormat.xlExcel8
  
  
  'oApp.Visible = True ' affiche excel
  'oApp.UserControl = True ' rend excel à l'utilisateur (ne le ferme pas à la fin de la fonction)
  
  'oApp.ActiveWorkbook.Close False, sFileName
  oApp.Quit
  'Close the Recordset
  'rs.Close
  
  If Not oSheet Is Nothing Then
    Set oSheet = Nothing
  End If
  Set oBook = Nothing
  Set oApp = Nothing
 
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
err_ExportSpreadContentToExcel:

  Screen.MousePointer = vbDefault
  
  Set oSheet = Nothing
  Set oBook = Nothing
  Set oApp = Nothing
  
  MsgBox "Erreur durant l'export : " & Err & vbLf & Err.Description, vbCritical
  Exit Sub
  Resume Next
  
End Sub

'##ModelId=5C8A67EC016A
Public Sub ImportBCAC(CommonDialog1 As CommonDialog, ProgressBar1 As ProgressBar)
  Dim nomTable As String, libTable As String, CleTable As Long, typeTable As Integer
  Dim Connexion As String
  Dim xlSource As DataAccess
  Dim rsXL As ADODB.Recordset
  Dim rsTableloi As ADODB.Recordset
  Dim rsCR As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean, fReplaceExistingTable As Boolean
  Dim f As ADODB.field, FieldName As String
  Dim n As Integer, bookmark As Variant, cleTableBCAC As Long
    
  If DroitAdmin = False Then Exit Sub
  
  ' demande le nom du fichier xls
  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Import d'un table du BCAC"
  CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
  
  If CommonDialog1.filename = "" _
     Or CommonDialog1.filename = "*.xls" _
     Or CommonDialog1.filename = "*.xlsx" _
     Or CommonDialog1.filename = "*.*" Then
    Exit Sub
  End If

  On Error GoTo GestionErreur
  
  ' demande le nom de la table
  nomTable = CommonDialog1.FileTitle
  n = InStrRev(nomTable, ".")
  If n <> 0 Then
    nomTable = Left(nomTable, n - 1)
  End If
  
  n = 0
  fReplaceExistingTable = False
  Do
    libTable = InputBox("Entrez un libellé pour cette table :" & IIf(n <> 0, vbLf & "=> Le nom de la table doit être unique !", ""), "Nom de la table", nomTable)
    
    ' test l'unicité du nom
    If m_dataHelper.GetParameterAsLongWithParam("SELECT Count(*) FROM ListeTableLoi WHERE LIBTABLE=?", libTable) > 0 Then
      n = 1
      If MsgBox("Une table existe avec ce nom. Voulez-vous la remplacer ?", vbQuestion + vbYesNo) = vbYes Then
        n = 0
        fReplaceExistingTable = True
      End If
    End If
  Loop While n = 1
  
  If libTable = "" Then Exit Sub
  
  Dim m_Logger As New clsLogger
  
  m_Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurImport.log"
  m_Logger.CreateLog "Import " & CommonDialog1.filename & " dans la table " & libTable

  ' demande le type de table
  If MsgBox("La table '" & libTable & "' concerne-t-elle les personnes en Incapacité ?", vbQuestion + vbYesNo) = vbYes Then
    typeTable = cdTypeTableCoeffBCACIncap
    m_Logger.EcritTraceDansLog "La table concerne les personnes en Incapacité"
  Else
    typeTable = cdTypeTableCoeffBCACInval
    m_Logger.EcritTraceDansLog "La table concerne les personnes en Invalidité"
  End If
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Value = 0
  ProgressBar1.Max = 100
  ProgressBar1.Refresh
  
  Screen.MousePointer = vbHourglass
 
  ' cree une transaction
  fFoundError = False
  m_dataSource.BeginTrans
  
  ' ouvre les tables de destinations
  Set rsCR = m_dataSource.OpenRecordset("ProvisionBCAC", table)
  Set rsTableloi = m_dataSource.OpenRecordset("SELECT * FROM ListeTableLoi", Dynamic)
  
  If fReplaceExistingTable Then
    cleTableBCAC = m_dataHelper.GetParameterAsLongWithParam("SELECT TABLECLE FROM ListeTableLoi WHERE LIBTABLE=?", libTable)
    
    m_dataHelper.Multi_Find rsTableloi, "TABLECLE=" & cleTableBCAC
    If rsTableloi.EOF = False Then
      ' met à jour le type de table
      rsTableloi.fields("TYPETABLE") = typeTable
      rsTableloi.Update
    End If
    
    ' vide la table existante
    m_dataSource.Execute "DELETE FROM ProvisionBCAC WHERE CleTable=" & cleTableBCAC
    
    m_Logger.EcritTraceDansLog "L'import remplace la table existante '" & libTable & "'"
  Else
    Dim cls As clsListeTableLoi
    
    Set cls = New clsListeTableLoi
    
    ' creer la table dans ListeTableLoi
    'rsTableloi.AddNew
    
'    rsTableloi.fields("LIBTABLE") = libTable
'    rsTableloi.fields("NOMTABLE") = NomTable
'    rsTableloi.fields("TYPETABLE") = typeTable
'
'    rsTableloi.Update
'
'    ' recupere la cle de la table
'    cleTableBCAC = rsTableloi.fields("TABLECLE")
  
    cls.m_LIBTABLE = libTable
    cls.m_NOMTABLE = nomTable
    cls.m_TYPETABLE = typeTable
    cls.m_TableUtilisateur = True
    
    cls.Save m_dataSource
    
    cleTableBCAC = cls.m_TABLECLE
  End If
  
  ' ouvre la feuille excel
  Set xlSource = New DataAccess
  
  ' chaine de connexion ADO pour Excel
  'Connexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.filename & ";Extended Properties=" & cdExcelExtendedProperties & ";Persist Security Info=False"
  
  ' chaine de connexion
  
  If UCase(Right(CommonDialog1.filename, 4)) = ".XLS" Then
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  ElseIf UCase(Right(CommonDialog1.filename, 5)) = ".XLSX" Then
    Connexion = Replace(ConnectionStringXlsx, "%1", CommonDialog1.filename)
  Else
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  End If
  
  xlSource.Connect Connexion

  
  Set rsXL = xlSource.OpenRecordset("SELECT * FROM TableBCAC WHERE Age IS NOT NULL", Snapshot)
  
  If rsXL.EOF Then
    ProgressBar1.Max = 1
    m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
  Else
    rsXL.MoveLast
    rsXL.MoveFirst
  
    ProgressBar1.Max = rsXL.RecordCount + 1
    
    ' lit les enregistrements
    Do Until rsXL.EOF
      bRejet = False
      'If (rsXL.AbsolutePosition Mod 10) = 0 Then
        ' affiche la position
        ProgressBar1.Value = rsXL.AbsolutePosition
        ProgressBar1.Refresh
      'End If
      
      For Each f In rsXL.fields
        If f.Name <> "Age" And Not IsNull(f.Value) Then
          rsCR.AddNew
          
          FieldName = "CleTable"
          rsCR.fields("CleTable") = cleTableBCAC
          FieldName = "Age"
          rsCR.fields("Age") = rsXL.fields("Age")
          FieldName = "Anciennete"
          rsCR.fields("Anciennete") = CLng(mID(f.Name, 4)) ' AncXX
          FieldName = "Provision"
          rsCR.fields("Provision") = m_dataHelper.GetDouble2(f.Value)
          
          rsCR.Update
        End If
      Next
      
      rsXL.MoveNext
    Loop
  End If
  
  Call m_Logger.EcritTraceDansLog(rsXL.RecordCount & " lignes dans le fichier " & CommonDialog1.filename)
  
  Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
  
  m_Logger.EcritTraceDansLog "Fin Import de la table '" & libTable & "'"
  
  If fFoundError Then
    m_dataSource.RollbackTrans
  Else
    m_dataSource.CommitTrans
  End If
  
  rsXL.Close
  Set rsXL = Nothing
  
  rsTableloi.Close
  Set rsTableloi = Nothing
  
  rsCR.Close
  Set rsCR = Nothing
  
  xlSource.Disconnect
  
  Set xlSource = Nothing
  
  Screen.MousePointer = vbDefault
  
  ProgressBar1.Visible = False
  
  ' affichage des erreurs
  m_Logger.AfficheErreurLog
  
  Exit Sub
  
GestionErreur:
  If rsXL Is Nothing Then
    If Err = -2147467259 Then Resume
      
    m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
  Else
    Select Case Err
      Case 3265
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import ")
      
      Case 3421
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  fFoundError = True
  Resume Next
End Sub

'##ModelId=5C8A67EC01A8
Public Function ImportGeneriqueAuto(filename As String, nomTable As String, NumPeriodeConcernee As Long, m_Logger As clsLogger) As Boolean
  
  Dim Connexion As String
  Dim xlSource As DataAccess
  Dim rsXL As ADODB.Recordset
  Dim rsCR As ADODB.Recordset
  Dim rsCRArch As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean
  Dim f As ADODB.field, FieldName As String
  Dim n As Integer, bookmark As Variant
  Dim archiveOk As Boolean
  
  If Not m_dataSourceArchive Is Nothing Then
    If m_dataSourceArchive.Connected Then
      archiveOk = True
    Else
      archiveOk = False
    End If
  End If
    
  If DroitAdmin = False Then Exit Function
  'If NumPeriodeConcernee <= 0 Then Exit Sub
  
  ImportGeneriqueAuto = False

  On Error GoTo GestionErreur
  
  ' cree une transaction
  fFoundError = False
  m_dataSource.BeginTrans
  If archiveOk Then m_dataSourceArchive.BeginTrans
  
  If NumPeriodeConcernee <= 0 Then
    m_dataSource.Execute "DELETE FROM " & nomTable
    If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM " & nomTable
  Else
    m_dataSource.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & NumPeriodeConcernee
    If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & NumPeriodeConcernee
  End If
  
  ' ouvre les tables de destianation
  Set rsCR = m_dataSource.OpenRecordset("" & nomTable, table)
  If archiveOk Then Set rsCRArch = m_dataSourceArchive.OpenRecordset("" & nomTable, table)
  
  ' ouvre la feuille excel
  Set xlSource = New DataAccess
  
  ' chaine de connexion
  
  If UCase(Right(filename, 4)) = ".XLS" Then
    Connexion = Replace(ConnectionStringXls, "%1", filename)
  ElseIf UCase(Right(filename, 5)) = ".XLSX" Then
    Connexion = Replace(ConnectionStringXlsx, "%1", filename)
  Else
    Connexion = Replace(ConnectionStringXls, "%1", filename)
  End If
  
  xlSource.Connect Connexion

  Dim idField As Integer
  
  If NumPeriodeConcernee > 0 Then
    idField = IIf(nomTable = "ParamRentes", 4, 2)
  Else
    idField = 0
  End If
  Set rsXL = xlSource.OpenRecordset("SELECT * FROM " & nomTable & " WHERE " & rsCR.fields(idField).Name & " IS NOT NULL", Snapshot)
  
  idField = IIf(idField >= 2, idField - 2, idField)
  
  If Not rsXL Is Nothing Then
    If rsXL.EOF Then
      m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
    Else
      rsXL.MoveLast
      rsXL.MoveFirst
          
      ' lit les enregistrements
      Do Until rsXL.EOF
        bRejet = False
                
        If Not IsNull(rsXL.fields(idField)) Then
          rsCR.AddNew
          If archiveOk Then rsCRArch.AddNew
          
          For Each f In rsCR.fields
            FieldName = f.Name
            If f.Name = "NumPeriode" Then
              f.Value = NumPeriodeConcernee
            ElseIf f.Name = "GroupeCle" Then
              f.Value = GroupeCle
            Else
              f.Value = rsXL.fields(f.Name)
            End If
          Next
          
          If archiveOk Then
            For Each f In rsCRArch.fields
              FieldName = f.Name
              If f.Name = "NumPeriode" Then
                f.Value = NumPeriodeConcernee
              ElseIf f.Name = "GroupeCle" Then
                f.Value = GroupeCle
              Else
                f.Value = rsXL.fields(f.Name)
              End If
            Next
          End If
          
          If fFoundError = True Then
            rsCR.CancelUpdate
            If archiveOk Then rsCRArch.CancelUpdate
            Exit Do
          Else
            rsCR.Update
            If archiveOk Then rsCRArch.Update
          End If
        End If
        
        rsXL.MoveNext
      Loop
    End If
  End If
  
  Call m_Logger.EcritTraceDansLog(rsXL.RecordCount & " lignes dans le fichier " & filename)
  
  If fFoundError Then
    Call m_Logger.EcritTraceDansLog(">>>>> Fichier rejetté à cause des erreurs durant l'import !")
    
    m_dataSource.RollbackTrans
    If archiveOk Then m_dataSourceArchive.RollbackTrans
    
    ImportGeneriqueAuto = False
  Else
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
    
    m_dataSource.CommitTrans
    If archiveOk Then m_dataSourceArchive.CommitTrans
    
    ImportGeneriqueAuto = True
  End If
  
  m_Logger.EcritTraceDansLog "Fin Import de la table '" & nomTable & "'"
  
  If Not rsXL Is Nothing Then
    rsXL.Close
    Set rsXL = Nothing
  End If
  
  rsCR.Close
  Set rsCR = Nothing
  
  If archiveOk Then
    rsCRArch.Close
    Set rsCRArch = Nothing
  End If
  
  xlSource.Disconnect
  
  Set xlSource = Nothing
    
  ' affichage des erreurs
  'm_Logger.AfficheErreurLog
  
  Exit Function
  
GestionErreur:
  If rsXL Is Nothing Then
    If Err = -2147217865 Then
      m_Logger.EcritTraceDansLog "   Format Incorrect : le fichier " & filename & " ne correspond pas au format de la table '" & nomTable & "'. " & Err.Description
    Else
      m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
    End If
  Else
    Select Case Err
      Case 3265
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import ")
      
      Case 3421
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rsXL.AbsolutePosition)
      
      Case -2147217887
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' une valeur doit être précisée (NULL interdit) - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  fFoundError = True
  Resume Next
End Function


'##ModelId=5C8A67EC0206
Public Function ImportGenerique(CommonDialog1 As CommonDialog, ProgressBar1 As ProgressBar, nomTable As String, NumPeriodeConcernee As Long) As Boolean
  Dim Connexion As String
  Dim xlSource As DataAccess
  Dim rsXL As ADODB.Recordset
  Dim rsCR As ADODB.Recordset
  Dim rsCRArch As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean
  Dim f As ADODB.field, FieldName As String
  Dim n As Integer, bookmark As Variant
  Dim archiveOk As Boolean
  
  If Not m_dataSourceArchive Is Nothing Then
    If m_dataSourceArchive.Connected Then
      archiveOk = True
    Else
      archiveOk = False
    End If
  End If
    
  If DroitAdmin = False Then Exit Function
  'If NumPeriodeConcernee <= 0 Then Exit Sub
  
  ImportGenerique = False
  
  On Error GoTo GestionErreur
  
  ' demande le nom du fichier xls
  CommonDialog1.filename = "*.xls"
  CommonDialog1.DefaultExt = ".xls"
  CommonDialog1.DialogTitle = "Import de la table '" & nomTable & "'"
  CommonDialog1.filter = "Fichiers Excel|*.xls|Fichiers Excel 2007|*.xlsx|All Files|*.*"
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoChangeDir
  CommonDialog1.ShowOpen
  
  If CommonDialog1.filename = "" _
     Or CommonDialog1.filename = "*.xls" _
     Or CommonDialog1.filename = "*.xlsx" _
     Or CommonDialog1.filename = "*.*" Then
    Exit Function
  End If

  
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Value = 0
  ProgressBar1.Max = 100
  ProgressBar1.Refresh
  
  Screen.MousePointer = vbHourglass
 
  Dim m_Logger As New clsLogger
  
  m_Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurImport.log"
  m_Logger.CreateLog "Import " & CommonDialog1.filename & " dans la table '" & nomTable & "'"

  ' cree une transaction
  fFoundError = False
  m_dataSource.BeginTrans
  If archiveOk Then m_dataSourceArchive.BeginTrans
  
  If NumPeriodeConcernee <= 0 Then
    m_dataSource.Execute "DELETE FROM " & nomTable
    If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM " & nomTable
  Else
    m_dataSource.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & NumPeriodeConcernee
    If archiveOk Then m_dataSourceArchive.Execute "DELETE FROM " & nomTable & " WHERE GroupeCle=" & GroupeCle & " AND NumPeriode=" & NumPeriodeConcernee
  End If
  
  ' ouvre les tables de destianation
  Set rsCR = m_dataSource.OpenRecordset("" & nomTable, table)
  If archiveOk Then Set rsCRArch = m_dataSourceArchive.OpenRecordset("" & nomTable, table)
  
  ' ouvre la feuille excel
  Set xlSource = New DataAccess
  
  ' chaine de connexion ADO pour Excel
  'Connexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.filename & ";Extended Properties=" & cdExcelExtendedProperties & ";Persist Security Info=False"
  
  ' chaine de connexion
  
  If UCase(Right(CommonDialog1.filename, 4)) = ".XLS" Then
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  ElseIf UCase(Right(CommonDialog1.filename, 5)) = ".XLSX" Then
    Connexion = Replace(ConnectionStringXlsx, "%1", CommonDialog1.filename)
  Else
    Connexion = Replace(ConnectionStringXls, "%1", CommonDialog1.filename)
  End If
  
  xlSource.Connect Connexion

  Dim idField As Integer
  
  If NumPeriodeConcernee > 0 Then
    idField = IIf(nomTable = "ParamRentes", 4, 2)
  Else
    idField = 0
  End If
  Set rsXL = xlSource.OpenRecordset("SELECT * FROM " & nomTable & " WHERE " & rsCR.fields(idField).Name & " IS NOT NULL", Snapshot)
  
  idField = IIf(idField >= 2, idField - 2, idField)
  
  If Not rsXL Is Nothing Then
    If rsXL.EOF Then
      ProgressBar1.Max = 1
      m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
    Else
      rsXL.MoveLast
      rsXL.MoveFirst
    
      ProgressBar1.Max = rsXL.RecordCount + 1
      
      ' lit les enregistrements
      Do Until rsXL.EOF
        bRejet = False
        If (rsXL.AbsolutePosition Mod 5) = 0 Then
          ' affiche la position
          ProgressBar1.Value = rsXL.AbsolutePosition
          ProgressBar1.Refresh
        End If
        
        If Not IsNull(rsXL.fields(idField)) Then
          rsCR.AddNew
          If archiveOk Then rsCRArch.AddNew
          
          For Each f In rsCR.fields
            FieldName = f.Name
            If f.Name = "NumPeriode" Then
              f.Value = NumPeriodeConcernee
            ElseIf f.Name = "GroupeCle" Then
              f.Value = GroupeCle
            Else
              f.Value = rsXL.fields(f.Name)
            End If
          Next
          
          If archiveOk Then
            For Each f In rsCRArch.fields
              FieldName = f.Name
              If f.Name = "NumPeriode" Then
                f.Value = NumPeriodeConcernee
              ElseIf f.Name = "GroupeCle" Then
                f.Value = GroupeCle
              Else
                f.Value = rsXL.fields(f.Name)
              End If
            Next
          End If
          
          If fFoundError = True Then
            rsCR.CancelUpdate
            If archiveOk Then rsCRArch.CancelUpdate
            Exit Do
          Else
            rsCR.Update
            If archiveOk Then rsCRArch.Update
          End If
        End If
        
        rsXL.MoveNext
      Loop
    End If
  End If
  
  Call m_Logger.EcritTraceDansLog(rsXL.RecordCount & " lignes dans le fichier " & CommonDialog1.filename)
  
  If fFoundError Then
    Call m_Logger.EcritTraceDansLog(">>>>> Fichier rejetté à cause des erreurs durant l'import !")
    
    m_dataSource.RollbackTrans
    If archiveOk Then m_dataSourceArchive.RollbackTrans
    
    ImportGenerique = False
  Else
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
    
    m_dataSource.CommitTrans
    If archiveOk Then m_dataSourceArchive.CommitTrans
    
    ImportGenerique = True
  End If
  
  m_Logger.EcritTraceDansLog "Fin Import de la table '" & nomTable & "'"
  
  If Not rsXL Is Nothing Then
    rsXL.Close
    Set rsXL = Nothing
  End If
  
  rsCR.Close
  Set rsCR = Nothing
  
  If archiveOk Then
    rsCRArch.Close
    Set rsCRArch = Nothing
  End If
  
  xlSource.Disconnect
  
  Set xlSource = Nothing
  
  Screen.MousePointer = vbDefault
  
  ProgressBar1.Visible = False
  
  ' affichage des erreurs
  m_Logger.AfficheErreurLog
  
  Exit Function
  
GestionErreur:
  If Err = 32755 Then
    Resume Next
  End If
  
  If rsXL Is Nothing Then
    If Err = -2147217865 Then
      m_Logger.EcritTraceDansLog "   Format Incorrect : le fichier " & CommonDialog1.filename & " ne correspond pas au format de la table '" & nomTable & "'. " & Err.Description
    Else
      m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
    End If
  Else
    Select Case Err
      Case 3265
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import ")
      
      Case 3421
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rsXL.AbsolutePosition)
      
      Case -2147217887
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' une valeur doit être précisée (NULL interdit) - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  
  fFoundError = True
  Resume Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Genere un pivot en utilisant Excel
'
'##ModelId=5C8A67EC0264
Public Sub GenerateReport(prmTemp As ADODB.Recordset, prmPage As String, prmCol As String, prmRow As String, prmData As String, prmFile As String)

On Error GoTo Err_GenerateReport
   'VARIABLE DECLARATION

   Dim xlApp    As Excel.Application
   Dim xlBook   As Excel.Workbook
   Dim xlSheet  As Excel.Worksheet
   Dim xlSheet1 As Excel.Worksheet
   Dim rstemp   As ADODB.Recordset
   Dim intX     As Integer
   Dim intY     As Integer
   Set xlApp = New Excel.Application
   Set xlBook = xlApp.Workbooks.Add
   Set xlSheet = xlBook.Worksheets.Add
   xlSheet.Name = "Pivot"

   Set rstemp = prmTemp

   'DUMP THE RECORDSET TO EXCEL

   For intY = 0 To rstemp.fields.Count - 1
      xlSheet.Cells(intX + 1, intY + 1).Value = _
         rstemp.fields(intY).Name
   Next intY


   intX = intX + 1
   While Not rstemp.EOF
      For intY = 0 To rstemp.fields.Count - 1
         xlSheet.Cells(intX + 1, intY + 1).Value = _
            rstemp.fields(intY).Value
      Next intY
      rstemp.MoveNext
      intX = intX + 1
   Wend

   'DATA DUMPED, SO WE HAVE THE DATA ON WHICH TO PIVOT


   'ADDING A NEW WORKSHEET FOR THE PIVOT TABLE

   Set xlSheet1 = xlBook.Worksheets.Add
   xlSheet1.Name = "Report"

   'CREATING THE PIVOT TABLE

   xlBook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:= _
      "Pivot!R1C1:R" & rstemp.RecordCount + 1 _
      & "C" & rstemp.fields.Count).CreatePivotTable _
      TableDestination:=xlSheet1.Range("A9"), _
      tableName:="PivotTable1"

   xlSheet1.PivotTables("PivotTable1").SmallGrid = False

   'SETTING THE PAGE LEVEL FIELDS FROM THE PARAMETERS PASSED

   If Len(prmPage) > 0 Then
      For intX = 1 To Len(prmPage)
         With xlSheet1.PivotTables("PivotTable1").PivotFields( _
            rstemp.fields(CInt(mID$(prmPage, intX, 1))).Name)
            .Orientation = xlPageField
            .Position = intX
         End With
      Next intX
   End If

   'SETTING THE COL LEVEL FIELDS FROM THE PARAMETERS PASSED

   If Len(prmCol) > 0 Then
      For intX = 1 To Len(prmCol)
         With xlSheet1.PivotTables("PivotTable1").PivotFields( _
            rstemp.fields(CInt(mID$(prmCol, intX, 1))).Name)
            .Orientation = xlColumnField
            .Position = intX
         End With
      Next intX
   End If

   'SETTING THE ROW LEVEL FIELDS FROM THE PARAMETERS PASSED

   If Len(prmRow) > 0 Then
      For intX = 1 To Len(prmRow)
         With xlSheet1.PivotTables("PivotTable1").PivotFields( _
            rstemp.fields(CInt(mID$(prmRow, intX, 1))).Name)
            .Orientation = xlRowField
            .Position = intX
         End With
      Next intX
   End If

   'SETTING THE DATA FIELDS FROM THE PARAMETERS PASSED

   If Len(prmData) > 0 Then
      For intX = 1 To Len(prmData)
         With xlSheet1.PivotTables("PivotTable1").PivotFields( _
            rstemp.fields(CInt(mID$(prmData, intX, 1))).Name)
            .Orientation = xlDataField
            .Position = 1
         End With
      Next intX
   End If
   'HIDING THE PIVOTTABLE COMMANDBAR

   xlApp.CommandBars("PivotTable").Visible = False

   xlSheet1.Cells.EntireColumn.AutoFit
   xlSheet1.Range("A1").Select
   xlApp.DisplayAlerts = False
   'DELETING THE SHEET WITH THE SOURCE DATA - SO THAT NO ONE

   'CAN MODIFY

   xlSheet.Delete
   xlApp.DisplayAlerts = True
   xlSheet1.Range("A1").Select

   'SAVING THE EXCEL SHEET

   xlBook.SaveAs prmFile
   xlApp.Visible = True


Exit_GenerateReport:
   'xlBook.Close

   Set xlApp = Nothing
   Set xlBook = Nothing
   Set xlSheet = Nothing
   Set xlSheet1 = Nothing
   Set rstemp = Nothing
   Exit Sub

Err_GenerateReport:
   xlBook.Close
   Set xlApp = Nothing
   Set xlBook = Nothing
   Set xlSheet = Nothing
   Set xlSheet1 = Nothing
   Set rstemp = Nothing
   Err.Raise vbObjectError + 1500, "modReport.GenerateReport", _
      Err.Description

End Sub


'##ModelId=5C8A67EC02D1
Public Sub ExportBCACToExcelPivot(CleTable As Long)

On Error GoTo Err_optDet

   Dim sSql            As String
   Dim sPageLevelPivot As String
   Dim sColLevelPivot  As String
   Dim sRowLevelPivot  As String
   Dim sDataPivot      As String
   Dim ExistFile
   Dim rstemp As ADODB.Recordset

   Screen.MousePointer = vbHourglass

   'CONSTRUCT YOUR SQL QUERY HERE

   sSql = "SELECT Anciennete, Age, Provision FROM ProvisionBCAC WHERE CleTable=" & CleTable & " ORDER BY Anciennete, Age"

   'CREATE THE RECORDSET rstemp HERE

   '(BEFORE THAT CREATE THE CONNECTION TO THE DATA SOURCE)

   '******************************************

   Set rstemp = m_dataSource.OpenRecordset(sSql, Disconnected)

   'DELETING THE PREVIOUS REPORT IF IT EXISTS

   ExistFile = Dir("C:\Report.xls")
   If ExistFile <> "" Then Kill ("C:\Report.xls")

   'THIS IS WHERE THE FLEXIBILITY COMES IN, PASS THE PARAMETERS AS

   'FIELD INDEXES OF THE RECORDSET FOR THE GROUPING

   sPageLevelPivot = "0"
   sColLevelPivot = "1"
   'sRowLevelPivot = "2345"    'any combination can be passed

   sRowLevelPivot = "0"
   sDataPivot = "2"
   Call GenerateReport(rstemp, sPageLevelPivot, sColLevelPivot, _
                       sRowLevelPivot, sDataPivot, "C:\Report.xls")
   Screen.MousePointer = 0
Exit_optDet:

   Set rstemp = Nothing
   Exit Sub

Err_optDet:
   Screen.MousePointer = 0
   'VARIOUS ERROR HANDLERS

   Resume Exit_optDet
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Export SQL Query to new Excel Book
'
'##ModelId=5C8A67EC0300
Public Sub ExportQueryResultToExcel(SrcDB As DataAccess, sSql As String, sFileName As String, sSheetName As String, Optional ByRef spr As fpSpread = Nothing, Optional nom_zone As String = "", Optional autoMode As Boolean = False, Optional autoLogger As clsLogger)
  
  Dim strSQL As String, cs As String
  
  If Not autoMode Then
    Screen.MousePointer = vbHourglass
  End If
  
  On Error Resume Next
  
  Kill sFileName
  
  On Error GoTo err_ExportQueryResultToExcel
  
  'Start a new workbook in Excel
  Dim oApp As New Excel.Application
  Dim oBook As Excel.Workbook
  Dim oSheet As Excel.Worksheet
  
  Set oBook = oApp.Workbooks.Add
  Set oSheet = oBook.Worksheets(1)
  
  oSheet.Name = sSheetName
  
  Dim rs As ADODB.Recordset
  
  Set rs = SrcDB.OpenRecordset(sSql, Snapshot)
  
  'Add the field names in row 1
  Dim i As Integer, iNumCols As Integer
  
  iNumCols = rs.fields.Count
  For i = 1 To iNumCols
    oSheet.Cells(1, i).Value = rs.fields(i - 1).Name
  Next
  
  'Add the data starting at cell A2
  oSheet.Range("A2").CopyFromRecordset rs
  
  'Format the header row as bold and autofit the columns
  With oSheet.Range("A1").Resize(1, iNumCols)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .EntireColumn.AutoFit
  End With
  
  ' Fige les volets
  oSheet.Range("A2").Select
  oApp.ActiveWindow.FreezePanes = True
  
  ' couleur de fond des colonnes
  If Not IsMissing(spr) And Not spr Is Nothing Then
    spr.Row = 1
    For i = 1 To iNumCols - 1
      spr.Col = i
      If spr.BackColor <> vbWhite And spr.BackColor <> vbBlack Then
        cs = ExcelColumnName(i)
        oSheet.Range(cs & "2:" & cs & "65535").Select
        'oSheet.Range(cs & "2:" & cs & spr.MaxRows + 1).Select
        With oApp.Selection.Interior
            If spr.BackColor > 0 Then
              .color = spr.BackColor
              .Pattern = xlSolid
            End If
        End With
      End If
    Next
    
'    cs = ExcelColumnName(iNumCols - 1)
'    spr.Col = 1
'    For i = 2 To rs.RecordCount
'      oSheet.Range("A" & i & ":" & cs & i).Select
'      With oApp.Selection.Interior
'          spr.Row = i
'          .color = spr.BackColor
'          .Pattern = xlSolid
'      End With
'    Next
  End If
  
  ' Entete sur fond gris
  oSheet.Rows("1:1").Select
  With oApp.Selection.Interior
      .ColorIndex = 15
      .Pattern = xlSolid
  End With
  oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1)).Select
  
  ' nom de la zone de données
  If nom_zone <> "" Then
    ' nomme la zone de données
    'oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(rs.RecordCount + 1, iNumCols)).Select
    oBook.Names.Add Name:=nom_zone, RefersToR1C1:="='" & sSheetName & "'!C1:C" & iNumCols, Visible:=True
  End If
  
  ' largeur des colonnes
  oSheet.Columns("A:" & ExcelColumnName(iNumCols)).EntireColumn.AutoFit
  
  ' Sauvegarde
  oSheet.SaveAs sFileName    ', XlFileFormat.xlOpenXMLWorkbook   ', FileFormat:=56  ', XlFileFormat.xlExcel8
  
  If Not autoMode Then
    oApp.Visible = True ' affiche excel
    oApp.UserControl = True ' rend excel à l'utilisateur (ne le ferme pas à la fin de la fonction)
  Else
    oApp.Visible = True
    oApp.UserControl = True
    
    'oBook.Close
    'Set oSheet = Nothing
    'Set oBook = Nothing
    'oApp.Quit
    'Set oApp = Nothing
  End If
  
  'Close the Recordset
  rs.Close
  
  'Set oSheet = Null
  'Set oBook = Null
  
  If Not autoMode Then
    Screen.MousePointer = vbDefault
  End If
  
  Exit Sub
  
err_ExportQueryResultToExcel:

  If Not autoMode Then
    Screen.MousePointer = vbDefault
    MsgBox "Erreur durant l'export : " & Err & vbLf & Err.Description, vbCritical
  Else
    autoLogger.EcritTraceDansLog "Erreur durant l'export : " & Err & vbLf & Err.Description
    
    'oSheet.SaveAs sFileName
    'oApp.Quit
    'Set oSheet = Nothing
    'Set oBook = Nothing
    'Set oApp = Nothing
  End If
  
  Exit Sub
  Resume Next
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Creation du fichier top de signalisation
'
'##ModelId=5C8A67EC038D
Public Sub CreationFichierSignalisation()
  
  Dim sFichierTop As String
  Dim fErreur As Boolean
  
  
  On Error GoTo err_export
  
  
  sFichierTop = GetSettingIni(SectionName, "Dir", "FichierTop", "#")
  If sFichierTop = "#" Then
    
    ' erreur de parametrage
    MsgBox "Vous devez spécifier le chemin complet du fichier 'top' dans le fichier de parametre " & vbLf & sFichierIni & vbLf & "Section [DB], Entrée FichierTop ", vbCritical
  
  Else
        
    ' signal la présence des données
    Dim FileNumber As Integer
    
    FileNumber = FreeFile   ' Get unused file
    
    fErreur = False
    Open sFichierTop For Output As #FileNumber   ' Create file name.
    Print #FileNumber, "OK " & Now   ' Output text.
    Close #FileNumber         ' Close file.
    
    If fErreur = False Then
      MsgBox "Le fichier de signalisation a été généré avec succès !", vbInformation
    Else
      MsgBox "Impossible de créer le fichier de signalisation !", vbCritical
    End If
    
  End If

  
  Exit Sub


err_export:
  
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  
  fErreur = True
  
  Resume Next

End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Création de la structure des loi de maintien avant import.
' Types de table autorisé : cdTypeTable_LoiMaintienIncapacite, cdTypeTable_LoiPassage,
'cdTypeTable_LoiMaintienInvalidite
'
' Attention : la fonction ne vérifie pas si la table existe déjà.
'
'##ModelId=5C8A67EC039C
Public Function CreateTable(typeTable As Integer, nomTable As String, destDB As DataAccess) As Boolean
  
  Dim szSQL As String
  
  
  
  szSQL = "CREATE TABLE " & nomTable & "( "
  
  If typeTable = cdTypeTable_LoiMaintienIncapacite Then
    
    szSQL = szSQL & " Age smallint NOT NULL, Anc0 float NULL, Anc1 float NULL, Anc2 float NULL, Anc3 float NULL, " _
                  & " Anc4 float NULL, Anc5 float NULL, Anc6 float NULL, Anc7 float NULL, Anc8 float NULL, " _
                  & " Anc9 float NULL, Anc10 float NULL, Anc11 float NULL, Anc12 float NULL, Anc13 float NULL, " _
                  & " Anc14 float NULL, Anc15 float NULL, Anc16 float NULL, Anc17 float NULL, Anc18 float NULL, " _
                  & " Anc19 float NULL, Anc20 float NULL, Anc21 float NULL, Anc22 float NULL, Anc23 float NULL, " _
                  & " Anc24 float NULL, Anc25 float NULL, Anc26 float NULL, Anc27 float NULL, Anc28 float NULL, " _
                  & " Anc29 float NULL, Anc30 float NULL, Anc31 float NULL, Anc32 float NULL, Anc33 float NULL, " _
                  & " Anc34 float NULL, Anc35 float NULL, Anc36 float NULL, Comment varchar(255) NULL, " _
                  & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
                  
  ElseIf typeTable = cdTypeTable_BaremeAnneeStatutaire Then
  
  '###
    szSQL = szSQL & " StatId int NOT NULL, " _
                  & " TableAnnee int NOT NULL, " _
                  & " Collect nvarchar(10) NOT NULL, " _
                  & " TypeSinistre nvarchar(10) NOT NULL, " _
                  & " Sexe int NOT NULL, " _
                  & " AgeMalade int NOT NULL, " _
                  & " Semaine int NOT NULL, " _
                  & " PM_AT float NULL, " _
                  & " PM_CLD float NULL, " _
                  & " PM_CLM float NULL, " _
                  & " PM_CLM_CLD float NULL, " _
                  & " PM_MO float NULL, " _
                  & " PM_MO_CLM float NULL, " _
                  & " PM_MO_CLD float NULL, " _
                  & " PM_Total float NULL, " _
                  & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(StatId Asc) )"
    
  ElseIf typeTable = cdTypeTable_LoiMaintienInvalidite Then
  
'    szSQL = szSQL & " Age smallint NOT NULL, Anc0 float NULL, Anc1 float NULL, Anc2 float NULL, Anc3 float NULL, " _
                  & " Anc4 float NULL, Anc5 float NULL, Anc6 float NULL, Anc7 float NULL, Anc8 float NULL, " _
                  & " Anc9 float NULL, Anc10 float NULL, Anc11 float NULL, Anc12 float NULL, Anc13 float NULL, " _
                  & " Anc14 float NULL, Anc15 float NULL, Anc16 float NULL, Anc17 float NULL, Anc18 float NULL, " _
                  & " Anc19 float NULL, Anc20 float NULL, Anc21 float NULL, Anc22 float NULL, Anc23 float NULL, " _
                  & " Anc24 float NULL, Anc25 float NULL, Anc26 float NULL, Anc27 float NULL, Anc28 float NULL, " _
                  & " Anc29 float NULL, Anc30 float NULL, Anc31 float NULL, Anc32 float NULL, Anc33 float NULL, " _
                  & " Anc34 float NULL, Anc35 float NULL, Anc36 float NULL, Anc37 float NULL, Anc38 float NULL, " _
                  & " Anc39 float NULL, Anc40 float NULL, Anc41 float NULL, Anc42 float NULL, Anc43 float NULL, " _
                  & " Anc44 float NULL, Anc45 float NULL, Anc46 float NULL, Anc47 float NULL, Comment varchar(255) NULL, " _
                  & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
  
 ' modif du 10/08/2017  RS AM prologation jusqu'à Anc102 (au lieu Anc47)
    szSQL = szSQL & " Age smallint NOT NULL, " _
    & " Anc0 float Null, Anc1 float Null, Anc2 float Null, Anc3 float Null, Anc4 float Null, Anc5 float Null, Anc6 float Null, Anc7 float Null, Anc8 float Null, Anc9 float Null, Anc10 float Null, " _
    & " Anc11 float Null, Anc12 float Null, Anc13 float Null, Anc14 float Null, Anc15 float Null, Anc16 float Null, Anc17 float Null, Anc18 float Null, Anc19 float Null, Anc20 float Null, " _
    & " Anc21 float Null, Anc22 float Null, Anc23 float Null, Anc24 float Null, Anc25 float Null, Anc26 float Null, Anc27 float Null, Anc28 float Null, Anc29 float Null, Anc30 float Null, " _
    & " Anc31 float Null, Anc32 float Null, Anc33 float Null, Anc34 float Null, Anc35 float Null, Anc36 float Null, Anc37 float Null, Anc38 float Null, Anc39 float Null, Anc40 float Null, " _
    & " Anc41 float Null, Anc42 float Null, Anc43 float Null, Anc44 float Null, Anc45 float Null, Anc46 float Null, Anc47 float Null, Anc48 float Null, Anc49 float Null, Anc50 float Null, " _
    & " Anc51 float Null, Anc52 float Null, Anc53 float Null, Anc54 float Null, Anc55 float Null, Anc56 float Null, Anc57 float Null, Anc58 float Null, Anc59 float Null, Anc60 float Null, " _
    & " Anc61 float Null, Anc62 float Null, Anc63 float Null, Anc64 float Null, Anc65 float Null, Anc66 float Null, Anc67 float Null, Anc68 float Null, Anc69 float Null, Anc70 float Null, " _
    & " Anc71 float Null, Anc72 float Null, Anc73 float Null, Anc74 float Null, Anc75 float Null, Anc76 float Null, Anc77 float Null, Anc78 float Null, Anc79 float Null, Anc80 float Null, " _
    & " Anc81 float Null, Anc82 float Null, Anc83 float Null, Anc84 float Null, Anc85 float Null, Anc86 float Null, Anc87 float Null, Anc88 float Null, Anc89 float Null, Anc90 float Null, " _
    & " Anc91 float Null, Anc92 float Null, Anc93 float Null, Anc94 float Null, Anc95 float Null, Anc96 float Null, Anc97 float Null, Anc98 float Null, Anc99 float Null, Anc100 float Null, "

    szSQL = szSQL & " Anc101 float Null, Anc102 float Null, Comment varchar(255) NULL, " _
    & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
   
  
  
  ElseIf typeTable = cdTypeTable_LoiPassage Then

    szSQL = szSQL & " Age smallint NOT NULL, Anc0 float NULL, Anc1 float NULL, Anc2 float NULL, Anc3 float NULL, " _
                  & " Anc4 float NULL, Anc5 float NULL, Anc6 float NULL, Anc7 float NULL, Anc8 float NULL, " _
                  & " Anc9 float NULL, Anc10 float NULL, Anc11 float NULL, Anc12 float NULL, Anc13 float NULL, " _
                  & " Anc14 float NULL, Anc15 float NULL, Anc16 float NULL, Anc17 float NULL, Anc18 float NULL, " _
                  & " Anc19 float NULL, Anc20 float NULL, Anc21 float NULL, Anc22 float NULL, Anc23 float NULL, " _
                  & " Anc24 float NULL, Anc25 float NULL, Anc26 float NULL, Anc27 float NULL, Anc28 float NULL, " _
                  & " Anc29 float NULL, Anc30 float NULL, Anc31 float NULL, Anc32 float NULL, Anc33 float NULL, " _
                  & " Anc34 float NULL, Anc35 float NULL, Comment varchar(255) NULL, " _
                  & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
  
  ElseIf typeTable = cdTypeTable_LoiDependance Then

    szSQL = szSQL & " Age smallint NOT NULL, " _
    & " Anc0 float Null, Anc1 float Null, Anc2 float Null, Anc3 float Null, Anc4 float Null, Anc5 float Null, Anc6 float Null, Anc7 float Null, Anc8 float Null, Anc9 float Null, Anc10 float Null, " _
    & " Anc11 float Null, Anc12 float Null, Anc13 float Null, Anc14 float Null, Anc15 float Null, Anc16 float Null, Anc17 float Null, Anc18 float Null, Anc19 float Null, Anc20 float Null, " _
    & " Anc21 float Null, Anc22 float Null, Anc23 float Null, Anc24 float Null, Anc25 float Null, Anc26 float Null, Anc27 float Null, Anc28 float Null, Anc29 float Null, Anc30 float Null, " _
    & " Anc31 float Null, Anc32 float Null, Anc33 float Null, Anc34 float Null, Anc35 float Null, Anc36 float Null, Anc37 float Null, Anc38 float Null, Anc39 float Null, Anc40 float Null, " _
    & " Anc41 float Null, Anc42 float Null, Anc43 float Null, Anc44 float Null, Anc45 float Null, Anc46 float Null, Anc47 float Null, Anc48 float Null, Anc49 float Null, Anc50 float Null, " _
    & " Anc51 float Null, Anc52 float Null, Anc53 float Null, Anc54 float Null, Anc55 float Null, Anc56 float Null, Anc57 float Null, Anc58 float Null, Anc59 float Null, Anc60 float Null, " _
    & " Anc61 float Null, Anc62 float Null, Anc63 float Null, Anc64 float Null, Anc65 float Null, Anc66 float Null, Anc67 float Null, Anc68 float Null, Anc69 float Null, Anc70 float Null, " _
    & " Anc71 float Null, Anc72 float Null, Anc73 float Null, Anc74 float Null, Anc75 float Null, Anc76 float Null, Anc77 float Null, Anc78 float Null, Anc79 float Null, Anc80 float Null, " _
    & " Anc81 float Null, Anc82 float Null, Anc83 float Null, Anc84 float Null, Anc85 float Null, Anc86 float Null, Anc87 float Null, Anc88 float Null, Anc89 float Null, Anc90 float Null, " _
    & " Anc91 float Null, Anc92 float Null, Anc93 float Null, Anc94 float Null, Anc95 float Null, Anc96 float Null, Anc97 float Null, Anc98 float Null, Anc99 float Null, Anc100 float Null, "

    szSQL = szSQL & " Anc101 float Null, Anc102 float Null, Anc103 float Null, Anc104 float Null, Anc105 float Null, Anc106 float Null, Anc107 float Null, Anc108 float Null, Anc109 float Null, Anc110 float Null, " _
    & " Anc111 float Null, Anc112 float Null, Anc113 float Null, Anc114 float Null, Anc115 float Null, Anc116 float Null, Anc117 float Null, Anc118 float Null, Anc119 float Null, Anc120 float Null, " _
    & " Anc121 float Null, Anc122 float Null, Anc123 float Null, Anc124 float Null, Anc125 float Null, Anc126 float Null, Anc127 float Null, Anc128 float Null, Anc129 float Null, Anc130 float Null, " _
    & " Anc131 float Null, Anc132 float Null, Anc133 float Null, Anc134 float Null, Anc135 float Null, Anc136 float Null, Anc137 float Null, Anc138 float Null, Anc139 float Null, Anc140 float Null, " _
    & " Anc141 float Null, Anc142 float Null, Anc143 float Null, Anc144 float Null, Anc145 float Null, Anc146 float Null, Anc147 float Null, Anc148 float Null, Anc149 float Null, Anc150 float Null, " _
    & " Anc151 float Null, Anc152 float Null, Anc153 float Null, Anc154 float Null, Anc155 float Null, Anc156 float Null, Anc157 float Null, Anc158 float Null, Anc159 float Null, Anc160 float Null, " _
    & " Anc161 float Null, Anc162 float Null, Anc163 float Null, Anc164 float Null, Anc165 float Null, Anc166 float Null, Anc167 float Null, Anc168 float Null, Anc169 float Null, Anc170 float Null, " _
    & " Anc171 float Null, Anc172 float Null, Anc173 float Null, Anc174 float Null, Anc175 float Null, Anc176 float Null, Anc177 float Null, Anc178 float Null, Anc179 float Null, Anc180 float Null, " _
    & " Anc181 float Null, Anc182 float Null, Anc183 float Null, Anc184 float Null, Anc185 float Null, Anc186 float Null, Anc187 float Null, Anc188 float Null, Anc189 float Null, Anc190 float Null, " _
    & " Anc191 float Null, Anc192 float Null, Anc193 float Null, Anc194 float Null, Anc195 float Null, Anc196 float Null, Anc197 float Null, Anc198 float Null, Anc199 float Null, Anc200 float Null, " _
    & " Anc201 float Null, Anc202 float Null, Anc203 float Null, Anc204 float Null, Anc205 float Null, Anc206 float Null, Anc207 float Null, Anc208 float Null, Anc209 float Null, Anc210 float Null, " _
    & " Anc211 float Null, Anc212 float Null, Anc213 float Null, Anc214 float Null, Anc215 float Null, Anc216 float Null, Anc217 float Null, Anc218 float Null, Anc219 float Null, Anc220 float Null, " _
    & " Anc221 float Null, Anc222 float Null, Anc223 float Null, Anc224 float Null, Anc225 float Null, Anc226 float Null, Anc227 float Null, Anc228 float Null, Anc229 float Null, Anc230 float Null, " _
    & " Anc231 float Null, Anc232 float Null, Anc233 float Null, Anc234 float Null, Anc235 float Null, Anc236 float Null, Anc237 float Null, Anc238 float Null, Anc239 float Null, Anc240 float Null, Comment varchar(255) NULL, " _
    & " CONSTRAINT PK_" & nomTable & " PRIMARY KEY NONCLUSTERED(Age Asc) )"
  
  
  Else
    
    CreateTable = False
    Exit Function
  
  End If

  On Error GoTo err_CreateTable
  
  destDB.Execute szSQL
  
  CreateTable = True
  
  Exit Function
  
err_CreateTable:
  MsgBox "Erreur lors de la création de la table " & nomTable & " : " & Err.Description, vbCritical
  CreateTable = False
End Function


'##ModelId=5C8A67ED0003
Public Function ImportTableMortalite(NomFichier As String, NomTableDest As String, CleTableDest As Long, NomTableSrc As String) As Boolean
  Dim Connexion As String
  Dim xlSource As DataAccess
  Dim rsXL As ADODB.Recordset
  Dim rsCR As ADODB.Recordset
  Dim NbRejet As Long, bRejet As Boolean
  Dim fFoundError As Boolean
  Dim f As ADODB.field, FieldName As String
  Dim n As Integer, bookmark As Variant
    
  If DroitAdmin = False Then Exit Function
  'If NumPeriodeConcernee <= 0 Then Exit Sub
  
  ImportTableMortalite = False
  
  On Error GoTo GestionErreur
  
  Screen.MousePointer = vbHourglass
 
  Dim m_Logger As New clsLogger
  
  m_Logger.FichierLog = m_logPath & "\" & GetWinUser & "_ErreurImport.log"
  m_Logger.CreateLog "Import " & NomFichier & " : table '" & NomTableSrc & "'"

  ' cree une transaction
  fFoundError = False
  
  ' ouvre les tables de destianation
  Set rsCR = m_dataSource.OpenRecordset("" & NomTableDest, table)
  
  ' ouvre la feuille excel
  Set xlSource = New DataAccess
  
  ' chaine de connexion ADO pour Excel
  'Connexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CommonDialog1.filename & ";Extended Properties=" & cdExcelExtendedProperties & ";Persist Security Info=False"
  
  ' chaine de connexion
  
  If UCase(Right(NomFichier, 4)) = ".XLS" Then
    Connexion = Replace(ConnectionStringXls, "%1", NomFichier)
  ElseIf UCase(Right(NomFichier, 5)) = ".XLSX" Then
    Connexion = Replace(ConnectionStringXlsx, "%1", NomFichier)
  Else
    Connexion = Replace(ConnectionStringXls, "%1", NomFichier)
  End If
  
  xlSource.Connect Connexion

  Dim idField As Integer
  
  If NomTableDest = "TableMortalite" Or NomTableDest = "MortIncap" Or NomTableDest = "MortInval" Then
    idField = 1
  Else
    idField = 0
  End If
  
  Set rsXL = xlSource.OpenRecordset("SELECT * FROM " & NomTableSrc & " WHERE " & rsCR.fields(idField).Name & " IS NOT NULL", Snapshot)
  
  If Not rsXL Is Nothing Then
    If rsXL.EOF Then
      m_Logger.EcritTraceDansLog "   Aucun enregistrement trouvé"
    Else
      
      ' lit les enregistrements
      Do Until rsXL.EOF
        bRejet = False
        
        If Not IsNull(rsXL.fields(0)) Then
          rsCR.AddNew
          
          For Each f In rsCR.fields
            FieldName = f.Name
            If f.Name = "CleTable" Then
              f.Value = CleTableDest
            ElseIf f.Name = "NomTable" Then
              f.Value = NomTableSrc
            Else
              f.Value = rsXL.fields(f.Name)
            End If
          Next
          
          If fFoundError = True Then
            rsCR.CancelUpdate
            Exit Do
          Else
            rsCR.Update
          End If
        End If
        
        rsXL.MoveNext
      Loop
    End If
  End If
  
  If Not (rsXL Is Nothing) Then
    Call m_Logger.EcritTraceDansLog(rsXL.RecordCount & " lignes dans le fichier " & NomFichier)
  End If
  
  If fFoundError Then
    Call m_Logger.EcritTraceDansLog(">>>>> Fichier rejetté à cause des erreurs durant l'import !")
    
    ImportTableMortalite = False
  Else
    Call m_Logger.EcritTraceDansLog(NbRejet & " rejet" & IIf(NbRejet = 0, "", "s") & " durant l'import")
    
    ImportTableMortalite = True
  End If
  
  m_Logger.EcritTraceDansLog "Fin Import de la table '" & NomTableDest & "'"
  
  If Not rsXL Is Nothing Then
    rsXL.Close
    Set rsXL = Nothing
  End If
  
  rsCR.Close
  Set rsCR = Nothing
  
  xlSource.Disconnect
  
  Set xlSource = Nothing
  
  Screen.MousePointer = vbDefault
  
  ' affichage des erreurs
  m_Logger.AfficheErreurLog
  
  Exit Function
  
GestionErreur:
  If rsXL Is Nothing Then
    If Err = -2147217865 Then
      m_Logger.EcritTraceDansLog "   Format Incorrect : le fichier " & NomFichier & " ne correspond pas au format de la table '" & NomTableDest & "'. " & Err.Description
    Else
      m_Logger.EcritTraceDansLog "   Erreur " & Err & " : " & Err.Description
    End If
  Else
    Select Case Err
      Case 3265
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' introuvable dans le fichier d'import ")
      
      Case 3421
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' type de donnée non correcte - Ligne " & rsXL.AbsolutePosition)
      
      Case -2147217887
        Call m_Logger.EcritTraceDansLog("   Erreur " & Err & " : Colonne '" & FieldName & "' une valeur doit être précisée (NULL interdit) - Ligne " & rsXL.AbsolutePosition)
      
      Case Else
        m_Logger.EcritTraceDansLog "Erreur " & Err & " : " & Err.Description & " - Ligne " & rsXL.AbsolutePosition
    End Select
  End If
  fFoundError = True
  Resume Next
End Function


