VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmListeLotSASP3I 
   Caption         =   "Jeux de données disponibles"
   ClientHeight    =   3915
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   11595
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   8370
      TabIndex        =   4
      Top             =   3555
      Visible         =   0   'False
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox lblAvancement 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4005
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "----"
      Top             =   3420
      Width           =   4335
   End
   Begin VB.CommandButton btnUtiliser 
      Caption         =   "&Utiliser ce jeux"
      Height          =   420
      Left            =   45
      TabIndex        =   1
      Top             =   3330
      Width           =   1935
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "&Fermer"
      Height          =   420
      Left            =   2025
      TabIndex        =   0
      Top             =   3330
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9630
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Base de données source"
      FileName        =   "*.mdb"
      Filter          =   "*.mdb"
   End
   Begin MSAdodcLib.Adodc dtaPeriode 
      Height          =   330
      Left            =   8505
      Top             =   3375
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   582
      ConnectMode     =   1
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   "SELECT * FROM P3IUser.P3IUser.Assure WHERE RECNO=-1"
      Caption         =   "dtaPeriode"
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
   Begin FPSpreadADO.fpSpread sprListe 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11040
      _Version        =   524288
      _ExtentX        =   19473
      _ExtentY        =   5741
      _StockProps     =   64
      BackColorStyle  =   1
      DAutoSizeCols   =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   1000000
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmListeLotSASP3I.frx":0000
      UserResize      =   1
      VirtualMode     =   -1  'True
      VisibleCols     =   10
      VisibleRows     =   100
      AppearanceStyle =   0
   End
End
Attribute VB_Name = "frmListeLotSASP3I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6818017C"

Option Explicit

'##ModelId=5C8A681802A5
Private frmNumeroLot As Long

' base de donnee
'##ModelId=5C8A681802C4
Private frmDatabaseFileName As String

'##ModelId=5C8A681802F3
Private frmDataSource As DataAccess
'##ModelId=5C8A681802F6
Private frmDataHelper As DataHelper
'

'##ModelId=5C8A68180303
Private Sub SetNumeroLot()
  frmNumeroLot = 0
  
  If sprListe.ActiveRow < 0 Then Exit Sub
  
  If sprListe.MaxRows = 0 Then Exit Sub
  
  sprListe.Row = sprListe.ActiveRow
  sprListe.Col = 1
  
  frmNumeroLot = CLng(sprListe.text)
End Sub

'##ModelId=5C8A68180322
Private Sub btnClose_Click()
  ret_code = -1
  Unload Me
End Sub

'##ModelId=5C8A68180332
Private Sub btnUtiliser_Click()

  Dim res As Long
  Dim sqlStr As String
  Dim rsCount As ADODB.Recordset
  
  ' import depuis le jeux de données sélectioné puis fermeture de la fenetre
  SetNumeroLot
  
  If frmNumeroLot = 0 Then Exit Sub
  
  If MsgBox("Le lot n°" & frmNumeroLot & " va être importé. Les données vont être écrasées si elles existent déjà dans P3I." & vbLf & "Voulez-vous continuer ?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
  
  On Error GoTo err_btnUtiliser
  
  Dim cLogger As clsLogger
  Set cLogger = New clsLogger
  
  cLogger.FichierLog = sReadIniFile("Dir", "LogPath", "##", 255, sFichierIni) & GetWinUser & "_CopyLot.log"
  cLogger.CreateLog "Import du lot " & frmNumeroLot
  
  cLogger.EcritTraceDansLog ""
  
  lblAvancement.text = "Suppression des données précédentes du lot " & frmNumeroLot & "..."
  
  'm_dataSource.Execute "DELETE FROM P3ITRACE WHERE NUTRAITP3I=" & frmNumeroLot
  'm_dataSource.Execute "DELETE FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot
  
Delete_P3ITRACE:
        sqlStr = "DELETE TOP (1000) FROM P3ITRACE WHERE NUTRAITP3I = " & frmNumeroLot
        m_dataSource.Execute sqlStr
        
        Set rsCount = m_dataSource.OpenRecordset("SELECT COUNT(*)FROM P3ITRACE WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
        res = rsCount(0).Value
        
        DoEvents
        'frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        'DoEvents
        
        If res > 0 Then GoTo Delete_P3ITRACE
        
Delete_P3IPROVCOLL:
        sqlStr = "DELETE TOP (1000) FROM P3IPROVCOLL WHERE NUTRAITP3I = " & frmNumeroLot
        m_dataSource.Execute sqlStr
        
        Set rsCount = m_dataSource.OpenRecordset("SELECT COUNT(*)FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
        res = rsCount(0).Value
        
        DoEvents
        'frmPeriode.lblStatus.Caption = statusMessage + " -- Assurées à supprimer : " & res
        'DoEvents
        
        If res > 0 Then GoTo Delete_P3IPROVCOLL
  
  
  m_dataSource.Execute "DELETE FROM P3ILOGTRAIT WHERE NUTRAITP3I=" & frmNumeroLot
  
  DoEvents
  
  'CopyLot "P3ILOGTRAIT", cLogger
  'CopyLot "P3IPROVCOLL", cLogger
  CopyLot_ADO "P3ILOGTRAIT", cLogger
  CopyLot_ADO "P3IPROVCOLL", cLogger
  
  m_dataSource.Execute "UPDATE P3IPROVCOLL SET Selected=0 WHERE NUTRAITP3I=" & frmNumeroLot
  
  lblAvancement.text = "Import terminé. Cliquer sur Fermer pour quitter"
  
  cLogger.EcritTraceDansLog "Import terminé. Cliquer sur Fermer pour quitter"
  
  cLogger.AfficheErreurLog False
  
  Exit Sub
  
err_btnUtiliser:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie une date sous forme de long en date
' format YYYYMMDD
'
'##ModelId=5C8A68180341
Private Function ConvertDate(lDate As Long) As Date
  Dim str As String
  
  '                    YYYYMMDD
  str = Format(lDate, "00000000")
  
  ConvertDate = DateSerial(Left(str, 4), mID(str, 5, 2), Right(str, 2))
 
End Function

'##ModelId=5C8A68180361
Private Function CheckDate(lDate As Long) As Boolean
  Dim str As String, d As Date
  
  On Error GoTo err_CheckDate
  
  '                    YYYYMMDD
  str = Format(lDate, "00000000")
  
  d = DateSerial(Left(str, 4), mID(str, 5, 2), Right(str, 2))
  
  CheckDate = True
  
  Exit Function
  
err_CheckDate:
  CheckDate = False
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Convertie une date sous forme de long en date
' format YYYYMMDD
'
'##ModelId=5C8A68180390
Private Function ConvertTime(lTime As Long) As Date
  Dim str As String
  
  '                    HHMMSS
  str = Format(lTime, "000000")
  
  ConvertTime = TimeSerial(Left(str, 2), mID(str, 3, 2), Right(str, 2))
 
End Function

'##ModelId=5C8A681803BF
Private Function CheckTime(lTime As Long) As Boolean
  Dim str As String, d As Date
  
  On Error GoTo err_CheckTime
  
  '                    HHMMSS
  str = Format(lTime, "000000")
  
  d = TimeSerial(Left(str, 2), mID(str, 3, 2), Right(str, 2))
  
  CheckTime = True
  
  Exit Function
  
err_CheckTime:
  CheckTime = False
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'Private Sub CopyLot(sTable As String, cLogger As clsLogger)
'  If frmNumeroLot = 0 Then Exit Sub
'
'  On Error GoTo err_CopyLot
'
'  Dim rsIn As ADODB.Recordset, rsOut As ADODB.Recordset, i As Integer, f As ADODB.Field, nb As Long, bOk As Boolean
'  Dim pos As Long, nbRecord As Long
'
'  lblAvancement.text = "Copie des données de la table " & sTable & "..."
'  DoEvents
'
'  cLogger.EcritTraceDansLog "Copie des données de la table " & sTable & "..."
'
''  Set rsIn = frmDataSource.OpenRecordset("SELECT count(*) FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
''  If Not rsIn.EOF Then
''    nbRecord = rsIn.fields(0)
''  End If
''  rsIn.Close
'
'  If sTable <> "P3IPROVCOLL" Then
'    Set rsIn = frmDataSource.OpenRecordset("SELECT * FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
'  Else
'    ' pour prevenir le crash du client Oracle 8i sur les champs CHAR vide, on utilise la fonction TRIM() qui renvoie NULL si le champ est vide
'    Set rsIn = frmDataSource.OpenRecordset("SELECT CDCOMPAGNIE, CDAPPLI, TRIM(CDPRODUIT) as CDPRODUIT, TRIM(NUCONTRA) as NUCONTRA, " _
'                                           & " NUPRET, TRIM(NUDOSSIERPREST) as NUDOSSIERPREST, TRIM(NUSOUSDOSSIERPREST) as NUSOUSDOSSIERPREST, " _
'                                           & " TRIM(NUCONTRATGESTDELEG)as NUCONTRATGESTDELEG, CDGARAN, NUREFMVT, CDTYPMVT, NUTRAITE, " _
'                                           & " CDCATADHESION, TRIM(LBCATADHESION) as LBCATADHESION, TRIM(CDOPTION) as CDOPTION, " _
'                                           & " TRIM(LBOPTIONCON) as LBOPTIONCON, TRIM(LBSOUSCR) as LBSOUSCR, TRIM(IDASSUREAGI) as IDASSUREAGI, " _
'                                           & " IDASSURE, TRIM(LBASSURE) as LBASSURE, DTNAISSASS, TRIM(CDSEXASSURE) as CDSEXASSURE, " _
'                                           & " TRIM(IDRENTIERAGI) as IDRENTIERAGI, TRIM(IDRENTIER) as IDRENTIER, TRIM(LBRENTIER) as LBRENTIER, " _
'                                           & " DTNAISSREN, TRIM(IDCORENTIERAGI) as IDCORENTIERAGI, IDCORENTIER, TRIM(LBCORENTIER) as LBCORENTIER, " _
'                                           & " DTNAISSCOR, DTSURVSIN, AGESURVSIN, DURPOSARR, ANCARRTRA, TRIM(CDPERIODICITE) as CDPERIODICITE, " _
'                                           & " CDTYPTERME, DTEFFREN, DTLIMPRO, DTDERREG, DTDEBPER, DTFINPER, TRIM(CDPROEVA) as CDPROEVA, ANNEADHE, " _
'                                           & " TRIM(CDSINCON) as CDSINCON, TRIM(CDMISINV) as CDMISINV, DTMISINV, NUETABLI, TXREVERSION, DTCALCULPROV, " _
'                                           & " DTTRAITPROV, DTCREATI, DTSIGBIA, DTDECSIN, TRIM(CDRISQUE) as CDRISQUE, TRIM(LBRISQUE) as LBRISQUE, " _
'                                           & " CDSITUATSIN, DTSITUATSIN, TRIM(CDPRETRATTSIN) as CDPRETRATTSIN, TRIM(CDCTGPRT) as CDCTGPRT, " _
'                                           & " TRIM(LBCTGPRT) as LBCTGPRT, DTPREECH, DTDERECH, TRIM(CDPERIODICITEECH) as CDPERIODICITEECH, " _
'                                           & " MTECHEANCE1, DTDEBPERECH1, DTFINPERECH1, MTECHEANCE2, DTDEBPERECH2, DTFINPERECH2, MTECHEANCE3, " _
'                                           & " DTDEBPERECH3, DTFINPERECH3, TRIM(CDTYPAMO) as CDTYPAMO, TRIM(LBTYPAMO) as LBTYPAMO, TXINVPEC, " _
'                                           & " DTDEBPERPIP, DTFINPERPIP, DTSAISIEPERJUSTIF, DTDEBPERJUSTIF, DTFINPERJUSTIF, DTDEBDERPERRGLTADA, " _
'                                           & " DTFINDERPERRGLTADA, DTDERPERRGLTADA, MTDERPERRGLTADA, DTDEBDERPERRGLTADC, DTFINDERPERRGLTADC, " _
'                                           & " DTDERPERRGLTADC, MTDERPERRGLTADC, TRIM(CDSINPREPROV) as CDSINPREPROV, MTTOTREGLEICIV, DTDEBPROV, " _
'                                           & " DTFINPROV, TRIM(INDBASREV) as INDBASREV, MTPREANN ,  MTPREREV, MTPREMAJ, MTPRIREG, MTPRIRE1, " _
'                                           & " MTPRIRE2, TRIM(CDMONNAIE) as CDMONNAIE, TRIM(CDPAYS) as CDPAYS, CDAPPLISOURCE, TRIM(LBCOMLIG) as LBCOMLIG, " _
'                                           & " NUTRAITP3I, NUENRP3I FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
'  End If
'
'  Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot, Dynamic)
'
'  ProgressBar1.Visible = True
'  ProgressBar1.Min = 0
'  ProgressBar1.Value = 0
'  ProgressBar1.Max = frmDataHelper.GetParameterAsDouble("SELECT count(*) FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot)
'
'  pos = 0
'  Do Until rsIn.EOF
'    bOk = True
'
'    If ProgressBar1.Max < rsIn.RecordCount Then
'      ProgressBar1.Max = rsIn.RecordCount
'    End If
'
'    pos = pos + 1
'    ProgressBar1.Value = pos
'
'    rsOut.AddNew
'
'    For i = 0 To rsIn.fields.Count - 1
'
'      Set f = rsIn.fields(i)
'
'      rsOut.fields(f.Name).Value = f.Value
'
'    Next
'
'    If sTable = "P3IPROVCOLL" Then
'      rsOut.fields("DataVersion") = eInitiale
'    End If
'
'    If bOk Then
'      rsOut.Update
'    Else
'      rsOut.CancelUpdate
'    End If
'
'    rsIn.MoveNext
'  Loop
'
'  rsIn.Close
'  rsOut.Close
'
'  cLogger.EcritTraceDansLog ""
'
'  lblAvancement.text = vbNullString
'
'  ProgressBar1.Value = 0
'  ProgressBar1.Visible = False
'
'  Exit Sub
'
'err_CopyLot:
'  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
'  Resume Next
'End Sub



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie un lot de données Oracle dans nos tables SQL Server
'
'##ModelId=5C8A68190015
Private Sub CopyLot_ADO(sTable As String, cLogger As clsLogger)
  If frmNumeroLot = 0 Then Exit Sub
  
  Screen.MousePointer = vbHourglass
  
  On Error GoTo err_CopyLot
  
  Dim rsIn As ADODB.Recordset, i As Integer, nb As Long, bOk As Boolean
  Dim pos As Long, nbRecord As Long
  Dim rq As String
  
  lblAvancement.text = "Copie des données de la table " & sTable & "..."
  DoEvents
  
  cLogger.EcritTraceDansLog "Copie des données de la table " & sTable & "..."
  
'  Set rsIn = frmDataSource.OpenRecordset("SELECT count(*) FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
'  If Not rsIn.EOF Then
'    nbRecord = rsIn.fields(0)
'  End If
'  rsIn.Close
  
  If sTable <> "P3IPROVCOLL" Then
    Set rsIn = frmDataSource.OpenRecordset("SELECT * FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot, Snapshot)
  Else
    ' pour prevenir le crash du client Oracle 8i sur les champs CHAR vide, on utilise la fonction TRIM() qui renvoie NULL si le champ est vide
    rq = "SELECT CDCOMPAGNIE, CDAPPLI, TRIM(CDPRODUIT) as CDPRODUIT, TRIM(NUCONTRA) as NUCONTRA, " _
          & " NUPRET, TRIM(NUDOSSIERPREST) as NUDOSSIERPREST, TRIM(NUSOUSDOSSIERPREST) as NUSOUSDOSSIERPREST, " _
          & " TRIM(NUCONTRATGESTDELEG)as NUCONTRATGESTDELEG, CDGARAN, NUREFMVT, CDTYPMVT, NUTRAITE, " _
          & " CDCATADHESION, TRIM(LBCATADHESION) as LBCATADHESION, TRIM(CDOPTION) as CDOPTION, " _
          & " TRIM(LBOPTIONCON) as LBOPTIONCON, TRIM(LBSOUSCR) as LBSOUSCR, TRIM(IDASSUREAGI) as IDASSUREAGI, " _
          & " IDASSURE, TRIM(LBASSURE) as LBASSURE, DTNAISSASS, TRIM(CDSEXASSURE) as CDSEXASSURE, " _
          & " TRIM(IDRENTIERAGI) as IDRENTIERAGI, TRIM(IDRENTIER) as IDRENTIER, TRIM(LBRENTIER) as LBRENTIER, " _
          & " DTNAISSREN, TRIM(IDCORENTIERAGI) as IDCORENTIERAGI, IDCORENTIER, TRIM(LBCORENTIER) as LBCORENTIER, " _
          & " DTNAISSCOR, DTSURVSIN, AGESURVSIN, DURPOSARR, ANCARRTRA, TRIM(CDPERIODICITE) as CDPERIODICITE, " _
          & " CDTYPTERME, DTEFFREN, DTLIMPRO, DTDERREG, DTDEBPER, DTFINPER, TRIM(CDPROEVA) as CDPROEVA, ANNEADHE, " _
          & " TRIM(CDSINCON) as CDSINCON, TRIM(CDMISINV) as CDMISINV, DTMISINV, NUETABLI, TXREVERSION, DTCALCULPROV, " _
          & " DTTRAITPROV, DTCREATI, DTSIGBIA, DTDECSIN, TRIM(CDRISQUE) as CDRISQUE, TRIM(LBRISQUE) as LBRISQUE, " _
          & " CDSITUATSIN, DTSITUATSIN, TRIM(CDPRETRATTSIN) as CDPRETRATTSIN, TRIM(CDCTGPRT) as CDCTGPRT, " _
          & " TRIM(LBCTGPRT) as LBCTGPRT, DTPREECH, DTDERECH, TRIM(CDPERIODICITEECH) as CDPERIODICITEECH, " _
          & " MTECHEANCE1, DTDEBPERECH1, DTFINPERECH1, MTECHEANCE2, DTDEBPERECH2, DTFINPERECH2, MTECHEANCE3, " _
          & " DTDEBPERECH3, DTFINPERECH3, TRIM(CDTYPAMO) as CDTYPAMO, TRIM(LBTYPAMO) as LBTYPAMO, TXINVPEC, " _
          & " DTDEBPERPIP, DTFINPERPIP, DTSAISIEPERJUSTIF, DTDEBPERJUSTIF, DTFINPERJUSTIF, DTDEBDERPERRGLTADA, " _
          & " DTFINDERPERRGLTADA, DTDERPERRGLTADA, MTDERPERRGLTADA, DTDEBDERPERRGLTADC, DTFINDERPERRGLTADC, " _
          & " DTDERPERRGLTADC, MTDERPERRGLTADC, TRIM(CDSINPREPROV) as CDSINPREPROV, MTTOTREGLEICIV, DTDEBPROV, " _
          & " DTFINPROV, TRIM(INDBASREV) as INDBASREV, MTPREANN ,  MTPREREV, MTPREMAJ, MTPRIREG, MTPRIRE1, " _
          & " MTPRIRE2, TRIM(CDMONNAIE) as CDMONNAIE, TRIM(CDPAYS) as CDPAYS, CDAPPLISOURCE, "
    
    ' Evol 2010 - Lot 2
    rq = rq & " TRIM(CDCATINV) as CDCATINV, TRIM(LBCATINV) as LBCATINV,  TRIM(CDCONTENTIEUX) as CDCONTENTIEUX, TRIM(NUSINISTRE) as NUSINISTRE, " _
          & " TRIM(CDCHOIXPREST) as CDCHOIXPREST, TRIM(LBCHOIXPREST) as LBCHOIXPREST, MTCAPSSRISQ, TRIM(FLAMORTISSABLE) as FLAMORTISSABLE, "
    
    rq = rq & " TRIM(LBCOMLIG) as LBCOMLIG, NUTRAITP3I, NUENRP3I FROM P3IPROVCOLL WHERE NUTRAITP3I=" & frmNumeroLot
      
    Set rsIn = frmDataSource.OpenRecordset(rq, Snapshot)
  End If
  
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Value = 0
  ProgressBar1.Max = frmDataHelper.GetParameterAsDouble("SELECT count(*) FROM " & sTable & " WHERE NUTRAITP3I=" & frmNumeroLot)
  
  
  ' creation de la requete d'insert
  Dim cmdText As String, cmdValues As String, f As ADODB.field
  
  cmdText = "INSERT INTO " & sTable & "("
  cmdValues = " VALUES("
  
  For Each f In rsIn.fields
    cmdText = cmdText & f.Name & ", "
    cmdValues = cmdValues & "?, "
  Next
  
  If sTable <> "P3IPROVCOLL" Then
    cmdText = Left(cmdText, Len(cmdText) - 2) & ")"
    cmdValues = Left(cmdValues, Len(cmdValues) - 2) & ")"
  Else
    cmdText = cmdText & "DataVersion)"
    cmdValues = cmdValues & eInitiale & ")"
  End If
  
  cmdText = cmdText & cmdValues
  
  ' ajout des enregistrements du lot
  pos = 0
  nbRecord = 0
  Do Until rsIn.EOF
    Dim cmd As ADODB.Command
    
    If ProgressBar1.Max < rsIn.RecordCount Then
      ProgressBar1.Max = rsIn.RecordCount
    End If

    pos = pos + 1
    If pos Mod 9 = 0 Then
      ProgressBar1.Value = pos
      DoEvents
    End If
    
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = m_dataSource.Connection
    cmd.CommandType = adCmdText
    cmd.CommandText = cmdText
    
    For Each f In rsIn.fields
      Dim prm As ADODB.Parameter, l As Long, t As ADODB.DataTypeEnum
      
      Select Case f.Type
        Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar
          If IsNull(f.Value) Then
            l = 1
          Else
            l = Len(f.Value)
            If l = 0 Then
              l = 1
            End If
          End If
          t = adVarChar
        
        Case adNumeric, adDecimal, adDouble
          l = 0
          t = adDouble
       
        Case Else
          l = 0
          t = f.Type
      End Select
      
      Set prm = cmd.CreateParameter("@" & f.Name, t, adParamInput, l, f.Value)
      
      cmd.Parameters.Append prm
    Next
    
    Err = 0
    cmd.Execute
    If Err = 0 Then
      nbRecord = nbRecord + 1
    End If
    
    Set cmd = Nothing
    
    rsIn.MoveNext
  Loop
  
  rsIn.Close
  
  cLogger.EcritTraceDansLog nbRecord & " copié(s) dans la table " & sTable
  cLogger.EcritTraceDansLog ""
  
  lblAvancement.text = vbNullString
  
  ProgressBar1.Value = 0
  ProgressBar1.Visible = False
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
err_CopyLot:
  Screen.MousePointer = vbDefault
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Screen.MousePointer = vbHourglass
  
  Resume Next
End Sub


'##ModelId=5C8A68190044
Private Sub Form_Load()
  ' chargement du masque du spread
  sprListe.LoadFromFile App.Path & "\ListeJeuxDonnees.ss7"
  
  frmDatabaseFileName = GetSettingIni(SectionName, "DB", "SASP3IConnectionString", "#")
  If frmDatabaseFileName = "#" Then
    MsgBox "La chaine de connexion à SASP3I n'est pas spécifier !" & vbLf & "Paramètre SASP3IConnectionString dans [DB]", vbCritical
    Unload Me
  End If
  
  Set frmDataSource = New DataAccess
  If frmDataSource.Connect(frmDatabaseFileName) = False Then
    MsgBox "Impossible d'ouvrir SASP3I !" & vbLf & "Connection : " & frmDatabaseFileName, vbCritical
    Unload Me
    Exit Sub
  End If
  
  'Set theDB = m_dataSource.Connection
  Set frmDataHelper = frmDataSource.CreateHelper
    
  frmDataSource.SetDatabase dtaPeriode
  
  RefreshListe
End Sub

'##ModelId=5C8A68190054
Private Sub RefreshListe()
  
  Dim rq As String, rs As ADODB.Recordset
  Dim filter As String
  Dim i As Integer
  
  On Error GoTo err_RefreshListe
  
  Screen.MousePointer = vbHourglass
  
  ' fabrique le titre de la fenetre
  Me.Caption = "Jeux de données disponibles dans SASP3I"
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  DoEvents
  
  sprListe.Visible = False
  sprListe.ReDraw = False
  
  ' Virtual mode pour la rapidité
  sprListe.VirtualMode = False
  'sprListe.VirtualMode = True
  'sprListe.VirtualMaxRows = -1
  sprListe.MaxRows = 0
  'sprListe.VScrollSpecial = True
  'sprListe.VScrollSpecialType = 0
  
  ' fabrique la requete de remplissage du spread : periode pour le groupe en cours
  If frmDataHelper.SqlMode = SQLServer Then
    rq = "SELECT NUTRAITP3I, DTTRAIT as [Date], CONVERT(CHAR(8), HHTRAIT, 8) as [Heure], NUTRAITP3I as [Identifiant du lot], RTRIM(NUTRAIT) as [N° traitement AGI], " _
         & " DTDEBPER as [Debut], DTFINPER as [Fin], NBLIGTRAIT as [nb lignes], MTTRAIT as [Montant], NUTRAIT, IDTABLESAS " _
         & " FROM P3ILOGTRAIT ORDER BY NUTRAITP3I DESC"
  ElseIf frmDataHelper.SqlMode = Oracle Then
'    rq = "SELECT P.NUTRAITP3I, TO_DATE(TO_CHAR(P.DTTRAIT), 'YYYYMMDD') AS ""Date"", P.HHTRAIT AS ""Heure"", " _
'        & " P.NUTRAITP3I AS ""Identifiant du lot"", TO_DATE(TO_CHAR(P.DTDEBPER), 'YYYYMMDD') AS ""Début""," _
'        & " TO_DATE(TO_CHAR(P.DTFINPER), 'YYYYMMDD') AS ""Fin"", P.NUTRAIT AS ""N° traitement AGI"",  P.NBLIGTRAIT AS ""Nb lignes"", " _
'        & " P.MTTRAIT AS ""Montant"" FROM P3ILOGTRAIT P ORDER BY P.NUTRAITP3I DESC "
    rq = "SELECT P.NUTRAITP3I, " _
          & "CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTTRAIT),7,2),'/'),CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTTRAIT),5,2),'/'),SUBSTR(TO_CHAR(P.DTTRAIT),1,4))) AS ""Date""," _
          & "CONCAT(CONCAT(SUBSTR(TO_CHAR(P.HHTRAIT),1,2),':'),CONCAT(CONCAT(SUBSTR(TO_CHAR(P.HHTRAIT),3,2),':'),SUBSTR(TO_CHAR(P.HHTRAIT),5,2))) AS ""Heure""," _
          & "P.NUTRAITP3I AS ""Identifiant du lot""," _
          & "CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTDEBPER),7,2),'/'),CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTDEBPER),5,2),'/'),SUBSTR(TO_CHAR(P.DTDEBPER),1,4))) AS ""Début""," _
          & "CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTFINPER),7,2),'/'),CONCAT(CONCAT(SUBSTR(TO_CHAR(P.DTFINPER),5,2),'/'),SUBSTR(TO_CHAR(P.DTFINPER),1,4))) AS ""Fin""," _
          & "TRIM(P.NUTRAIT) AS ""N° traitement AGI"", " _
          & "P.NBLIGTRAIT AS ""Nb lignes"", P.MTTRAIT AS ""Montant"", NUTRAIT, IDTABLESAS " _
          & "FROM P3ILOGTRAIT P ORDER BY P.NUTRAITP3I DESC"
  Else
    MsgBox "SQLMode inconnu pour le SASP3I !", vbCritical
    Exit Sub
  End If
          
  dtaPeriode.RecordSource = frmDataHelper.ValidateSQL(rq)
  dtaPeriode.Refresh
    
  Set sprListe.DataSource = dtaPeriode
      
  ' mets à jours les n° de ligne dans le spread
  If dtaPeriode.Recordset.EOF = False Then
    dtaPeriode.Recordset.MoveLast
    dtaPeriode.Recordset.MoveFirst
  
    sprListe.MaxRows = dtaPeriode.Recordset.RecordCount
    'sprListe.VirtualMaxRows = dtaPeriode.Recordset.RecordCount
  
    dtaPeriode.Recordset.MoveFirst
  Else
    sprListe.MaxRows = 0
    'sprListe.VirtualMaxRows = 0
    sprListe.ColWidth(1) = 0
    sprListe.Visible = True
    sprListe.ReDraw = True

    Screen.MousePointer = vbDefault
    
    GoTo pas_de_donnee
  End If
  
  ' cache la colonne RECNO
  sprListe.ColWidth(1) = 0
     
  For i = 2 To sprListe.MaxCols
    sprListe.ColWidth(i) = sprListe.MaxTextColWidth(i) + 2
  Next i
 
  sprListe.BlockMode = True
  
  sprListe.Row = -1
  sprListe.Row = -1
  
  sprListe.Col = 1
  sprListe.Col2 = sprListe.MaxCols - 1
  sprListe.TypeHAlign = TypeHAlignCenter
  
  sprListe.Col = sprListe.MaxCols
  sprListe.Col2 = sprListe.MaxCols
  sprListe.TypeHAlign = TypeHAlignLeft
  
  sprListe.BlockMode = False
  
pas_de_donnee:

  On Error GoTo 0
  
  ' affiche le spread (vitesse)
  sprListe.Visible = True
  sprListe.ReDraw = True
  
  lblAvancement.text = vbNullString
  
  Screen.MousePointer = vbDefault

  Exit Sub

err_RefreshListe:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A68190063
Private Sub Form_Resize()
  Dim topbtn As Integer
  
  If Me.WindowState = vbMinimized Then Exit Sub
  
  ' place la liste
  sprListe.top = 30
  sprListe.Left = 30
  sprListe.Width = Me.Width - 130
 
  topbtn = Me.ScaleHeight - btnHeight
  
  sprListe.Height = Maximum(topbtn - 100, 0)
  
  PlacePremierBoutton btnUtiliser, topbtn
  
  PlaceBoutton btnClose, btnUtiliser, topbtn
  
  lblAvancement.top = topbtn + 100
  lblAvancement.Left = btnClose.Left + btnClose.Width + 50
  
  ProgressBar1.top = btnClose.top
  ProgressBar1.Left = lblAvancement.Left + lblAvancement.Width + 50
End Sub

'##ModelId=5C8A68190082
Private Sub Form_Unload(Cancel As Integer)
  frmDataSource.Disconnect
End Sub

'##ModelId=5C8A681900B1
Private Sub sprListe_DblClick(ByVal Col As Long, ByVal Row As Long)
  ' NE PAS ENLEVER : evite l'entree en mode edition dans une cellule
End Sub

'##ModelId=5C8A681900E0
Private Sub sprListe_DataColConfig(ByVal Col As Long, ByVal DataField As String, ByVal DataType As Integer)
  If dtaPeriode.Recordset.fields(Col - 1).Properties("BASECOLUMNNAME").Value = "Commentaire" Then
    sprListe.Col = Col
    sprListe.Row = -1
    sprListe.CellType = CellTypeEdit
    sprListe.TypeMaxEditLen = 255
  End If
End Sub

