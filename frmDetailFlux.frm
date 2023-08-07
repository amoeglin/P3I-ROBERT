VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "_fpSPR80.OCX"
Begin VB.Form frmDetailFlux 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Details..."
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8025
   Icon            =   "frmDetailFlux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Ajouter"
      Height          =   375
      Left            =   6750
      TabIndex        =   6
      Top             =   2655
      Width           =   1215
   End
   Begin VB.CommandButton btnUndelete 
      Caption         =   "&Récuperer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6750
      TabIndex        =   5
      Top             =   1755
      Width           =   1215
   End
   Begin VB.CommandButton btnDoublon 
      Caption         =   "&Doublon"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6750
      TabIndex        =   4
      Top             =   1350
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Supprimer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6750
      TabIndex        =   3
      Top             =   945
      Width           =   1215
   End
   Begin FPSpreadADO.fpSpread sprEdit 
      Height          =   5550
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   6585
      _Version        =   524288
      _ExtentX        =   11615
      _ExtentY        =   9790
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      RowHeaderDisplay=   2
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "frmDetailFlux.frx":1BB2
      AppearanceStyle =   0
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   6750
      TabIndex        =   1
      Top             =   5265
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   6750
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDetailFlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6817013E"
Option Explicit

'##ModelId=5C8A68170238
Public NumeroLot As Long
'##ModelId=5C8A68170257
Public numeroEnregistrement As Long

'##ModelId=5C8A68170267
Public frmAutomate As Boolean ' true : action automatique
'##ModelId=5C8A68170286
Public frmAction As Long ' 1=Modifier, 2=Supprimer, 3=Doublon, 4=Undelete, 5=Ajouter
'


'##ModelId=5C8A68170296
Private Sub btnAdd_Click()
  frmAction = 5
  OKButton.Enabled = True
  numeroEnregistrement = m_dataHelper.GetParameterAsDouble("SELECT MAX(NUENRP3I) FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot) + 1
  
  m_dataSource.Execute "INSERT INTO P3IPROVCOLL(NUTRAITP3I, NUENRP3I, DataVersion) VALUES (" & NumeroLot & ", " & numeroEnregistrement & ", 1)"
  
  EcritTrace Nothing, eAjouter
  
  DisplayRecord
End Sub

'##ModelId=5C8A681702A5
Private Sub btnDelete_Click()
  
  If frmAutomate = False Then
    If MsgBox("Voulez-vous supprimer cet enregistrement ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  frmAction = 2
  btnUndelete.Enabled = True

End Sub


'##ModelId=5C8A681702B5
Private Sub btnDoublon_Click()
  
  If frmAutomate = False Then
    If MsgBox("Voulez-vous marquer cet enregistrement comme Doublon ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  frmAction = 3
  btnUndelete.Enabled = True

End Sub


'##ModelId=5C8A681702D4
Private Sub btnUndelete_Click()
  
  If frmAutomate = False Then
    If MsgBox("Voulez-vous récuperer cet enregistrement ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  frmAction = 4
  OKButton.Enabled = True

End Sub

'##ModelId=5C8A681702E4
Private Sub CancelButton_Click()
  ret_code = -1
  Unload Me
End Sub


'##ModelId=5C8A681702F3
Private Sub Form_Activate()
  Screen.MousePointer = vbHourglass
  
  Me.Visible = Not frmAutomate
    
  DisplayRecord
  
  If frmAutomate = True Then
        
    Select Case frmAction
      Case 2
        btnDelete_Click
        OKButton_Click
        
      Case 3
        btnDoublon_Click
        OKButton_Click
        
      Case 4
        btnUndelete_Click
        OKButton_Click
    End Select
  End If

  Screen.MousePointer = vbDefault
End Sub


'##ModelId=5C8A68170322
Private Sub DisplayRecord()
  Dim i As Integer, depart As Integer
  Dim rs As ADODB.Recordset
  
  Me.Caption = "Lot " & NumeroLot & " - Ligne " & numeroEnregistrement
  
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Snapshot)
  
  If Not IsNull(rs.fields("DataVersion").Value) Then
    Select Case rs.fields("DataVersion").Value
      Case eInitiale
        Me.Caption = Me.Caption & " (Originale)"
        btnUndelete.Enabled = False
        btnDelete.Enabled = True
        btnDoublon.Enabled = True
        OKButton.Enabled = True
  
      Case eModifie
        Me.Caption = Me.Caption & " (Modifiée)"
        btnUndelete.Enabled = False
        btnDelete.Enabled = True
        btnDoublon.Enabled = True
        OKButton.Enabled = True
  
      Case eAjouter
        Me.Caption = Me.Caption & " (Ajoutée)"
        btnUndelete.Enabled = False
        btnDelete.Enabled = True
        btnDoublon.Enabled = True
        OKButton.Enabled = True
  
      Case eSupprimer
        Me.Caption = Me.Caption & " (Supprimée)"
        OKButton.Enabled = False
        btnUndelete.Enabled = True
        btnDelete.Enabled = False
        btnDoublon.Enabled = False
  
      Case eDoublon
        Me.Caption = Me.Caption & " (Doublon)"
        OKButton.Enabled = False
        btnUndelete.Enabled = True
        btnDelete.Enabled = False
        btnDoublon.Enabled = False
        
      Case Else
        Me.Caption = Me.Caption & " (Etat inconnu)"
    End Select
  End If
  
  sprEdit.MaxRows = rs.fields.Count - 3
  sprEdit.DAutoSizeCols = DAutoSizeColsBest
    
  sprEdit.ScrollBars = ScrollBarsBoth
  sprEdit.ScrollBarExtMode = True
  
  depart = 3 ' on ignore les cles
  For i = depart To rs.fields.Count - 1
    ' libellé
    sprEdit.Col = SpreadHeader
    sprEdit.Row = i + 1 - depart
    sprEdit.text = Replace(rs.fields(i).Name, vbLf, " ")
    
    If UCase(rs.fields(i).Name) = "SELECTED" Then
      sprEdit.RowHidden = True
    End If
    
    'If rs.fields(i).Name = "NUREFMVT" Then Stop
    
'    If rs.fields(i).Properties("BASETABLENAME").Value = "Utilisateur" _
'       And rs.fields(i).Properties("BASECOLUMNNAME").Value = "Nom" Then
'      sprEdit.Col = 1
'      sprEdit.Row = i + 1
'      sprEdit.Lock = True
'    End If
'
'    If rs.fields(i).Properties("BASECOLUMNNAME").Value = "PrestationAnnulee" Then
'      sprEdit.Col = 1
'      sprEdit.Row = i + 1
'      sprEdit.Lock = True
'    End If
'
'    If rs.fields(i).Properties("BASETABLENAME").Value = "DetailRegroupement" _
'       And rs.fields(i).Properties("BASECOLUMNNAME").Value = "TypeGarantie" Then
'      ' type de garantie
'      sprEdit.Col = 1
'      sprEdit.Row = i + 1
'      sprEdit.CellType = CellTypeComboBox
'
'      sprEdit.TypeComboBoxClear sprEdit.Col, sprEdit.Row
'      sprEdit.TypeComboBoxEditable = False
'
'      Dim rsT As ADODB.Recordset
'      Set rsT = m_dataSource.OpenRecordset("SELECT CleTypeGarantie, Libelle FROM CRUser.TypeGarantie ORDER BY Libelle", Snapshot)
'      Do Until rsT.EOF
'        sprEdit.TypeComboBoxIndex = -1
'        sprEdit.TypeComboBoxString = rsT.fields("CleTypeGarantie") & " - " & rsT.fields("Libelle")
'        If Not rs.EOF Then
'          If rsT.fields("CleTypeGarantie") = rs.fields("Type de Garantie") Then
'            sprEdit.TypeComboBoxCurSel = sprEdit.TypeComboBoxCount - 1
'          End If
'        End If
'        rsT.MoveNext
'      Loop
'      rsT.Close
'    Else
      ' valeur
      sprEdit.Col = 1
      sprEdit.Row = i + 1 - depart
      Select Case rs.fields(i).Type
        Case adDate, adDBDate, adDBTimeStamp
'          sprEdit.CellType = CellTypeEdit
'          sprEdit.TypeEditMultiLine = False
          sprEdit.CellType = CellTypeDate
          sprEdit.TypeDateCentury = True
          sprEdit.TypeDateFormat = TypeDateFormatDDMMYY
          sprEdit.TypeDateMin = "01011900"
          sprEdit.TypeDateMax = "01012200"
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i).Value) Then
              sprEdit.text = rs.fields(i).Value
            Else
              sprEdit.text = vbNullString
            End If
          End If
        
        Case adSmallInt, adTinyInt, adBigInt, adInteger
          sprEdit.CellType = CellTypeNumber
          sprEdit.TypeNumberDecPlaces = 0
          sprEdit.TypeNumberShowSep = False
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i)) Then
              sprEdit.text = rs.fields(i)
            Else
              sprEdit.text = ""
            End If
          End If
          
        Case adDouble
          sprEdit.CellType = CellTypeNumber
          sprEdit.TypeNumberDecPlaces = 2
          sprEdit.TypeNumberShowSep = True
          sprEdit.TypeNumberMin = -1E+16
          sprEdit.TypeNumberMax = 1E+16
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i)) Then
              sprEdit.text = rs.fields(i)
            Else
              sprEdit.text = ""
            End If
          End If
        
        Case adNumeric
          sprEdit.CellType = CellTypeNumber
          sprEdit.TypeNumberDecPlaces = rs.fields(i).NumericScale
          sprEdit.TypeNumberShowSep = False
          sprEdit.TypeNumberMin = -1E+16
          sprEdit.TypeNumberMax = 1E+16
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i)) Then
              sprEdit.text = CDbl(rs.fields(i).Value)
            Else
              sprEdit.text = ""
            End If
          End If
        
        Case adChar, adVarChar, adVarWChar
          sprEdit.CellType = CellTypeEdit
          sprEdit.TypeEditMultiLine = IIf(rs.fields(i).Name = "Commentaire", True, False)
          sprEdit.TypeMaxEditLen = rs.fields(i).DefinedSize
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i)) Then
              sprEdit.text = Trim(rs.fields(i))
            Else
              sprEdit.text = ""
            End If
          End If
          If rs.fields(i).Name = "Commentaire" Then
            sprEdit.RowHeight(sprEdit.Row) = 56
          End If
          
        Case adBoolean
          sprEdit.CellType = CellTypeCheckBox
          sprEdit.TypeCheckType = TypeCheckTypeNormal
          sprEdit.TypeCheckCenter = True
          If Not IsNull(rs.fields(i)) Then
            sprEdit.Value = rs.fields(i)
          End If
        
          
        Case Else
          MsgBox "Type Inconnu pour le champ '" & Replace(rs.fields(i).Name, vbLf, " ") & "' Type=" & rs.fields(i).Type, vbInformation
          sprEdit.CellType = CellTypeEdit
          sprEdit.TypeEditMultiLine = False
          If Not rs.EOF Then
            If Not IsNull(rs.fields(i)) Then
              sprEdit.text = rs.fields(i)
            Else
              sprEdit.text = ""
            End If
          End If
      End Select
      sprEdit.TypeHAlign = TypeHAlignLeft
'    End If
        
  Next
  
  rs.Close
End Sub

'##ModelId=5C8A68170332
Private Sub Form_Load()
  frmAutomate = False
  frmAction = 0
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copie une ligne en lui donnant le statut modifiee
'
'##ModelId=5C8A68170351
Private Sub CopyLigne(etat As Integer)
  If numeroEnregistrement = 0 Or NumeroLot = 0 Then Exit Sub
  
  On Error GoTo err_CopyLigne
    
  Dim sSql As String
  
  sSql = "INSERT INTO P3IPROVCOLL (  NUTRAITP3I, NUENRP3I  , DataVersion  ,  CDCOMPAGNIE  ,  CDAPPLI  ,  CDPRODUIT  ," _
        & "  NUCONTRA  , NUPRET  , NUDOSSIERPREST  , NUSOUSDOSSIERPREST  , NUCONTRATGESTDELEG  , CDGARAN  ," _
        & "  NUREFMVT  , CDTYPMVT  , NUTRAITE  , CDCATADHESION  ,  LBCATADHESION  ,  CDOPTION  , LBOPTIONCON  ," _
        & "  LBSOUSCR  , IDASSUREAGI  ,  IDASSURE  , LBASSURE  , DTNAISSASS  , CDSEXASSURE  ,  IDRENTIERAGI  ," _
        & "  IDRENTIER  ,  LBRENTIER  ,  DTNAISSREN  , IDCORENTIERAGI  , IDCORENTIER  ,  LBCORENTIER  ,  DTNAISSCOR  ," _
        & "  DTSURVSIN  ,  AGESURVSIN  , DURPOSARR  ,  ANCARRTRA  ,  CDPERIODICITE  ,  CDTYPTERME  , DTEFFREN  ," _
        & "  DTLIMPRO  , DTDERREG  , DTDEBPER  , DTFINPER  , CDPROEVA  , ANNEADHE  , CDSINCON  , CDMISINV  ," _
        & "  DTMISINV  , NUETABLI  , TXREVERSION  ,  DTCALCULPROV  , DTTRAITPROV  ,  DTCREATI  , DTSIGBIA  ," _
        & "  DTDECSIN  , CDRISQUE  , LBRISQUE  , CDSITUATSIN  ,  DTSITUATSIN  ,  CDPRETRATTSIN  ,  CDCTGPRT  ," _
        & "  LBCTGPRT  , DTPREECH  , DTDERECH  , CDPERIODICITEECH  , MTECHEANCE1  ,  DTDEBPERECH1  , DTFINPERECH1  ," _
        & "  MTECHEANCE2  ,  DTDEBPERECH2  , DTFINPERECH2  , MTECHEANCE3  ,  DTDEBPERECH3  , DTFINPERECH3  ," _
        & "  CDTYPAMO  , LBTYPAMO  , TXINVPEC  , DTDEBPERPIP  ,  DTFINPERPIP  ,  DTSAISIEPERJUSTIF  ,  DTDEBPERJUSTIF  ," _
        & "  DTFINPERJUSTIF  , DTDEBDERPERRGLTADA  , DTFINDERPERRGLTADA  , DTDERPERRGLTADA  ,  MTDERPERRGLTADA  ," _
        & "  DTDEBDERPERRGLTADC  , DTFINDERPERRGLTADC  , DTDERPERRGLTADC  ,  MTDERPERRGLTADC  ,  CDSINPREPROV  ," _
        & "  MTTOTREGLEICIV  , DTDEBPROV  ,  DTFINPROV  ,  INDBASREV  ,  MTPREANN  , MTPREREV  , MTPREMAJ  , " _
  
  sSql = sSql & " CDCATINV, LBCATINV,  CDCONTENTIEUX, NUSINISTRE, CDCHOIXPREST, LBCHOIXPREST, MTCAPSSRISQ, " _
        & "  MTPRIREG  , MTPRIRE1  , MTPRIRE2  , CDMONNAIE  ,  CDPAYS  , CDAPPLISOURCE, FLAMORTISSABLE,  LBCOMLIG  )"
  
  sSql = sSql & " SELECT  TOP 1 NUTRAITP3I , NUENRP3I  , " & etat & "  ,  CDCOMPAGNIE  ,  CDAPPLI  ,  CDPRODUIT  ," _
        & "  NUCONTRA  , NUPRET  , NUDOSSIERPREST  , NUSOUSDOSSIERPREST  , NUCONTRATGESTDELEG  , CDGARAN  ," _
        & "  NUREFMVT  , CDTYPMVT  , NUTRAITE  , CDCATADHESION  ,  LBCATADHESION  ,  CDOPTION  , LBOPTIONCON  ," _
        & "  LBSOUSCR  , IDASSUREAGI  ,  IDASSURE  , LBASSURE  , DTNAISSASS  , CDSEXASSURE  ,  IDRENTIERAGI  ," _
        & "  IDRENTIER  ,  LBRENTIER  ,  DTNAISSREN  , IDCORENTIERAGI  , IDCORENTIER  ,  LBCORENTIER  ,  DTNAISSCOR  ," _
        & "  DTSURVSIN  ,  AGESURVSIN  , DURPOSARR  ,  ANCARRTRA  ,  CDPERIODICITE  ,  CDTYPTERME  , DTEFFREN  ," _
        & "  DTLIMPRO  , DTDERREG  , DTDEBPER  , DTFINPER  , CDPROEVA  , ANNEADHE  , CDSINCON  , CDMISINV  ," _
        & "  DTMISINV  , NUETABLI  , TXREVERSION  ,  DTCALCULPROV  , DTTRAITPROV  ,  DTCREATI  , DTSIGBIA  ," _
        & "  DTDECSIN  , CDRISQUE  , LBRISQUE  , CDSITUATSIN  ,  DTSITUATSIN  ,  CDPRETRATTSIN  ,  CDCTGPRT  ," _
        & "  LBCTGPRT  , DTPREECH  , DTDERECH  , CDPERIODICITEECH  , MTECHEANCE1  ,  DTDEBPERECH1  , DTFINPERECH1  ," _
        & "  MTECHEANCE2  ,  DTDEBPERECH2  , DTFINPERECH2  , MTECHEANCE3  ,  DTDEBPERECH3  , DTFINPERECH3  ," _
        & "  CDTYPAMO  , LBTYPAMO  , TXINVPEC  , DTDEBPERPIP  ,  DTFINPERPIP  ,  DTSAISIEPERJUSTIF  ,  DTDEBPERJUSTIF  ," _
        & "  DTFINPERJUSTIF  , DTDEBDERPERRGLTADA  , DTFINDERPERRGLTADA  , DTDERPERRGLTADA  ,  MTDERPERRGLTADA  ," _
        & "  DTDEBDERPERRGLTADC  , DTFINDERPERRGLTADC  , DTDERPERRGLTADC  ,  MTDERPERRGLTADC  ,  CDSINPREPROV  ," _
        & "  MTTOTREGLEICIV  , DTDEBPROV  ,  DTFINPROV  ,  INDBASREV  ,  MTPREANN  , MTPREREV  , MTPREMAJ, "
  
  sSql = sSql & " CDCATINV, LBCATINV,  CDCONTENTIEUX, NUSINISTRE, CDCHOIXPREST, LBCHOIXPREST, MTCAPSSRISQ, " _
        & "  MTPRIREG , MTPRIRE1, MTPRIRE2, CDMONNAIE, CDPAYS, CDAPPLISOURCE, FLAMORTISSABLE , LBCOMLIG "
  
  sSql = sSql & " FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC"
  
  m_dataSource.Execute sSql

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Old method
'
'  Dim rsIn As ADODB.Recordset, rsOut As ADODB.Recordset, i As Integer
'
'  Set rsIn = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Snapshot)
'  Set rsOut = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=-1", Dynamic)
'
'  rsOut.AddNew
'
'  For i = 0 To rsIn.fields.Count - 1
'    If rsOut.fields(i).Name = "DataVersion" Then
'      rsOut.fields("DataVersion") = etat
'    Else
'      rsOut.fields(rsOut.fields(i).Name) = rsIn.fields(i)
'    End If
'  Next
'
'  rsOut.Update
'
'  rsIn.Close
'  rsOut.Close
  
  Exit Sub
  
err_CopyLigne:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


'##ModelId=5C8A6817039F
Private Function EcritTrace(rs As ADODB.Recordset, etat As Integer) As Boolean
  ' ajoute une ligne dans la trace avec les données
  Dim rsTrace As ADODB.Recordset, comment As String, depart As Integer, txtChamps As String, i As Integer
  
  On Error GoTo err_EcritTrace
  
  Set rsTrace = m_dataSource.OpenRecordset("SELECT * FROM P3ITRACE WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement, Dynamic)
  
  rsTrace.AddNew
  
  
  rsTrace.fields("NUTRAITP3I") = NumeroLot
  rsTrace.fields("NUENRP3I") = numeroEnregistrement
  rsTrace.fields("DateModif") = Now
  
  ' copie les anciennes données (appeler EcritTrace() avant d'effectuer les changements)
  If etat <> eAjouter Then
    For i = 3 To rs.fields.Count - 1
      If UCase(rs.fields(i).Name) <> "SELECTED" Then
        rsTrace.fields(rs.fields(i).Name) = rs.fields(i)
      End If
    Next
  End If
  
  ' test le code action ou test si il y a eu un changement de valeur
  Select Case etat
    Case eUndelete ' Undelete
      ' modifie simplement le statut de la ligne
      rsTrace.fields("TypeChangement").Value = "M"
      rsTrace.fields("Commentaire").Value = "Récupération de la ligne"
    
    Case eModifie ' Modifier
      rsTrace.fields("TypeChangement").Value = "M"
      comment = "Modification de la ligne"
      
      txtChamps = vbNullString
      sprEdit.Col = 1
      depart = 3 ' on ignore les cles
      ' changements ?
      Dim bChange As Boolean
      For i = 3 To rs.fields.Count - 1
        sprEdit.Row = i - depart + 1
        bChange = False
        
        If UCase(rs.fields(i).Name) <> "SELECTED" Then
          Select Case rs.fields(i).Type
            Case adDate, adDBDate, adDBTimeStamp
              If sprEdit.text = vbNullString And Not IsNull(rs.fields(i).Value) Then
                bChange = True
              ElseIf sprEdit.text = vbNullString And IsNull(rs.fields(i).Value) Then
                bChange = False ' a cause du CDate
              ElseIf sprEdit.text <> vbNullString And IsNull(rs.fields(i).Value) Then
                bChange = True
              ElseIf CDate(sprEdit.text) <> rs.fields(i).Value Then
                bChange = True
              End If
            
            Case adChar, adVarChar, adVarWChar
              If Trim(sprEdit.Value) <> Trim(rs.fields(i).Value) _
                Or ((IsEmpty(rs.fields(i).Value) Or IsNull(rs.fields(i).Value)) And Trim(sprEdit.text) <> vbNullString) Then
                bChange = True
              End If
            
            Case adNumeric
              If sprEdit.text <> "" And IsNull(rs.fields(i).Value) Then
                bChange = True
              ElseIf sprEdit.text = "" And IsNull(rs.fields(i).Value) Then
                bChange = bChange
              ElseIf CDbl(sprEdit.text) <> CDbl(rs.fields(i).Value) Then
                bChange = True
              End If
            
            Case Else
              If sprEdit.Value <> rs.fields(i).Value Or ((IsEmpty(rs.fields(i).Value) Or IsNull(rs.fields(i).Value)) And sprEdit.text <> vbNullString) Then
                bChange = True
              End If
          End Select
        End If
        
        If bChange = True Then
          txtChamps = txtChamps & rs.fields(i).Name & ", "
        End If
      Next
    
      If txtChamps <> vbNullString Then
        If Right(txtChamps, 2) = ", " Then
          txtChamps = Left(txtChamps, Len(txtChamps) - 2)
        End If
        comment = comment & " - " & txtChamps
      End If
      rsTrace.fields("Commentaire").Value = comment
    
    Case eSupprimer ' Supprimer
      rsTrace.fields("TypeChangement").Value = "S"
      rsTrace.fields("Commentaire").Value = "Suppression de la ligne"
      
    Case eDoublon ' Doublon
      rsTrace.fields("TypeChangement").Value = "D"
      rsTrace.fields("Commentaire").Value = "Suppression (Doublon) de la ligne"
  
    Case eAjouter ' Ajout
      rsTrace.fields("TypeChangement").Value = "A"
      rsTrace.fields("Commentaire").Value = "Ajout de la ligne"
  
  End Select
  
  
  If txtChamps <> vbNullString Or rsTrace.fields("TypeChangement").Value = "S" _
     Or rsTrace.fields("TypeChangement").Value = "D" Or rsTrace.fields("TypeChangement").Value = "A" Then
    rsTrace.Update
  Else
    rsTrace.CancelUpdate
  End If
  
  rsTrace.Close
  
  EcritTrace = bChange
  
  Exit Function
  
err_EcritTrace:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
  
End Function


'##ModelId=5C8A68180005
Public Sub OKButton_Click()
  On Error GoTo err_OK
  
  Screen.MousePointer = vbHourglass
  
  Dim rs As ADODB.Recordset, i As Integer, depart As Integer
  
  Set rs = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Dynamic)
    
  ' test l'etat de la ligne
  If Not IsNull(rs.fields("DataVersion").Value) Then
    Select Case rs.fields("DataVersion").Value
      Case eSupprimer, eDoublon
        ret_code = 0
        Unload Me
    End Select
  End If
    
  ' test le code action ou test si il y a eu un changement de valeur
  Select Case frmAction
    Case eUndelete
      ' modifie simplement le statut de la ligne
      If rs.fields("DataVersion").Value > eModifie Then
        EcritTrace rs, eUndelete
        
        rs.fields("DataVersion").Value = eModifie
        rs.Update
      End If
    
    Case eInitiale, eModifie, eAjouter
      If rs.fields("DataVersion").Value <> eInitiale Then
        ' modifie simplement le statut de la ligne
        
        If frmAction <> eAjouter Then
          EcritTrace rs, eModifie
        End If
        
        sprEdit.Col = 1
        depart = 3 ' on ignore les cles
        
        ' changements ?
        For i = 3 To rs.fields.Count - 1
          sprEdit.Row = i - depart + 1
          If sprEdit.Value <> rs.fields(i).Value Or ((IsNull(rs.fields(i).Value) Or IsEmpty(rs.fields(i).Value)) And sprEdit.text <> vbNullString) Then
            Select Case rs.fields(i).Type
              Case adDate, adDBDate, adDBTimeStamp
                If sprEdit.text = vbNullString Or IsEmpty(sprEdit.text) Then
                  rs.fields(i).Value = Null
                Else
                  rs.fields(i).Value = sprEdit.text
                End If
              
              Case adChar, adVarChar, adVarWChar
                rs.fields(i).Value = Trim(sprEdit.text)
              
              Case adSmallInt, adTinyInt, adBigInt, adInteger, adDouble, adNumeric
                If sprEdit.text = vbNullString Then
                  rs.fields(i).Value = Null
                Else
                  rs.fields(i).Value = m_dataHelper.GetDouble2(sprEdit.Value)
                End If
                        
              Case Else
                rs.fields(i).Value = sprEdit.Value
            End Select
          End If
        Next
          
        rs.fields("DataVersion").Value = eModifie
        
        
        If frmAction = eAjouter Then
          rs.fields("LBCOMLIG").Value = "Ligne ajoutee"
        End If
        
        rs.Update
      Else
        rs.Close
        
        ' copie la ligne et change son statut
        CopyLigne eModifie
        
        Set rs = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Dynamic)
        
        EcritTrace rs, eModifie
        
        sprEdit.Col = 1
        depart = 3 ' on ignore les cles
        
        ' changements ?
        For i = 3 To rs.fields.Count - 1
          sprEdit.Row = i - depart + 1
          If sprEdit.text <> "" And IsNull(rs.fields(i).Value) Then
            rs.fields(i).Value = sprEdit.text
          ElseIf rs.fields(i).Type = adNumeric Then
            If sprEdit.text = "" Then
              If Not IsNull(rs.fields(i).Value) Then
                rs.fields(i).Value = 0
              End If
            ElseIf CDbl(sprEdit.text) <> CDbl(rs.fields(i).Value) Then
              rs.fields(i).Value = sprEdit.text
            End If
          ElseIf sprEdit.text <> rs.fields(i).Value Then
            rs.fields(i).Value = sprEdit.text
          End If
        Next
        
        rs.fields("DataVersion").Value = eModifie
        rs.Update
      End If
    
    Case 2 ' Supprimer
      If rs.fields("DataVersion").Value <> eInitiale Then
        ' modifie simplement le statut de la ligne
        EcritTrace rs, eSupprimer
        
        rs.fields("DataVersion").Value = eSupprimer
        rs.Update
      Else
        rs.Close
        
        ' copie la ligne et change son statut
        CopyLigne eSupprimer
        
        Set rs = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Dynamic)

        EcritTrace rs, eSupprimer

        'rs.fields("DataVersion").Value = eSupprimer
        'rs.Update
      End If
      
    Case 3 ' Doublon
      If rs.fields("DataVersion").Value <> eInitiale Then
        ' modifie simplement le statut de la ligne
        EcritTrace rs, eDoublon
        
        rs.fields("DataVersion").Value = eDoublon
        rs.Update
      Else
        rs.Close
        
        ' copie la ligne et change son statut
        CopyLigne eDoublon
        
        Set rs = m_dataSource.OpenRecordset("SELECT * FROM P3IPROVCOLL WHERE NUTRAITP3I=" & NumeroLot & " AND NUENRP3I=" & numeroEnregistrement & " ORDER BY DataVersion DESC", Dynamic)

        EcritTrace rs, eDoublon

        'rs.fields("DataVersion").Value = eDoublon
        'rs.Update
        
      End If
  
  End Select
  
  rs.Close
  
  Screen.MousePointer = vbDefault
  
  ret_code = 0
  Unload Me
  
  Exit Sub
  
err_OK:
  Screen.MousePointer = vbDefault
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
  Screen.MousePointer = vbHourglass
End Sub
