VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDisplayLog 
   Caption         =   "Erreurs durant l'import..."
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   Icon            =   "frmDisplayLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExcel 
      Caption         =   "Excel"
      Height          =   285
      Left            =   7740
      TabIndex        =   5
      Top             =   30
      Width           =   960
   End
   Begin VB.TextBox lblLog 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   7665
   End
   Begin VB.CommandButton btnImprimer 
      Caption         =   "&Imprimer"
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   6960
      Width           =   1740
   End
   Begin RichTextLib.RichTextBox edtLog 
      Height          =   6540
      Left            =   45
      TabIndex        =   2
      Top             =   315
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11536
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDisplayLog.frx":1BB2
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   1
      Top             =   6840
      Width           =   8655
   End
   Begin VB.CommandButton btnFermer 
      Caption         =   "&Fermer et Continuer"
      Default         =   -1  'True
      Height          =   330
      Left            =   6840
      TabIndex        =   0
      Top             =   6960
      Width           =   1740
   End
End
Attribute VB_Name = "frmDisplayLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gWidth As Integer
Private gHeight As Integer
Private gEdtHeight As Integer

Public FichierLog As String
Public FichierLog_FileName As String
Public bFichierLogIsText As Boolean

Public m_sFichierIni As String
'

Private Sub btnExcel_Click()
  Dim xl As Excel.Application
  Dim books As Excel.Workbooks
  Dim srcBook As Excel.Workbook
  Dim sheet As Excel.Worksheet
  Dim sFileName As String
  
  On Error GoTo err_ExcelFile
  
  If bFichierLogIsText Then
    sFileName = sReadIniFile("Dir", "ExportPath", App.Path, 300, m_sFichierIni)
    If Right(sFileName, 1) <> "\" Then
      sFileName = sFileName & "\"
    End If
    sFileName = sFileName & FichierLog_FileName

    edtLog.SaveFile sFileName, rtfText
  Else
    sFileName = FichierLog
  End If
  
  Set xl = New Excel.Application
  
  xl.Visible = True
  
  Set books = xl.Workbooks
    
  books.OpenText sFileName, xlWindows, 1, xlDelimited, xlDoubleQuote, False, False, False, False, False, True, "="
  Set srcBook = xl.ActiveWorkbook
  Set sheet = srcBook.ActiveSheet
  
  sheet.Columns("A:A").EntireColumn.AutoFit

  Exit Sub

err_ExcelFile:
  MsgBox "Erreur " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub

Private Sub btnFermer_Click()
  Unload Me
End Sub

Private Sub btnImprimer_Click()
  On Error GoTo err_print
  
  Dim s As String
  
  Printer.FontName = "Arial"
  Printer.FontSize = 10
  
  s = lblLog & vbLf & vbLf
  
  edtLog.SelStart = 0
  edtLog.SelLength = 0
  edtLog.SelText = s
  
  edtLog.SelPrint Printer.hDC
  
  'Printer.NewPage
  Printer.EndDoc
  
  edtLog.SelStart = 0
  edtLog.SelLength = Len(s)
  edtLog.SelText = ""
  
  Exit Sub
  
err_print:
  MsgBox "Error " & Err & vbLf & Err.Description, vbCritical
  Resume Next
End Sub


Private Sub Form_Initialize()

  bFichierLogIsText = False
  
End Sub

Private Sub Form_Load()
  On Error GoTo err_Load
  ' Centre la fenetre
  Left = (Screen.Width - Width) / 2
  Top = (Screen.Height - Height) / 2
  
  If bFichierLogIsText Then
    lblLog = "Informations"
    
    edtLog.text = FichierLog
  Else
    lblLog = "Contenu du fichier des erreurs : " & FichierLog
    
    edtLog.text = ""
    edtLog.LoadFile FichierLog
  End If
  
  gWidth = Me.Width
  gHeight = Me.Height
  gEdtHeight = edtLog.Height
  
  Exit Sub
  
err_Load:
  edtLog.text = edtLog.text & vbLf & "Erreur durant la lecture du fichier : " & Err & " - " & Err.Description
  Resume Next
End Sub

Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub
  
  If Me.Width < gWidth Then
    Me.Width = gWidth
  End If
  
  If Me.Height < gHeight Then
    Me.Height = gHeight
  End If
  
  lblLog.Left = 0
  lblLog.Width = Me.Width - 180 - btnExcel.Width
  
  btnExcel.Left = Me.Width - 150 - btnExcel.Width
  
  edtLog.Left = 0
  edtLog.Width = Me.Width - 150
  edtLog.Height = Me.Height - gHeight + gEdtHeight
  
  Frame1.Left = 0
  Frame1.Top = edtLog.Top + edtLog.Height
  Frame1.Width = Me.Width
  
  btnImprimer.Left = 50
  btnImprimer.Top = edtLog.Top + edtLog.Height + 150
  
  btnFermer.Left = Me.Width - btnFermer.Width - 200
  btnFermer.Top = edtLog.Top + edtLog.Height + 150
End Sub
