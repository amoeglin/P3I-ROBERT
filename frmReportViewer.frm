VERSION 5.00
Begin VB.Form frmReportViewer 
   Caption         =   "report viewer..."
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11220
   Icon            =   "frmReportViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox CRViewer1 
      Height          =   4095
      Left            =   840
      ScaleHeight     =   4035
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A6829034E"

Option Explicit

'##ModelId=5C8A682A0071
Public m_report As CRAXDRT.Report
'

'##ModelId=5C8A682A007F
Private Sub Form_Load()
  On Error GoTo err_LanceEtat
  
  Screen.MousePointer = vbHourglass
  
  WindowState = vbMaximized
  
  'CRViewer1.ReportSource = m_report
  
  'CRViewer1.ViewReport
  'Do While CRViewer1.IsBusy
    'DoEvents
  'Loop
  
  Screen.MousePointer = vbDefault
  
  Exit Sub
  
err_LanceEtat:
  Screen.MousePointer = vbDefault
  DisplayError
  Resume Next
End Sub


'##ModelId=5C8A682A008F
Public Sub Form_Unload(Cancel As Integer)
  'Do Until CRViewer1.ViewCount <= 1
  '  CRViewer1.CloseView CRViewer1.ViewCount
  'Loop
  
  'Do While CRViewer1.IsBusy = True
  '  DoEvents
  'Loop
  
  'Set m_report = Nothing
End Sub


'##ModelId=5C8A682A00AE
Private Sub Form_Resize()
  CRViewer1.top = 0
  CRViewer1.Left = 0
  CRViewer1.Width = Me.ScaleWidth
  CRViewer1.Height = Me.ScaleHeight
End Sub


'##ModelId=5C8A682A00BE
Private Sub DisplayError()
  Dim str As String
  
  str = "Erreur " & Err & vbLf & Err.Description
    
  MsgBox str, vbCritical
End Sub

'
'Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'  m_report.PrinterSetupEx Me.hwnd
'End Sub

