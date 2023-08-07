VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public FichierLog As String
Public bErreurDansLog As Boolean
Public modeAuto As Boolean
'

Public Sub Reset()
  FichierLog = ""
  bErreurDansLog = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' création du fichier log
'
Public Sub CreateLog(titre As String, Optional fErase As Boolean = True)
  If FichierLog = "" Then
    MsgBox "CreateLog() : FichierLog non renseigné!", vbCritical
    Exit Sub
  End If
  
  If fErase = True Then
    On Error Resume Next
    Kill FichierLog
  Else
    Call EcritTraceDansLog("")
    Call EcritTraceDansLog("--------------------------------------------------")
    Call EcritTraceDansLog("")
  End If
  
  If Trim$(titre) <> "" Then
    Call EcritTraceDansLog(titre)
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ecrit un message dans le log
'
Public Sub EcritTraceDansLog(text As String)
  Dim f As Integer
  
  On Error GoTo err_EcritTraceDansLog
  
  f = FreeFile(0)
  
  bErreurDansLog = True
  
  Open FichierLog For Append As #f
    Print #f, Now() & " - " & text
  Close #f

  Exit Sub
  
err_EcritTraceDansLog:

  If Not modeAuto Then
    MsgBox "Erreur " & Err & " durant l'écriture dans " & FichierLog & vbLf & Err.Description, vbCritical
  End If
  
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' affichage des erreurs
'
Public Sub AfficheErreurLog(Optional fAsk As Boolean = True, Optional sMessage As String = "")
  If sMessage = "" Then
    sMessage = "Voulez-vous consultez le fichier " & FichierLog & " qui contient les erreurs apparues durant le traitement ?"
  Else
    sMessage = Replace(sMessage, "<FichierLog>", FichierLog)
  End If
  
  If fAsk Then
    If MsgBox(sMessage, vbInformation + vbYesNo) = vbNo Then
      Exit Sub
    End If
  End If
  
  Dim frm As New frmDisplayLog
  
  frm.FichierLog = FichierLog
  
  frm.Show vbModal
  
  Set frm = Nothing
End Sub

