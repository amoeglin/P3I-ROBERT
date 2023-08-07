Attribute VB_Name = "Module_NbMois"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"5C8A67DC0050"
Option Explicit

' Nombre de mois entre 2 dates (on se place en fin de mois)
'##ModelId=5C8A67DC0169
Public Function NbMois(DateDebut As Date, DateFin As Date) As Integer
  NbMois = 0
  If Year(DateDebut) < Year(DateFin) Then
  NbMois = 12 - Month(DateDebut) + 12 * (Year(DateFin) - Year(DateDebut) - 1) + Month(DateFin)
     ElseIf Year(DateDebut) = Year(DateFin) Then
     NbMois = Month(DateFin) - Month(DateDebut)
  End If
  
  If NbMois < 0 Then
    NbMois = 0
  Else
  End If
End Function


