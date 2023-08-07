Attribute VB_Name = "iDuvGlobal"
Option Explicit

Private Function Maximum(a As Double, b As Double) As Double
  If a > b Then
    Maximum = a
  Else
    Maximum = b
  End If
End Function

' calcule la largeur des colonnes
' ObjWnd  : form contenant l'objet listview
' ObjList : objet listview
Public Sub LargeurAutomatique(ObjWnd As Form, ObjList As Object)
  Dim ColWidth() As Double
  Dim i As Long, j As Long
    
  ' calcule la largeur des colonnes
  ReDim ColWidth(ObjList.ColumnHeaders.Count + 5) As Double
  
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

Public Function BuildDateLimit(debut As String, Fin As String) As String
  Dim Limit As String
  
  Limit = "BETWEEN #" & Format(CDate(debut), "mm/dd/yyyy") & "# AND #" & Format(CDate(Fin), "mm/dd/yyyy") & "#"
  
  BuildDateLimit = Limit
End Function

