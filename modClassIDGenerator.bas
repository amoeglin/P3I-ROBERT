Attribute VB_Name = "modClassIDGenerator"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F85491100EA"
Option Explicit
'##ModelId=3F85491100FA
Function GetNextClassDebugID() As Long
    'class ID generator
    Static lClassDebugID As Long
    lClassDebugID = lClassDebugID + 1
    GetNextClassDebugID = lClassDebugID
End Function

