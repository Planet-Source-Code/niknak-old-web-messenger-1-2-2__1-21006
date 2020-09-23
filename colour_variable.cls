VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colour_variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Public variable_name As String
Public variable_description As String
Public variable_colour As Integer

'SAVES THE COLVAR MEMBERS TO THE SYSTEM REGISTRY
Public Sub save_vars()
    SaveSetting App.ProductName, variable_name, "variable_colour", variable_colour
End Sub

'LOADS THE COLVAR MEMBERS FROM THE SYSTEM REGISTRY
Public Sub load_vars()
    If GetSetting(App.ProductName, variable_name, "variable_colour") <> "" Then
        variable_colour = Val(GetSetting(App.ProductName, variable_name, "variable_colour"))
    End If
End Sub

