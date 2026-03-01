VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Delegate to Auto_Open for shortcut registration
    ' Auto_Open handles the actual setup so it can also be called
    ' when the Add-In is loaded via VBA at runtime
    Auto_Open
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Auto_Close
End Sub
