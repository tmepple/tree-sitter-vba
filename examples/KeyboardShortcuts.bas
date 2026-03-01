Attribute VB_Name = "KeyboardShortcuts"
Option Explicit

' Auto_Open - Called automatically when the Add-In is loaded
' Registers all global keyboard shortcuts
Sub Auto_Open()
    ' Ctrl+Shift+B → Borrowing Base calculator
    Application.OnKey "^+b", "CalcBorrowingBaseShortcut"

    ' Add more shortcuts here as you migrate from PERSONAL.XLSB
    ' Mac key syntax notes:
    '   ^ = Control
    '   + = Shift
    '   % = Option (Alt on Windows)
    '   For Command key on Mac, use the key directly (limited support)
    '
    ' Example: Application.OnKey "^%l", "YourMacroName"
End Sub

' Auto_Close - Called automatically when the Add-In is unloaded
' Unregisters all keyboard shortcuts to clean up
Sub Auto_Close()
    ' Reset shortcuts to default behavior
    Application.OnKey "^+b", ""
End Sub
