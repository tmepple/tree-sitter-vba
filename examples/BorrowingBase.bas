Attribute VB_Name = "BorrowingBase"
Option Explicit

Private Const INVENTORY_RATE As Double = 0.75
Private Const AR_RATE As Double = 0.8
Private Const INVENTORY_PREFIX As String = "120"

' CalcBorrowingBase - Main entry point (called from Ribbon button)
' Ribbon callback signature requires the control parameter
Public Sub CalcBorrowingBase(control As IRibbonControl)
    RunBorrowingBase
End Sub

' CalcBorrowingBaseShortcut - Entry point for keyboard shortcut (no parameters)
Public Sub CalcBorrowingBaseShortcut()
    RunBorrowingBase
End Sub

' RunBorrowingBase - Core logic, called by both entry points
Private Sub RunBorrowingBase()
    Dim wsBalance As Worksheet
    Dim wsAR As Worksheet
    Dim inventoryTotal As Double
    Dim arTotal As Double
    Dim borrowingBase As Double

    ' Validate workbook has required sheets
    Set wsBalance = FindSheet("BalanceSheet")
    If wsBalance Is Nothing Then
        MsgBox "Sheet 'BalanceSheet' not found in the active workbook.", vbExclamation
        Exit Sub
    End If

    Set wsAR = FindSheet("ARAgingSummary")
    If wsAR Is Nothing Then
        MsgBox "Sheet 'ARAgingSummary' not found in the active workbook.", vbExclamation
        Exit Sub
    End If

    ' Calculate components
    inventoryTotal = CalcInventory(wsBalance)
    arTotal = CalcEligibleAR(wsAR)

    If inventoryTotal < 0 Or arTotal < 0 Then
        ' Error already shown by the sub-functions
        Exit Sub
    End If

    Dim inventoryComponent As Double
    Dim arComponent As Double
    inventoryComponent = inventoryTotal * INVENTORY_RATE
    arComponent = arTotal * AR_RATE
    borrowingBase = inventoryComponent + arComponent

    ' Display result
    Dim msg As String
    msg = "Borrowing Base Calculation" & vbCrLf & vbCrLf
    msg = msg & "Inventory (120x accounts): " & FormatAsCurrency(inventoryTotal) & vbCrLf
    msg = msg & "  x " & Format(INVENTORY_RATE, "0%") & " = " & FormatAsCurrency(inventoryComponent) & vbCrLf & vbCrLf
    msg = msg & "Eligible AR (Current + 30 + 60): " & FormatAsCurrency(arTotal) & vbCrLf
    msg = msg & "  x " & Format(AR_RATE, "0%") & " = " & FormatAsCurrency(arComponent) & vbCrLf & vbCrLf
    msg = msg & "Borrowing Base: " & FormatAsCurrency(borrowingBase)

    MsgBox msg, vbInformation, "Borrowing Base"
End Sub

' CalcInventory - Sum all account balances where account code starts with "120"
' Scans BalanceSheet column A for account codes, sums corresponding column B values
' Returns -1 on error
Private Function CalcInventory(ws As Worksheet) As Double
    Dim matchedRows As Collection
    Dim total As Double
    Dim i As Long
    Dim rowNum As Variant

    ' Find all rows with account codes starting with "120"
    ' Account codes appear as e.g. "12000 - Inventory", "12020 - Inventory - WIP"
    Set matchedRows = FindRowsByPrefix(ws, 1, INVENTORY_PREFIX)

    If matchedRows.Count = 0 Then
        MsgBox "No inventory accounts (prefix '" & INVENTORY_PREFIX & "') found on BalanceSheet.", vbExclamation
        CalcInventory = -1
        Exit Function
    End If

    total = 0
    For Each rowNum In matchedRows
        Dim cellVal As Variant
        cellVal = ws.Cells(CLng(rowNum), 2).Value
        If IsNumeric(cellVal) Then
            total = total + CDbl(cellVal)
        End If
    Next rowNum

    CalcInventory = total
End Function

' CalcEligibleAR - Sum Current + 30-day + 60-day AR from the Total row
' AR column order (fixed positional): A=Name, B=Current, C=30-day, D=60-day, E=90-day, F=>90, G=Total
' Returns -1 on error
Private Function CalcEligibleAR(ws As Worksheet) As Double
    Dim totalRow As Long
    Dim arCurrent As Double
    Dim ar30 As Double
    Dim ar60 As Double

    totalRow = FindTotalRow(ws, 1)

    If totalRow = 0 Then
        MsgBox "Could not find 'Total' row on ARAgingSummary.", vbExclamation
        CalcEligibleAR = -1
        Exit Function
    End If

    ' Columns B, C, D = Current, 30-day, 60-day (positional, headers are dynamic dates)
    Dim val As Variant

    val = ws.Cells(totalRow, 2).Value
    arCurrent = IIf(IsNumeric(val), CDbl(val), 0)

    val = ws.Cells(totalRow, 3).Value
    ar30 = IIf(IsNumeric(val), CDbl(val), 0)

    val = ws.Cells(totalRow, 4).Value
    ar60 = IIf(IsNumeric(val), CDbl(val), 0)

    CalcEligibleAR = arCurrent + ar30 + ar60
End Function
