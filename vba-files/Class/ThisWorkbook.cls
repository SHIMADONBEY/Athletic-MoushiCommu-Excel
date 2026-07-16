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
    shtInternalRegistry.Cells(1, 2).Value = Now
    shtInternalRegistry.Cells(2, 2).Value = Application.ScreenUpdating
    shtInternalRegistry.Cells(3, 2).Value = Application.Calculation
    shtInternalRegistry.Cells(4, 2).Value = Application.EnableEvents
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If VarType(shtInternalRegistry.Cells(2, 2).Value) = vbBoolean Then
        Application.ScreenUpdating = shtInternalRegistry.Cells(2, 2).Value
    Else
        Application.ScreenUpdating = True
    End If

    If VarType(shtInternalRegistry.Cells(4, 2).Value) = vbBoolean Then
        Application.EnableEvents = shtInternalRegistry.Cells(4, 2).Value
    Else
        Application.EnableEvents = True
    End If

    Dim varCalc As Variant: varCalc = shtInternalRegistry.Cells(3, 2).Value
    Select Case varCalc
        Case xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic
            Application.Calculation = CLng(varCalc)
        Case Else
            Application.Calculation = xlCalculationAutomatic
    End Select

    shtInternalRegistry.Cells(1, 2).ClearContents
    shtInternalRegistry.Cells(2, 2).ClearContents
    shtInternalRegistry.Cells(3, 2).ClearContents
    shtInternalRegistry.Cells(4, 2).ClearContents
End Sub
