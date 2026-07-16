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
    Application.ScreenUpdating = shtInternalRegistry.Cells(2, 2).Value
    Application.Calculation = shtInternalRegistry.Cells(3, 2).Value
    Application.EnableEvents = shtInternalRegistry.Cells(4, 2).Value

    shtInternalRegistry.Cells(1, 2).ClearContents
    shtInternalRegistry.Cells(2, 2).ClearContents
    shtInternalRegistry.Cells(3, 2).ClearContents
    shtInternalRegistry.Cells(4, 2).ClearContents
End Sub
