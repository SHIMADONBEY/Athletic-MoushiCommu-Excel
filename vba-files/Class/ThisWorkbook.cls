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
    shtInternalRegistory.Cells(1, 2).Value = Now
    shtInternalRegistory.Cells(2, 2).Value = Application.ScreenUpdating
    shtInternalRegistory.Cells(3, 2).Value = Application.Calculation
    shtInternalRegistory.Cells(4, 2).Value = Application.EnableEvents
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ScreenUpdating = shtInternalRegistory.Cells(2, 2).Value
    Application.Calculation = shtInternalRegistory.Cells(3, 2).Value
    Application.EnableEvents = shtInternalRegistory.Cells(4, 2).Value

    shtInternalRegistory.Cells(1, 2).ClearContents
    shtInternalRegistory.Cells(2, 2).ClearContents
    shtInternalRegistory.Cells(3, 2).ClearContents
    shtInternalRegistory.Cells(4, 2).ClearContents
End Sub
