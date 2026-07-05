VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtTopPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'namespace=vba-files/Class
Option Explicit

Public Sub btnAggregateTeamClick()
    On Error GoTo error_btnAggregateTeamClick
        MacroProcessorOperator.GetProcessor().StartProcessing
        Dim vShapeButton As ShapeButton: Set vShapeButton = New ShapeButton
        vShapeButton.Initialize "btnAggregateTeam_front", "btnAggregateTeam_back"
        vShapeButton.ClickDown
        Application.Wait Now + TimeValue("00:00:01")

        Dim vEntryBooks As Collection: Set vEntryBooks = WorkbookPicker.Pick()
        TeamAggregateController.AggregateTeamList vEntryBooks
    On Error GoTo 0
    GoTo finally_btnAggregateTeamClick
error_btnAggregateTeamClick:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
finally_btnAggregateTeamClick:
    If Not vShapeButton Is Nothing Then vShapeButton.ClickUp
    MacroProcessorOperator.DeleteProcessor
End Sub

Public Sub btnAggregateEntryClick()
    On Error GoTo error_btnAggregateEntryClick
        MacroProcessorOperator.GetProcessor().StartProcessing
        Dim vShapeButton As ShapeButton: Set vShapeButton = New ShapeButton
        vShapeButton.Initialize "btnAggregateEntry_front", "btnAggregateEntry_back"
        vShapeButton.ClickDown
        Application.Wait Now + TimeValue("00:00:01")

        Dim vEntryBooks As Collection: Set vEntryBooks = WorkbookPicker.Pick()
        EntrySheetAggregateController.AggregateEntryList vEntryBooks
    On Error GoTo 0
    GoTo finally_btnAggregateEntryClick
error_btnAggregateEntryClick:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
finally_btnAggregateEntryClick:
    If Not vShapeButton Is Nothing Then vShapeButton.ClickUp
    MacroProcessorOperator.DeleteProcessor
End Sub
