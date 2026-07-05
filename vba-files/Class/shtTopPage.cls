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
    Dim vShapeButton As ShapeButton: Set vShapeButton = New ShapeButton
    vShapeButton.Initialize "btnAggregateTeam_Front", "btnAggregateTeam_Back"
    vShapeButton.ClickDown
    Application.Wait Now + TimeValue("00:00:02")
    vShapeButton.ClickUp
End Sub