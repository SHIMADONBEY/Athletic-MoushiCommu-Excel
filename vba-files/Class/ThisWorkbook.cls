VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private m_MacroProcessor As MacroProcessor

Private Sub Workbook_Open()
    Set m_MacroProcessor = New MacroProcessor
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set m_MacroProcessor = Nothing
End Sub

Public Function GetMacroProcessor() As MacroProcessor
    Set GetMacroProcessor = m_MacroProcessor
End Function
