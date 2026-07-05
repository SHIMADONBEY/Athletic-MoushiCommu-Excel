VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "MacroProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/Process
Option Explicit

Private m_ScreenUpdatingState As Boolean
Private m_CalculationState As XlCalculation
Private m_EnableEventsState As Boolean

Private Sub Class_Initialize()
    m_ScreenUpdatingState = Application.ScreenUpdating
    m_CalculationState = Application.Calculation
    m_EnableEventsState = Application.EnableEvents
End Sub

Private Sub Class_Terminate()
    Me.RestoreState
End Sub

Public Sub StartProcessing()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

Public Sub ToggleScreenUpdating()
    Application.ScreenUpdating = Not Application.ScreenUpdating
End Sub

Public Sub RestoreState()
    Application.ScreenUpdating = m_ScreenUpdatingState
    Application.Calculation = m_CalculationState
    Application.EnableEvents = m_EnableEventsState
End Sub

Public Property Get ScreenUpdatingState() As Boolean
    ScreenUpdatingState = m_ScreenUpdatingState
End Property

Public Property Get CalculationState() As XlCalculation
    CalculationState = m_CalculationState
End Property

Public Property Get EnableEventsState() As Boolean
    EnableEventsState = m_EnableEventsState
End Property

