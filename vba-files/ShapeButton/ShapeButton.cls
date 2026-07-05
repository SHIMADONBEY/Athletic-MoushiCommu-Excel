VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/ShapeButton
Option Explicit

Private m_ButtonFront As Shape
Private m_ButtonBack As Shape

Private m_Initialized As Boolean
Private m_Clicked As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
    m_Clicked = False
    Set m_ButtonFront = Nothing
    Set m_ButtonBack = Nothing
End Sub

Private Sub Class_Terminate()
    If m_Clicked Then
        m_ButtonFront.Visible = True
    End If

    Set m_ButtonFront = Nothing
    Set m_ButtonBack = Nothing
End Sub

Public Sub Initialize(ByVal ButtonFrontName As String, ByVal ButtonBackName As String)
    Set m_ButtonFront = ActiveSheet.Shapes(ButtonFrontName)
    Set m_ButtonBack = ActiveSheet.Shapes(ButtonBackName)
    m_Initialized = True
End Sub

Public Sub ClickDown()
    If Not m_Initialized Then Exit Sub
    m_ButtonFront.Visible = False
    m_Clicked = True
End Sub

Public Sub ClickUp()
    If Not m_Initialized Then Exit Sub
    If Not m_Clicked Then Exit Sub
    m_ButtonFront.Visible = True
    m_Clicked = False
End Sub
