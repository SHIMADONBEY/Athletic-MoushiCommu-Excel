VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "TeamListColumnConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/Team/Config
Option Explicit

Private m_ColumnOrder As Long
Private m_Title As String
Private m_Address As String
Private m_DataType As String

Private m_Initialized As Boolean

Private Sub Class_Initialize()
    m_Initialized = False
End Sub

Public Sub Initialize(ByVal ColumnOrder As Long, ByVal Title As String, ByVal Address As String, ByVal DataType As String)
    m_ColumnOrder = ColumnOrder
    m_Title = Title
    m_Address = Address
    m_DataType = DataType
    m_Initialized = True
End Sub

Public Property Get ColumnOrder() As Long
    If Not m_Initialized Then Err.Raise vbObjectError + 1, TypeName(Me) & ".ColumnOrder", TypeName(Me) & ".ColumnOrder not initialized"
    ColumnOrder = m_ColumnOrder
End Property

Public Property Get Title() As String
    If Not m_Initialized Then Err.Raise vbObjectError + 1, TypeName(Me) & ".Title", TypeName(Me) & ".Title not initialized"
    Title = m_Title
End Property

Public Property Get Address() As String
    If Not m_Initialized Then Err.Raise vbObjectError + 1, TypeName(Me) & ".Address", TypeName(Me) & ".Address not initialized"
    Address = m_Address
End Property

Public Property Get DataType() As String
    If Not m_Initialized Then Err.Raise vbObjectError + 1, TypeName(Me) & ".DataType", TypeName(Me) & ".DataType not initialized"
    DataType = m_DataType
End Property

Public Function Self() As TeamListColumnConfiguration
    Set Self = Me
End Function
