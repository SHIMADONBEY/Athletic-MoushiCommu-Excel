Attribute VB_Name = "MacroProcessorOperator"
'namespace=vba-files/Process
Option Explicit

' NOTE: 緊急時に備え、`ResetProcessor` メソッドを呼び出すことで、マクロプロセッサの状態を初期化できるようにしています。
'       そのため、`Option Private Module` は設定せず、外部からアクセス可能な状態にしています。

Private m_Processor As MacroProcessor

' NOTE: この関数は、マクロプロセッサのインスタンスを取得するためのものです。
'       マクロの登録や実行できないようにするため、ダミー引数を追加しています。
Public Function GetProcessor(Optional ByVal DummyFlag As Boolean = True) As MacroProcessor
    If m_Processor Is Nothing Then
        Set m_Processor = New MacroProcessor
    End If
    Set GetProcessor = m_Processor
End Function

' NOTE: このサブルーチンは、マクロプロセッサのインスタンスを削除するためのものです。
'       マクロの登録や実行できないようにするため、ダミー引数を追加しています。
Public Sub DeleteProcessor(Optional ByVal DummyFlag As Boolean = True)
    Set m_Processor = Nothing
End Sub

' NOTE: このサブルーチンは、マクロプロセッサの状態をリセットするためのものです。
'       緊急時に備え、`ResetProcessor` メソッドを呼び出すことで、マクロプロセッサの状態を初期化できます。
Public Sub ResetProcessor()
    If IsEmpty(shtInternalRegistry.Cells(2, 2).Value) Then
        Application.ScreenUpdating = True
    ElseIf Not VarType(shtInternalRegistry.Cells(2, 2).Value) = vbBoolean Then
        Application.ScreenUpdating = True
    Else
        Application.ScreenUpdating = CBool(shtInternalRegistry.Cells(2, 2).Value)
    End If

    If IsEmpty(shtInternalRegistry.Cells(4, 2).Value) Then
        Application.EnableEvents = True
    ElseIf Not VarType(shtInternalRegistry.Cells(4, 2).Value) = vbBoolean Then
        Application.EnableEvents = True
    Else 
        Application.EnableEvents = CBool(shtInternalRegistry.Cells(4, 2).Value)
    End If

    Application.Calculation = shtInternalRegistry.Cells(3, 2).Value
End Sub
