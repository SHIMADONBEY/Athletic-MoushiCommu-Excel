Attribute VB_Name = "MacroProcessorOperator"
'namespace=vba-files/Process
Option Explicit

' NOTE: 緊急時に備え、`ResetProcessor` メソッドを呼び出すことで、マクロプロセッサの状態を初期化できるようにしています。
'       そのため、`Option Private Module` は設定せず、外部からアクセス可能な状態にしています。

Private m_Processor As MacroProcessor

' NOTE: この関数は、マクロプロセッサのインスタンスを取得するためのものです。
'       マクロの登録や実行できないようにするため、ダミー引数を追加しています。
Public Function GetProcessor(Optional ByVal Switch As Boolean = True) As MacroProcessor
    If m_Processor Is Nothing Then
        Set m_Processor = New MacroProcessor
    End If
    Set GetProcessor = m_Processor
End Function

' NOTE: このサブルーチンは、マクロプロセッサのインスタンスを削除するためのものです。
'       マクロの登録や実行できないようにするため、ダミー引数を追加しています。
Public Sub DeleteProcessor(Optional ByVal Switch As Boolean = True)
    Set m_Processor = Nothing
End Sub

' NOTE: このサブルーチンは、マクロプロセッサの状態をリセットするためのものです。
'       緊急時に備え、`ResetProcessor` メソッドを呼び出すことで、マクロプロセッサの状態を初期化できます。
Public Sub ResetProcessor()
    Application.ScreenUpdating = shtInternalRegistory.Cells(2, 2).Value
    Application.Calculation = shtInternalRegistory.Cells(3, 2).Value
    Application.EnableEvents = shtInternalRegistory.Cells(4, 2).Value
End Sub
