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
    If VarType(shtInternalRegistry.Cells(2, 2).Value) = vbBoolean Then
        Application.ScreenUpdating = CBool(shtInternalRegistry.Cells(2, 2).Value)
    Else
        ' Boolean 型でない場合は、デフォルトの True に設定します。
        Application.ScreenUpdating = True
    End If

    If VarType(shtInternalRegistry.Cells(4, 2).Value) = vbBoolean Then
        Application.EnableEvents = CBool(shtInternalRegistry.Cells(4, 2).Value)
    Else
        ' Boolean 型でない場合は、デフォルトの True に設定します。
        Application.EnableEvents = True
    End If

    Dim vCalculationMode As Variant: vCalculationMode = shtInternalRegistry.Cells(3, 2).Value

    Select Case vCalculationMode
    Case xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic
        ' XlCalculation の定数値を使用して、計算モードを設定します。
        Application.Calculation = CLng(vCalculationMode)
    Case Else
        ' XlCalculation の定数値以外の場合は、デフォルトの自動計算モードに設定します。
        Application.Calculation = xlCalculationAutomatic
    End Select
End Sub
