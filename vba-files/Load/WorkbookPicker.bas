Attribute VB_Name = "WorkbookPicker"
'namespace=vba-files/Load
Option Explicit
Option Private Module

Public Function Pick() As Collection
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim vSelectedFiles As Collection: Set vSelectedFiles = New Collection
    With fd
        .Title = "申込書を選んでください."
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = True
        If .Show = -1 Then
            Dim vFilePath As Variant
            For Each vFilePath In .SelectedItems
                vSelectedFiles.Add vFilePath
            Next vFilePath
        End If
    End With
    Set Pick = vSelectedFiles
End Function
