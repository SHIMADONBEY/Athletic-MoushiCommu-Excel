Attribute VB_Name = "TeamAggregateController"
'namespace=vba-files/Team
Option Explicit
Option Private Module

Private Const ENTRY_BOOK_TEAM_SHEET_NAME As String = "(1)鑑"

Public Sub AggregateTeamList(ByVal EntryBookFiles As Collection)
    If EntryBookFiles.Count = 0 Then
        MsgBox "申込が選択されていません。", vbExclamation
        Exit Sub
    End If

    Dim wsTeamList As Worksheet
    Set wsTeamList = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    
    Dim vColumnList As Collection: Set vColumnList = TeamListColumnConfigRepository.ReadAll()

    InitializeTeamListSheet wsTeamList, vColumnList

    Dim vCurrentRow As Long: vCurrentRow = 2
    Dim vFilePath As Variant
    For Each vFilePath In EntryBookFiles
        Dim vExtractedCount As Long: vExtractedCount = ExtractTeamsFromEntryBook(vFilePath, wsTeamList, vCurrentRow, vColumnList)

        If vExtractedCount > 0 Then
            Debug.Print "申込書からチーム情報を抽出しました: " & vFilePath
            wsTeamList.Cells(vCurrentRow, 1).Value = vFilePath
        End If

        vCurrentRow = vCurrentRow + vExtractedCount
    Next vFilePath
End Sub

Private Sub InitializeTeamListSheet( _
        ByVal TeamListSheet As Worksheet, _
        ByVal ColumnList As Collection _
)
    TeamListSheet.Name = "TeamList" & "_" & Format(Now, "yyyymmdd_hhmmss")

    Dim vTitleArray() As Variant
    ReDim vTitleArray(1 To ColumnList.Count + 1)

    Dim vIndex As Long: vIndex = 1
    Dim vColumnConfig As TeamListColumnConfiguration

    vTitleArray(vIndex) = "Source File"
    For Each vColumnConfig In ColumnList
        vTitleArray(vIndex + 1) = vColumnConfig.Title
        vIndex = vIndex + 1
    Next vColumnConfig

    With TeamListSheet.Range(TeamListSheet.Cells(1, 1), TeamListSheet.Cells(1, UBound(vTitleArray)))
        .Value = vTitleArray
        .Font.Bold = True
        .Interior.Color = RGB(192, 240, 200)
    End With
End Sub

Private Function ExtractTeamsFromEntryBook( _
        ByVal EntryBookFilePath As String, _
        ByVal TeamListSheet As Worksheet, _
        ByRef CurrentRow As Long, _
        ByVal ColumnList As Collection _
) As Long
    Dim wbEntryBook As Workbook
    
    On Error Resume Next
    Set wbEntryBook = Workbooks.Open(EntryBookFilePath, ReadOnly:=True)
    On Error GoTo 0

    If wbEntryBook Is Nothing Then
        Debug.Print "申込書を開けません: " & EntryBookFilePath
        ExtractTeamsFromEntryBook = 0
        Exit Function
    End If

    Dim wsTeamSheet As Worksheet
    On Error Resume Next
    Set wsTeamSheet = wbEntryBook.Sheets(ENTRY_BOOK_TEAM_SHEET_NAME)
    On Error GoTo 0

    If wsTeamSheet Is Nothing Then
        Debug.Print "チームシートが見つかりません: " & EntryBookFilePath
        wbEntryBook.Close SaveChanges:=False
        ExtractTeamsFromEntryBook = 0
        Exit Function
    End If

    Dim vColumnConfig As TeamListColumnConfiguration
    Dim vColumnIndex As Long: vColumnIndex = 2
    For Each vColumnConfig In ColumnList
        Dim vCellValue As Variant
        vCellValue = wsTeamSheet.Range(vColumnConfig.Address).Value

        Select Case vColumnConfig.DataType
        Case "文字列"
            TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = "'" & CStr(vCellValue)
        Case "数値"
            If IsNumeric(vCellValue) Then
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CDec(vCellValue)
            Else
                Debug.Print "数値変換エラー: " & vCellValue & " in " & EntryBookFilePath
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CVErr(xlErrValue)
            End If
        Case "日付"
            If IsDate(vCellValue) Then
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CDate(vCellValue)
                TeamListSheet.Cells(CurrentRow, vColumnIndex).NumberFormat = "yyyy/mm/dd"
            Else
                Debug.Print "日付変換エラー: " & vCellValue & " in " & EntryBookFilePath
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CVErr(xlErrValue)
            End If
        Case "整数"
            If IsNumeric(vCellValue) Then
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CLng(vCellValue)
            Else
                Debug.Print "整数変換エラー: " & vCellValue & " in " & EntryBookFilePath
                TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CVErr(xlErrValue)
            End If
        Case Else
            ' 設定されていないのは読み取りされないので無視
            TeamListSheet.Cells(CurrentRow, vColumnIndex).Value = CVErr(xlErrNA)
        End Select
        vColumnIndex = vColumnIndex + 1
    Next vColumnConfig

    wbEntryBook.Close SaveChanges:=False

    ExtractTeamsFromEntryBook = 1
End Function
