Attribute VB_Name = "EntrySheetAggregateController"
'namespace=vba-files/Entry
Option Explicit
Option Private Module

Private Enum EntryListColumn
    SourceFilePath = 1
    TeamNameFull 
    TeamNameShort
    AthleteIndex
    EntryIndex
    EntryCount
    Gender
    Bib
    AthleteName
    AthletePhonetic
    Age
    TeamName
    TeamPhonetic
    Region
    BirthDate
    EntryEvent
    Record
    Comment
    Fee
    EntryEventId
    EntryEventType
    EntryEventComment
    EntryRecordOrder
    LastColumn                          ' 最後の列を示すためのダミー値
End Enum

Private Enum EntryBookListColumn
    AthleteIndexNum = 1
    EntryIndex
    NotUsed1                            ' 未使用の列
    AthleteIndex
    Gender
    Bib
    AthleteName
    AthletePhonetic
    Age
    TeamName
    Region
    BirthDate
    EntryEvent
    Record
    Comment
    Fee
    EntryEventId
    EntryEventType
    EntryEventComment
    EntryRecordOrder
    LastColumn                          ' 最後の列を示すためのダミー値
End Enum

Private Const MAX_ENTRIES_PER_BOOK As Long = 500
Private Const ENTRY_BOOK_LIST_TOP_ROW As Long = 4
Private Const ENTRY_BOOK_LIST_SHEET_NAME As String = "（自動）エントリー名簿"
Private Const ENTRY_BOOK_TEAM_SHEET_NAME As String = "(1)鑑"

Public Sub AggregateEntryList(ByVal EntryBookFiles As Collection)
    If EntryBookFiles.Count = 0 Then
        Debug.Print "申込書が選択されていません。"
        Exit Sub
    End If

    Dim wsEntryList As Worksheet
    Set wsEntryList = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    InitializeEntryListSheet wsEntryList

    Dim vCurrentRow As Long: vCurrentRow = 2
    Dim vFilePath As Variant
    For Each vFilePath In EntryBookFiles
        Dim vExtractedCount As Long: vExtractedCount = ExtractEntriesFromEntryBook(vFilePath, wsEntryList, vCurrentRow)
        vCurrentRow = vCurrentRow + vExtractedCount
    Next vFilePath

    Debug.Print "エントリーリストの集計が完了しました。合計 " & (vCurrentRow - 2) & " 件のエントリーが集計されました。"
End Sub

Private Sub InitializeEntryListSheet(ByVal EntryListSheet As Worksheet)
    EntryListSheet.Name = "EntryList" & "_" & Format(Now, "yyyymmdd_hhmmss")

    Dim aTitle As Variant: aTitle = Array( _
            "Source File", _
            "チーム名（正式）", _
            "チーム名（略称）", _
            "名簿No.", _
            "種目番号", _
            "エントリー回数", _
            "性別", _
            "Bib", _
            "競技者氏名", _
            "競技者カナ名", _
            "学年", _
            "所属", _
            "所属カナ", _
            "都道府県", _
            "生年月日", _
            "出場種目", _
            "参考記録", _
            "エントリー備考", _
            "参加料", _
            "種目ID", _
            "種目種別", _
            "エントリー情報", _
            "参考記録（数値）" _
    )

    With EntryListSheet.Range(EntryListSheet.Cells(1, 1), EntryListSheet.Cells(1, UBound(aTitle) + 1))
        .Value = aTitle
        .Font.Bold = True
        .Interior.Color = RGB(192, 240, 200)
    End With
End Sub

Private Function ExtractEntriesFromEntryBook( _
        ByVal EntryBookFilePath As String, _
        ByVal EntryListSheet As Worksheet, _
        ByVal StartRow As Long _
) As Long

    Dim wbEntryBook As Workbook

    On Error Resume Next
    Set wbEntryBook = Workbooks.Open(EntryBookFilePath, ReadOnly:=True)
    On Error GoTo 0

    If wbEntryBook Is Nothing Then
        Debug.Print "申込書を開けません: " & EntryBookFilePath
        ExtractEntriesFromEntryBook = 0
        Exit Function
    End If

    Dim wsTeamSheet As Worksheet
    Dim wsEntryListSheet As Worksheet
    On Error Resume Next
    Set wsTeamSheet = wbEntryBook.Worksheets(ENTRY_BOOK_TEAM_SHEET_NAME)
    Set wsEntryListSheet = wbEntryBook.Worksheets(ENTRY_BOOK_LIST_SHEET_NAME)
    On Error GoTo 0

    If wsTeamSheet Is Nothing Then
        Debug.Print "チームシートが見つかりません: " & EntryBookFilePath
        wbEntryBook.Close SaveChanges:=False
        ExtractEntriesFromEntryBook = 0
        Exit Function
    ElseIf wsEntryListSheet Is Nothing Then
        Debug.Print "エントリーリストシートが見つかりません: " & EntryBookFilePath
        wbEntryBook.Close SaveChanges:=False
        ExtractEntriesFromEntryBook = 0
        Exit Function
    End If

    Dim vTeamNameFull As String: vTeamNameFull = wsTeamSheet.Range("E5").Value
    Dim vTeamNameShort As String: vTeamNameShort = wsTeamSheet.Range("E6").Value
    Dim vEntryCount As Dictionary: Set vEntryCount = New Dictionary

    Dim vWroteRowCount As Long
    Dim vCurrentRow As Long: vCurrentRow = StartRow
    Dim vCursorRow As Long

    For vCursorRow = ENTRY_BOOK_LIST_TOP_ROW To ENTRY_BOOK_LIST_TOP_ROW + MAX_ENTRIES_PER_BOOK - 1
        Dim rCursorRow As Range
        Set rCursorRow = wsEntryListSheet.Cells(vCursorRow, 1).Resize(1, EntryBookListColumn.LastColumn)
    
        Dim vWrote As Long
        vWrote = WriteEntryToEntryListSheet(EntryListSheet, vCurrentRow, rCursorRow)
        vWroteRowCount = vWroteRowCount + vWrote
        If vWrote > 0 Then
            Dim vAthleteIndex As String: vAthleteIndex = Trim(rCursorRow.Cells(1, EntryBookListColumn.AthleteIndex).Value)
            If Not vEntryCount.Exists(vAthleteIndex) Then
                vEntryCount.Add vAthleteIndex, 0
            End If

            vEntryCount(vAthleteIndex) = vEntryCount(vAthleteIndex) + 1
            
            EntryListSheet.Cells(vCurrentRow, EntryListColumn.SourceFilePath).Value = EntryBookFilePath
            ' TODO: チーム名の正式名称と略称を取得する参照先をハードコーディングではなく、設定シートから取得するように変更する.
            '       別Issueで対応する.
            EntryListSheet.Cells(vCurrentRow, EntryListColumn.TeamNameFull).Value = vTeamNameFull
            EntryListSheet.Cells(vCurrentRow, EntryListColumn.TeamNameShort).Value = vTeamNameShort

            EntryListSheet.Cells(vCurrentRow, EntryListColumn.EntryCount).Value = vEntryCount(vAthleteIndex)

            vCurrentRow = vCurrentRow + vWrote
        End If
    Next vCursorRow

    wbEntryBook.Close SaveChanges:=False

    Debug.Print "Extracted " & vWroteRowCount & " entries from: " & EntryBookFilePath
    ExtractEntriesFromEntryBook = vWroteRowCount
End Function

Private Function WriteEntryToEntryListSheet( _
        ByVal EntryListSheet As Worksheet, _
        ByVal TargetRow As Long, _
        ByVal SourceRow As Range _
) As Long

    Dim vAthleteIndex As String: vAthleteIndex = Trim(SourceRow.Cells(1, EntryBookListColumn.AthleteIndex).Value)
    If vAthleteIndex = "" Then
        WriteEntryToEntryListSheet = 0
        Exit Function
    End If

    With EntryListSheet.Cells(TargetRow, 1).Resize(1, EntryListColumn.LastColumn - 1)
        .Cells(1, EntryListColumn.AthleteIndex).Value           = SourceRow.Cells(1, EntryBookListColumn.AthleteIndex).Value
        .Cells(1, EntryListColumn.EntryIndex).Value             = SourceRow.Cells(1, EntryBookListColumn.EntryIndex).Value
        .Cells(1, EntryListColumn.Gender).Value                 = SourceRow.Cells(1, EntryBookListColumn.Gender).Value
        .Cells(1, EntryListColumn.Bib).Value                    = "'" & SourceRow.Cells(1, EntryBookListColumn.Bib).Value
        .Cells(1, EntryListColumn.AthleteName).Value            = SourceRow.Cells(1, EntryBookListColumn.AthleteName).Value
        .Cells(1, EntryListColumn.AthletePhonetic).Value        = SourceRow.Cells(1, EntryBookListColumn.AthletePhonetic).Value
        .Cells(1, EntryListColumn.Age).Value                    = "'" & SourceRow.Cells(1, EntryBookListColumn.Age).Value
        .Cells(1, EntryListColumn.TeamName).Value               = SourceRow.Cells(1, EntryBookListColumn.TeamName).Value

        ' TODO: 所属カナの列は、申込書の仕様上、未実装の列のため、空文字を設定する. 別Issueで対応する.
        .Cells(1, EntryListColumn.TeamPhonetic).Value           = ""

        .Cells(1, EntryListColumn.Region).Value                 = SourceRow.Cells(1, EntryBookListColumn.Region).Value
        
        .Cells(1, EntryListColumn.EntryEvent).Value             = SourceRow.Cells(1, EntryBookListColumn.EntryEvent).Value
        .Cells(1, EntryListColumn.Record).Value                 = "'" & SourceRow.Cells(1, EntryBookListColumn.Record).Value
        .Cells(1, EntryListColumn.Comment).Value                = SourceRow.Cells(1, EntryBookListColumn.Comment).Value
        .Cells(1, EntryListColumn.EntryEventId).Value           = SourceRow.Cells(1, EntryBookListColumn.EntryEventId).Value
        .Cells(1, EntryListColumn.EntryEventType).Value         = SourceRow.Cells(1, EntryBookListColumn.EntryEventType).Value
        .Cells(1, EntryListColumn.EntryEventComment).Value      = SourceRow.Cells(1, EntryBookListColumn.EntryEventComment).Value
        .Cells(1, EntryListColumn.EntryRecordOrder).Value       = SourceRow.Cells(1, EntryBookListColumn.EntryRecordOrder).Value

        ' データ変換の必要がある列の処理
        ' 生年月日列は、申込書の仕様上、日付型で入力されることが想定されているが、
        ' Excelの仕様上、日付型で入力されていない場合もあるため、日付型に変換してから設定する.
        If IsDate(SourceRow.Cells(1, EntryBookListColumn.BirthDate).Value) Then
            .Cells(1, EntryListColumn.BirthDate).Value = CDate(SourceRow.Cells(1, EntryBookListColumn.BirthDate).Value)
            .Cells(1, EntryListColumn.BirthDate).NumberFormat = "yyyy/mm/dd"
        Else
            .Cells(1, EntryListColumn.BirthDate).Value = SourceRow.Cells(1, EntryBookListColumn.BirthDate).Value
        End If

        ' エントリー料金列は、申込書の仕様上、数値型で入力されることが想定されているが、
        ' Excelの仕様上、数値型で入力されていない場合もあるため、数値型に変換してから設定する.
        If IsNumeric(SourceRow.Cells(1, EntryBookListColumn.Fee).Value) Then
            .Cells(1, EntryListColumn.Fee).Value = CDec(SourceRow.Cells(1, EntryBookListColumn.Fee).Value)
        Else
            .Cells(1, EntryListColumn.Fee).Value = SourceRow.Cells(1, EntryBookListColumn.Fee).Value
        End If
    End With

    WriteEntryToEntryListSheet = 1
End Function
