Attribute VB_Name = "TeamListColumnConfigRepository"
'namespace=vba-files/Team/Config
Option Explicit
Option Private Module

Public Function ReadAll() As Collection
    Dim vRows As Collection: Set vRows = New Collection
    Dim vShtConfigIndices As Dictionary: Set vShtConfigIndices = New Dictionary
    
    Dim lsoConfig As ListObject: Set lsoConfig = shtConfiguration.ListObjects("tblFields")

    If lsoConfig Is Nothing Then
        Set ReadAll = vRows
        Exit Function
    End If

    Dim rConfigurationRecords As Range: Set rConfigurationRecords = lsoConfig.DataBodyRange

    If rConfigurationRecords Is Nothing Then
        Set ReadAll = vRows
        Exit Function
    End If

    Dim rRecordRow As Range
    For Each rRecordRow In rConfigurationRecords.Rows
        Dim vRecord As TeamListColumnConfiguration: Set vRecord = ReadRecord(rRecordRow)
        If Not vRecord Is Nothing Then
            Dim vCollectionIndex As String: vCollectionIndex = vRecord.Title
            If vShtConfigIndices.Exists(vCollectionIndex) Then
                Debug.Print "列名が重複しています。列名：" & vCollectionIndex & " 行番号：" & rRecordRow.Row
                vShtConfigIndices(vCollectionIndex) = vShtConfigIndices(vCollectionIndex) + 1
                vRows.Add vRecord, vCollectionIndex & "_" & vShtConfigIndices(vCollectionIndex)
            Else
                vShtConfigIndices.Add vCollectionIndex, 0
                vRows.Add vRecord, vCollectionIndex
            End If
        End If
    Next rRecordRow
    
    Set ReadAll = vRows
End Function

Private Function ReadRecord(RecordRow As Range) As TeamListColumnConfiguration
    Dim vCells As Variant: vCells = RecordRow.Value

    If vCells(1, 1) = "" Then
        Set ReadRecord = Nothing
        Exit Function
    End If

    With New TeamListColumnConfiguration
        .Initialize vCells(1, 1), vCells(1, 2), vCells(1, 3), vCells(1, 4)
        Set ReadRecord = .Self
    End With
End Function
