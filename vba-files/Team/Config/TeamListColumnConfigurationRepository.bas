Attribute VB_Name = "TeamListColumnConfigurationRepository"
'namespace=vba-files/Team/Config
Option Explicit
Option Private Module

Public Function ReadAll() As Collection
    Dim vRows As Collection: Set vRows = New Collection
    
    Dim rRecordRow As Range
    For Each rRecordRow In shtConfiguration.Range("tblFields").ListRows
        Dim vRecord As TeamListColumnConfiguration
        Set vRecord = ReadRecord(rRecordRow)
        If Not vRecord Is Nothing Then
            vRows.Add vRecord, vRecord.Title
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
