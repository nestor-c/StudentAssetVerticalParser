Sub addStudentData()
    Dim workingData As Range
    Dim newEntry As Variant
    Dim assetSize As Integer
    Dim assets As Variant
    Dim nextRow As Range
    ' Select student data
    Set workingData = selectStudentData()
    ' Create new entry based on data
    newEntry = createEntry(workingData)
    'Get the next available row and resize to length appropriate for data
    Set nextRow = newRow(newEntry, Worksheets("VerticalStudentData"))
    assets = GrabAssets(workingData)
    assetSize = returnVariantSize(assets)

    For counter = 1 To assetSize
        nextRow = newRow(newEntry, Worksheets("VerticalStudentData"))
        nextRow = newEntry
        Cells(nextRow.Row, nextRow.Columns.Count + 1) = assets(counter)
    Next
End Sub
'
'


Function selectStudentData() As Range
    Dim UsedRange As Range
    Dim newRow As Range
    Dim test As Variant
    Set UsedRange = ActiveSheet.UsedRange
    Set UsedRange = UsedRange.Resize(rowSize:=UsedRange.Rows.Count - 2) _
                    .Offset(rowoffset:=2)
    Set selectStudentData = UsedRange
End Function

Function newRow(entry As Variant, entrySheet As Worksheet) As Range
  Set newRow = Cells(entrySheet.UsedRange.Rows.Count + 1, 1) _
        .Resize(ColumnSize:=returnVariantSize(entry))

End Function

Function returnVariantSize(source As Variant) As Integer
    Dim i As Integer
    Dim ele As Variant
    i = 0
    For Each ele In source
        i = i + 1
    Next
    returnVariantSize = i
End Function

Private Function createEntry(source As Range) As Variant
    Dim entryArr As Variant
    Dim tempArr(1 To 4) As String

    ReDim entryArr(1 To 4)

    tempArr(1) = GrabName(source)
    tempArr(2) = GrabID(source)
    tempArr(3) = GrabGrade(source)
    tempArr(4) = GrabBirthdate(source)

    For i = 1 To 4 Step 1
        entryArr(i) = tempArr(i)
    Next i

    createEntry = entryArr
End Function

Private Function GrabName(source As Range) As String
    GrabName = source.Cells(1, 1).Value
End Function

Private Function GrabID(source As Range) As String
    GrabID = source.Cells(1, 7).Value
End Function

Private Function GrabGrade(source As Range) As String
    GrabGrade = source.Cells(1, 10).Value
End Function

Private Function GrabBirthdate(source As Range) As String
    GrabBirthdate = source.Cells(1, 14).Value
End Function

Private Function GrabIssueDate(source As Range) As Variant
    '---The array to be returned
    Dim dateArr As Variant
    '---
    '---Grab ranges of issue date values---
    Dim IssueDates As Range
    Set IssueDates = source.Offset(rowoffset:=2)
    Set IssueDates = IssueDates.Columns(6)
    Set IssueDates = IssueDates.SpecialCells(xlCellTypeConstants)
'---Create contigous Range with IssueDates---
    Dim rangeEle As Range
    ReDim dateArr(1 To IssueDates.Count)
    Dim i As Integer
    i = 0
    For Each rangeEle In IssueDates
        i = i + 1
        dateArr(i) = rangeEle
    Next
    GrabIssueDate = dateArr
End Function

' Returns array containing every asset range
Function GrabAssets(source As Range) As Variant
    Dim assetArr As Variant
    Dim i As Integer
    Dim assetEle As Range
    Dim assetColumn As Range
    Set assetColumn = source.Columns(1)
    Set assetColumn = assetColumn.Offset(rowoffset:=2)
    Set assetColumn = assetColumn.SpecialCells(xlCellTypeConstants)
    ReDim assetArr(1 To assetColumn.Count)
    i = 0
    For Each assetEle In assetColumn
        i = i + 1
        assetArr(i) = assetEle
    Next
    GrabAssets = assetArr
End Function

' Returns array containing every barcode range
Function GrabBarcodes(source As Range) As Variant
    Dim barcodeArr As Variant
    Dim i As Integer
    Dim barcodeEle As Range
    Dim assetColumn As Range
    Set assetColumn = source.Columns(1)
    Set assetColumn = assetColumn.Offset(rowoffset:=2)
    Set assetColumn = assetColumn.SpecialCells(xlCellTypeConstants)
    ReDim assetArr(1 To assetColumn.Count)
    i = 0
    For Each assetEle In assetColumn
        i = i + 1
        assetArr(i) = assetEle
    Next
    GrabAssets = assetArr
End Function





