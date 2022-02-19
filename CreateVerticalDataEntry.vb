Sub selectStudentData()
    Dim UsedRange As Range
    Dim NewRow As Range
    Dim test As Variant
    Set UsedRange = ActiveSheet.UsedRange
    Set UsedRange = UsedRange.Resize(rowSize:=UsedRange.Rows.Count - 2)
    Set UsedRange = UsedRange.Offset(rowoffset:=2)
    Set NewRow = Cells(ActiveSheet.UsedRange.Rows.Count + 1, 1)
    Set NewRow = NewRow.Resize(ColumnSize:=returnVariantSize(createEntry(UsedRange)))
    NewRow = createEntry(UsedRange)
    
End Sub

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





