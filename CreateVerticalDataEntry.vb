Sub selectStudentData()
    Dim UsedRange As Range
    Set UsedRange = ActiveSheet.UsedRange
    Set UsedRange = UsedRange.Resize(rowSize:=UsedRange.Rows.count - 2)
    Set UsedRange = UsedRange.Offset(rowoffset:=2)
End Sub

Function createEntry(source As Range) As Range
    Dim Entry As Range
    Dim Name As String
    Dim StudentID As String
    Dim Grade As String
    Dim Birthdate As String
    
    
    Set Name = GrabName(source)
    Set StudentID = GrabID(source)
    Set Grade = GrabGrade(source)
    Set Birthdate = GrabBirthdate(source)
    Set Entry = Array(Name, StudentID, Grade, Birthdate)
    
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
    ReDim dateArr(1 To IssueDates.count)
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
    ReDim assetArr(1 To assetColumn.count)
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
    ReDim assetArr(1 To assetColumn.count)
    i = 0
    For Each assetEle In assetColumn
        i = i + 1
        assetArr(i) = assetEle
    Next
    GrabAssets = assetArr
End Function


