'Sub CreateNewEntry()
'
'End Sub
'




Sub selectData()
    Dim NumberOfEntries As Integer
    Dim UsedRange As Range
    Set UsedRange = ActiveSheet.UsedRange
    Set UsedRange = UsedRange.Resize(rowSize:=UsedRange.Rows.Count - 2)
    Set UsedRange = UsedRange.Offset(rowoffset:=2)
   UsedRange.Activate
   MsgBox GrabBirthdate(UsedRange)

'

'
End Sub


'Sub createArray()
'    Dim data As String(entries,)
'
'End Sub

Function GrabName(source As Range) As String
    GrabName = source.Cells(1, 1).Value
End Function

Function GrabID(source As Range) As String
    GrabID = source.Cells(1, 7).Value
End Function

Function GrabGrade(source As Range) As String
    GrabGrade = source.Cells(1, 10).Value
End Function

Function GrabBirthdate(source As Range) As String
    GrabBirthdate = source.Cells(1, 14).Value
End Function

Function GrabAssets(source As Range) As Variant
    Dim tempArray As Variant
    Dim AssetColumn As Range
    Set AssetColumn = source.Columns(1).Resize(rowSize:=source.Columns(1).Rows.Count - 2)
'    Set issueDate = issueDate.Offset(rowoffset:=2)
'    Set issueDate = issueDate.SpecialCells(xlCellTypeConstants)
'    NumberOfEntries = issueDate.Count
End Function
'Sub CreateNewEntry()
'
'End Sub
'




Sub selectData()
    Dim NumberOfEntries As Integer
    Dim UsedRange As Range
    Set UsedRange = ActiveSheet.UsedRange
    Set UsedRange = UsedRange.Resize(rowSize:=UsedRange.Rows.Count - 2)
    Set UsedRange = UsedRange.Offset(rowoffset:=2)
   UsedRange.Activate
   MsgBox GrabBirthdate(UsedRange)

'

'
End Sub


'Sub createArray()
'    Dim data As String(entries,)
'
'End Sub

Function GrabName(source As Range) As String
    GrabName = source.Cells(1, 1).Value
End Function

Function GrabID(source As Range) As String
    GrabID = source.Cells(1, 7).Value
End Function

Function GrabGrade(source As Range) As String
    GrabGrade = source.Cells(1, 10).Value
End Function

Function GrabBirthdate(source As Range) As String
    GrabBirthdate = source.Cells(1, 14).Value
End Function

Function GrabAssets(source As Range) As Variant
    Dim tempArray As Variant
    Dim AssetColumn As Range
    Set AssetColumn = source.Columns(1).Resize(rowSize:=source.Columns(1).Rows.Count - 2)
'    Set issueDate = issueDate.Offset(rowoffset:=2)
'    Set issueDate = issueDate.SpecialCells(xlCellTypeConstants)
'    NumberOfEntries = issueDate.Count
End Function
