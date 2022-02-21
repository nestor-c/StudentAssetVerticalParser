Sub SplitPagesToSheets()
    Dim FirstRow As Range
    Dim PageBreaks As HPageBreaks
    Dim pb As HPageBreak
    Dim BottomRight As Range
    Dim TopLeft As Range
    Dim PageWorkArea As Range
    Set TopLeft = Range("A1")
    Set PageBreaks = ActiveSheet.HPageBreaks

    For Each pb In PageBreaks
            Set BottomRight = BottomRightCell(pb)
            Set PageWorkArea = Sheets(1).Range(TopLeft, BottomRight)
            Set TopLeft = TopLeftCell(pb)
            CopyDataToNewSheet PageWorkArea
            Sheets(1).Activate
    Next
            CopyDataToNewSheet Range(TopLeft, Sheets(1).UsedRange.SpecialCells(xlCellTypeLastCell))
End Sub

Function TopLeftCell(pb As HPageBreak) As Range
    Dim FirstColumn As Integer
    FirstColumn = 1
    Set TopLeftCell = Cells(pb.Location.Row, FirstColumn)
End Function

Function BottomRightCell(pageBreak As HPageBreak) As Range
    Dim LastColumn As Integer
    Dim LastRow As Range
    LastColumn = ActiveSheet.UsedRange.Columns.Count
    Set LastRow = pageBreak.Location.Offset(rowoffset:=-1)
    Set BottomRightCell = Cells(LastRow.Row, LastColumn)
End Function

Sub CopyDataToNewSheet(data As Range)
    Dim NewSheet As Worksheet
    Set NewSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    data.Copy (Sheets(NewSheet.Index).Range("A1"))
End Sub
