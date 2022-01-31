Sub selectData()
    Set usedRange = ActiveSheet.usedRange
    Set usedRange = usedRange.Resize(rowSize:=usedRange.Rows.Count - 2)
    Set usedRange = usedRange.Offset(Rowoffset:=2)
    usedRange.Activate
End Sub

    Dim usedRange As Range
