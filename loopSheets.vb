Function loopSheets()
    Dim studentData As Variant
    AddDataPage.AddDataPage
    For Each Sheet In ActiveWorkbook.sheets
        If Sheet.index = 1 Or Sheet.index = 2 Then
            GoTo NextSheet
        Else
            Sheet.Activate
            studentData = createEntry.createStudentData
            sheets(1).Activate
            For i = 1 To UBound(studentData, 1)
                rows(2).Insert xlShiftDown
                copyVariantRow studentData, i
            Next i

        End If
NextSheet:
    Next Sheet
End Function

Sub copyVariantRow(myVariant As Variant, index)
    For i = 1 To UBound(myVariant, 2)
        Cells(2, i).Value = myVariant(index, i)
        If (i = 6) Then
            Cells(2, i).NumberFormat = "00000"
        End If

    Next i
End Sub

