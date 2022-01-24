
Sub AddDataPage()
    Dim header As Range
    Dim headerInputs As New Collection
    Dim sheetName As String
    sheetName = "VerticalStudentData"
    Set header = Range("A1:O1")
    header = Array("Name", "Asset", "Cost", "Student", "ID", "Grade", "Due", "Date", "Item", "Birthday", _
                    "Barcode", "Condition", "Comment", "School", "chk-out")
      CreateNewSheet sheetName
     header.Copy Destination:=Range("A1")
     
End Sub

Sub CreateNewSheet(sheetName As String)
Dim sh As Worksheet
    For Each sh In Worksheets
        If Application.Proper(sh.Name) = Application.Proper(sheetName) Then
            Sheets(sheetName).Activate
            Exit Sub
        End If
    Next
     Sheets.Add(After:=Sheets(1)).Name = sheetName
End Sub
