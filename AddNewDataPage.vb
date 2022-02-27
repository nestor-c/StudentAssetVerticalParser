Sub AddDataPage()
    Dim header As Range
    Dim headerInputs As New Collection
    Dim sheetName As String
    sheetName = "VerticalStudentData"
    Set header = Range("A1:E1")
    header = Array("Name", "ID", "Grade", "Issue Date", "Asset Description")
    CreateNewSheet sheetName
    header.copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
     
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


