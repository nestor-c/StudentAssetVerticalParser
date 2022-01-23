Sub AssetsVerticalFormat()
    
    AddDataPage
        AddHeader
    AddShDataToDataPg

End Sub

Sub AddDataPage()
    Dim header As Range
    Dim headerInputs As New Collection
    Set header = Range("A1:M1")
    headerInputs.Add ("Name")
    headerInputs.Add ("Asset")
    headerInputs.Add ("Cost")
    headerInputs.Add ("Student")
    headerInputs.Add ("ID")
    headerInputs.Add ("Grade")
    headerInputs.Add ("Due")
    headerInputs.Add ("Date")
    headerInputs.Add ("Item")
    headerInputs.Add ("Birthday")
    headerInputs.Add ("Barcode")
    headerInputs.Add ("Condition")
    headerInputs.Add ("Comment")
    headerInputs.Add ("School")
    headerInputs.Add ("chk-out")
    Dim myInputs As String
    
    
    Sheets.Add(After:=Sheets(1)).Name = "VerticalStudentAssetsData"
    Dim
    For Each myInput In headerInputs
        
        header.Columns(1).Value = myInput
    Next
    
End Sub
