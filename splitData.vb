Sub SplitData()
'Updateby20140617
Dim WorkRng As Range
Dim xRow As Range
Dim SplitRow As Integer

On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
SplitRow = Application.InputBox("Split Row Num", xTitleId, 5, Type:=1)

Set xRow = WorkRng.Rows(1)
Application.ScreenUpdating = False
For i = 1 To WorkRng.Rows.Count Step SplitRow
    resizeCount = SplitRow
    If (WorkRng.Rows.Count - xRow.Row + 1) < SplitRow Then resizeCount = WorkRng.Rows.Count - xRow.Row + 1
    xRow.Resize(resizeCount).Copy
    Application.Worksheets.Add after:=Application.Worksheets(Application.Worksheets.Count)
    Application.ActiveSheet.Range("A1").PasteSpecial
    Set xRow = xRow.Offset(SplitRow)
Next
Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub