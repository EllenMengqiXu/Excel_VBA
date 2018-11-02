Attribute VB_Name = "CPV"
Sub copytopstvalues()

    lastRow = ActiveSheet.Range("Y" & Rows.Count).End(xlUp).Row
    ActiveSheet.Range("B12:Z" & lastRow).Copy
    Worksheets("Sheet4").Cells.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
End Sub

