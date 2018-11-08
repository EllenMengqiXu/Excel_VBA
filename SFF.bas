Attribute VB_Name = "SFF"
Sub SFF()
    
    Dim x As Long
    Dim d As String
      
    'copy filtered records
    Workbooks("Fruit.xlsx").Worksheets("apple").Range("A2").CurrentRegion.Copy
    'create a new workbook
    Set NewBook = Workbooks.Add
    'paste filtered records to the new workbook
    NewBook.Worksheets("Sheet1").Paste
        
    'define the location of the new Workbook
    Path = "C:\Users\exu\Desktop\Forecast\1108\Sent\"
    'filename comes from D2
    filename = Range("D2")
    'save the new workbook under the defined path and file name
    NewBook.saveas filename:=Path & filename & "_apple_" & "_" & Format(Now(), "MM-DD-YY") & ".xls", FileFormat:=xlNormal
    
    'define column N equals to the sum of column R through column AU
    Application.ScreenUpdating = False

        x = Range("A" & Rows.Count).End(xlUp).Row
        d = "=SUM(RC[4]:RC[33])"
        Range("N2").Resize(x - 1).FormulaR1C1 = d

    Application.ScreenUpdating = True
    
    'how to find the RC code?
    'open a new sheet, go to N2 and input formula "=sum(R2:AU2)"
    'go to file --> options --> formula, and check R1C1 referene style
    'click ok, then it will show you the R1C1 reference for the above formula
    
    'save file again
    ActiveSheet.Save

End Sub
