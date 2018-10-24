Attribute VB_Name = "createSheet"
Sub CreateSheet()

    Dim WS As Worksheet
    Set WS = Sheets.Add
    Sheets.Add.Name = "test"
    
End Sub
