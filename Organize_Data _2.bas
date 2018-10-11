Attribute VB_Name = "Combination"
Sub BDG()

    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim x As Long
    Dim d As String
    
    'KEEP RELEVANT COLUMNS
    currentColumn = 1
    While currentColumn <= ActiveSheet.UsedRange.Columns.Count
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "Apple" Then keepColumn = True
        If columnHeading = "Banana" Then keepColumn = True
        If columnHeading = "Car" Then keepColumn = True
        If columnHeading = "Dog" Then keepColumn = True
        If columnHeading = "Eifel Tower" Then keepColumn = True
        If columnHeading = "Fog" Then keepColumn = True
        If columnHeading = "Gaggle" Then keepColumn = True
        If columnHeading = "Happy" Then keepColumn = True
        If columnHeading = "Ice Cream" Then keepColumn = True
        If columnHeading = "Joker" Then keepColumn = True
        If columnHeading = "Kangaroo" Then keepColumn = True
        If columnHeading = "Limo" Then keepColumn = True


        If keepColumn Then
        'IF YES THEN SKIP TO THE NEXT COLUMN,
            currentColumn = currentColumn + 1
        Else
        'IF NO DELETE THE COLUMN
            ActiveSheet.Columns(currentColumn).Delete
        End If

        'LASTLY AN ESCAPE IN CASE THE SHEET HAS NO COLUMNS LEFT
        If (ActiveSheet.UsedRange.Address = "$A$1") And (ActiveSheet.Range("$A$1").Text = "") Then Exit Sub
    Wend
    'Insert a new column whose head called "Bike Dog" and cell formuna as K2-L2
    Application.ScreenUpdating = False

    Range("M1").EntireColumn.Insert shift:=xlToRight
    Range("M1").Value = "Bike Dog"
    x = Range("A" & Rows.Count).End(xlUp).Row
    d = "=K2-L2"
    Range("M2").Resize(x - 1).Formula = d

    Application.ScreenUpdating = True
    'FILTER COLUMN 8 AS DATE OF THIS MONTH AND COLUMN 10 WHOSE VALUE IS COMFORTABLE
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter field:=8, Criteria1:=">" & Application.EoMonth(Now, -1), Criteria2:="<=" & Application.EoMonth(Now, 0)
    .UsedRange.AutoFilter field:=10, Criteria1:="COMFORTABLE"
    
    End With
    'SORT DATA BASED ON M COLUMN IN DESCENDING (FROM OLDEST TO NEWEST)
    Range("A1").CurrentRegion.Sort Key1:=Range("M1"), Order1:=xlDescending, Header:=xlYes
    'SAVE FILE AS .XLS FORMAT AS CSV FORMAT CANNOT SAVE EXCEL FUNCTION.
    ActiveWorkbook.SaveAs FileFormat:=xlWorkbookNormal

End Sub
