Attribute VB_Name = "combination"
Sub CampaignEndingSoon()

    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim StartDate As Long, EndDate As Long
    StartDate = DateSerial(Year(Date), Month(Date), Day(Date))
    EndDate = DateSerial(Year(Date), Month(Date), Day(Date) + 4)
    
    'Multiple Filters
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter Field:=16, Criteria1:=">0"
    .UsedRange.AutoFilter Field:=21, Criteria1:="Servable"
    
    'Filter Date Range from today(current day in excel function is TODAY(),and it is date in VBA)
    .UsedRange.AutoFilter Field:=9, Criteria1:=">=" & StartDate, Criteria2:="<=" & EndDate
    End With
    
    'SortingAscending
    Range("A1").CurrentRegion.Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlYes
    
    'Keep Relevant Columns
    currentColumn = 1
    While currentColumn <= ActiveSheet.UsedRange.Columns.Count
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "lASTNAME" Then keepColumn = True
        If columnHeading = "FIRSTNAME" Then keepColumn = True
        If columnHeading = "STUDENT ID" Then keepColumn = True
        If columnHeading = "AGE" Then keepColumn = True
        If columnHeading = "GENDER" Then keepColumn = True
        If columnHeading = "TUITION" Then keepColumn = True
        If columnHeading = "LIVING EXPENSE" Then keepColumn = True
        If columnHeading = "BALANCE" Then keepColumn = True

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
     'SAVE FILE AS .XLS FORMAT AS CSV FORMAT CANNOT SAVE EXCEL FUNCTION.
    ActiveWorkbook.SaveAs FileFormat:=xlWorkbookNormal
    
End Sub
