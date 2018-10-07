Attribute VB_Name = "KeepCols"
Sub deleteIrrelevantColumns()
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String

    currentColumn = 1
    While currentColumn <= ActiveSheet.UsedRange.Columns.Count
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "Advertiser Name" Then keepColumn = True
        If columnHeading = "Advertiser ID" Then keepColumn = True
        If columnHeading = "Sales Representative(s)" Then keepColumn = True
        If columnHeading = "Account Manager" Then keepColumn = True
        If columnHeading = "Campaign ID" Then keepColumn = True
        If columnHeading = "Campaign Name" Then keepColumn = True
        If columnHeading = "Campaign Start Date" Then keepColumn = True
        If columnHeading = "Campaign End Date" Then keepColumn = True
        If columnHeading = "CPL" Then keepColumn = True
        If columnHeading = "Servability Status" Then keepColumn = True
        If columnHeading = "Campaign Balance" Then keepColumn = True
        If columnHeading = "Current Servable Balance" Then keepColumn = True


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

End Sub

