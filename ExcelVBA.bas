Attribute VB_Name = "Excel_VBA"
Sub deleteIrrelevantColumns()
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String

    currentColumn = 1
    While currentColumn <= ActiveSheet.UsedRange.Columns.Count
        columnHeading = ActiveSheet.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "Name" Then keepColumn = True
        If columnHeading = "ID" Then keepColumn = True
        If columnHeading = "SR" Then keepColumn = True
        If columnHeading = "AM" Then keepColumn = True
        If columnHeading = "C_ID" Then keepColumn = True
        If columnHeading = "C_Name" Then keepColumn = True
        If columnHeading = "Start Date" Then keepColumn = True
        If columnHeading = "End Date" Then keepColumn = True
        If columnHeading = "CPL" Then keepColumn = True
        If columnHeading = "Active" Then keepColumn = True
        If columnHeading = "Balance" Then keepColumn = True
        If columnHeading = "Current Active Balance" Then keepColumn = True


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
Sub InsertBudgetDiff()

Dim x As Long
Dim d As String

Application.ScreenUpdating = False

Range("M1").EntireColumn.Insert shift:=xlToRight
Range("M1").Value = "Budget Difference"
x = Range("A" & Rows.Count).End(xlUp).Row
d = "=K2-L2"
Range("M2").Resize(x - 1).Formula = d

Application.ScreenUpdating = True


End Sub

Sub multifilter_PR()
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter field:=8, Criteria1:=">" & Application.EoMonth(Now, -1), Criteria2:="<=" & Application.EoMonth(Now, 0)
    .UsedRange.AutoFilter field:=10, Criteria1:="Active"

    End With

End Sub

Sub finalsort_PR()

    Range("A1").CurrentRegion.Sort Key1:=Range("M1"), Order1:=xlDescending, Header:=xlYes

End Sub
