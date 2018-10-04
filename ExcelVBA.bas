EXCEL VBA
Tools —>macro—>Visual Basic Editor
Step into code —>command+shift+I
Date —> #06/02/2012#

* Remove rows when the first column cell is blank (manually: select entire column and press ctrl+g —> special: blank —>right click delete —> entire row

Sub conditional_remove_blankrows()

    On Error Resume Next
    Columns("A").SpecialCells(xlBlanks).EntireRow.Delete
End Sub

* Remove blank row

Sub remove_blank_row()

Dim rng
Set rng = Nothing

For Each i In Range("A1:A1178")
    If Application.CountA(i.EntireRow) = 0 Then
        If rng Is Nothing Then
            Set rng = i
        Else
            Set rng = Union(rng, i)
        End If
    End If
Next i

rng.EntireRow.Delete

End Sub

* Remove duplicates and leave the unique value (in one column)

Sub Remove_Duplicates()
    
    Sheets("Sheet1").Range("A1:A10").RemoveDuplicates Columns:=1, Header:=xlYes
    
End Sub

* Remove duplicates and leave the unique value (in all columns)

Sub Remove_Duplicates()
    
    Sheets("Sheet2").Range("A1:F10").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
    
End Sub

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

