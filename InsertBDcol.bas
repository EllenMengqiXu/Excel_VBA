Attribute VB_Name = "InsertBDcol"
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

