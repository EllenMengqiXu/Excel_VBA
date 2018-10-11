Attribute VB_Name = "Combination"
Sub TDR()
    
    'Delete the first row
    ActiveSheet.Range("1:1").Delete
    'Unwrap all active Rows.
    ActiveSheet.Rows.WrapText = False
    'Delete columns
    ActiveSheet.Range("O:P").Delete
    'Sort worksheet baed on column B value by ascending (date from oldest to newest) order.
    Range("A1").CurrentRegion.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
    'Filter column 12 whose value contains (6), * stands for wildcard.
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter field:=12, Criteria1:="(6)*"
    End With
    'Delete filtered Data.
    ActiveSheet.AutoFilter.Range.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
    'Show data after Being Deleted.
    ActiveSheet.ShowAllData
    
End Sub

