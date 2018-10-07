Attribute VB_Name = "multifilters"
Sub multifilter_PR()

    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter field:=8, Criteria1:=">" & Application.EoMonth(Now, -1), Criteria2:="<=" & Application.EoMonth(Now, 0)
    .UsedRange.AutoFilter field:=10, Criteria1:="Servable"
    
    End With

End Sub

