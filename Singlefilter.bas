Attribute VB_Name = "Singlefilter"
Sub singlefilter_PR()
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter field:=13, Criteria1:="<=0"
    
    End With

End Sub

