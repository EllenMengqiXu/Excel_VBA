Attribute VB_Name = "sortDesceding"
Sub sort_PR()
    
    Range("A1").CurrentRegion.Sort Key1:=Range("M1"), Order1:=xlDescending, Header:=xlYes

End Sub

