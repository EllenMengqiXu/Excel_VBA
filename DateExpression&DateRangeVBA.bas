Attribute VB_Name = "Second_DateFilter"
'Advanced Excel Funtion to express date:
'AS of 10/11/2018
'year: TEXT(TODAY(),"YYYY") --> 2018
'year: TEXT(TODAY(),"YY") --> 18
'month: TEXT(TODAY(),"MM") --> 10
'month: TEXT(TODAY(),"MMMM") --> October
'day: TEXT(TODAY(),"DD") --> 10
'day: TEXT(TODAY(),"DDDD") --> Thursday
'month/day/year: TEXT(TODAY(),"MM/DD/YYYY") --> 10/11/2018

'Use the below programm to filter a data range from today to next 5 days which is 6 days.
Sub AboutDate()

    Dim StartDate As Long, EndDate As Long
    StartDate = DateSerial(Year(Date), Month(Date), Day(Date))
    EndDate = DateSerial(Year(Date), Month(Date), Day(Date) + 5)
    
    With ActiveSheet
    .AutoFilterMode = False
    .UsedRange.AutoFilter
    .UsedRange.AutoFilter Field:=9, Criteria1:=">=" & StartDate, Criteria2:="<=" & EndDate
    
    End With
    
End Sub
