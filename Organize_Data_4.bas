Attribute VB_Name = "PausedDayTillToday"
Sub PDTT()

    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute this macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
    
    'set up the begging of current month and today's date
        StartDate = DateSerial(Year(Date), Month(Date), 1)
        EndDate = DateSerial(Year(Date), Month(Date), Day(Date))
    'filter record whose not comfortable and date range between the first day of current month and today
        With ActiveSheet
        .AutoFilterMode = False
        .UsedRange.AutoFilter
        .UsedRange.AutoFilter Field:=14, Criteria1:="Not comfortable"
        .UsedRange.AutoFilter Field:=11, Criteria1:=">=" & StartDate, Criteria2:="<=" & EndDate
        End With
    
    'sort data as from oldest to newest
        Range("A1").CurrentRegion.Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
    
    'copy data and paste to a new sheet and rename it as L1 Value
        Set wh = Worksheets(ActiveSheet.Name)
        ActiveSheet.Copy After:=Worksheets(Sheets.Count)
        If wh.Range("A1").Value <> "" Then
        ActiveSheet.Name = wh.Range("L1").Value
        End If
        wh.Activate
    
    'move to the new sheet, which next to the current sheet
        Worksheets(ActiveSheet.Index + 1).Select
    
    'highlighted column L whose value counted
        lastRow = ActiveSheet.Range("L" & Rows.Count).End(xlUp).Row
        Range("L2:L" & lastRow).Interior.Color = vbYellow
    
    
        Case vbNo
        GoTo Quit:
    End Select

Quit:

End Sub

