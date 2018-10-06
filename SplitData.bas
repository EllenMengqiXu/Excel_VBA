Attribute VB_Name = "Split_PR"
Sub SplitData()
    Const AMCol = "D"
    Const HeaderRow = 1
    Const FirstRow = 2
    Dim SrcSheet As Worksheet
    Dim TrgSheet As Worksheet
    Dim SrcRow As Long
    Dim LastRow As Long
    Dim TrgRow As Long
    Dim AM As String
    Application.ScreenUpdating = False
    Set SrcSheet = ActiveSheet
    LastRow = SrcSheet.Cells(SrcSheet.Rows.Count, AMCol).End(xlUp).Row
    For SrcRow = FirstRow To LastRow
        AM = SrcSheet.Cells(SrcRow, AMCol).Value
        Set TrgSheet = Nothing
        On Error Resume Next
        Set TrgSheet = Worksheets(AM)
        On Error GoTo 0
        If TrgSheet Is Nothing Then
            Set TrgSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            TrgSheet.Name = AM
            SrcSheet.Rows(HeaderRow).Copy Destination:=TrgSheet.Rows(HeaderRow)
        End If
        TrgRow = TrgSheet.Cells(TrgSheet.Rows.Count, AMCol).End(xlUp).Row + 1
        SrcSheet.Rows(SrcRow).Copy Destination:=TrgSheet.Rows(TrgRow)
    Next SrcRow
    Application.ScreenUpdating = True
End Sub
