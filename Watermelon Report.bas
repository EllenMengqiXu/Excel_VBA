Attribute VB_Name = "CS"
Sub CS()
    Const SPCol = "D"
    Const HeaderRow = 1
    Const FirstRow = 2
    
    Dim LastRow As Long
    Dim r As Range
    Dim sh As Worksheet
    
    Dim SrcSheet As Worksheet
    Dim TrgSheet As Worksheet
    Dim SrcRow As Long
    Dim LR As Long
    Dim TrgRow As Long
    Dim SP As String
    
    Dim LC As Long
    Dim WS_Count As Integer
    Dim I As Integer
    'select multiple data ranges
    Set sh = ActiveSheet
    With sh
        LastRow = Range("A" & Rows.Count).End(xlUp).Row
        Set r = Union(.Range("A2:N" & LastRow), .Range("PL2:PX" & LastRow))
    End With
    'copy selected range and create a new workbook and paste
    r.Copy
    Set NewBook = Workbooks.Add
    NewBook.Worksheets("Sheet1").Paste
    'unwarp data
    ActiveSheet.Rows.WrapText = False
    'filter data with an array
    With ActiveSheet
        .AutoFilterMode = False
        .UsedRange.AutoFilter
        .UsedRange.AutoFilter Field:=4, Criteria1:=Array("apple", "banana", "car", "dog", "engineer", "fire", "google"), Operator:=xlFilterValues
    End With
    'create a new sheet nearby and paste filtered data
    Sheets.Add.Name = "New"
    Sheets("Sheet1").Activate
    Range("A1").CurrentRegion.Copy
    Worksheets("New").Paste
    Application.CutCopyMode = False
    Worksheets("New").Activate
'Split data based on Column D's values
    Application.ScreenUpdating = False
    Set SrcSheet = ActiveSheet
    LR = SrcSheet.Cells(SrcSheet.Rows.Count, SPCol).End(xlUp).Row
    For SrcRow = FirstRow To LR
        SP = SrcSheet.Cells(SrcRow, SPCol).Value
        Set TrgSheet = Nothing
        On Error Resume Next
        Set TrgSheet = Worksheets(SP)
        On Error GoTo 0
        If TrgSheet Is Nothing Then
            Set TrgSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
            TrgSheet.Name = SP
            SrcSheet.Rows(HeaderRow).Copy Destination:=TrgSheet.Rows(HeaderRow)
        End If
        TrgRow = TrgSheet.Cells(TrgSheet.Rows.Count, SPCol).End(xlUp).Row + 1
        SrcSheet.Rows(SrcRow).Copy Destination:=TrgSheet.Rows(TrgRow)
    Next SrcRow
    Application.ScreenUpdating = True
            
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For I = 1 To WS_Count
        
        Sheets(I).Select
    'insert sum formulas on a specific cell and set font and number format
        LC = Range("A" & Rows.Count).End(xlUp).Row
        Cells(LC + 1, 25).Formula = "=SUM(Y2:Y" & LC & ")"
        Cells(LC + 1, 25).Font.FontStyle = "Bold"
        Cells(LC + 1, 25).NumberFormat = "$#,##0.00"
        Cells(LC + 1, 26).Formula = "=SUM(Z2:Z" & LC & ")"
        Cells(LC + 1, 26).Font.FontStyle = "Bold"
        Cells(LC + 1, 26).NumberFormat = "$#,##0.00"

            
    Next
    'define path and save file as name plus date.    
    Dim Path As String
    Path = "C:\Users\exu\Desktop\examples\"
    ActiveWorkbook.saveas filename:=Path & "fruits_Updates" & "_" & Format(Now(), "MM-DD-YY") & ".xls", FileFormat:=xlNormal

       
End Sub

