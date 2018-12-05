Attribute VB_Name = "Ellen"
Sub apple()

    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute apple macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
    
        Sheets.Add.Name = "New"
    
        Sheets("aple").Select
    
        lastRow = ActiveSheet.Range("Y" & Rows.Count).End(xlUp).Row
        ActiveSheet.Range("B12:Z" & lastRow).Copy
        Worksheets("New").Cells.PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
        Sheets("New").Select
    
        ActiveSheet.Range("B:B").Delete
    
        ActiveSheet.Range("G:H,J:J,W:W").NumberFormat = "m/d/yyyy"
    
        Application.ScreenUpdating = False

            Range("F1").EntireColumn.Insert shift:=xlToRight
            Range("L1").EntireColumn.Insert shift:=xlToRight
            Range("R1").EntireColumn.Insert shift:=xlToRight
            Range("X1").EntireColumn.Insert shift:=xlToRight
            Range("F1").Value = "orange"
            Range("L1").Value = "orange"
            Range("R1").Value = "orange"
            Range("X1").Value = "orange"
        
        Application.ScreenUpdating = True
    
        Case vbNo
        GoTo Quit:
    End Select

Quit:
End Sub
