Attribute VB_Name = "SP"
Sub SP()

    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute SP macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
        Dim CurFile As String
        Dim DestWB As Workbook
        Dim a As Variant
        
        Const DirLoc As String = "C:\sp\"
        
        Application.ScreenUpdating = False
        
        For Each a In Array("apple", "banana", "car", "dog", "engineer", "fire", "google")
            
        Set DestWB = Workbooks.Add(xlWorksheet)
        
        CurFile = Dir(DirLoc & "*.xls")
        
        Do While CurFile <> vbNullString
        Dim OrigWB As Workbook
        
        Set OrigWB = Workbooks.Open(filename:=DirLoc & CurFile, ReadOnly:=True)
        On Error Resume Next
        OrigWB.Sheets(a).Copy After:=DestWB.Sheets(DestWB.Sheets.Count)
        
        CurFile = Left(Left(CurFile, Len(CurFile) - 4), 4) 'Limits to valid sheet names
        'and removes ".xls"
        
        DestWB.Sheets(DestWB.Sheets.Count).Name = CurFile
        
        OrigWB.Close SaveChanges:=False
        
        CurFile = Dir
        Loop
        
        Application.DisplayAlerts = False
        Application.DisplayAlerts = True
        filename = a
        Path = "C:\sp\Temp\"
        ActiveWorkbook.saveas filename:=Path & filename & "su" & Format(Now(), "MM-DD-YY") & ".xls", FileFormat:=xlNormal
        
        Next a
        
        Application.ScreenUpdating = True
        Case vbNo
        GoTo Quit:
    End Select

Quit:

End Sub

