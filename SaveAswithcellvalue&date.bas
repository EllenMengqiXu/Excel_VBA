Attribute VB_Name = "saveas"
Sub savefile()
    
    Dim Path As String
    Dim filename As String
	'clarify where you want to save your file
    Path = "C:\Users\exu\Desktop\love\"
	'clarify your filename by choosing from excel sheet
    filename = Range("A1")
	'use & to connect words and date,'mm' means 10, 'mmm' means Oct, and 'mmmm' menas October.  
    ActiveWorkbook.saveas filename:=Path & filename & "_love_" & Format(Now(), "mmm") & "_" & Format(Now(), "DD-MM-YY") & ".xls", FileFormat:=xlNormal
    
End Sub
