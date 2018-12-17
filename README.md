Excel_VBA
===

This project focuses on basic VBA programming used to generate reports and is not finished yet. Generally, it is accumulated when there is new knowledges need me to learn.
---

TO BE CONTINUED...

*what you can get from here:*

Message Box
```
Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute apple macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
        Case vbNo
        GoTo Quit:
    End Select
Quit:        
```
Create Newsheet
```
Sheets.Add.Name = "New"
```
Define last Row
```
LastRow = ActiveSheet.Range("Y" & Rows.Count).End(xlUp).Row
```
Delete Column
```
ActiveSheet.Range("B:B").Delete
```
Delete the first row
```
ActiveSheet.Range("1:1").Delete
```
Define Cell format
```
ActiveSheet.Range("G:H,J:J,W:W").NumberFormat = "m/d/yyyy"
```
Insert Cols and name first cell
```
Range("X1").EntireColumn.Insert shift:=xlToRight
Range("X1").Value = "orange"
```
Copy paste special
```
ActiveSheet.Range("I4:I26").Copy
        ActiveSheet.Range("C4:C26").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
ActiveSheet.Range("H28:H34").Copy
        ActiveSheet.Range("D28:D34").PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False        
```
Clear contents
```
ActiveSheet.Range("D71:H77").ClearContents
```
Add value to a cell
```
Cells(4, 13).Value = Cells(4, 13).Value + 1
```
Define last Column
```
LastColumn1 = ActiveSheet.Range("IV28").End(xlToLeft).Column
LastColumn2 = ActiveSheet.Range("IV34").End(xlToLeft).Column
```
Copy Range with defined cell
```
ActiveSheet.Range(Cells(28, copyfirst), Cells(34, copylast)).Copy
```
Paste to next available column
```
ActiveSheet.Range(Cells(28, copyfirst + 1), Cells(34, copylast + 1)).PasteSpecial Paste:=xlPasteFormulas
```
Copy range and paste to the last blank range
```
ActiveSheet.Range("A5:A11").Copy
ActiveSheet.Range("IV5").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
```  
Select multiple data ranges
```
Set r = Union(.Range("A2:N" & LastRow), .Range("PL2:PX" & LastRow))
```
Unwarp data
```
ActiveSheet.Rows.WrapText = False
```
Filter data with an array
```
With ActiveSheet
        .AutoFilterMode = False
        .UsedRange.AutoFilter
        .UsedRange.AutoFilter Field:=4, Criteria1:=Array("apple", "banana", "car", "dog", "engineer", "fire", "google"), Operator:=xlFilterValues
End With
```
Insert sum formulas on a specific cell and set font and number format
```
LC = Range("A" & Rows.Count).End(xlUp).Row
Cells(LC + 1, 25).Formula = "=SUM(Y2:Y" & LC & ")"
Cells(LC + 1, 25).Font.FontStyle = "Bold"
Cells(LC + 1, 25).NumberFormat = "$#,##0.00"
```
Insert column and add formula
```
Sub InsertBudgetDiff()
Dim x As Long
Dim d As String

Application.ScreenUpdating = False
Range("M1").EntireColumn.Insert shift:=xlToRight
Range("M1").Value = "Budget Difference"
x = Range("A" & Rows.Count).End(xlUp).Row
d = "=K2-L2"
Range("M2").Resize(x - 1).Formula = d
Application.ScreenUpdating = True

End Sub
```

SORT DATA BASED ON M COLUMN IN DESCENDING (FROM OLDEST TO NEWEST)
```
Range("A1").CurrentRegion.Sort Key1:=Range("M1"), Order1:=xlDescending, Header:=xlYes
```
highlighted column L whose value counted
```
lastRow = ActiveSheet.Range("L" & Rows.Count).End(xlUp).Row
Range("L2:L" & lastRow).Interior.Color = vbYellow
```
Save As
```
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
```
**Apple Report**

Message Box

Create Newsheet

Copy range from sheet A to sheet B
* define last Row

Delete Column

Define cell format

Insert Cols and name first cell

**Pineapple Report**

Message Box

Define cell format

copy paste special

Clear contents

Add value to a cell

**Orange Report**

Message Box

Define last Column

Copy Range with defined cell

Paste to next available column

**Banana Report**

Message Box

Copy range and paste to the last blank range

**Watermelon Report**

select multiple data ranges

copy selected range and create a new workbook and paste

unwarp data

filter data with an array

create a new sheet nearby and paste filtered data

Split data based on Column D's values

insert sum formulas on a specific cell and set font and number format

define path and save file as name plus date

**Kiwi Report**

Delete the first row

Unwrap all active Rows

Delete columns

Sort worksheet baed on column B value by ascending (date from oldest to newest) order

Filter column 12 whose value contains (6), * stands for wildcard

Delete filtered Data

Show data after Being Deleted

**Date Report**

KEEP RELEVANT COLUMNS

CHECK WHETHER TO KEEP THE COLUMN
* IF YES THEN SKIP TO THE NEXT COLUMN,
* IF NO DELETE THE COLUMN

LASTLY AN ESCAPE IN CASE THE SHEET HAS NO COLUMNS LEFT

Insert a new column whose head called "Bike Dog" and cell formuna as K2-L2

FILTER COLUMN 8 AS DATE OF THIS MONTH AND COLUMN 10 WHOSE VALUE IS COMFORTABLE

SORT DATA BASED ON M COLUMN IN DESCENDING (FROM OLDEST TO NEWEST)

SAVE FILE AS .XLS FORMAT AS CSV FORMAT CANNOT SAVE EXCEL FUNCTION

**Pearl Report**

Multiple Filters

Filter Date Range from today(current day in excel function is TODAY(),and it is date in VBA)

SortingAscending

Keep Relevant Columns

* IF YES THEN SKIP TO THE NEXT COLUMN,
* IF NO DELETE THE COLUMN

LASTLY AN ESCAPE IN CASE THE SHEET HAS NO COLUMNS LEFT

SAVE FILE AS .XLS FORMAT

**Resberry Report**

set up the begging of current month and today's date

filter record whose not comfortable and date range between the first day of current month and today

sort data as from oldest to newest

copy data and paste to a new sheet and rename it as L1 Value

move to the new sheet, which next to the current sheet

highlighted column L whose value counted

**Blackberry Report**

copy filtered records

create a new workbook

paste filtered records to the new workbook

define the location of the new Workbook

filename comes from D2

save the new workbook under the defined path and file name

define column N equals to the sum of column R through column AU

save file again

**Durian Report**

Select same name tab such as a,b,c from different reports under same directory

Store these tabs in destination workbooks and tabs' name are from the first four letter of their original workbook

save the destination workbooks as a.xls, b.xls and c.xls.
