Excel_VBA
===

This project focuses on basic VBA programming used to generate reports and is not finished yet. Generally, it is accumulated when there is new knowledges need me to learn.
---
TO BE CONTINUED

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
define last Row
```
LastRow = ActiveSheet.Range("Y" & Rows.Count).End(xlUp).Row
```
Delete Column
```
ActiveSheet.Range("B:B").Delete
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
copy paste special
```
ActiveSheet.Range("I4:I26").Copy
        ActiveSheet.Range("C4:C26").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
ActiveSheet.Range("H28:H34").Copy
        ActiveSheet.Range("D28:D34").PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False        
```
clear cntents
```
ActiveSheet.Range("D71:H77").ClearContents
```
add value to a cell
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
select multiple data ranges
```
Set r = Union(.Range("A2:N" & LastRow), .Range("PL2:PX" & LastRow))
```
unwarp data
```
ActiveSheet.Rows.WrapText = False
```
filter data with an array
```
With ActiveSheet
        .AutoFilterMode = False
        .UsedRange.AutoFilter
        .UsedRange.AutoFilter Field:=4, Criteria1:=Array("apple", "banana", "car", "dog", "engineer", "fire", "google"), Operator:=xlFilterValues
End With
```
insert sum formulas on a specific cell and set font and number format
```
LC = Range("A" & Rows.Count).End(xlUp).Row
Cells(LC + 1, 25).Formula = "=SUM(Y2:Y" & LC & ")"
Cells(LC + 1, 25).Font.FontStyle = "Bold"
Cells(LC + 1, 25).NumberFormat = "$#,##0.00"
```
define path and save file as name plus date
```
Dim Path As String
Path = "C:\Users\exu\Desktop\examples\"
ActiveWorkbook.saveas filename:=Path & "fruits_Updates" & "_" & Format(Now(), "MM-DD-YY") & ".xls", FileFormat:=xlNormal
```
* remove rows and columns
* filter and sort
* remove duplicates
* keep relevant columns
* split data into several worksheets
* insert column and add formula
* highlight value counted
* massage box before running macro
* copy paste to a new sheet and rename
* save as file specific: path and contains: cell value and date
* copy from specific row and paste to last column
* copy from last column and paste to next availalbe column
* copy and paste special such as values, formulas, and number format
* and so on......

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
