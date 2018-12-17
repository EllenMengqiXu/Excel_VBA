Excel_VBA
===

This project focuses on basic VBA programming used to generate reports and is not finished yet. Generally, it is accumulated when there is new knowledges need me to learn.
---
TO BE CONTINUED

*what you can get from here:

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
* define last Row
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

Define Cell format

Insert Cols and name first cell

