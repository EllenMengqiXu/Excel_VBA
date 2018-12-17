Attribute VB_Name = "DCU_T"
Sub DCU_T()

    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute TDR macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes

        ActiveSheet.Range("D2").Value = Format(Now() - 1, "dd-mmm")
        ActiveSheet.Range("E2").Value = Format(Now(), "dd-mmm")
        ActiveSheet.Range("F2").Value = Format(Now() + 1, "dd-mmm")
        ActiveSheet.Range("G2").Value = Format(Now() + 2, "dd-mmm")
        ActiveSheet.Range("H2").Value = Format(Now() + 3, "dd-mmm")
        ActiveSheet.Range("H1").Value = Format(Now() + 6, "dd-mmm")
    
        ActiveSheet.Range("I4:I26").Copy
        ActiveSheet.Range("C4:C26").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
        ActiveSheet.Range("D8:H10").ClearContents
        ActiveSheet.Range("D16:H26").ClearContents
    
        ActiveSheet.Range("H28:H34").Copy
        ActiveSheet.Range("D28:D34").PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
        ActiveSheet.Range("E28:H34").ClearContents
    
        Cells(4, 13).Value = Cells(4, 13).Value + 1
    
        ActiveSheet.Range("D37:H51").ClearContents
        ActiveSheet.Range("D54:H62").ClearContents
    
        ActiveSheet.Range("I64:I66").Copy
        ActiveSheet.Range("C64:C66").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
        ActiveSheet.Range("D64:H66").ClearContents
        ActiveSheet.Range("D68:H68").ClearContents
        ActiveSheet.Range("E69:H69").ClearContents
    
        ActiveSheet.Range("I79").Copy
        ActiveSheet.Range("C79").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    
        ActiveSheet.Range("D71:H77").ClearContents
        ActiveSheet.Range("E79:H79").ClearContents
        
        Case vbNo
        GoTo Quit:
    End Select

Quit:
    
End Sub
