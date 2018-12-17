Attribute VB_Name = "CPLAC"
Sub CPnextAvaiCol()
    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute CP macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
    
        ActiveSheet.Range("A5:A11").Copy
        ActiveSheet.Range("IV5").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
        
        ActiveSheet.Range("A13:A19").Copy
        ActiveSheet.Range("IV13").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
        ActiveSheet.Range("A25:A31").Copy
        ActiveSheet.Range("IV25").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
        
        ActiveSheet.Range("A33:A39").Copy
        ActiveSheet.Range("IV33").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
        ActiveSheet.Range("A44:A50").Copy
        ActiveSheet.Range("IV44").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValues
        
        ActiveSheet.Range("A52:A58").Copy
        ActiveSheet.Range("IV52").End(xlToLeft).Offset(, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
        Application.CutCopyMode = False
        
        Case vbNo
        GoTo Quit:
    End Select

Quit:
        
End Sub
