Attribute VB_Name = "CLCtN"
Sub copylastcoltonext()

    Dim Msg As String, Ans As Variant

    Msg = "Would you like to execute DCU_TtoM macro?"

    Ans = MsgBox(Msg, vbYesNo)

    Select Case Ans

        Case vbYes
        Dim copyfirst As Integer
        Dim copylast As Integer
        
        copyfirst = ActiveSheet.Range("IV28").End(xlToLeft).Column
        copylast = ActiveSheet.Range("IV34").End(xlToLeft).Column
    
        ActiveSheet.Range(Cells(28, copyfirst), Cells(34, copylast)).Copy
    
        ActiveSheet.Range(Cells(28, copyfirst + 1), Cells(34, copylast + 1)).PasteSpecial Paste:=xlPasteFormulas
        ActiveSheet.Range(Cells(28, copyfirst), Cells(34, copylast)).PasteSpecial Paste:=xlPasteValues
    
    
        ActiveSheet.Cells(79, copyfirst).Copy
        ActiveSheet.Cells(79, copyfirst + 1).PasteSpecial Paste:=xlPasteFormulas
        
        Application.CutCopyMode = False
    
        Cells(4, 13).Value = Cells(4, 13).Value + 1
        
        Case vbNo
        GoTo Quit:
    End Select

Quit:
    
End Sub


