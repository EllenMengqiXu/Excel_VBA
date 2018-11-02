Attribute VB_Name = "CLCtN"
Sub copylastcoltonext()
    
    Dim copyfirst As Integer
    Dim copylast As Integer
        
    copyfirst = ActiveSheet.Range("IV28").End(xlToLeft).Column
    copylast = ActiveSheet.Range("IV34").End(xlToLeft).Column
    
    ActiveSheet.Range(Cells(28, copyfirst), Cells(34, copylast)).Copy
    
    ActiveSheet.Range(Cells(28, copyfirst + 1), Cells(34, copylast + 1)).PasteSpecial Paste:=xlPasteFormulas
    ActiveSheet.Range(Cells(28, copyfirst), Cells(34, copylast)).PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
    
End Sub


