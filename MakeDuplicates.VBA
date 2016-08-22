Option Explicit 
Sub MakeDuplicates() 
    Dim r As Long, lr As Long, sr As Long, er As Long, rr As Long, s As String, c As Long 
    Application.ScreenUpdating = False 
    lr = Cells(Rows.Count, 1).End(xlUp).Row 
    For r = lr To 1 Step -1 
        sr = r 
        er = r + 23 
        Rows(r + 1).Resize(23).Insert 
        Rows(r).Copy Rows(r + 1).Resize(23) 
        s = Cells(r, 1).Value 
        c = 1 
        For rr = sr To er Step 1 
            Cells(rr, 1).Value = c & s
            c = c + 1 
        Next rr 
    Next r 
    Application.ScreenUpdating = True 
End Sub 
