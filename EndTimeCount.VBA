Private Sub EndTimeCount()

Dim i As Integer
Dim j As Integer
Dim k As Integer

For j = 0 To 729

	k=7
	
	For i = (j*24+1) + 1 To (j*24+1) + 24
	 

			Cells(i, 6) = Cells(i, k).Value
			
			k=k+1
				   
	Next i
	
Next j



End Sub
