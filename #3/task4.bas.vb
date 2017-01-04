Private Sub CommandButton1_Click()
  Worksheets("2").Activate
  For i = 1 To 34
    Cells(i, 1) = Worksheets("2").Cells(i, 1)
    Cells(i, 2) = "Òåìà ¹" + CStr(i)
  Next i
End Sub

Private Sub CommandButton2_Click()
Worksheets("2").Activate
 For i = 1 To 1000
   Randomize
   r = Int((34 - 1 + 1) * Rnd + 1)
   t1 = Cells(r, 1)
   r2 = Int((34 - 1 + 1) * Rnd + 1)
   Cells(r, 1) = Cells(r2, 1)
   Cells(r2, 1) = t1
   
 Next i
 
  For i = 1 To 1000
   Randomize
   r = Int((34 - 1 + 1) * Rnd + 1)
   t1 = Cells(r, 2)
   
   r2 = Int((34 - 1 + 1) * Rnd + 1)
   Cells(r, 2) = Cells(r2, 2)
   
   Cells(r2, 2) = t1
   
   
 Next i
 
End Sub
