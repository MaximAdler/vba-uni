﻿Private Sub CommandButton1_Click()

  Dim m As Integer
  Dim n As Integer
  Dim r As Integer
  Dim k As Integer
  Dim p As Integer
  Dim col(1 To 4) As Long
  
  Worksheets("30").Activate
  
  col(1) = RGB(0, 0, 51)
  col(2) = RGB(0, 113, 51)
  col(3) = RGB(196, 113, 51)
  col(4) = RGB(196, 152, 95)
  m = TextBox1.Value
  n = TextBox2.Value
  r = TextBox3.Value
  k = TextBox4.Value
  p = TextBox5.Value
  
  Range(Cells(1, 1), Cells(m, n)).Interior.Color = RGB(0, 152, 95)
  For i = 1 To r
    x1 = Int(((m - k) - 1 + 1) * Rnd + 1)
    y1 = Int(((n - p) - 1 + 1) * Rnd + 1)
    Range(Cells(x1, y1), Cells(x1 + k, y1 + p)).Interior.Color = col(Int((4 - 1 + 1) * Rnd + 1))
  Next i
  
  Range(Cells(30, 1), Cells(30, 1)).Interior.Color = RGB(0, 152, 95)
  Range(Cells(31, 1), Cells(31, 1)).Interior.Color = col(1)
  Range(Cells(32, 1), Cells(32, 1)).Interior.Color = col(2)
  Range(Cells(33, 1), Cells(33, 1)).Interior.Color = col(3)
  Range(Cells(34, 1), Cells(34, 1)).Interior.Color = col(4)
  Cells(30, 2) = "0"
  Cells(31, 2) = "0"
  Cells(32, 2) = "0"
  Cells(33, 2) = "0"
  Cells(34, 2) = "0"
  
  For i = 1 To m
    For j = 1 To n
      If Range(Cells(i, j), Cells(i, j)).Interior.Color = RGB(0, 152, 95) Then Cells(30, 2) = Cells(30, 2) + 1
      If Range(Cells(i, j), Cells(i, j)).Interior.Color = col(1) Then Cells(31, 2) = Cells(31, 2) + 1
      If Range(Cells(i, j), Cells(i, j)).Interior.Color = col(2) Then Cells(32, 2) = Cells(32, 2) + 1
      If Range(Cells(i, j), Cells(i, j)).Interior.Color = col(3) Then Cells(33, 2) = Cells(33, 2) + 1
      If Range(Cells(i, j), Cells(i, j)).Interior.Color = col(4) Then Cells(34, 2) = Cells(34, 2) + 1
    Next j
  Next i
  
End Sub
