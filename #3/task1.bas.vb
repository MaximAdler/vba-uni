Private Sub CommandButton1_Click()
  Worksheets("1").Activate
  
   Dim N As Integer
   Dim k() As Double
   Dim kk() As Double
   Dim b() As Double
   Dim dd() As Double
   
   N = TextBox2.Value
   TextBox1.Text = ""
   ReDim k(1 To N, 1 To N) As Double
   ReDim kk(1 To N, 1 To N) As Double
   ReDim b(1 To N) As Double
   ReDim dd(1 To N) As Double
   
   For i = 1 To N
     For j = 1 To N
       k(i, j) = Cells(i, j)
     Next j
     b(i) = Cells(i, N + 1)
   Next i
   
   Worksheets("1.1").Activate
   ActiveSheet.Cells.Clear
   
   
   If k(1, 1) <> 1 And k(1, 1) <> 0 Then
     d = k(1, 1)
     For j = 1 To N
       k(1, j) = k(1, j) / d
     Next j
   End If
   
   For i = 1 To N
     If k(i, i) <> 1 And k(i, i) <> 0 Then
     d = k(i, i)
     For j = i To N
       k(i, j) = k(i, j) / d
     Next j
      b(i) = b(i) / d
     End If
     For j = i + 1 To N
       d = k(j, i)
       d = (-1) * d
       For l = i To N
         k(j, l) = k(j, l) + k(i, l) * d
       Next l
       b(j) = b(j) + b(i) * d
     Next j
   Next i
     
   'îáðàòíûé õîä
   j = N
   TextBox1.Text = ""
   Do
   If j = N Then
     dd(j) = b(j) / k(j, j)
   Else
     x3 = b(j)
    For i = j + 1 To N
     x3 = x3 - k(j, i) * dd(i)
    Next i
     dd(j) = x3 / k(j, j)
   End If
    TextBox1.Text = TextBox1.Text + "X" + CStr(j) + "=" + CStr(dd(j)) + " "
    j = j - 1
   Loop Until j = 0
End Sub

Private Sub CommandButton2_Click()
    Worksheets("1").Activate
   Dim N As Integer
   Dim k() As Double
   Dim kk() As Double
   Dim b() As Double
   Dim dd() As Double


   N = TextBox2.Value
   TextBox1.Text = ""
   ReDim k(1 To N, 1 To N) As Double
   ReDim kk(1 To N, 1 To N) As Double
   ReDim b(1 To N) As Double
   ReDim dd(1 To N) As Double
   
   For i = 1 To N
     For j = 1 To N
       k(i, j) = Cells(i, j)
     Next j
     b(i) = Cells(i, N + 1)
   Next i
   
   det = Module1.detn(k, N)
   TextBox1.Text = TextBox1.Text + "det=" + CStr(det)
   If det = 0 Then
     TextBox1.Text = TextBox1.Text + " íåò ðåøåíèÿ"
   End If
   For i = 1 To N
    
       For l = 1 To N
         For m = 1 To N
           kk(l, m) = k(l, m) 'copy array
         Next m
    Next l
    
    
    For j = 1 To N
      kk(j, i) = b(j)
    Next j
      dd(i) = detn(kk, N)
     TextBox1.Text = TextBox1.Text + "det" + CStr(i) + "=" + CStr(dd(i)) + " "
   Next i
   Sum = 0
   For i = 1 To N
   
     If dd(i) <> 0 Then
       TextBox1.Text = TextBox1.Text + "     X" + CStr(i) + "=" + CStr(dd(i) / det) + " | "
       Sum = Sum + dd(i) * k(1, i) / det
    End If
      
   Next i
   TextBox1.Text = TextBox1.Text + " sum1 = " + CStr(Sum)
End Sub
