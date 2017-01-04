Private Sub CommandButton1_Click()

 Worksheets("3").Activate
 
 Dim xs As Double
 Dim step As Double
 Dim xe As Double
 Dim f() As Double
 Dim x As Double
 Dim x0 As Double
 
 N = TextBox1.Value
 e = TextBox2.Value
 xs = TextBox3.Value
 xe = TextBox4.Value
 x = TextBox5.Value
 
 
 
 
 ReDim f(N)
 For i = 0 To N
   f(i) = Cells(i + 2, 1)
 Next i
 step = (xe - xs) / 100
 For i = 0 To 100
   Cells(i + 1, 4) = xs + i * step
   Cells(i + 1, 5) = Module2.getF(f, xs + i * step)
 Next i
 
 
 
 Do
 x0 = x
 x = x0 - Module2.getF(f, x) / Module2.getFF(f, x)
 Loop While Abs(x0 - x) > e
 Label4.Caption = " Êîðåíü " + CStr(x)
End Sub