﻿Private Sub CommandButton1_Click()
    
    Dim n As Integer
    Dim rez As Boolean
    
    Worksheets("34").Activate

  
  n = TextBox1.Value
  sum = 0
  
  For i = 1 To n
   sum = sum + Cells(1, i)
  Next i
  
  
  For i = 1 To n
    sum2 = 0
    For j = 1 To n
      sum2 = sum2 + Cells(i, j)
    Next j
    If sum2 <> sum Then
      MsgBox ("íå ìàã³÷íèé êâàäðàò")
      Exit Sub
    End If
  Next i
  
  sum2 = 0
  For i = 1 To n
    sum2 = sum2 + Cells(i, i)
  Next i
  If sum2 <> sum Then
      MsgBox ("íå ìàã³÷íèé êâàäðàò")
      Exit Sub
  End If
  
  sum2 = 0
  For i = 1 To n
    sum2 = sum2 + Cells(i, n - i + 1)
  Next i
  If sum2 <> sum Then
      MsgBox ("íå ìàã³÷íèé êâàäðàò")
      Exit Sub
          Else
        MsgBox "ìàã³÷íèé êâàäðàò"
  End If
End Sub

