﻿Private Sub CommandButton1_Click()
 Dim x As String
 Dim res As Long
 
  x = TextBox1.text
  res = 0
  amount = 0
  TextBox2.text = ""
  
  If x = "" Then
    MsgBox "Íå ïðàâèëüíî ââåäåíî çíà÷åííÿ"
    Exit Sub
  End If
  If Not IsNumeric(x) Then
    MsgBox "Íå ïðàâèëüíî ââåäåíî çíà÷åííÿ"
    Exit Sub
  End If
  If x < 0 Then
    MsgBox "Íå ïðàâèëüíî ââåäåíî çíà÷åííÿ"
    Exit Sub
  End If
  
  For a = 0 To x \ 1
    For b = 0 To x \ 2
      For c = 0 To x \ 5
        For D = 0 To x \ 10
          res = a + b * 2 + c * 5 + D * 10
          If res = x Then
            TextBox2.text = TextBox2.text + " Âàð³àíò:" + " 1 ºâðî - " + CStr(a) + "; 3 ºâðî - " + CStr(b) + "; 5 ºâðî - " + CStr(c) + "; 10 ºâðî - " + CStr(D)
            amount = amount + 1
          End If
        Next D
      Next c
    Next b
  Next a
  Label3.Caption = "Âñüîãî âàð³àíò³â " + CStr(amount)
End Sub