Private Sub CommandButton1_Click()

 Dim a As String
 Dim system As String
 Dim integ As Long
 Dim notInteg As Double
 Dim integ1 As String
 Dim a1 As Integer
 Dim a2 As Integer
 Dim l As Integer
 Dim k As Integer
 Dim i As Integer
 Dim res1 As String
 Dim res As String
 Dim res2 As String
 Dim remainder As Integer
 Dim result As String
 Dim result1 As String
 Dim result2 As String
 Dim result3 As String
 
 
 a1 = 0
 a2 = 0
 a = TextBox1.text
 If Val(a) = 0 And a <> "0" Then
     MsgBox "Âè ââåëè íå êîðåêòí³ äàí³!"
     Exit Sub
 End If
 If a < 0 Or a = Empty Then
 MsgBox "Âè ââåëè íå êîðåêòí³ äàí³!"
     Exit Sub
 End If
 
n:
 system = TextBox2.text
 If Val(system) = 0 And system <> "0" Then
     MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
     Exit Sub
 End If
 If system < 0 Or system * 1 <> Int(system) Or system = Empty Then
  MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
     Exit Sub
 End If
 
 If system <= 1 Or system > 65 Then
     MsgBox "Â òàêó ñèñòåìó íåìîæëèâî ïåðåâåñòè ÷èñëî!"
     Exit Sub
 End If
 If a < 0 Then
     a1 = 1
 End If
 
 integ = Fix(a)
 notInteg = a - Fix(a)
 integ1 = CStr(integ)
 notInteg = Trim(notInteg)
 
 Do While integ <> 0
 k = integ Mod system
 integ = Int(integ / system)
 
 Select Case k
     Case Is >= 10
         For i = 10 To 34 Step 1
             If k = i Then
                 res1 = Chr(i + 55)
             End If
         Next i
     Case Is >= 35
         For i = 35 To 59 Step 1
             If k = i Then
                 res1 = "A" + Chr(i + 30)
             End If
         Next i
     Case Is >= 60
         For i = 60 To 84 Step 1
             If k = i Then
                 res1 = "b" + Chr(i + 5)
             End If
         Next i
     Case Else
         res1 = str(k)
 End Select
 res = res + res1
 Loop
 
 If notInteg <> 0 Then
 Do While notInteg <> 0
 If a2 > 3 Then GoTo r
 a2 = a2 + 1
 remainder = Fix(notInteg * system)
 notInteg = notInteg * Fix(system) - Fix(notInteg * system)
 Select Case remainder
     Case Is >= 10
     For i = 10 To 34 Step 1
         If remainder = i Then
         res2 = Chr(i + 55)
         End If
     Next i
     Case Is >= 35
     For i = 35 To 59 Step 1
     If remainder = i Then
     res2 = "A" + Chr(i + 30)
    End If
     Next i
     Case Is >= 60
     For i = 60 To 84 Step 1
     If remainder = i Then
     res2 = "b" + Chr(i + 5)
    End If
     Next i
     Case Else
     res2 = str(remainder)
 End Select
 result = result + res2
 Loop
 End If
 
r:
 
 res = Replace(res, " ", "")
 For i = Len(res) To 1 Step (-1)
 result1 = result1 + Mid(res, i, 1)
 Next i
 
 result = Replace(result, " ", "")
 result3 = result1
 If a1 = 1 Then
 result3 = "-" + result1
 End If
 
 If notInteg <> 0 Then
 result3 = result3 + "," + result
 End If
 result3 = Replace(result3, " ", "")
 Label4.Caption = result3
 Exit Sub
 
k:
 MsgBox "Äàíó îïåðàö³þ íå ìîæëèâî âèêîíàòè. Ïîìèëêà: " & Err.Description
 
End Sub