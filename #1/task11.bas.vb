﻿Function Delay(Pause As Single)
    Dim Start As Single
    Start = Timer
    Do While Timer < Start + Pause
    DoEvents
    Loop
End Function

Private Sub CommandButton1_Click()

Dim m As String
Dim n As String
Dim i As Integer
Dim j As Integer
Dim pramok() As Integer
Dim x As Integer
Dim y As Integer
Dim xodx(8) As Integer
Dim xody(8) As Integer
Dim k1 As Integer
Dim k2 As Integer
Dim result As Integer

Worksheets("32").Activate
ActiveSheet.Cells.ClearContents
Range(Cells(1, 1), Cells(56000, 3)) = Null
ActiveSheet.Cells.Font.Name = "Times New Roman"
b:
m = TextBox1.text
If Val(m) = 0 And m <> "0" Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If m < 0 Or m * 1 <> Int(m) Or m = Empty Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If m > 50 Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If m < 5 Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
m = CInt(m)
a:
n = TextBox2.text
If Val(n) = 0 And n <> "0" Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If n < 0 Or n * 1 <> Int(n) Or n = Empty Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If n > 50 Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If n < 5 Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
n = CInt(n)

ReDim pramok(m, n) As Integer

For i = 1 To m
For j = 1 To n
pramok(i, j) = 0
Next j
Next i

For i = 1 To m
For j = 1 To n
ActiveSheet.Cells(i, j) = pramok(i, j)
Next j
Next i

c:
x = TextBox3.Value
If Val(x) = 0 And x <> "0" Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If x < 0 Or x * 1 <> Int(x) Or x = Empty Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If x > m Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If

D:
y = TextBox4.Value
If Val(y) = 0 And y <> "0" Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If y < 0 Or y * 1 <> Int(y) Or y = Empty Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If
If y > n Then
MsgBox "Íåïðàâèëüíî ââåäåíî çíà÷åííÿ!"
Exit Sub
End If

xodx(1) = 1
xody(1) = 2
xodx(2) = 1
xody(2) = -2
xodx(3) = -1
xody(3) = -2
xodx(4) = -1
xody(4) = 2
xodx(5) = 2
xody(5) = 1
xodx(6) = 2
xody(6) = -1
xodx(7) = -2
xody(7) = -1
xodx(8) = -2
xody(8) = 1

For k1 = 1 To m * n
    pramok(x, y) = 100
For i = 1 To 8
    If (x - xodx(i)) > 0 And (x - xodx(i)) <= m And (y - xody(i)) > 0 And (y - xody(i)) <= n Then
        pramok(x - xodx(i), y - xody(i)) = pramok(x - xodx(i), y - xody(i)) - 1
    End If
Next
k2 = 100
For i = 1 To 8
If (x - xodx(i)) >= 1 And (x - xodx(i)) <= m And (y - xody(i)) > 0 And (y - xody(i)) <= n And (x - xodx(i)) <> 0 Then
result = pramok(x - xodx(i), y - xody(i))
If k2 > result Then
k2 = result
j = i
End If
End If
Next
ActiveSheet.Cells(x, y) = k1
ActiveSheet.Cells(x, y).Interior.ColorIndex = 20
x = x - xodx(j)
y = y - xody(j)

Delay (0.3)
Next
End Sub