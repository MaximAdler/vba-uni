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
Dim x As Integer
Dim y As Integer
Dim elem As Long
Dim mas() As Long
Dim l As Long
Dim q As Long
Dim s As Long
Dim number As Integer
Dim radius As Integer
Dim side As String


Worksheets("31").Activate
ActiveSheet.Cells.ClearContents
ActiveSheet.Cells.Font.Name = "Times New Roman"
b:
    m = TextBox1.text
m = Trim(m)
m = Replace(m, ".", ",")
If Val(m) = 0 And m <> "0" Then
MsgBox "m ââåäåíî íåïðàâèëüíî!"
Exit Sub
End If
If m < 0 Or m * 1 <> Int(m) Or m = Empty Then
MsgBox "m ââåäåíî íåïðàâèëüíî!"
Exit Sub
End If
m = Int(m)
If m > 50 Then
MsgBox "m ââåäåíî íåïðàâèëüíî!"
Exit Sub
End If


a:
n = TextBox2.text
n = Trim(n)
n = Replace(n, ".", ",")
If Val(n) = 0 And n <> "0" Then
MsgBox "Íå ïðàâèëüíî ââåäåíî n!"
Exit Sub
End If
If n < 0 Or n * 1 <> Int(n) Or n = Empty Then
MsgBox "Íå ïðàâèëüíî ââåäåíî n!"
Exit Sub
End If
n = Int(n)
If n > 50 Then
MsgBox "Íå ïðàâèëüíî ââåäåíî n!"
Exit Sub
End If

elem = m * n
ReDim mas(elem) As Long

l = 1
q = 1
Do While q <= elem
s = 0
For i = 1 To Int(l / 2)
    If (l Mod i) = 0 Then
        s = s + 1
    End If
Next i
If s >= 3 And s <= 5 Then
    mas(q) = l
    q = q + 1
End If
l = l + 1
Loop

c:
x = TextBox3.Value
x = Trim(x)
x = Replace(x, ".", ",")
If Val(x) = 0 And x <> "0" Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If
If x < 0 Or x * 1 <> Int(x) Or x = Empty Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If
x = Int(x)
If x > m Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If

D:
y = TextBox4.Value
y = Trim(y)
y = Replace(y, ".", ",")
If Val(y) = 0 And y <> "0" Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If
If y < 0 Or y * 1 <> Int(y) Or y = Empty Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If
y = Int(y)
If y > n Then
MsgBox "Íå ïðàâèëüíî ââåäåíî äàí³!"
Exit Sub
End If

ActiveSheet.Cells(x, y) = mas(1)
i = x
j = y - 1
number = 2
radius = 1

e:
side = TextBox5.text
side = Trim(side)


If side = 1 Then
Do While number < elem
Do While Abs(x - i) <= radius
    If i <= m And j <= n And j > 0 And i > 0 Then
    ActiveSheet.Cells(i, j) = mas(number)
    Delay (0.3)
    number = number + 1
    End If
    i = i - 1
Loop
i = i + 1
j = j + 1

Do While Abs(y - j) <= radius
    If i <= m And j <= n And i > 0 And j > 0 Then
    ActiveSheet.Cells(i, j) = mas(number)
    Delay (0.3)
    number = number + 1
    End If
    j = j + 1
Loop
i = i + 1
j = j - 1

Do While Abs(x - i) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
ActiveSheet.Cells(i, j) = mas(number)
Delay (0.3)
number = number + 1
End If
i = i + 1
Loop
j = j - 1
i = i - 1
radius = radius + 1
Do While Abs(y - j) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
ActiveSheet.Cells(i, j) = mas(number)
Delay (0.3)
number = number + 1
End If
j = j - 1
Loop
j = j + 1
i = i - 1
Loop

Else


Do While number < elem
Do While Abs(x - i) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
    ActiveSheet.Cells(i, j) = mas(number)
    Delay (0.3)
    number = number + 1
End If
i = i + 1
Loop
i = i - 1
j = j + 1
Do While Abs(y - j) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
ActiveSheet.Cells(i, j) = mas(number)
Delay (0.3)
number = number + 1
End If
j = j + 1
Loop
i = i - 1
j = j - 1
Do While Abs(x - i) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
ActiveSheet.Cells(i, j) = mas(number)
Delay (0.3)
number = number + 1
End If
i = i - 1
Loop
j = j - 1
i = i + 1
radius = radius + 1
Do While Abs(y - j) <= radius
If i <= m And j <= n And i > 0 And j > 0 Then
ActiveSheet.Cells(i, j) = mas(number)
Delay (0.3)
number = number + 1
End If
j = j - 1
Loop
j = j + 1
i = i + 1
Loop
End If

End Sub