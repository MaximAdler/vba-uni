﻿Private Sub CommandButton1_Click()

Dim x1 As Variant
Dim y1 As Variant
Dim x2 As Variant
Dim y2 As Variant
Dim i As Integer
Dim j As Integer
Dim x3 As Integer
Dim y3 As Integer
Dim a As Long
Dim b As Long
Dim num As Long
Dim stp As Integer
Dim num2 As Integer
Dim k As Integer

On Error GoTo k

Worksheets("39").Activate
ActiveSheet.Cells.ClearContents
ActiveSheet.Cells.Font.Name = "Times New Roman"
ActiveSheet.Cells.Font.ColorIndex = 1
ActiveSheet.Cells.Interior.ColorIndex = -4142


x1 = TextBox1.text
x1 = Trim(x1)
If Val(x1) = 0 And x1 <> "0" Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
If x1 <= 0 Or x1 * 1 <> Int(x1) Or x1 = Empty Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
x1 = CInt(x1)
x1 = Int(x1)

y1 = TextBox2.text
y1 = Trim(y1)
If Val(y1) = 0 And y1 <> "0" Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
If y1 <= 0 Or y1 * 1 <> Int(y1) Or y1 = Empty Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
y1 = CInt(y1)
y1 = Int(y1)

x2 = TextBox3.text
x2 = Trim(x2)
If Val(x2) = 0 And x2 <> "0" Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
If x2 <= 0 Or x2 * 1 <> Int(x2) Or x2 = Empty Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
x2 = CInt(x2)
If x2 = x1 Then
MsgBox "Çàäàéòå ³íùó êîîðäèíàòó, õ1 = õ2"
Exit Sub
End If
x2 = Int(x2)

y2 = TextBox4.text
y2 = Trim(y2)
If Val(y2) = 0 And y2 <> "0" Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
If y2 <= 0 Or y2 * 1 <> Int(y2) Or y2 = Empty Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
y2 = CInt(y2)
If y2 = y1 Then
MsgBox "Çàäàéòå ³íùó êîîðäèíàòó, y1 = y2"
Exit Sub
End If
y2 = Int(y2)


If y2 < y1 Then
y3 = y1
y1 = y2
y2 = y3
x3 = x1
x1 = x2
x2 = x3
End If

If x1 < x2 Then
a = x2 - x1 + 1
b = y2 - y1 + 1

For i = x2 + 2 To x2 + 1 + a
For j = y1 To y2
ActiveSheet.Cells(i, j) = Round(1 * Rnd(), 0)
Next j
Next i

ActiveSheet.Cells(x1, y1) = 1
num = 1
For i = x1 To x2
For j = y1 To y2
num = ActiveSheet.Cells(i, j)
If j + 1 <= y2 Then
    If ActiveSheet.Cells(i, j + 1) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i, j + 1) = "" Then ActiveSheet.Cells(i, j + 1) = num + 1
End If

If i + 1 <= x2 Then
    If ActiveSheet.Cells(i + 1, j) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i + 1, j) = "" Then ActiveSheet.Cells(i + 1, j) = num + 1
End If

If i + 1 <= x2 And j + 1 <= y2 Then
    If ActiveSheet.Cells(i + 1, j + 1) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i + 1, j + 1) = "" Then
        If ActiveSheet.Cells(i + a + 2, j + 1) = 1 Then ActiveSheet.Cells(i + 1, j + 1) = num + 1
    End If
End If

Next j
Next i

stp = ActiveSheet.Cells(x2, y2)
num2 = stp
i = x2
j = y2
ActiveSheet.Cells(i, j).Interior.ColorIndex = 20

For k = 1 To (num2 - 1)
    If ActiveSheet.Cells(i - 1, j) = stp - 1 Then
        ActiveSheet.Cells(i - 1, j).Interior.ColorIndex = 20
        i = i - 1
        stp = stp - 1
        GoTo N1
    End If

    If ActiveSheet.Cells(i, j - 1) = stp - 1 Then
        ActiveSheet.Cells(i, j - 1).Interior.ColorIndex = 20
        j = j - 1
        stp = stp - 1
        GoTo N1
    End If
    
    If ActiveSheet.Cells(i - 1, j - 1) = stp - 1 Then
        ActiveSheet.Cells(i - 1, j - 1).Interior.ColorIndex = 20
        j = j - 1
        i = i - 1
        stp = stp - 1
        GoTo N1
    End If
N1:
Next k



Else

a = x1 - x2 + 1
b = y2 - y1 + 1


For i = x1 + 2 To x1 + 1 + a
For j = y1 To y2
ActiveSheet.Cells(i, j) = Round(1 * Rnd(), 0)
Next j
Next i

ActiveSheet.Cells(x2, y2) = 1
num = 1
For i = x2 To x1 - 1
For j = y2 To y1 Step (-1)
num = ActiveSheet.Cells(i, j)
If j - 1 >= y1 Then
    If ActiveSheet.Cells(i, j - 1) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i, j - 1) = "" Then ActiveSheet.Cells(i, j - 1) = num + 1
End If

If i + 1 >= x2 Then
    If ActiveSheet.Cells(i + 1, j) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i + 1, j) = "" Then ActiveSheet.Cells(i + 1, j) = num + 1
End If

If i + 1 >= x2 And j - 1 >= y1 Then
    If ActiveSheet.Cells(i + 1, j - 1) > ActiveSheet.Cells(i, j) + 1 Or ActiveSheet.Cells(i + 1, j - 1) = "" Then
        If ActiveSheet.Cells(i + a + 2, j - 1) = 1 Then ActiveSheet.Cells(i + 1, j - 1) = num + 1
    End If
End If
Next j
Next i

stp = ActiveSheet.Cells(x1, y1)
num2 = stp
i = x1
j = y1
ActiveSheet.Cells(i, j).Interior.ColorIndex = 20

For k = 1 To (num2 - 1)
    If ActiveSheet.Cells(i - 1, j) = stp - 1 Then
        ActiveSheet.Cells(i - 1, j).Interior.ColorIndex = 20
        i = i - 1
        stp = stp - 1
    End If

    If ActiveSheet.Cells(i, j + 1) = stp - 1 Then
        ActiveSheet.Cells(i, j + 1).Interior.ColorIndex = 20
        j = j + 1
        stp = stp - 1
    End If
    
    If ActiveSheet.Cells(i - 1, j + 1) = stp - 1 Then
        ActiveSheet.Cells(i - 1, j + 1).Interior.ColorIndex = 20
        j = j + 1
        i = i - 1
        stp = stp - 1
    End If
Next k
End If

k:
MsgBox "Çà òàêèõ äàíèõ êîîðäèíàò òî÷îê ïðîõ³ä íå ìîæëèâèé."

End Sub
