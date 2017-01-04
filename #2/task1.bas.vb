Function min(a1() As Integer, a2() As Integer, k As Integer) As Integer
Dim t1 As Integer
Dim t2 As Integer
Dim I As Integer
Dim res As Integer
t1 = a1(1)
For I = 2 To k
If a1(I) < t1 Then
t1 = a1(I)
res = I
End If
Next I
t2 = a2(res)
For I = 1 To k
If a1(I) = t1 And t2 > a2(I) Then
t2 = a2(I)
res = I
End If
Next I
min = res
End Function

Function max(a1() As Integer, a2() As Integer, a As Integer, k As Integer) As Integer
Dim t1 As Integer
Dim t2 As Integer
Dim res As Integer
Dim I As Integer
t1 = 0
For I = 1 To k
If a1(I) > t1 And a1(I) <= a Then
t1 = a1(I)
res = I
End If
Next I

If t1 <> 0 Then
GoTo M1
Else
t1 = 100
For I = 1 To k
If a1(I) < t1 And a1(I) > a Then
t1 = a1(I)
res = I
End If
Next I
End If

M1:
t2 = a2(res)
For I = 1 To k
If a1(I) = t1 And t2 < a2(I) Then
t2 = a2(I)
res = I
End If
Next I

max = res

End Function

Private Sub CommandButton1_Click()

Dim k$, I%, a%, result$, bestResult$, otherResult, p%, s%, b$, iLastRow&
Dim n() As Integer
Dim q() As Integer
Dim M() As Integer
Dim num() As Integer

iLastRow = Cells(Rows.count, 1).End(xlUp).Row

k = TextBox1.text
bestResult = ""
otherResult = ""

' Tests
    If Not IsNumeric(k) Then
        MsgBox "Íåïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If
    If k < 0 Then
        MsgBox "Íåïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If
    If k = "" Then
        MsgBox "Íåïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If
    If k > iLastRow - 1 Then
        MsgBox "Ê á³ëüøå í³æ ê³ëüê³ñòü ì³øê³â"
        Exit Sub
    End If
' End Tests

Worksheets("1").Activate

I = 1
Do While ActiveSheet.Cells(I + 1, 1) <> ""
    I = I + 1
Loop
a = I - 1
ReDim n(a) As Integer
ReDim q(a) As Integer
ReDim M(a) As Integer
ReDim num(a) As Integer

For I = 1 To a
    n(I) = ActiveSheet.Cells(I + 1, 1)
    q(I) = ActiveSheet.Cells(I + 1, 2)
    M(I) = ActiveSheet.Cells(I + 1, 3)
Next I
num(1) = min(q, M, a)
q(num(1)) = 100
ActiveSheet.Cells(2, 4) = n(num(1))

For I = 2 To a
    num(I) = max(q, M, M(num(I - 1)), a)
    M(num(I - 1)) = 100
    q(num(I)) = 100
    ActiveSheet.Cells(I + 1, 4) = n(num(I))
    result = result + CStr(n(num(I))) + " "
Next I


For p = Len(result) + 1 - k * 2.26 To Len(result)
    b = Mid(result, p, 1)
        bestResult = bestResult + b
Next p
For s = 1 To Len(result) - k * 2.26
    b = Mid(result, s, 1)
        otherResult = otherResult + b
Next s

   Label3.Caption = "Íàéêðàù³ ì³øêè: " + bestResult + "        ²íø³ ì³øêè: " + otherResult
   
   ' Change before checking
   Open "C:\Users\Admin\Desktop\s12\new.txt" For Output As #1
        Print #1, "Íàéêðàù³ ì³øêè: " + bestResult + "        ²íø³ ì³øêè: " + otherResult
   Close #1

End Sub