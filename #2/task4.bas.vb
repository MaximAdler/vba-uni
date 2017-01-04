Private Sub CommandButton1_Click()

Dim n$, h&, t&, o&, result$
Dim ed(10)
Dim ed2(10)
Dim ed3(10)
Dim ed4(10)

n = TextBox1.text

' Tests

    If Not IsNumeric(n) Then
        MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If
    If n > 1000 Then
        MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If
    If n < 1 Then
        MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³"
        Exit Sub
    End If

' End Tests

ed(1) = "îäèí"
ed(2) = "äâà"
ed(3) = "òðè"
ed(4) = "÷îòèðè"
ed(5) = "ï'ÿòü"
ed(6) = "ø³ñòü"
ed(7) = "ñ³ì"
ed(8) = "â³ñ³ì"
ed(9) = "äåâ'ÿòü"
ed(10) = "äåñÿòü"

ed2(1) = "îäèíàäöÿòü"
ed2(2) = "äâàíàäöÿòü"
ed2(3) = "òðèíàäöÿòü"
ed2(4) = "÷îòèðíàäöÿòü"
ed2(5) = "ï'ÿòíàäöÿòü"
ed2(6) = "ø³ñòíàäöÿòü"
ed2(7) = "ñ³ìíàäöàòü"
ed2(8) = "â³ñ³ìíàäöÿòü"
ed2(9) = "äåâ'ÿòíàäñÿòü"
ed2(10) = "äâàäöÿòü"

ed3(1) = "äåñÿòü"
ed3(2) = "äâàäöÿòü"
ed3(3) = "òðèäöÿòü"
ed3(4) = "ñîðîê"
ed3(5) = "ï'ÿòäåñÿò"
ed3(6) = "ø³ñòäåñÿò"
ed3(7) = "ñ³ìäåñÿò"
ed3(8) = "â³ñ³ìäåñÿò"
ed3(9) = "äåâ'ÿíîñòî"
ed3(10) = "ñòî"

ed4(1) = "ñòî"
ed4(2) = "äâ³ñò³"
ed4(3) = "òðèñòà"
ed4(4) = "÷îòèðèñòà"
ed4(5) = "ï'ÿòñîò"
ed4(6) = "ø³ñòñîò"
ed4(7) = "ñ³ìñîò"
ed4(8) = "â³ñ³ìñîò"
ed4(9) = "äåâ'ÿòñîò"
ed4(10) = "òèñÿ÷à"

h = n \ 100
t = (n Mod 100) \ 10
o = n Mod 10

result = ""

If h > 0 Then result = result + ed4(h) + " "
If t > 1 Then result = result + ed3(t) + " "
If t = 1 And o = 0 Then result = result + ed3(t) + " "
If t = 1 And o > 0 Then result = result + ed2(o) + " "
If t = 0 And o > 0 Then result = result + ed(o) + " "
If t > 1 And o > 0 Then result = result + ed(o) + " "

TextBox2.text = result

End Sub
