﻿Private Sub CommandButton1_Click()

Dim V As String
Dim weight100 As Integer
Dim weight200 As Integer
Dim weight300 As Integer
Dim weight500 As Integer
Dim weight1000 As Integer
Dim weight1200 As Integer
Dim weight1400 As Integer
Dim weight1500 As Integer
Dim weight2000 As Integer
Dim weight3000 As Integer
Dim k As Long
Dim str As Long

str = 2
k = 0

On Error GoTo e

V = TextBox1.Value
If Val(V) = 0 And V <> "0" Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If

If V <= 0 Or V * 1 <> Int(V) Or V = Empty Then
MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
Exit Sub
End If
V = Int(V)

For weight100 = 0 To Int(V / 100)
For weight200 = 0 To Int(V / 200)
For weight300 = 0 To Int(V / 300)
For weight500 = 0 To Int(V / 500)
For weight1000 = 0 To Int(V / 1000)
For weight1200 = 0 To Int(V / 1200)
For weight1400 = 0 To Int(V / 1400)
For weight1500 = 0 To Int(V / 1500)
For weight2000 = 0 To Int(V / 2000)
For weight3000 = 0 To Int(V / 3000)
If weight100 * 100 + weight200 * 200 + weight300 * 300 + weight500 * 500 + weight1000 * 1000 + weight1200 * 1200 + weight1400 * 1400 + weight1500 * 1500 + weight2000 * 2000 + weight3000 * 3000 = V Then
k = k + 1
End If
Next weight3000
Next weight2000
Next weight1500
Next weight1400
Next weight1200
Next weight1000
Next weight500
Next weight300
Next weight200
Next weight100

Label3.Caption = "Ê³ëüê³ñòü ñïîñîá³â ñêëàñòè äàíó âàãó: " + CStr(k)
Exit Sub

e:
If Err.number = 6 Then
MsgBox "Ââåä³òü ìåíøå ÷èñëî!"
Exit Sub
End If

End Sub
