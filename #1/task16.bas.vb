﻿Private Sub CommandButton1_Click()

Dim s As String
Dim bTime As String
Dim dTime As String
Dim director() As Integer
Dim booker() As Integer
Dim worker() As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim num() As Integer
Dim colA As Integer
Dim colB As Integer
Dim colC As Integer

Worksheets("38").Activate
ActiveSheet.Cells.Font.Name = "Times New Roman"
ActiveSheet.Cells(1, 1) = "Ïð³çâèùå ïðàö³âíèêà"
ActiveSheet.Cells(1, 2) = "×àñ ó áóõãàëòåðà"
ActiveSheet.Cells(1, 3) = "×àñ ó ô³íàíñîâîãî äèðåêòîðà"
ActiveSheet.Cells(1, 4) = "Ïîðÿäîê ó ÷åðç³"

s = TextBox1.text
bTime = TextBox2.text
dTime = TextBox3.text

colA = ActiveSheet.Columns("A").Rows(65536).End(xlUp).Row + 1
colB = ActiveSheet.Columns("B").Rows(65536).End(xlUp).Row + 1
colC = ActiveSheet.Columns("C").Rows(65536).End(xlUp).Row + 1
ActiveSheet.Cells(colA, 1) = s
ActiveSheet.Cells(colB, 2) = bTime
ActiveSheet.Cells(colC, 3) = dTime

If s = "" Then
    MsgBox "Çàïîâí³òü äàí³"
    Exit Sub
End If
If bTime = "" Then
    MsgBox "Çàïîâí³òü äàí³"
    Exit Sub
End If
If dTime = "" Then
    MsgBox "Çàïîâí³òü äàí³"
    Exit Sub
End If
If Not IsNumeric(dTime) Then
    MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³"
    Exit Sub
End If
If Not IsNumeric(bTime) Then
    MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³"
    Exit Sub
End If

i = 1
Do While ActiveSheet.Cells(i + 1, 1) <> ""
i = i + 1
Loop
k = i - 1
ReDim director(k) As Integer
ReDim booker(k) As Integer
ReDim worker(k) As String
ReDim num(k) As Integer

For i = 1 To k
worker(i) = ActiveSheet.Cells(i + 1, 1)
director(i) = ActiveSheet.Cells(i + 1, 3)
booker(i) = ActiveSheet.Cells(i + 1, 2)
Next i
num(1) = min(booker, director, k)
booker(num(1)) = 100
ActiveSheet.Cells(2, 4) = worker(num(1))

For i = 2 To k
num(i) = max(booker, director, director(num(i - 1)), k)
director(num(i - 1)) = 100
booker(num(i)) = 100
ActiveSheet.Cells(i + 1, 4) = worker(num(i))
Next i

End Sub
