﻿Private Sub CommandButton1_Click()

Dim m As String
Dim n As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim mas(100, 100) As Byte
Dim sq As Integer
Dim sq1 As Integer
Dim sq2 As Integer
Dim i1 As Integer
Dim j1 As Integer
Dim i2 As Integer
Dim j2 As Integer
Dim i3 As Integer
Dim j3 As Integer
Dim stolbec As Integer
Dim s As Integer
Dim i4 As Integer

s = 0

Worksheets("33").Activate
ActiveSheet.Cells.ClearContents
Range(Cells(2, 2), Cells(56000, 3)) = Null
ActiveSheet.Cells.Font.Name = "Times New Roman"

b:
m = TextBox1.text
If Val(m) = 0 And m <> "0" Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If
If m < 0 Or m * 1 <> Int(m) Or m = Empty Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If
m = CInt(m)
If m > 100 Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If


a:
n = TextBox2.text
If Val(n) = 0 And n <> "0" Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If
If n < 0 Or n * 1 <> Int(n) Or n = Empty Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If
n = CInt(n)
If n > 100 Then
    MsgBox "Íåïðàâèëüí³ äàí³!"
    Exit Sub
End If



For i = 1 To m
    For j = 1 To n
        ActiveSheet.Cells(i, j) = Round(Rnd, 0)
    Next j
Next i
For i = 1 To m
    For j = 1 To n
        mas(i, j) = ActiveSheet.Cells(i, j)
    Next j
Next i



sq = 0
For i = 1 To m
    For j = 1 To n
If mas(i, j) = 1 Then
        i1 = i
        j1 = 0
        k = 0
        sq1 = 1000
Do While mas(i1, j + j1) = 1
       k = 0
    Do While mas(i1, j + j1) = 1
        i2 = i
        i1 = i1 + 1
        k = k + 1
    Loop
    i1 = i
    If k < sq1 Then
        sq1 = k
    End If
    j1 = j1 + 1
    sq2 = sq1 * j1
    If sq2 > sq Then
        sq = sq2
        j2 = j + j1 - 1
        stolbec = j
    End If
Loop
    If sq > s Then
    s = sq
    i3 = i2
    j3 = j2
    End If
End If
Next j
Next i


i4 = s / (j3 - stolbec + 1) + i3 - 1
ActiveSheet.Range(Cells(i3, stolbec), Cells(i4, j3)).Interior.ColorIndex = 20
Label4.Caption = "Ê³ëüê³ñòü = " + CStr(s)

Exit Sub
End Sub
