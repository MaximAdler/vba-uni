﻿Private Sub CommandButton1_Click()

Dim text As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim k1 As Integer
Dim w As Integer
Dim words() As String
Dim words2() As String
Dim n As Integer
Dim num As Integer
Dim useNum As Integer

k = 0
n = 5
num = 1
useNum = 1



Worksheets("40").Activate
ActiveSheet.Cells.ClearContents
ActiveSheet.Cells.Font.Name = "Times New Roman"
ActiveSheet.Cells.Font.ColorIndex = 1
ActiveSheet.Cells.Interior.ColorIndex = -4142

If ActiveSheet.ChartObjects.Count > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

text = TextBox1.text

text = Replace(text, "? ", " ")
text = Replace(text, "?", " ")
text = Replace(text, "! ", " ")
text = Replace(text, "!", " ")
text = Replace(text, ". ", " ")
text = Replace(text, ".", " ")
text = Replace(text, ", ", " ")
text = Replace(text, " - ", " ")
text = Replace(text, ": ", " ")
text = Replace(text, "; ", " ")

text = Trim(text)

For i = 1 To Len(text)
    If text = " " Then
        MsgBox "Ââåä³òü òåêñò!"
        Exit Sub
    End If
    If IsNumeric(Mid(text, i, 1)) Then
        MsgBox "Òåêñò íå ïîâèíåí ì³ñòèòè ÷èñëà"
        Exit Sub
    End If
    If Mid(text, i, 2) = "  " Then
        MsgBox "Òåêñò íå ïîâèíåí ì³ñòèòè ê³ëüêà ïðîá³ë³â!"
        Exit Sub
    End If
Next i


For i = 1 To Len(text)
    If Mid(text, i, 1) = " " Then
    w = w + 1
    End If
Next i
w = w + 1

ReDim words(w) As String


j = 1
k = 1
For i = 1 To Len(text)
        Do While Mid(text, j, 1) <> " "
            words(k) = words(k) & Mid(text, j, 1)
            If j >= Len(text) Then GoTo b
            j = j + 1
        Loop
    j = j + 1
    k = k + 1
    i = j
Next i

b:
ReDim words2(w, w) As String

For i = 1 To w
For j = 1 To w
If words(i) = words(j) And i <> j Then
words(j) = "  "
useNum = useNum + 1
End If

If words(i) <> "  " Then
words2(num, 1) = words(i)
words2(num, 2) = useNum
End If
Next j
If words(i) <> "  " Then
num = num + 1
End If
useNum = 1
Next i

n = 5
For k = 1 To w
    ActiveSheet.Cells(n, 1) = words2(k, 1)
    ActiveSheet.Cells(n, 2) = words2(k, 2)
    n = n + 1
Next k

n = 5
For k = 1 To w
If words2(k, 2) <> "" Then
   If words2(k, 2) > 1 Then
    ActiveSheet.Cells(n, 4) = words2(k, 1)
    ActiveSheet.Cells(n, 5) = words2(k, 2)
    n = n + 1
    End If
End If
Next k

    Range(Cells(5, 4), Cells(n - 1, 5)).Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SetSourceData Source:=Range(Cells(5, 4), Cells(n - 1, 5))
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MajorUnit = 0.5
    ActiveChart.Axes(xlValue).MajorUnit = 1
    ActiveChart.Legend.Select
    Selection.Delete


End Sub