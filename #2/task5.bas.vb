Private Sub CommandButton1_Click()
Dim s$, I%, intfh%, j%, M%, n%, big%, k%, min%, xmin%, ymin%
Dim file As Variant
Dim f As Variant
Dim matr(100, 100) As Variant
Dim w() As String
Dim x1 As Variant
Dim y1 As Variant
Dim x2 As Variant
Dim y2 As Variant
Dim way() As Integer
Dim a() As Boolean
Dim way2() As Integer

I = 1
min = 1000
big = 1000

Worksheets("3").Activate
ActiveSheet.Cells.ClearContents
ActiveSheet.Cells.Font.Size = 11
ActiveSheet.Cells.Font.Name = "Times New Roman"
ActiveSheet.Cells.Font.ColorIndex = 1
ActiveSheet.Cells.Interior.ColorIndex = -4142

file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

If file <> False Then
    f = file
Else
    MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
    Exit Sub
End If



intfh = FreeFile()
Open f For Input As #intfh

Do While Not EOF(intfh)
Line Input #intfh, s
w = Split(s, " ")
    For j = 1 To (UBound(w) + 1) Step 1
        matr(I, j) = w(j - 1)
        ActiveSheet.Cells(I, j) = matr(I, j)
    Next j
I = I + 1
Loop

M = I - 1
n = j - 1

x1 = TextBox1.Value
If Val(x1) = 0 And x1 <> "0" Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If
If x1 <= 0 Or x1 * 1 <> Int(x1) Or x1 = Empty Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If
x1 = CInt(x1)
If x1 > M Then
    MsgBox ("Âåëèêå ÷èñëî, ââåä³òü ìåíøå ÷èñëî!")
    Exit Sub
End If

y1 = TextBox3.Value

If Val(y1) = 0 And y1 <> "0" Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If
If y1 <= 0 Or y1 * 1 <> Int(y1) Or y1 = Empty Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If

y1 = CInt(y1)

If y1 > n Then
    MsgBox ("Âåëèêå ÷èñëî, ââåä³òü ìåíøå ÷èñëî!")
    Exit Sub
End If


x2 = TextBox4.Value

If Val(x2) = 0 And x2 <> "0" Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³, ñïðîáóéòå ùå ðàç!"
    Exit Sub
End If
If x2 <= 0 Or x2 * 1 <> Int(x2) Or x2 = Empty Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If

x2 = CInt(x2)

If x2 > M Then
    MsgBox ("Âåëèêå ÷èñëî, ââåä³òü ìåíøå ÷èñëî!")
    Exit Sub
End If

y2 = TextBox2.Value

If Val(y2) = 0 And y2 <> "0" Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If
If y2 <= 0 Or y2 * 1 <> Int(y2) Or y2 = Empty Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If

y2 = CInt(y2)

If y2 > n Then
    MsgBox ("Âåëèêå ÷èñëî, ââåä³òü ìåíøå ÷èñëî!")
    Exit Sub
End If

ReDim way(M, n) As Integer
For I = 1 To M
    For j = 1 To n
        way(I, j) = ActiveSheet.Cells(I, j)
    Next j
Next I

ReDim a(M, n) As Boolean
For I = 1 To M
    For j = 1 To n
        a(I, j) = True
    Next j
Next I

ReDim way2(M, n) As Integer
For I = 1 To M
    For j = 1 To n
        way2(I, j) = big
    Next j
Next I

way2(x1, y1) = way(x1, y1)
ActiveSheet.Cells(M + x1 + 2, y1) = way2(x1, y1)
a(x1, y1) = False
I = 1
j = 1


For k = 1 To M * n

    If y1 < n Then
        If a(x1, y1 + 1) = True Then
            If way(x1, y1 + 1) + way2(x1, y1) < way2(x1, y1 + 1) Then
                way2(x1, y1 + 1) = way(x1, y1 + 1) + way2(x1, y1)
            End If
        End If
    End If
        
        
    If x1 < M Then
        If a(x1 + 1, y1) = True Then
            If way(x1 + 1, y1) + way2(x1, y1) < way2(x1 + 1, y1) Then
                way2(x1 + 1, y1) = way(x1 + 1, y1) + way2(x1, y1)
            End If
        End If
    End If
    
    If y1 > 1 Then
        If a(x1, y1 - 1) = True Then
            If way(x1, y1 - 1) + way2(x1, y1) < way2(x1, y1 - 1) Then
                way2(x1, y1 - 1) = way(x1, y1 - 1) + way2(x1, y1)
            End If
        End If
    End If
    
    If x1 > 1 Then
        If a(x1 - 1, y1) = True Then
            If way(x1 - 1, y1) + way2(x1, y1) < way2(x1 - 1, y1) Then
                way2(x1 - 1, y1) = way(x1 - 1, y1) + way2(x1, y1)
            End If
        End If
    End If
    min = 1000
    For x = 1 To M
    For y = 1 To n
        If a(x, y) = True Then
        If way2(x, y) < min Then
            min = way2(x, y)
            xmin = x
            ymin = y
        End If
        End If
    Next
    Next
    a(xmin, ymin) = False
    x1 = xmin
    y1 = ymin
    ActiveSheet.Cells(M + x1 + 2, y1) = way2(x1, y1)
Next

ActiveSheet.Cells(x2 + M + 2, y2).Interior.ColorIndex = 28
For k = 1 To M * n
    If y2 < n Then
        If way2(x2, y2 + 1) = way2(x2, y2) - way(x2, y2) Then
            ActiveSheet.Cells(x2 + M + 2, y2 + 1).Interior.ColorIndex = 28
            y2 = y2 + 1
        End If
    End If
    If x2 < M Then
        If way2(x2 + 1, y2) = way2(x2, y2) - way(x2, y2) Then
            ActiveSheet.Cells(x2 + M + 3, y2).Interior.ColorIndex = 28
            x2 = x2 + 1
        End If
    End If
    
    If x2 > 1 Then
        If way2(x2 - 1, y2) = way2(x2, y2) - way(x2, y2) Then
            ActiveSheet.Cells(x2 + M + 1, y2).Interior.ColorIndex = 28
            x2 = x2 - 1
        End If
    End If

    If y2 > 1 Then
    If way2(x2, y2 - 1) = way2(x2, y2) - way(x2, y2) Then
    ActiveSheet.Cells(x2 + M + 2, y2 - 1).Interior.ColorIndex = 28
    y2 = y2 - 1
    End If
    End If
Next k
    
Close #intfh

End Sub