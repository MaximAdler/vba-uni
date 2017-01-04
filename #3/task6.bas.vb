Private Sub CommandButton1_Click()

Dim f As Variant
Dim file As Variant
Dim n As Variant
Dim comb() As Integer
Dim text() As String
Dim s$, I%, st%, num%, j%, ind%, intfh%, k%, r%, combin$



Worksheets("2").Activate
ActiveSheet.Cells.ClearContents
ActiveSheet.Cells.Font.Name = "Times New Roman"
ActiveSheet.Cells.Font.ColorIndex = 1
ActiveSheet.Cells.Interior.ColorIndex = -4142

I = 1

file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

If file <> False Then
    f = file
Else
    MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
    Exit Sub
End If

intfh = FreeFile()
Open f For Input As intfh
    Do While Not EOF(intfh)
        Line Input #intfh, s
        text = Split(s, Chr(13))
        For j = 1 To (UBound(text) + 1)
            ActiveSheet.Cells(I, j) = text(j - 1)
        Next j
        I = I + 1
    Loop
Close #intfh

    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), TrailingMinusNumbers:= _
        True
st = I - 2

n = TextBox1.text
n = Trim(n)
n = Replace(n, ".", ",")

If Val(n) = 0 And n <> "0" Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If
If n <= 0 Or n * 1 <> Int(n) Or n = Empty Then
    MsgBox "Âè ââåëè íåêîðåêòí³ äàí³!"
    Exit Sub
End If

n = CInt(n)

If n > st Then
    MsgBox ("Âåëèêå ÷èñëî, ââåä³òü ìåíøå ÷èñëî!")
    Exit Sub
End If


ReDim comb(n) As Integer
For I = 1 To n
    comb(I) = I
Next
num = 1


file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé äëÿ çàïèñó ôàéë")

If file <> False Then
    f = file
Else
    MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
    Exit Sub
End If

intfh = FreeFile()
Open f For Output As #intfh

For I = n To 1 Step -1
Do While comb(I) <= st
    If I <> n Then
        For k = I + 1 To n
            comb(k) = comb(k - 1) + 1
        Next k
    End If
For r = comb(I) To st - n + I
    For j = 1 To n
        ActiveSheet.Cells(num, j + 5) = comb(j)
        combin = combin & " " & ActiveSheet.Cells(num, j + 5)
    Next j
    num = num + 1
    comb(n) = comb(n) + 1
    Print #intfh, combin
    combin = ""
Next r
comb(I) = comb(I) + 1
Loop
comb(I - 1) = comb(I - 1) + 1
Next I
Close #intfh

Exit Sub

End Sub
