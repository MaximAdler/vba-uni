Private Sub CommandButton1_Click()

Dim str As Variant, I As Integer, b() As Integer, a() As String, s As Variant
Dim Nums1 As Variant, Nums2 As Variant, Nums3 As Variant, Nums4 As Variant, Nums5 As Variant, n As Variant, otvet As Variant

here:
file = Application.GetOpenFilename("Text Files(*.txt), *.txt", 0, "Îáåð³òü ôàéë ç ÷èñëîì, ïîäàíèì ðèìñüêèìè öèôðàìè")
If file <> False Then
    f = file
Else
    MsgBox ("Îáåð³òü ôàéë, áóäü-ëàñêà")
    GoTo here
End If
intfh = FreeFile()
Open f For Input As #intfh
    Do While Not EOF(intfh)
        Line Input #intfh, str
    Loop
Close #intfh

ReDim a(1 To Len(str)), b(0 To Len(str))
For I = 1 To Len(str)
    a(I) = Mid(str, I, 1)
    If (a(I) <> "I") And (a(I) <> "V") And (a(I) <> "X") And (a(I) > "L") And (a(I) <> "C") And (a(I) <> "D") And (a(I) <> "M") Then
        MsgBox "Íåìàº ðèìñüêîãî ÷èñëà"
    Else
        If a(I) = "I" Then b(I) = 1
        If a(I) = "V" Then b(I) = 5
        If a(I) = "X" Then b(I) = 10
        If a(I) = "L" Then b(I) = 50
        If a(I) = "C" Then b(I) = 100
        If a(I) = "D" Then b(I) = 500
        If a(I) = "M" Then b(I) = 1000
    End If
Next I

For I = 1 To Len(str)
 s = s + b(I)
 If (I > 1) And (b(I - 1) < b(I)) Then
 s = s - 2 * b(I - 1)
 End If
Next
here1:
file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäå çàïèñóâàòèñü ÷èñëî àðàáñüêèìè öèôðàìè")
If file <> False Then
    f1 = file
Else
    MsgBox ("Áóäü-ëàñêà, îáåð³òü ôàéë")
    GoTo here1
End If

intfh = FreeFile()
Open f1 For Output As #intfh
    Print #intfh, s
Close #intfh


here7:
file = Application.GetOpenFilename("Text Files(*.txt), *.txt", 0, "Îáåð³òü ôàéë ç ÷èñëîì, ÿêå íåîáõ³äíî çàïèñàòè ñëîâàìè ")
If file <> False Then
    f2 = file
Else
    MsgBox ("Îáåð³òü ôàéë, áóäü-ëàñêà")
    GoTo here7
End If
intfh = FreeFile()
Open f2 For Input As #intfh
    Do While Not EOF(intfh)
    Line Input #intfh, n
    Loop
Close #intfh

Nums1 = Array(" ", "îäèí", "äâà", "òðè", "÷îòèðè", "ï'ÿòü", "ø³ñòü", "ñ³ì", "â³ñ³ì", "äåâ'ÿòü")
Nums2 = Array(" ", "äåñÿòü ", "äâàäöÿòü ", "òðèäöÿòü ", "ñîðîê ", "ï'ÿòäåñÿò ", "ø³ñòäåñÿò ", "ñ³ìäåñÿò ", "â³ñ³ìäåñÿò ", "äåâ'ÿíîñòî ")
Nums3 = Array(" ", "ñòî ", "äâ³ñò³ ", "òðèñòà ", "÷îòèðèñòà ", "ï'ÿòñîò ", "ø³ñòñîò ", "ñ³ìñîò ", "â³ñ³ìñîò ", "äåâ'ÿòñîò ")
Nums4 = Array(" ", "îäíà ", "äâ³ ", "òðè ", "÷îòèðè ", "ï'ÿòü ", "ø³ñòü ", "ñ³ì ", "â³ñ³ì ", "äåâ'ÿòü ")
Nums5 = Array("äåñÿòü", "îäèíàäöÿòü", "äâàíàäöÿòü", "òðèíàäöÿòü", "÷îòèðíàäöÿòü", "ï'ÿòíàäöÿòü", "ø³ñòíàäöÿòü", "ñ³ìíàäöÿòü", "â³ñ³ìíàäöÿòü", "äåâ'ÿòíàäöÿòü")
' ðîçä³ëÿºìî íà ðîçðÿäè
    ed = Class(n, 1)
   dec = Class(n, 2)
   sot = Class(n, 3)
   tys = Class(n, 4)
dectys = Class(n, 5)

here2:
Select Case dectys
    Case 0
        If tys = 0 Then GoTo here3
        If tys = 1 Then tys_txt = Nums4(tys) & "òèñÿ÷à "
        If tys = 2 Or tys = 3 Or tys = 4 Then tys_txt = Nums4(tys) & "òèñÿ÷³ "
        If tys = 5 Or tys = 6 Or tys = 7 Or tys = 8 Or tys = 9 Then tys_txt = Nums4(tys) & "òèñÿ÷ "
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9
        If tys = 0 Then dectys_txt = Nums2(dectys) & "òèñÿ÷ "
        If tys = 1 Then tys_txt = Nums2(dectys) & Nums4(tys) & "òèñÿ÷à "
        If tys = 2 Or tys = 3 Or tys = 4 Then tys_txt = Nums2(dectys) & Nums4(tys) & "òèñÿ÷³ "
        If tys = 5 Or tys = 6 Or tys = 7 Or tys = 8 Or tys = 9 Then tys_txt = Nums2(dectys) & Nums4(tys) & "òèñÿ÷ "
End Select

here3:
Select Case sot
Case 0
    If dec = 0 Then GoTo here4
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9
        sot_txt = Nums3(sot)
    End Select
Select Case dec
    Case 0
    dec_txt = ""
    Case 1
        If ed <> 0 Then dec_txt = Nums5(ed)
    Case 2, 3, 4, 5, 6, 7, 8, 9
        dec_txt = Nums2(dec)
End Select

here4:
If ed = 0 Then ed_txt = ""
If ed <> 0 Then ed_txt = Nums1(ed)

otvet = dectys_txt & tys_txt & sot_txt & dec_txt & ed_txt
' çàïèñ â³äïîâ³ä³
here10:
file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäå çàïèñóâàòèñü ÷èñëî cëîâàìè")
If file <> False Then
    f3 = file
Else
    MsgBox ("Áóäü-ëàñêà, îáåð³òü ôàéë")
    GoTo here10
End If

intfh = FreeFile()

Open f3 For Output As #intfh
    Print #intfh, otvet
Close #intfh

End Sub

Function Class(M, I)

    Class = Int(Int(M - (10 ^ I) * Int(M / (10 ^ I))) / 10 ^ (I - 1))

End Function
