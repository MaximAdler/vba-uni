﻿Private Sub CommandButton1_Click()
    
    Dim i As Single
    Dim j As Single
    Dim m As Single
    Dim n As Single
    Dim l As Single
    
    On Error GoTo Error
    Worksheets("15").Activate
    ActiveSheet.Cells.Clear
    ActiveSheet.Cells.Font.Name = "Times New Roman"
Error1:
    i = TextBox1.text
    If (i - CInt(i)) <> 0 Then
        MsgBox "×èñëî ðÿäê³â ïîâèííî áóòè ö³ëèì ÷èñëîì"
        GoTo Error1
    End If
    If i > 1048576 Then
        MsgBox "Çàäàíà ê³ëüê³ñòü ðÿäê³â á³ëüøà äîïóñòèìîãî çíà÷åííÿ"
        GoTo Error1
    End If
    If Sgn(i) = -1 Then
        MsgBox "×èñëî ðÿäê³â íå ìîæå áóòè â³ä'ºìíèì"
        GoTo Error1
    End If
    If i = 0 Then
        MsgBox "×èñëî ðÿäê³â íå ìîæå äîð³âíþâàòè 0"
        GoTo Error1
    End If
Error2:
    j = TextBox2.text
    If (j - CInt(j)) <> 0 Then
        MsgBox "Ê³ëüê³ñòü ðÿäê³â íå ìîæå áóòè íå ö³ëèì ÷èñëîì"
        GoTo Error2
    End If
    If j > 16384 Then
        MsgBox "Çàäàíà ê³ëüê³ñòü ñòîâïö³â á³ëüøà äîïóñòèìîãî çíà÷åííÿ"
        GoTo Error2
    End If
    If Sgn(j) = -1 Then
        MsgBox "×èñëî ñòîâïö³â íå ìîæå áóòè â³ä'ºìíèì"
        GoTo Error2
    End If
    If j = 0 Then
        MsgBox "Ê³ëüê³ñòü ñòîâïö³â íå ìîæå äîð³âíþâàòè íóëþ!"
        GoTo Error2
    End If
Error3:
    l = TextBox3.text
    If (l - CInt(l)) <> 0 Then
        MsgBox "Ê³ëüê³ñòü çíàê³â ï³ñëÿ êîìè íå ìîæå áóòè íå ö³ëèì ÷èñëîì"
        GoTo Error3
    End If
    If Sgn(l) = -1 Then
        MsgBox "Ê³ëüê³ñòü çíàê³â ï³ñëÿ êîìè íå ìîæå áóòè â³ä'ºìíîþ"
        GoTo Error3
    End If
    For m = 1 To i
        For n = 1 To j
            If (21.8 * n) / (3.2 * m) ^ 3 > 20 Then
                Cells(m, n).Formula = FormatNumber(11.3 * m / n, l)
                Else: Cells(m, n).Formula = FormatNumber((21.8 * n) / (3.2 * m) ^ 3, l)
            End If
            Next n
            Next m
            Exit Sub
Error:
            MsgBox "Ñïðîáóéòå ùå"
        End Sub
