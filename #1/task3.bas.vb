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
Private Sub CommandButton1_Click()
    
    Dim i, j, a, b, p, k, h, l As Long
    Dim n, m, min As Integer
    
    On Error GoTo Error
    Worksheets("19").Activate
    ActiveSheet.Cells.Clear
    ActiveSheet.Cells.Font.Name = "Times New Roman"
Error1:
    i = TextBox1.text
    
    If (i - CInt(i)) <> 0 Then
        MsgBox "Ââåä³òü ö³ëå ÷èñëî"
    End If
    If i > 1048576 Then
        MsgBox "Çàäàíå ÷èñëî ïåðåâèùóº ê³ëüê³ñòü ðÿäê³â íà ëèñò³"
    End If
    If Sgn(i) = -1 Then
        MsgBox "Ê³ëüê³ñòü ðÿäê³â íå ìîæå áóòè â³ä'ºìíèì ÷èñëîì"
    End If
    If i = 0 Then
        MsgBox "Ê³ëüê³ñòü ðÿäê³â íå ìîæå äîð³âíþâàòè íóëþ!"
    End If
    
Error2:
    j = TextBox2.text
    If (j - CInt(j)) <> 0 Then
        MsgBox "Ââåä³òü ö³ëå ÷èñëî"
    End If
    If j > 16384 Then
        MsgBox "Çàäàíå ÷èñëî ïåðåâèùóº ê³ëüê³ñòü ñòîâïö³â íà ëèñò³"
    End If
    If Sgn(j) = -1 Then
        MsgBox "Ê³ëüê³ñòü ñòîâïö³â íå ìîæå áóòè â³ä'ºìíèì ÷èñëîì"
    End If
    If j = 0 Then
        MsgBox "Ê³ëüê³ñòü ñòîâïö³â íå ìîæå äîð³âíþâàòè íóëþ"
    End If
    For a = 1 To i
        For b = 1 To j
            Cells(a, b).Formula = Int((100 * Rnd()) + 1)
            Next b
            Next a
            min = 100
            For a = 1 To i
                For b = 1 To j
                    For n = 0 To i
                        For m = 0 To j
                            If Abs(Cells(a, b) - Cells(a + n, b + m)) < min And m - n <> 0 Then
                                min = Abs(Cells(a, b) - Cells(a + n, b + m))
                                If a + n <= i And b + m <= j Then
                                    Cells(i + 2, 1) = Cells(a, b)
                                    Cells(i + 3, 1) = Cells(a + n, b + m)
                                End If
                            End If
                            Next m
                            Next n
                            Next b
                            Next a
                            
                            Exit Sub
Error:
                            MsgBox "Ïåðåâ³ðòå ïðàâèëüí³ñòü ââåäåíèõ äàíèõ!"
                        End Sub
                        
