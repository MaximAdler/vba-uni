﻿Private Sub CommandButton1_Click()
    
    Dim i&, j&, k&, l&, r&, y&, n&, m&
    Dim arr()
    
    jrr = TextBox1.text
    If jrr = "" Then
        MsgBox "Ââåä³òü ÷èñëî"
        Exit Sub
    End If
    If Not IsNumeric(jrr) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ!"
        Exit Sub
    End If
    If jrr < 0 Then
        MsgBox "Ââåä³òü äîäàòíå çíà÷åííÿ!"
        Exit Sub
    End If
    If jrr = 0 Then
        MsgBox "×èñëî ìàº áóòè á³ëüøå íóëÿ!"
        Exit Sub
    End If
    
    ReDim arr(1 To jrr)
    arr(1) = 3
    l = 1
    For i = 5 To jrr Step 2
        k = Abs(Int(-i ^ 0.5))
        For y = 1 To jrr
            If arr(y) >= k Then
                m = y
                Exit For
            End If
            Next y
            For n = 1 To m
                If i Mod arr(n) = 0 Then Exit For
                r = l + 1
                Next n
                If r - l = 1 Then
                    l = l + 1
                    arr(l) = i
                End If
                Next i
                Label3.Caption = Join(arr)
                
End Sub