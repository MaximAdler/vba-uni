﻿Private Sub CommandButton1_Click()
    
    Dim a As String
    Dim b As String
    Dim k As String
    Dim k2 As Long
    Dim k1 As Long
    Dim j As Long
    Dim i As Long
    Dim str As String
    
    a = TextBox1.text
    b = TextBox2.text
    k = TextBox3.text
    k2 = 0
    str = "Â³äïîâ³äü: "
    
    If a > b Then
        MsgBox "Íå ïðàâèëüíî ââåäåí³ äàí³!"
        Exit Sub
    End If
    If a = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ a"
        Exit Sub
    End If
    If b = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ b"
        Exit Sub
    End If
    If k = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ k"
        Exit Sub
    End If
    If Not IsNumeric(a) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ"
        Exit Sub
    End If
    If Not IsNumeric(b) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ"
        Exit Sub
    End If
    If Not IsNumeric(k) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ"
        Exit Sub
    End If
    
    For i = a To b
        If i = 1 Then
            k1 = 1
        Else
            k1 = 2
        End If
        For j = 2 To Int(i / 2)
            If i Mod j = 0 Then
                k1 = k1 + 1
            End If
            Next j
            If k1 = k Then
                str = str + " " + CStr(i) + " "
                k2 = k2 + 1
            End If
            Label5.Caption = str
            Next i
End Sub
