﻿Private Sub CommandButton1_Click()

    Dim n As String
    Dim s As String
    Dim result As Double
    
    n = TextBox1.text
    s = TextBox2.text
    
    If n = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ ÷èñëî"
        Exit Sub
    End If
    If s = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ ñòåï³íü"
        Exit Sub
    End If
    If Not IsNumeric(n) Then
        MsgBox "Íå ïðàâèëüíî ââåäåíî ÷èñëî"
        Exit Sub
    End If
    If Not IsNumeric(s) Then
        MsgBox "Íå ïðàâèëüíî ââåäåíî ñòåï³íü"
        Exit Sub
    End If
    
    result = CStr(n) ^ CStr(s)
    Label4.Caption = "Â³äïîâ³äü: " + CStr(result)

End Sub
