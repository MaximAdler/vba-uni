Private Sub CommandButton1_Click()
    
    Dim i As Integer
    Dim j As Integer
    Dim str1 As String
    Dim str2 As String
    Dim str As String
    
    str1 = TextBox1.text
    str2 = TextBox2.text
    str = "Çàäàíî: " + CStr(str1) + "/" + CStr(str2) + "  Ðåçóëüòàò: "
    
    If str1 = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ ÷èñåëüíèêà"
        Exit Sub
    End If
    If str2 = "" Then
        MsgBox "Ââåä³òü çíà÷åííÿ çíàìåííèêà"
        Exit Sub
    End If
    If Not IsNumeric(str1) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ!"
        Exit Sub
    End If
    If Not IsNumeric(str2) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ!"
        Exit Sub
    End If
    If str2 > 7 Then
        MsgBox "Íå çàäîâ³ëüíÿº óìîâó!"
        Exit Sub
    End If
    If str1 > str2 Then
        MsgBox "Íå çàäîâ³ëüíÿº óìîâó!"
        Exit Sub
    End If

    
    For i = 2 To str2
        For j = 1 To i - 1
            str = str + " " + CStr(j) + "/" + CStr(i)
        Next j
        Label2.Caption = str
    Next i
End Sub
