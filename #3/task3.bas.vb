Private Sub CommandButton1_Click()
    
    Dim p$, t1$, t2$
    p = TextBox1.Text
    t1 = TextBox2.Text
    t2 = TextBox3.Text
    
    If p = "" Then
        MsgBox "Ââåä³òü äàí³!"
        Exit Sub
    End If
    If t1 = "" Then
        MsgBox "Ââåä³òü äàí³!"
        Exit Sub
    End If
    If t2 = "" Then
        MsgBox "Ââåä³òü äàí³!"
        Exit Sub
    End If
    If p < 25 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    If t1 < 70 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    If p < 35 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    If p > 180 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    If t1 > 215 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    If t2 > 180 Then
        MsgBox "Çäàºòüñÿ, ëþäèí³ ïîãàíî!"
        Exit Sub
    End If
    
    If p > 80 And t1 > 135 And t2 > 95 Then
        Label6.Caption = Label6.Caption + "Ëþäèíà áðåøå!"
    Else
        Label6.Caption = Label6.Caption + "Ëþäèíà êàæå ïðàâäó!"
    End If
    
End Sub

