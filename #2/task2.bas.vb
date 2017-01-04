Private Sub CommandButton1_Click()

    Dim str$, I%, result$, complited$, b$, count%
    
    str = TextBox1.text
    
    str = Replace(str, " ", "")
    str = Replace(str, ", ", "")
    str = Replace(str, ".", "")
    str = Replace(str, ",", "")
    str = Trim(str)
    
    result = ""
    complited = ""
 
    For I = 1 To Len(str)
        count = 1
        b = Mid(str, I, 1)
        For j = I + 1 To Len(str)
            If b = Mid(str, j, 1) Then count = count + 1
        Next j
        If count = 1 And InStr(complited, b) = 0 Then
            result = result + b + " "
        Else
            complited = complited + b
        End If
    Next I
    Label2.Caption = "Íåïîâòîðþâàí³ ë³òåðè: " + CStr(result)

End Sub
