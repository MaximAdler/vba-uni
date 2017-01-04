Private Sub CommandButton1_Click()

    Dim str$, result$
    Dim arr() As String

    result = ""
    FF = FreeFile()

 'Change before checking
    Open "C:\Users\Admin\Desktop\s12\test.txt" For Input As #FF
        Do While Not EOF(FF)
            Line Input #FF, str
            arr = Split(str)
            For I = LBound(arr) To UBound(arr)
                If Left(arr(I), 1) = Right(arr(I), 1) Then
                    If Len(arr(I)) > 1 Then
                        result = result + arr(I) + " "
                    End If
                End If
            Next I
 
        Loop
    Close #FF
    
    TextBox1.text = result

End Sub
