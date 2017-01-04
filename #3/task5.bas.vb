Private Sub CommandButton1_Click()

  Dim MyDictionary As Object
  Set MyDictionary = CreateObject("Scripting.Dictionary")
  FF = FreeFile()
  
  ' Change before checking
    Open "C:\Users\Admin\Desktop\s12\test.txt" For Input As #FF
    Do While Not EOF(FF)
        Line Input #FF, s
        Dim larr() As String
        larr = Split(s, " ")
  
        For I = LBound(larr) To UBound(larr)
    
            With MyDictionary
                If Not .Exists(larr(I)) Then
                    .Add larr(I), 1
                Else
                    .Item(larr(I)) = .Item(larr(I)) + 1
                End If
            End With
        Next I
    Loop
  Close #FF
  
  
  For Each x In MyDictionary
    If MyDictionary.Item(x) > 1 Then
      TextBox1.text = TextBox1.text + CStr(x) + "(" + CStr(MyDictionary.Item(x)) + "), "
      
    End If
  Next
End Sub