Private Sub CommandButton1_Click()
    
    Dim w0&, i%, n%, j%
    Dim w1() As Long
    Dim c() As Long
    Dim dis() As Long
    Dim gain() As Long
    Dim result() As String
    
    Worksheets("human capital model").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Salary"
    ActiveSheet.Cells(1, 3) = "Expenses for relocation"
    ActiveSheet.Cells(1, 4) = "Discount"
    ActiveSheet.Cells(1, 5) = "Profit"
    ActiveSheet.Cells(1, 6) = "Choice"
    ActiveSheet.Cells(2, 1) = "Ukraine"
    ActiveSheet.Cells(2, 2) = 8000
    
    w0 = Range("B2").Value

    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim w1(n) As Long
    ReDim c(n) As Long
    ReDim dis(n) As Long
    ReDim gain(n) As Long
    ReDim result(n) As String
    
    For i = 2 To n
        w1(i) = ActiveSheet.Cells(i + 1, 2)
        c(i) = ActiveSheet.Cells(i + 1, 3)
        dis(i) = ActiveSheet.Cells(i + 1, 4)
        ActiveSheet.Cells(i + 1, 5) = ((w1(i) - w0) / (1 + dis(i))) - c(i)
        If ActiveSheet.Cells(i + 1, 5) > 0 Then
            ActiveSheet.Cells(i + 1, 6) = "Òàê"
        Else
            ActiveSheet.Cells(i + 1, 6) = "Í³"
        End If
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4), ActiveSheet.Cells(j + 1, 5), ActiveSheet.Cells(j + 1, 6)
            Next j
    Close #intfh
    
End Sub

Private Sub CommandButton11_Click()

    Dim w0&, i%, n%, j%
    Dim w1() As Long
    Dim c() As Long
    Dim gain() As Long
    Dim result() As String
    
    Worksheets("pull-push migration").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Salary"
    ActiveSheet.Cells(1, 3) = "Expenses for relocation"
    ActiveSheet.Cells(1, 4) = "Max profit"
    ActiveSheet.Cells(1, 5) = "Choice"
    ActiveSheet.Cells(2, 1) = "Ukraine"
    ActiveSheet.Cells(2, 2) = 8000
    
    w0 = Range("B2").Value
    
    '   For adding new item
    'colB = ActiveSheet.Columns("B").Rows(65536).End(xlUp).Row + 1
    'colC = ActiveSheet.Columns("C").Rows(65536).End(xlUp).Row + 1
    'ActiveSheet.Cells(colB, 2) = w1
    'ActiveSheet.Cells(colC, 3) = c
    
    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim w1(n) As Long
    ReDim c(n) As Long
    ReDim gain(n) As Long
    ReDim result(n) As String
    
    For i = 2 To n
        w1(i) = ActiveSheet.Cells(i + 1, 2)
        c(i) = ActiveSheet.Cells(i + 1, 3)
        ActiveSheet.Cells(i + 1, 4) = 0.8 * c(i) + 0.8 * (w1(i) - w0)
        If ActiveSheet.Cells(i + 1, 4) > 0 Then
            ActiveSheet.Cells(i + 1, 5) = "Òàê"
        Else
            ActiveSheet.Cells(i + 1, 5) = "Í³"
        End If
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4), ActiveSheet.Cells(j + 1, 5)
            Next j
    Close #intfh
    
End Sub

Private Sub CommandButton13_Click()

    Dim i%, n%, j%
    Dim pop() As Long
    Dim migr() As Long
    Dim result() As String
    
    Worksheets("assimilation model(import)").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Number of Ukrainian migrants"
    ActiveSheet.Cells(1, 3) = "Population"
    ActiveSheet.Cells(1, 4) = "The coefficient of assimilation"
    
     file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

    If file <> False Then
        f = file
    Else
        MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
        Exit Sub
    End If

    intfh = FreeFile()
    Open f For Input As intfh
            row_number = 0
            Do Until EOF(intfh)
                Line Input #intfh, s
                    LineItems = Split(s, ",")
                    Cells(2, 1).Activate
                    ActiveCell.Offset(row_number, 0).Value = LineItems(0)
                    ActiveCell.Offset(row_number, 1).Value = LineItems(1)
                    ActiveCell.Offset(row_number, 2).Value = LineItems(2)
                    row_number = row_number + 1
            Loop
    Close #intfh
    
    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim pop(n) As Long
    ReDim migr(n) As Long
    ReDim result(n) As String
    
    For i = 1 To n
        pop(i) = ActiveSheet.Cells(i + 1, 3)
        migr(i) = ActiveSheet.Cells(i + 1, 2)
        ActiveSheet.Cells(i + 1, 4) = migr(i) / pop(i)
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4)
            Next j
    Close #intfh

End Sub

Private Sub CommandButton15_Click()

    Dim i%, n%, j%
    Dim pop() As Long
    Dim migr() As Long
    Dim result() As String
    
    Worksheets("assimilation model").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Number of Ukrainian migrants"
    ActiveSheet.Cells(1, 3) = "Population"
    ActiveSheet.Cells(1, 4) = "The coefficient of assimilation"
    
    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim pop(n) As Long
    ReDim migr(n) As Long
    ReDim result(n) As String
    
    For i = 1 To n
        pop(i) = ActiveSheet.Cells(i + 1, 3)
        migr(i) = ActiveSheet.Cells(i + 1, 2)
        ActiveSheet.Cells(i + 1, 4) = migr(i) / pop(i)
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4)
            Next j
    Close #intfh

End Sub

Private Sub CommandButton2_Click()
 
    Dim w0&, i%, n%, j%
    Dim w1() As Long
    Dim c() As Long
    Dim dis() As Long
    Dim gain() As Long
    Dim result() As String
    
    Worksheets("human capital model(import)").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Salary"
    ActiveSheet.Cells(1, 3) = "Expenses for relocation"
    ActiveSheet.Cells(1, 4) = "Discount"
    ActiveSheet.Cells(1, 5) = "Profit"
    ActiveSheet.Cells(1, 6) = "Choice"
    ActiveSheet.Cells(2, 1) = "Ukraine"
    ActiveSheet.Cells(2, 2) = 8000
    
    w0 = Range("B2").Value

    file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

    If file <> False Then
        f = file
    Else
        MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
        Exit Sub
    End If

    intfh = FreeFile()
    Open f For Input As intfh
            row_number = 0
            Do Until EOF(intfh)
                Line Input #intfh, s
                    LineItems = Split(s, ",")
                    Cells(3, 1).Activate
                    ActiveCell.Offset(row_number, 0).Value = LineItems(0)
                    ActiveCell.Offset(row_number, 1).Value = LineItems(1)
                    ActiveCell.Offset(row_number, 2).Value = LineItems(2)
                    ActiveCell.Offset(row_number, 3).Value = LineItems(3)
                    row_number = row_number + 1
            Loop
    Close #intfh

    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim w1(n) As Long
    ReDim c(n) As Long
    ReDim dis(n) As Long
    ReDim gain(n) As Long
    ReDim result(n) As String
    
    For i = 2 To n
        w1(i) = ActiveSheet.Cells(i + 1, 2)
        c(i) = ActiveSheet.Cells(i + 1, 3)
        dis(i) = ActiveSheet.Cells(i + 1, 4)
        ActiveSheet.Cells(i + 1, 5) = ((w1(i) - w0) / (1 + dis(i))) - c(i)
        If ActiveSheet.Cells(i + 1, 5) > 0 Then
            ActiveSheet.Cells(i + 1, 6) = "Òàê"
        Else
            ActiveSheet.Cells(i + 1, 6) = "Í³"
        End If
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4), ActiveSheet.Cells(j + 1, 5), ActiveSheet.Cells(j + 1, 6)
            Next j
    Close #intfh
    
End Sub

Private Sub CommandButton5_Click()
Dim i%, n%, j%
    Dim imigrant() As Long
    Dim speak() As Long
    Dim result() As String
    
    Worksheets("diffusion migration(import)").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Ukrainian immigrants"
    ActiveSheet.Cells(1, 3) = "Ukrainian, that contact with others"
    ActiveSheet.Cells(1, 4) = "Coefficient"
    
    file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

    If file <> False Then
        f = file
    Else
        MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
        Exit Sub
    End If

    intfh = FreeFile()
    Open f For Input As intfh
            row_number = 0
            Do Until EOF(intfh)
                Line Input #intfh, s
                    LineItems = Split(s, ",")
                    Cells(2, 1).Activate
                    ActiveCell.Offset(row_number, 0).Value = LineItems(0)
                    ActiveCell.Offset(row_number, 1).Value = LineItems(1)
                    ActiveCell.Offset(row_number, 2).Value = LineItems(2)
                    row_number = row_number + 1
            Loop
    Close #intfh
    
    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim imigrant(n) As Long
    ReDim speak(n) As Long
    ReDim result(n) As String
    
    For i = 1 To n
        imigrant(i) = ActiveSheet.Cells(i + 1, 2)
        speak(i) = ActiveSheet.Cells(i + 1, 3)
        ActiveSheet.Cells(i + 1, 4) = imigrant(i) / speak(i)
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4)
            Next j
    Close #intfh
End Sub

Private Sub CommandButton7_Click()

    Dim i%, n%, j%
    Dim imigrant() As Long
    Dim speak() As Long
    Dim result() As String
    
    Worksheets("diffusion migration").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Ukrainian immigrants"
    ActiveSheet.Cells(1, 3) = "Ukrainian, that contact with others"
    ActiveSheet.Cells(1, 4) = "Coefficient"
    
    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim imigrant(n) As Long
    ReDim speak(n) As Long
    ReDim result(n) As String
    
    For i = 1 To n
        imigrant(i) = ActiveSheet.Cells(i + 1, 2)
        speak(i) = ActiveSheet.Cells(i + 1, 3)
        ActiveSheet.Cells(i + 1, 4) = imigrant(i) / speak(i)
    Next i

    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäå çàïèñóâàòèñü ÷èñëî àðàáñüêèìè öèôðàìè")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4)
            Next j
    Close #intfh
End Sub

Private Sub CommandButton9_Click()
    
    Dim w0&, i%, n%, s$, j%
    Dim f As Variant
    Dim file As Variant
    Dim w1() As Long
    Dim c() As Long
    Dim gain() As Long
    Dim result() As String
    
    Worksheets("pull-push migration(import)").Activate
    ActiveSheet.Cells(1, 1) = "Country"
    ActiveSheet.Cells(1, 2) = "Salary"
    ActiveSheet.Cells(1, 3) = "Expenses for relocation"
    ActiveSheet.Cells(1, 4) = "Max profit"
    ActiveSheet.Cells(1, 5) = "Choice"
    ActiveSheet.Cells(2, 1) = "Ukraine"
    ActiveSheet.Cells(2, 2) = 8000
    
    w0 = Range("B2").Value


    file = Application.GetOpenFilename("TextFiles(*.txt), *.txt", 0, "Îáåð³òü òåêñòîâèé ôàéë")

    If file <> False Then
        f = file
    Else
        MsgBox "Îáåð³òü òåêñòîâèé ôàéë"
        Exit Sub
    End If

    intfh = FreeFile()
    Open f For Input As intfh
            row_number = 0
            Do Until EOF(intfh)
                Line Input #intfh, s
                    LineItems = Split(s, ",")
                    Cells(3, 1).Activate
                    ActiveCell.Offset(row_number, 0).Value = LineItems(0)
                    ActiveCell.Offset(row_number, 1).Value = LineItems(1)
                    ActiveCell.Offset(row_number, 2).Value = LineItems(2)
                    row_number = row_number + 1
            Loop
    Close #intfh

    i = 1
    Do While ActiveSheet.Cells(i + 1, 1) <> ""
        i = i + 1
    Loop
    n = i - 1
    
    ReDim w1(n) As Long
    ReDim c(n) As Long
    ReDim gain(n) As Long
    ReDim result(n) As String
    
    For i = 2 To n
        w1(i) = ActiveSheet.Cells(i + 1, 2)
        c(i) = ActiveSheet.Cells(i + 1, 3)
        ActiveSheet.Cells(i + 1, 4) = 0.8 * c(i) + 0.8 * (w1(i) - w0)
        If ActiveSheet.Cells(i + 1, 4) > 0 Then
            ActiveSheet.Cells(i + 1, 5) = "Òàê"
        Else
            ActiveSheet.Cells(i + 1, 5) = "Í³"
        End If
    Next i
    
    file = Application.GetOpenFilename("Text files(*.txt),*.txt", 0, "Îáåð³òü ôàéë, êóäè áóäóòü çàâàíòàæåí³ äàí³!")
    If file <> False Then
        f1 = file
    Else
        Exit Sub
    End If
    intfh = FreeFile()
    Open f1 For Output As #intfh
            For j = 0 To 100
                Print #intfh, ActiveSheet.Cells(j + 1, 1), ActiveSheet.Cells(j + 1, 2), ActiveSheet.Cells(j + 1, 3), ActiveSheet.Cells(j + 1, 4), ActiveSheet.Cells(j + 1, 5)
            Next j
    Close #intfh



End Sub

Private Sub UserForm_Activate()
    Image1.Picture = LoadPicture("C:\Users\Admin\Desktop\Îðåë_Ì\pictures\oboi.jpg")
    Image2.Picture = LoadPicture("C:\Users\Admin\Desktop\Îðåë_Ì\pictures\oboi.jpg")
    Image3.Picture = LoadPicture("C:\Users\Admin\Desktop\Îðåë_Ì\pictures\oboi.jpg")
    Image4.Picture = LoadPicture("C:\Users\Admin\Desktop\Îðåë_Ì\pictures\oboi.jpg")
    Image5.Picture = LoadPicture("C:\Users\Admin\Desktop\Îðåë_Ì\pictures\oboi.jpg")
End Sub
