Private Sub CommandButton1_Click()
    
    Dim n As Long
    Dim m As Long
    Dim s As Long
    Dim i As Long
    Dim q As String
    
    i = 3
    Worksheets("11").Activate
    Range(Cells(2, 2), Cells(56000, 3)) = Null
    Cells(2, 2) = "n"
    Cells(2, 3) = "Valid m"
w1:
    q = TextBox2.text
    If q = "" Then
        MsgBox "Âè íå çàäàëè çíà÷åííÿ n"
        Exit Sub
    End If
    
    If Not IsNumeric(q) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ!"
        Exit Sub
    End If
    n = q
    If n = 0 Or n < 1 Then
        MsgBox "Íå ïðàâèëüíî çàäàíå íàòóðàëüíå ÷èñëî"
    End If
w2:
    q = TextBox1.text
    If q = "" Then
        MsgBox "Âè íå çàäàëè çíà÷åííÿ m"
        Exit Sub
    End If
    
    
    If Not IsNumeric(q) Then
        MsgBox "Ââåä³òü ÷èñëîâå çíà÷åííÿ!"
    End If
    m = q
    For s = n To 1 Step -1
        Cells(i, 2) = s
        i = i + 1
        Next s
        
        i = 3
        
        For s = 1 To n
            If m = (s + s) ^ 2 Then
                Cells(i, 3) = s
                i = i + 1
            End If
            Next s
            
            
        End Sub
        


' Js code

'function Ex1(n,m){
'  var arr = [];
'  var resultArr = [];
'  for(var i=1;i<n;i++){
'    arr.push (i)
'  }
'  for(var j=0;j<arr.length;j++){
'    if(arr[j].toString().length==1 && Math.pow(arr[j],2)==m){
'      resultArr.push(arr[j])
'    }
'    else if(arr[j].toString().length>1 && Math.pow(arr[j].toString().split('').reduce(function(a,b){return Number(a)+Number(b)}),2)==m){
'      resultArr.push(arr[j])
'    }
'  }
'  return resultArr.length>0?resultArr.join(','):'No result';
'};
