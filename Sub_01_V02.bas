Attribute VB_Name = "SubLib_01"
Sub SuperscriptBulk(inRange As Range, inArr As Variant)
    n = UBound(inArr)
    For i = 0 To n
        curr_inx = inArr(i)
        With inRange.Characters(start:=curr_inx, Length:=1).Font
            .Superscript = True
        End With
        
    Next i
End Sub
Sub SubscriptBulk(inRange As Range, inArr As Variant)
    n = UBound(inArr)
    For i = 0 To n
        curr_inx = inArr(i)
        With inRange.Characters(start:=curr_inx, Length:=1).Font
            .Subscript = True
            
        End With
        
    Next i
End Sub


Sub PrintColorCode()

End Sub
Sub ColorIf()
'Can run but need more features
    Dim arr_of_checker() As Variant
    arr_of_checker = Array("a", "e", "i", "o", "u")
    Dim inRange As Range
    Set inRange = Application.InputBox(Title:="Select Range", Prompt:="Please select Range for coloring", Type:=8)
    For Each curr_cell In inRange
        curr_string = curr_cell.Value
        first_ch = Left(curr_string, 1)
        If A_isInArr(arr_of_checker, first_ch) Then
            curr_cell.Interior.Color = vbYellow
        End If
        
    Next curr_cell
    
    
End Sub
Sub ColorFontFromTo()

End Sub
Sub ColorFontAt()
    Dim inRange As Range
    On Error Resume Next
    Set inRange = Application.InputBox(Prompt:="Please select your range for coloring", Type:=8)
    On Error GoTo 0
    If inRange Is Nothing Then Exit Sub
    
    
    inx = InputBox("Enter the index: ")
    'inx = 3
    For Each curr_cell In inRange
        With curr_cell.Characters(start:=inx, Length:=1).Font
            .Color = vbBlue
        End With
    Next curr_cell
    
    

End Sub
Sub ColorSubString()
'Add: More words
'Add: Custom color
'Add: Color multipleTimes
    Dim inRange As Range
    On Error Resume Next
    Set inRange = Application.InputBox(Prompt:="Please select your range for coloring", Type:=8)
    Set word_list = Application.InputBox(Prompt:="Please select your range for coloring", Type:=8)
    On Error GoTo 0
    If inRange Is Nothing Then Exit Sub
    If word_list Is Nothing Then Exit Sub
    
    wordToColor = word_list.Value
    n_word = Len(wordToColor)
    myColor = word_list.Font.Color
    
    For Each curr_cell In inRange
        curr_str = curr_cell.Value
        inx = InStr(1, curr_str, wordToColor, vbTextCompare)
        If inx <> 0 Then
            With curr_cell.Characters(start:=inx, Length:=n_word).Font
                .Color = myColor
            End With
        End If
        
    Next
    
    
    
    
    
End Sub
Sub ColorFontWith()
    

End Sub

Sub BoldFontAt()

End Sub

Sub BoldFontWith()

End Sub
