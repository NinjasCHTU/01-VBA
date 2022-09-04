Attribute VB_Name = "Lib_Dear1"
Function D_isItInWS(ws_name, elem, Optional n = 1000, Optional ans_type = 0)
'''''''''''''''''''''''''''''''''''''''''''''(Not Done)
'@@@@@@@@@@@@@@@@Dependency -> Space (St_SelectFromTL)
    Dim first_cell, ws_range As Range
    Set first_cell = Worksheets(ws_name).Range("A1")
    

'Stuck here: can't get the ws_range to work
    Set ws_range = St_SelectFromTL(first_cell, 5, 5)
    D_isItInWS = ws_range
    

End Function

Function D_isItInRange(rangeIn, elem)
'@@@@@@@@@@@@@@@@Dependency -> no
    res = False
    For Each curr_cell In rangeIn
        If curr_cell.Value = elem Then
            res = curr_cell.Address
        End If
        
        
    Next curr_cell
    D_isItInRange = res
    

End Function
Function D_AlphaSmallV(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(97 + i)
    Next i
    D_AlphaSmallV = WorksheetFunction.Transpose(arr)
End Function

Function D_AlphaSmallH(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(97 + i)
    Next i
    D_AlphaSmallH = arr

End Function

Function D_AlphaBigV(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(65 + i)
    Next i
    D_AlphaBigV = WorksheetFunction.Transpose(arr)

End Function
Function D_AlphaBigH(n)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr() As Variant
    ReDim arr(n - 1)
    For i = 0 To n - 1
        arr(i) = WorksheetFunction.Unichar(65 + i)
    Next i
    D_AlphaBigH = (arr)

End Function
Function D_Combination(ParamArray range_tab())
    '@@@@@@@@@@@@@@@@Dependency ->
' Lib_WSFunc(VB_Transpose)
    temp_tab = D_Combi2(range_tab(0), range_tab(1), 3)
    
    For i = 2 To UBound(range_tab)
        temp_tab = D_Combi2(temp_tab, range_tab(i), 3)
    Next i
    
    D_Combination = VB_Transpose(temp_tab)
End Function

Function D_Combi2(arr1, arr2, Optional choice = 2)
'@@@@@@@@@@@@@@@@Dependency ->
' Lib_WSFunc(VB_Transpose)
    If TypeOf arr1 Is Range Then
        row_n = arr1.Count
    Else
        row_n = UBound(arr1)
    End If
    
    If TypeOf arr2 Is Range Then
        col_n = arr2.Count
    Else
        col_n = UBound(arr2)
    End If
    
    
    Dim arr2d(3, 2) As Variant
    
    arr2d_new = D_2dArray(arr2d, row_n, col_n)
    Dim arr1d() As Variant
    ReDim Preserve arr1d(row_n * col_n - 1)
    
    'k is the index for arr1d
    k = 0
    '1&2 is the dimention
    For i = 0 To UBound(arr2d_new, 1)
        For j = 0 To UBound(arr2d_new, 2)
            If TypeOf arr1 Is Range Then
                new_elem = arr1(i + 1) & " " & arr2(j + 1)
            Else
                new_elem = arr1(i + 1) & " " & arr2(j + 1)
            End If
            
            arr2d_new(i, j) = new_elem
            arr1d(k) = new_elem
            k = k + 1
        Next j
    Next i
    
    If choice = 1 Then
        D_Combi2 = VB_Transpose(arr1d)
    ElseIf choice = 2 Then
        D_Combi2 = arr2d_new
    Else
        D_Combi2 = arr1d
    End If
    
    
End Function

Function D_2dArray(name, row, col)
  '@@@@@@@@@@@@@@@@Dependency -> No
    D_2dArray = ReDimPreserve(name, row - 1, col - 1)
End Function


Function D_sheetTOArr(table)
'Not Done
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim arr(11) As Variant
    row = table.Rows.Count
    col = table.Columns.Count
    arr02 = D_2dArray(arr, row, col)
    i = 0
    For Each c In table
        arr(i) = c
        i = i + 1
    Next c
    D_sheetTOArr = arr
End Function
Function make_key_num()
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim key_num As Dictionary
    Set key_num = New Dictionary
    key_num("C") = 1
    key_num("C#") = 2
    key_num("D") = 3
    key_num("D#") = 4
    key_num("E") = 5
    key_num("F") = 6
    key_num("F#") = 7
    key_num("G") = 8
    key_num("Ab") = 9
    key_num("A") = 10
    key_num("Bb") = 11
    key_num("B") = 12
    
    key_num("Db") = 2
    key_num("Eb") = 4
    key_num("Gb") = 7
    key_num("G#") = 9
    key_num("A#") = 11
    Set make_key_num = key_num

End Function
Function make_num_2_key()
  '@@@@@@@@@@@@@@@@Dependency -> No
    Dim num_2_key() As Variant
    num_2_key = Array("C", "C#", "D", "D#", "E", "F", "F#", "G", "Ab", "A", "Bb", "B")
    make_num_2_key = num_2_key

End Function

Function M_ChangeKey2(chord, shift)
      '@@@@@@@@@@@@@@@@Dependency -> No
    Dim key_num As Dictionary
    
    num_2_key = make_num_2_key
    Set key_num = make_key_num
    
    If TypeOf chord Is Range Then
        old_val = key_num(chord.Value)
    Else
        old_val = key_num(chord)
    
    End If
    
    new_val = (old_val + shift) Mod 12
    
    If new_val <= 0 Then
        new_val = 12 + new_val
    End If
    
    'M_ChangeKey2 = new_val
    M_ChangeKey2 = num_2_key(new_val - 1)
    
    'MsgBox (key_num("b"))
    
End Function
Function M_ChangeKeyVB(chord_txt, shift)
'$$$$$$$$$$$$$$$$$$$$$$ I can at the output format
'@@@@@@@@@@@@@@@@Dependency ->
    'Lib_String1
    each_chord = Split(chord_txt, " ")
   
    Dim newChordArr() As Variant
    ReDim newChordArr(UBound(each_chord))

    For i = LBound(each_chord) To UBound(each_chord)
        curr_chord = each_chord(i)
        
        con1 = (InStr(1, curr_chord, "b") > 0)
        con2 = (InStr(1, curr_chord, "#") > 0)
        
        If (InStr(1, curr_chord, "b") > 0) Or (InStr(1, curr_chord, "#") > 0) Then
            curr_key = S_TxtByInxVB(curr_chord, 1, 2)
            curr_add = S_UnLeftVB(curr_chord, 2)
            
        Else
            curr_key = S_TxtByInxVB(curr_chord, 1, 1)
            curr_add = S_UnLeftVB(curr_chord, 1)
        
        End If
        new_key = M_ChangeKey2(curr_key, shift)
        new_chord = new_key & curr_add
        newChordArr(i) = new_chord

    Next i

    
    new_chord_str = Join(newChordArr, " ")
    M_ChangeKeyVB = new_chord_str
    
    
    

End Function

'Add feature to also include number from the worksheet directly
'make it work with Array and Range
Function D_toDict(arr01, arr02)
    Dim return_dict As New Scripting.Dictionary
    'If Typeof(arr01) is Array then
    
    
End Function




Sub Test03()
    str01 = "[1,2,3,4]"
    str02 = S_RemoveAll(str01, ",")
    MsgBox (str02)
End Sub


