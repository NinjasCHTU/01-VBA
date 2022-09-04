Attribute VB_Name = "Lib_Array1"

Sub A_printArr(inArray)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim printCell, printCell_Temp As Range
    Dim arr01() As Variant
    'arr01 = Array(1, 2, 3, 4, 5, 6, 7)
    'n01 = 7
    n = UBound(inArray)
    'Set printCell_Temp = Range("K15")
    Set printCell = Application.InputBox(Title:="Print Array", Prompt:="Select Range to print out", Type:=8)
    For i = 0 To n
        Cells(printCell.row + i, printCell.column) = inArray(i)
    Next i
    
    
End Sub

Function A_toArray1d(selectRange)
'@@@@@@@@@@@@@@@@Dependency -> No
    Dim inRange As Range
    'Set selectRange = Application.InputBox(Title:="Import value to Arrays", Prompt:="Select Range to import ARRAYS value", Type:=8)
    n = selectRange.Cells.Count
    Dim outArray() As Variant
    ReDim outArray(n - 1)
    For i = 1 To n
        curr_val = selectRange.Cells(i).Value
        outArray(i - 1) = curr_val
    Next
    A_toArray1d = outArray
    
    
End Function

Function A_isInArr(arr_in, checker) As Boolean
'@@@@@@@@@@@@@@@@Dependency -> No
    For i = LBound(arr_in) To UBound(arr_in)
        If arr_in(i) = checker Then
            A_isInArr = True
            Exit Function
        End If

    Next i
    A_isInArr = False
    


End Function
'A_toArray2d = 2d Version but not combined with original Function
'must continue
Function A_toArray2d(selectRange)
'@@@@@@@@@@@@@@@@Dependency -> No
'But there is GONNA Be a Problem when it's 1 dimesion when I want to use to other part in VBA
    Dim inRange As Range
    'Set selectRange = Application.InputBox(Title:="Import value to Arrays", Prompt:="Select Range to import ARRAYS value", Type:=8)
    n_row = selectRange.Rows.Count
    n_col = selectRange.Columns.Count
    
    
    'A_toArray2d = n_row & "and" & n_col
    
    Dim outArray() As Variant
    ReDim outArray(n_row - 1, n_col - 1)
    
    'A_toArray2d = selectRange.Cells(2, 4).Value
    'A_toArray2d = outArray
    For i = 1 To n_row
        For j = 1 To n_col
        curr_val = selectRange.Cells(i, j).Value
        outArray(i - 1, j - 1) = curr_val
        Next j
    Next i
    A_toArray2d = outArray
    
    
End Function




Function A_TxtTO1dArr(inString)
'@@@@@@@@@@@@@@@@Dependency ->
'Lib_Dear1 (S_RemoveAll)

'declear 1d Array using PYTHON syntax
' Still have problem if I want the input to be INT (right now they are String)!!!!
    str02 = Mid(inString, 2, Len(inString) - 2)
    str03 = Split(str02, ",")
    'This loop is for remove space
    For i = LBound(str03) To UBound(str03)
        str03(i) = Trim(str03(i))
        str03(i) = S_RemoveAll(str03(i), """")
    Next
    
    A_TxtTO1dArr = str03
    
End Function

'to be Continue
'Still not dealing with removing " " eg "a"
Function A_TxtTO2dArr(inString)
'@@@@@@@@@@@@@@@@Dependency -> No
' Still have problem if I want the input to be INT (right now they are String)!!!!
    'Assuming retangular array
    'declear 2d Array using PYTHON syntax
    str_noBracket = Mid(inString, 2, Len(inString) - 2)
    str_noSpace = Replace(str_noBracket, " ", "")
    
    text_tab = Split(str_noSpace, "],")
    'to get the number of columns first
    str_removeL = Replace(text_tab(0), "[", "")
    str_removeR = Replace(str_removeL, "]", "")
    each_row_tab = Split(str_removeR, ",")

    col = UBound(each_row_tab) + 1
    row = UBound(text_tab) + 1
    'Declear the 2d out_arr
    Dim out_arr As Variant
    ReDim out_arr(row - 1, col - 1)



    For i = LBound(text_tab) To UBound(text_tab)
        str_removeL = Replace(text_tab(i), "[", "")
        str_removeR = Replace(str_removeL, "]", "")
        each_row_tab = Split(str_removeR, ",")
        

        For j = 0 To col - 1
            out_arr(i, j) = each_row_tab(j)
        Next j

        
    Next i
    
    A_TxtTO2dArr = out_arr
        
        
        
End Function

