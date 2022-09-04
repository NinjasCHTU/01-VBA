Attribute VB_Name = "Lib_SpaceANDSelect"

Function Sp_SelectFromTL(inCell, Optional n_row = 1, Optional n_col = 1)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim outRange As Range
    
    upLeft_row = inCell.row
    upLeft_col = inCell.column
    
    'temp = Range(Cells(inCell.row, inCell.col), Cells(inCell.row + n_row, inCell.column + n_col))
    
    Set outRange = Range(Cells(inCell.row, inCell.column), Cells(inCell.row + n_row - 1, inCell.column + n_col - 1))
    'Set outRange = Range(Cells(6, 1), Cells(8, 2))
    'Sp_SelectFromTL = inCell.row
    Sp_SelectFromTL = outRange
    
End Function

Function Sp_CombineToV(ParamArray arr_range() As Variant)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim temp_arr() As Variant
    Dim i As Integer
    i = 0
    For Each range_col In arr_range
        For Each elem In range_col
            ReDim Preserve temp_arr(0 To i)
            temp_arr(i) = elem
            i = i + 1
        Next
    Next
    Sp_CombineToV = WorksheetFunction.Transpose(temp_arr)
End Function

Function Sp_CombineToH(ParamArray arr_range() As Variant)
'@@@@@@@@@@@@@@@@Dependency -> no
    Dim temp_arr() As Variant
    Dim i As Integer
    i = 0
    For Each range_col In arr_range
        For Each elem In range_col
            ReDim Preserve temp_arr(0 To i)
            temp_arr(i) = elem
            i = i + 1
        Next
    Next
    Sp_CombineToH = (temp_arr)
End Function
Function Sp_to1DLine(inRange As Range, Optional direction = 0)
    n_area = inRange.Count
    Dim out_arr() As Variant
    ReDim Preserve out_arr(n_area - 1)
    
    i = 0
    If direction = 0 Then
        For Each curr_col In inRange.Columns
            For Each elem In curr_col.Value2
                out_arr(i) = elem
                i = i + 1
            Next
        Next
    Else
        For Each curr_row In inRange.Rows
            For Each elem In curr_row.Value2
                out_arr(i) = elem
                i = i + 1
            Next
        Next
    End If
    

    
    Sp_to1DLine = VB_Transpose(out_arr)
    
    
    

End Function
Function Sp_to2DTable(row, col, Optional direction = 1)
    '@@@@@@@@@@@@@@@@Dependency ->
End Function
Function Sp_to3DLayer()

End Function

'This is not Done
Function Sp_toDiagonal(arr_in)
    '@@@@@@@@@@@@@@@@Dependency ->
    n = arr_in.Count
    Dim arr() As Variant
    ReDim arr(n, n)
    i = 0
    For Each elem In arr_in
        arr(i, i) = arr_in(i)
        i = i + 1
    Next
    
    Sp_toDiagonal = arr
    

End Function

Function Sp_SelectSkipVB(arr, r, n)
    '@@@@@@@@@@@@@@@@Dependency ->
    n_arr = arr.Count
    m = n_arr \ n
    If r = n Then
        r = 0
    End If
    Dim temp_arr() As Variant
    ReDim temp_arr(m - 1)
    arr_i = 0
    
    For i = 1 To n_arr
        curr_r = i Mod n
        If curr_r = r Then
            
            temp_arr(arr_i) = arr(i)
            arr_i = arr_i + 1
        End If
    Next i
    Sp_SelectSkipVB = VB_Transpose(temp_arr)
    'r = is the remander selected r should be less than n
    'n = is the # of repeated cycles
    
End Function


