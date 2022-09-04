Attribute VB_Name = "Lib_WSFunc"
Function VB_Transpose(arr)
    VB_Transpose = WorksheetFunction.Transpose(arr)
End Function

'How to use sort in VBA
'https://excelchamps.com/vba/sort-range/



Function VB_Seq(n, Optional column = 1, Optional start = 1, Optional step = 1)
    res = WorksheetFunction.Sequence(n, column, start, step)
    VB_Seq = res
End Function

Function VB_SortBy(arr1 As Range, sort_by_this As Range, Optional order = 1)
    res = WorksheetFunction.SortBy(arr1, sort_by_this, order)
    VB_SortBy = res
    
End Function



