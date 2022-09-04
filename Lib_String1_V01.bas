Attribute VB_Name = "Lib_String1"


'Can expand Function to make it count many characters at once
Function S_count(inString, ch)
  '@@@@@@@@@@@@@@@@Dependency -> No
    n = Len(inString)
    n_ch = Len(ch)
    Count = 0
    For i = 1 To n
        curr_string = Mid(inString.Value, i, n_ch)
        If curr_string = ch Then
            Count = Count + 1
        End If
    Next
    S_count = Count
End Function

Function S_UnRightVB(text, n)
    S_UnRightVB = Left(text, Len(text) - n)
End Function

Function S_UnLeftVB(text, n)
    S_UnLeftVB = Right(text, Len(text) - n)
End Function
Function S_TxtByInxVB(text, start_inx, end_inx)
    S_TxtByInxVB = Mid(text, start_inx, end_inx - start_inx + 1)
End Function

Function S_Remove1(inString, ch)

End Function

Function S_RemoveAll(inString, ch)
  '@@@@@@@@@@@@@@@@Dependency -> No
    str02 = Replace(inString, ch, "")
    S_RemoveAll = str02
End Function

Function S_ReplaceBy(text, new_text As String, old_text_arr As Variant)
    'old_text_arr = array that contains old alphabets for the replacement
    n_text = Len(text)
    n_arr = UBound(old_text_arr)

    out_str = ""
    
    For i = 1 To n_text
        curr_ch = Mid(text, i, 1)
        For j = LBound(old_text_arr) To UBound(old_text_arr)
            If (curr_ch = old_text_arr(j)) Then
                curr_ch = new_text
            End If

        Next j
        out_str = out_str & curr_ch
    Next i
    S_ReplaceBy = out_str
End Function

Function S_UnDiaCriticVB(ch)
    Dim a_varyForm, e_varyForm, i_varyForm, o_varyForm, u_varyForm, y_varyForm, c_varyForm, n_varyForm As Variant
    a_varyForm = Array("ä", "á", "â", "à", "å", "ã")
    e_varyForm = Array("e", "ë", "é", "ê", "è")
    i_varyForm = Array("ï", "í", "î", "ì")
    o_varyForm = Array("è", "õ", "ô", "ò", "ó")
    u_varyForm = Array("ü", "ú", "û", "ù")
    y_varyForm = Array("ÿ")
    c_varyForm = Array("ç")
    n_varyForm = Array("ñ")
    
    un_a_str = S_ReplaceBy(ch, "a", a_varyForm)
    un_ae_str = S_ReplaceBy(un_a_str, "e", e_varyForm)
    un_aei_str = S_ReplaceBy(un_ae_str, "i", i_varyForm)
    un_aeo_str = S_ReplaceBy(un_aei_str, "o", o_varyForm)
    un_aeou_str = S_ReplaceBy(un_aeo_str, "u", u_varyForm)
    un_aeouy_str = S_ReplaceBy(un_aeou_str, "y", y_varyForm)
    un_aeouyc_str = S_ReplaceBy(un_aeouy_str, "c", c_varyForm)
    un_aeouycn_str = S_ReplaceBy(un_aeouyc_str, "n", n_varyForm)
    
    
    final_string = un_aeouycn_str
    S_UnDiaCriticVB = final_string
    
    
End Function
