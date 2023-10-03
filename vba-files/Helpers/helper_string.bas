Attribute VB_Name = "helper_string"




'namespace=vba-files\Helpers


Function NextLetter(Ref As String) As String
    If Asc(Ref) >= 65 And Asc(Ref) <= 89 Then     'Considers Letters A - Y (Not Z)
        T = Asc(Ref)
        T = T + 1
        T = Chr(T)
    End If
    NextLetter = T
End Function



Function RemoveLineBreak(Text As String) As String
    RemoveLineBreak = Trim(Replace(Replace(Text, Chr(10), ""), Chr(13), ""))
End Function



Public Function StringFormat(ByVal mask As String, ParamArray tokens()) As String

    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        mask = Replace(mask, "{" & i & "}", tokens(i))
        Next
        StringFormat = mask

End Function


