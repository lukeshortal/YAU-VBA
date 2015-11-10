Public Function ContainsNumber(WorkRng As Range) As Boolean

    Dim Rng As Range
    Dim varValue As Variant
    
    On Error Resume Next
    
    For Each Rng In WorkRng
        varValue = Rng.Value
        If (InStr(varValue, "1") Or InStr(varValue, "2") Or InStr(varValue, "3") Or InStr(varValue, "4") Or InStr(varValue, "5") Or InStr(varValue, "6") Or InStr(varValue, "7") Or InStr(varValue, "8") Or InStr(varValue, "9") Or InStr(varValue, "0")) Then
            ContainsNumber = True
        Else
            ContainsNumber = False
        End If
    Next

End Function
