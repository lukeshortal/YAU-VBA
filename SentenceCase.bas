Public Function SentenceCase(WorkRng As Range)

    Dim Rng As Range
    Dim varValue As Variant
    Dim blStart As Boolean
    Dim strMid As String
    Dim i As Integer
    
    On Error Resume Next
    
    For Each Rng In WorkRng
        varValue = Rng.Value
        blStart = True
        For i = 1 To Len(varValue)
            strMid = Mid(varValue, i, 1)
            Select Case strMid
                Case "."
                blStart = True
                Case "?"
                blStart = True
                Case "a" To "z"
                If blStart Then
                    strMid = UCase(strMid)
                    blStart = False
                End If
                Case "A" To "Z"
                If blStart Then
                    blStart = False
                Else
                    strMid = LCase(strMid)
                End If
            End Select
            Mid(varValue, i, 1) = strMid
        Next
        SentenceCase = varValue
    Next
    
End Function
