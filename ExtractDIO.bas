Public Function ExtractDIO(POSITION As Variant, LABEL As Variant, OnOff As Variant) As Variant

    'Extract the "ON" and "OFF" label from the yGet field "LABEL"
    '"POSITION" means that the label is reversed in the Centum configuration
    'OnOff is used in the query: set to "1" to return the "ON" label, set to "0" to return the "OFF" label
    
    If (OnOff = 1) Then
        If (POSITION = 1) Then
            ExtractDIO = Left(LABEL, InStr(LABEL, ",") - 1)
        ElseIf (POSITION = 2) Then
            ExtractDIO = Mid(LABEL, InStr(LABEL, ",,") + 2, (InStr(InStr(LABEL, ",,") + 2, LABEL, ",") + 2) - (InStr(LABEL, ",,") + 2) - 2)
        Else
            ExtractDIO = Null
        End If
        
    ElseIf (OnOff = 0) Then
        If (POSITION = 2) Then
            ExtractDIO = Left(LABEL, InStr(LABEL, ",") - 1)
        ElseIf (POSITION = 1) Then
            ExtractDIO = Mid(LABEL, InStr(LABEL, ",,") + 2, (InStr(InStr(LABEL, ",,") + 2, LABEL, ",") + 2) - (InStr(LABEL, ",,") + 2) - 2)
        Else
            ExtractDIO = Null
        End If
        
    Else
        ExtractDIO = Null
    End If
    
End Function
