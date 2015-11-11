Option Compare Database
Dim patternStr() As String
Dim numberPatTag As Integer

'----------------------------------------------------------------------------------------------------
' Public Function RegExpatternStr()
'   Build the patternStr() array that contains all regular expression patterns/replacement strings
'   Modify "strRegExTable" to select which table is used for populating the array
'   By using "GUI_Regex_testing", the expressions can be modified without altering the "master" table: "GUI_Regex"
'----------------------------------------------------------------------------------------------------
Public Function RegExpatternStr()

    Dim rstTableName As DAO.Recordset   'Your table
    Dim intArraySize As Integer         'The size of your array
    Dim iCounter As Integer             'Index of the array
    Dim strRegExTable As String
    
    strRegExTable = "GUI_Regex"
    'strRegExTable = "GUI_Regex_testing"
    Set rstTableName = CurrentDb.OpenRecordset("SELECT * FROM " & strRegExTable & " ORDER BY Priority")
    
    If Not rstTableName.EOF Then
    
        rstTableName.MoveFirst   'Ensure we begin on the first row
    
        'The size of the array should be equal to the number of rows in the table
        intArraySize = rstTableName.RecordCount

        iCounter = 1
        ReDim patternStr(intArraySize, 6) 'Need to size the array
    
        Do Until rstTableName.EOF
    
            patternStr(iCounter, 1) = rstTableName.Fields("1_Find")
            patternStr(iCounter, 2) = rstTableName.Fields("2_Replace")
            patternStr(iCounter, 3) = rstTableName.Fields("3_Replace")
            patternStr(iCounter, 4) = rstTableName.Fields("4_Replace")
            patternStr(iCounter, 5) = rstTableName.Fields("5_Replace")
            
    
            iCounter = iCounter + 1
            rstTableName.MoveNext
        Loop
    
    End If
    
    If IsObject(rstTableName) Then Set rstTableName = Nothing
    
    numberPatTag = intArraySize

End Function

'----------------------------------------------------------------------------------------------------
' Public Function RegexVBMulti(inputStr As String, TypInt As Integer) As String
'   Run a "regex.replace" over the input string for each pattern in patternStr()
'   Return the modified/replaced string
'----------------------------------------------------------------------------------------------------
Public Function RegexVBMulti(inputStr As String, TypInt As Integer) As String
    Dim regex As regexp
    Dim i, j As Integer

    Set regex = New regexp
    regex.IgnoreCase = True

    Call RegExpatternStr


    For i = 1 To numberPatTag
        regex.pattern = patternStr(i, 1)
        Set matches = regex.Execute(inputStr)
        If matches.Count > 0 Then
            If (patternStr(i, TypInt) <> "WRONG CHOICE") Then
                inputStr = regex.replace(inputStr, patternStr(i, TypInt))
            End If
        End If

    Next i

    RegexVBMulti = inputStr

    Set regex = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------
' Public Function RegexVBMultiPattern(inputStr As String, TypInt As Integer) As String
'   Run a "regex.replace" over the input string for each pattern in patternStr()
'   Return the pattern(s) that found matches
'----------------------------------------------------------------------------------------------------
Public Function RegexVBMultiPattern(inputStr As String, TypInt As Integer) As String
    Dim regex As regexp
    Dim outputStr As String
    Dim i, j As Integer

    Set regex = New regexp
    regex.IgnoreCase = True

    Call RegExpatternStr

    outputStr = ""
    For i = 1 To numberPatTag
        regex.pattern = patternStr(i, 1)
        Set matches = regex.Execute(inputStr)
        If matches.Count > 0 Then
            inputStr = regex.replace(inputStr, patternStr(i, TypInt))   'do the replacement
            outputStr = outputStr & ", " & patternStr(i, 1)             'return the pattern number
        End If

    Next i

    RegexVBMultiPattern = outputStr
    Set regex = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------
' Public Function RegexVB(inputStr As String, TypInt As Integer) As String
'   Run a "regex.replace" over the input string for each pattern in patternStr()
'   Break after a match is found
'   Return the modified/replaced string
'----------------------------------------------------------------------------------------------------
Public Function RegexVB(inputStr As String, TypInt As Integer) As String
    Dim regex As regexp
    Dim outputStr As String
    Dim i, j As Integer

    Set regex = New regexp
    regex.IgnoreCase = True
    
    Call RegExpatternStr


    For i = 1 To numberPatTag
        regex.pattern = patternStr(i, 1)
        Set matches = regex.Execute(inputStr)
        If matches.Count > 0 Then
            outputStr = regex.replace(inputStr, patternStr(i, TypInt))
            RegexVB = outputStr
            Exit For
        Else
            RegexVB = inputStr
        End If
        
    Next i
    
    Set regex = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------
' Public Function RegexVBPattern(inputStr As String, TypInt As Integer) As String
'   Run a "regex.replace" over the input string for each pattern in patternStr()
'   Break after a match is found
'   Return the pattern that found matches
'----------------------------------------------------------------------------------------------------
Public Function RegexVBPattern(inputStr As String, TypInt As Integer) As String
    Dim regex As regexp
    Dim outputStr As String
    Dim i, j As Integer

    Set regex = New regexp
    regex.IgnoreCase = True

    Call RegExpatternStr
    
    For i = 1 To numberPatTag
        regex.pattern = patternStr(i, 1)
        Set matches = regex.Execute(inputStr)
        If matches.Count > 0 Then
            'RegexVBPattern = i
            RegexVBPattern = patternStr(i, 1)
        Exit For
        Else
            RegexVBPattern = "none"
        End If
        
    Next i
    
    Set regex = Nothing
    
End Function



'----------------------------------------------------------------------------------------------------
' Public Function Check_Value(Value As Variant, strType As Variant) As String
'   Used for formatting the input string and
'   Will return “not set” if input is an invalid string or
'   Will return “999” if input is an invalid number
'----------------------------------------------------------------------------------------------------
Public Function Check_Value(Value As Variant, strType As Variant) As String

    'Check if "Value" has been set correctly.
    'i.e. is a number if it needs to be a number, or a string if it needs to be a string
    'strType is used in the query: set to "1" to check for/return a string, set to "0" to check for/return a number
    
    If (strType = 1) Then
        If (Nz(Value) = vbNullString) Then
            Check_Value = "not set"
        ElseIf (Value = "HOLD") Then
            Check_Value = "not set"
        ElseIf (Value = "-") Then
            Check_Value = "not set"
        Else
            Check_Value = Value
        End If
        
    ElseIf (strType = 0) Then
        If (Nz(Value) = vbNullString) Then
            Check_Value = "999"
        ElseIf (Value = "HOLD") Then                'if it contains a hold
            Check_Value = "999"
        ElseIf (Value = "-") Then
            Check_Value = "999"
        ElseIf (InStr(Value, ",")) Then
            Check_Value = replace(Value, ",", "")   'if it contains a comma
        Else
            Check_Value = Value
        End If
        
    Else
        Check_Value = Null
    End If

End Function
