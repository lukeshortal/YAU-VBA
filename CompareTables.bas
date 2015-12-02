Option Compare Database
Option Explicit

    
Public Sub CompareTables()

    '//---------------------------------------------------------------------//
    '//     Luke Shortal 2015-07-28
    '//     Revision 0.0
    '//     Revision 0.1
    '//---------------------------------------------------------------------//
    '//     This function takes two tables and compares the records
    '//     Would be useful to compare two revisions of an I/O or Modbus list
    '//
    '//     Set "strTableOld" as the name of the original table
    '//     Set "strTableNew" as the name of the new table
    '//     Set "strJoin" as the field on which to join the two tables
    '//
    '//     Set "strIncludeFields" to the fields that you want to include in the comparison
    '//     Set "intIncludeFields" to the numer of elemets in the "strIncludeFields" array
    '//
    '//     The function will build two queries and UNION them together:
    '//         A LEFT JOIN to find all new and modified records
    '//         A RIGHT JOIN to find all deleted and modified records
    '//
    '//     ToDo: Add "Build the xxxx SQL statement" to seperate function:
    '//         Essentially the same code is used twice to build LEFT/RIGHT statements.
    '//         Instead, create a seperate function and use "strJoinLRU" to differentiate query types
    '//     ToDo: Init "strIncludeFields" in seperate function (clear up clutter)
    '//     ToDo: Run "Set qdf" in seperate function (query creation may not always be required)
    '//
    '//---------------------------------------------------------------------//

    Dim strSQL As String
    Dim strJoin As String
    Dim strJoinLRU As String
    Dim strTableNew As String
    Dim strTableOld As String
    Dim strQueryName As String
    Dim strIncludeFields() As String
    
    Dim i As Integer
    Dim j As Integer
    Dim intIncludeFields As Integer

    Dim rsOld As DAO.Recordset
    'Dim rsNew As DAO.Recordset
    Dim db As DAO.Database
    Dim fld As DAO.Field
    Dim qdf As QueryDef
    
    '//---------------------------------------------------------------------//
    '       Define which tables to compare and which field to join
    '//---------------------------------------------------------------------//
    
    'SysDbAliasStr
    strTableOld = "SysDbAliasStr"
    strTableNew = "LS 99 SysDbAliasStr"
    strJoin = "signalName"

    'SysDbSignalStr
'    strTableOld = "SysDbSignalStr"
'    strTableNew = "LS 99 SysDbSignalStr"
'    strJoin = "signalName"

    '//---------------------------------------------------------------------//

    strQueryName = strTableNew
    
    '//---------------------------------------------------------------------//
    '       Define which fields to include in comparison
    '//---------------------------------------------------------------------//
    'SysDbAliasStr
    If (strTableOld = "SysDbAliasStr") Then
        intIncludeFields = 3
        ReDim strIncludeFields(intIncludeFields)
        
        i = 1
        strIncludeFields(i) = "proc_no"
        i = i + 1
        strIncludeFields(i) = "signalName"
        i = i + 1
        strIncludeFields(i) = "aliasName"

    End If
    
    'SysDbSignalStr
    If (strTableOld = "SysDbSignalStr") Then
        intIncludeFields = 5
        ReDim strIncludeFields(intIncludeFields)
        
        i = 1
        strIncludeFields(i) = "proc_no"
        i = i + 1
        strIncludeFields(i) = "signalName"
        i = i + 1
        strIncludeFields(i) = "sig_type"
        i = i + 1
        strIncludeFields(i) = "signalDesc"
        i = i + 1
        strIncludeFields(i) = "units"
'        i = i + 1
'        strIncludeFields(i) = "has_limits"
'        i = i + 1
'        strIncludeFields(i) = "lowLimit_ctl"
'        i = i + 1
'        strIncludeFields(i) = "lowLimit"
'        i = i + 1
'        strIncludeFields(i) = "highLimit_ctl"
'        i = i + 1
'        strIncludeFields(i) = "highimit"
'        i = i + 1
'        strIncludeFields(i) = "has_defaults"
'        i = i + 1
'        strIncludeFields(i) = "default_ctl"
'        i = i + 1
'        strIncludeFields(i) = "defVal"
'        i = i + 1
'        strIncludeFields(i) = "publisher_process_no"
'        i = i + 1
'        strIncludeFields(i) = "malf"

    End If
    '//---------------------------------------------------------------------//
    
    
    Set db = CurrentDb
    
    strSQL = "SELECT [" & strTableOld & "].* " & _
        "FROM [" & strTableOld & "];"
        'Debug.Print strSQL
    Set rsOld = db.OpenRecordset(strSQL)
    
'    strSQL = "SELECT [" & strTableNew & "].* " & _
'        "FROM [" & strTableNew & "];"
'        'Debug.Print strSQL
'    Set rsNew = db.OpenRecordset(strSQL)


    
    'Build the LEFT SQL statement
    strJoinLRU = "LEFT"
    strSQL = "SELECT "
    For Each fld In rsOld.Fields
        For j = 1 To intIncludeFields
            If (fld.Name = strIncludeFields(j)) Then                        'If fld.Name is a member of the fields to be included...
                strSQL = strSQL & "[" & strTableOld & "].[" & fld.Name & "], [" & strTableNew & "].[" & fld.Name & "], " & _
                        "IIf(Nz([" & strTableOld & "].[" & fld.Name & "])='','new', " & _
                        "IIf(Nz([" & strTableNew & "].[" & fld.Name & "])='','deleted', " & _
                        "IIf(Cstr([" & strTableOld & "].[" & fld.Name & "])=Cstr([" & strTableNew & "].[" & fld.Name & "]),'OK','modified'))) " & _
                        "AS [" & fld.Name & " OK], "
            End If

        Next j
    Next
    strSQL = strSQL & "FROM [" & strTableOld & "] " & strJoinLRU & " JOIN [" & strTableNew & "] ON [" & strTableOld & "]." & strJoin & " = [" & strTableNew & "]." & strJoin & " "
    strSQL = Replace(strSQL, ", FROM", " FROM")                             'Fix up the inevitable syntax error
    'Debug.Print strSQL
    
    
    'Build the RIGHT SQL statement
    strJoinLRU = "RIGHT"
    strSQL = strSQL & "UNION SELECT "                                       'The RIGHT is UNION'd with the LEFT
    For Each fld In rsOld.Fields
        For j = 1 To intIncludeFields
            If (fld.Name = strIncludeFields(j)) Then                        'If fld.Name is a member of the fields to be included...
                strSQL = strSQL & "[" & strTableOld & "].[" & fld.Name & "], [" & strTableNew & "].[" & fld.Name & "], " & _
                        "IIf(Nz([" & strTableOld & "].[" & fld.Name & "])='','new', " & _
                        "IIf(Nz([" & strTableNew & "].[" & fld.Name & "])='','deleted', " & _
                        "IIf(Cstr([" & strTableOld & "].[" & fld.Name & "])=Cstr([" & strTableNew & "].[" & fld.Name & "]),'OK','modified'))) " & _
                        "AS [" & fld.Name & " OK], "
            End If

        Next j
    Next
    strSQL = strSQL & "FROM [" & strTableOld & "] " & strJoinLRU & " JOIN [" & strTableNew & "] ON [" & strTableOld & "]." & strJoin & " = [" & strTableNew & "]." & strJoin & " "
    strSQL = Replace(strSQL, ", FROM", " FROM")                             'Fix up the inevitable syntax error
    'Debug.Print strSQL
    strSQL = strSQL & "ORDER BY "
        For j = 1 To intIncludeFields                                       'For each of the fields in our array
            For Each fld In rsOld.Fields
                If (fld.Name = strIncludeFields(j)) Then                    'And for only those that exist in the recordset
                    strSQL = strSQL & "[" & strTableOld & "].[" & strIncludeFields(j) & "],"
                    strSQL = strSQL & "[" & strTableNew & "].[" & strIncludeFields(j) & "]"
                    If (j = intIncludeFields) Then
                        strSQL = strSQL & ";"
                    Else
                        strSQL = strSQL & ", "
                    End If
                End If
            Next
        Next j
    'Debug.Print strSQL
    
    'Create compare query
    strQueryName = strTableNew & " Compare"
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & strQueryName & "'")) Then
        DoCmd.DeleteObject acQuery, strQueryName
    End If
    Set qdf = CurrentDb.CreateQueryDef(strQueryName, strSQL)
    
    Debug.Print "Created " & strQueryName
    
    'Create make table query for compare query
    strSQL = "SELECT [" & strQueryName & "].* INTO [" & strQueryName & " (tbl)] FROM [" & strQueryName & "];"
    strQueryName = strTableNew & " Compare (make)"
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & strQueryName & "'")) Then
        DoCmd.DeleteObject acQuery, strQueryName
    End If
    Set qdf = CurrentDb.CreateQueryDef(strQueryName, strSQL)
    
    Debug.Print "Created " & strQueryName


    Set qdf = Nothing
    Set fld = Nothing
    Set rsOld = Nothing
    'Set rsNew = Nothing
    Set db = Nothing
    
    
    
End Sub

