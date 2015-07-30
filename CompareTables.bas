Attribute VB_Name = "CompareTables"
Option Compare Database
Option Explicit

        
Public Sub Run()

    DoCmd.Save acModule, "CompareTables"
    CompareTables

End Sub

    
Public Sub CompareTables()

    '//---------------------------------------------------------------------//
    '//     Luke Shortal 2015-07-28
    '//     Revision 0.0
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
    strTableOld = "PCS IO (Rev8)"
    strTableNew = "PCS IO (Rev9)"
    'strTableOld = "SIS FGS IO (Rev7)"
    'strTableNew = "SIS FGS IO (Rev8)"
    strJoin = "COMPONENT_ID"
    
    'strTableOld = "PCS Wiring (Rev6)"
    'strTableNew = "PCS Wiring (Rev7)"
    'strTableOld = "SIS FGS Wiring (Rev5)"
    'strTableNew = "SIS FGS Wiring (Rev6)"
    'strJoin = "COMPONENT_IDWIRE_TAG"
    '//---------------------------------------------------------------------//

    strQueryName = strTableNew
    
    '//---------------------------------------------------------------------//
    '       Define which fields to include in comparison
    '//---------------------------------------------------------------------//
    If (strJoin = "COMPONENT_IDWIRE_TAG") Then
        intIncludeFields = 6
        ReDim strIncludeFields(intIncludeFields)
        
        i = 1
        strIncludeFields(i) = "MC_NAME"
        i = i + 1
        strIncludeFields(i) = "MC_TS_NAME"
        i = i + 1
        strIncludeFields(i) = "MC_TS_TERM_NO"
        i = i + 1
        strIncludeFields(i) = "WIRE_TAG"
        i = i + 1
        strIncludeFields(i) = "JB_CABLE_TYPE"
        i = i + 1
        strIncludeFields(i) = "JB_CABLE_NM"

    End If

    If (strJoin = "COMPONENT_ID") Then
        intIncludeFields = 10
        ReDim strIncludeFields(intIncludeFields)
        
        i = 1
        strIncludeFields(i) = "FCS"
        i = i + 1
        strIncludeFields(i) = "SIGNAL_ORIGIN"
        i = i + 1
        strIncludeFields(i) = "RACK_NODE"
        i = i + 1
        strIncludeFields(i) = "SLOT"
        i = i + 1
        strIncludeFields(i) = "CHANNEL"
        i = i + 1
        strIncludeFields(i) = "ICSS_TAG"
        i = i + 1
        strIncludeFields(i) = "TGCOMM"
        i = i + 1
        strIncludeFields(i) = "TGCOMM2"
        i = i + 1
        strIncludeFields(i) = "DIO_STATUS_1_LABEL"
        i = i + 1
        strIncludeFields(i) = "DIO_STATUS_0_LABEL"
        i = i + 1
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
                        "IIf([" & strTableOld & "].[" & fld.Name & "]=[" & strTableNew & "].[" & fld.Name & "],'OK', " & _
                        "IIf(Nz([" & strTableOld & "].[" & fld.Name & "])='','new', " & _
                        "IIf(Nz([" & strTableNew & "].[" & fld.Name & "])='','deleted','modified')))" & _
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
                        "IIf([" & strTableOld & "].[" & fld.Name & "]=[" & strTableNew & "].[" & fld.Name & "],'OK', " & _
                        "IIf(Nz([" & strTableOld & "].[" & fld.Name & "])='','new', " & _
                        "IIf(Nz([" & strTableNew & "].[" & fld.Name & "])='','deleted','modified')))" & _
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
    
    strQueryName = strTableNew & " Compare"
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
