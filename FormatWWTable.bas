Attribute VB_Name = "FormatWWTable"

Option Compare Database

Option Explicit

Public Function FormatWWXW()


    '//---------------------------------------------------------------------//
    '  Take the Cs_strProject_WW table and export it into a new table
    '  Add in the Modbus addresses where there were previously "*"
    '  Cs_strProject_WW table can then be used to cross reference between DIOTAG and %WB/%XB tables
    '
    '  If Cs_strProject_WW_Formatted already exists, it will be deleted by this function
    '  Set strWWXW to "WW" or "XW" depending on which table you wish to export
    '//---------------------------------------------------------------------//

    Dim strSQL As String
    Dim strProject As String
    Dim strWWXW As String
    Dim strAddresses() As String
    Dim strIPAddresses() As String
    Dim strCurrentFCS
    
    Dim i, j As Integer
    Dim tdf As TableDef

    Dim rs As DAO.Recordset
    Dim db As DAO.Database

    Set db = CurrentDb
    '//---------------------------------------------------------------------//
    '  Set strWWXW to "WW" or "XW"
    '//---------------------------------------------------------------------//
    strWWXW = "WW"
    'strWWXW = "XW"
    
    Debug.Print "Exporting " & strWWXW & "_Formatted Table"

                 
    strSQL = "SELECT " & strWWXW & ".* FROM " & strWWXW & ";"
        'Debug.Print strSQL

    Set rs = db.OpenRecordset(strSQL)
    rs.MoveLast                                                                             'Force complete retrieval to obtain a valid recordcount
    ReDim strAddresses(rs.RecordCount)
    ReDim strIPAddresses(rs.RecordCount)
    rs.MoveFirst                                                                            'Go back to the beginning before we start the loop
    
    '//---------------------------------------------------------------------//
    'Delete and recreate the WW Formatted table
    If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & strWWXW & "_Formatted' And Type In (1,4,6)")) Then
        DoCmd.DeleteObject acTable, "" & strWWXW & "_Formatted"
    End If
    Set tdf = CurrentDb.CreateTableDef("" & strWWXW & "_Formatted")
    With tdf
        .Fields.Append .CreateField("FCS", dbText)
        .Fields.Append .CreateField("IOADDR/SYSTAG", dbText)
        .Fields.Append .CreateField("BUFFER", dbText)
        .Fields.Append .CreateField("PROGRAM_NAME", dbText)
        .Fields.Append .CreateField("Size", dbText)
        .Fields.Append .CreateField("Port", dbText)
        .Fields.Append .CreateField("IP_ADDRESS", dbText)
        .Fields.Append .CreateField("STATION", dbText)
        .Fields.Append .CreateField("DEVICE_ADDRESS", dbText)
        .Fields.Append .CreateField("DATA_TYPE", dbText)
        .Fields.Append .CreateField("REVERSE", dbText)
        .Fields.Append .CreateField("SCAN", dbText)
        .Fields.Append .CreateField("COMMENT", dbText)
        .Fields.Append .CreateField("DIOTAG", dbText)
    End With
    db.TableDefs.Append tdf
       
        
    '//---------------------------------------------------------------------//
    'Loop through and edit records in recordset
    i = 1                                                                                   'Reset the loop values
    strCurrentFCS = "none"
    Do Until rs.EOF
    
        If (strCurrentFCS <> rs![FCS]) Then                                                 'Progress updates
            Debug.Print "Exporting " & rs![FCS]
        End If
        strCurrentFCS = rs![FCS]
        
        '//---------------------------------------------------------------------//
        'Edit the recordset records
        rs.Edit
        '[FCS]
        
        '[IOADDR/SYSTAG]
        
        '[BUFFER]
        
        '[PROGRAM_NAME]
        
        '[Size]
        
        '[Port]
        
        '[IP_ADDRESS]
        strIPAddresses(i) = rs![IP_ADDRESS]                                                 'Save the Addresses into an array
        If (Len(rs![IP_ADDRESS]) = 1) Then
            strIPAddresses(i) = strIPAddresses(i - 1)                                       'Get the previous address
        End If
        rs![IP_ADDRESS] = strIPAddresses(i)
        
        '[STATION]
        
        '[DEVICE_ADDRESS]
        strAddresses(i) = rs![DEVICE_ADDRESS]                                               'Save the Addresses into an array
        If (Len(rs![DEVICE_ADDRESS]) = 1) Then
            strAddresses(i) = Right(strAddresses(i - 1), Len(strAddresses(i - 1)) - 1) + 1  'Get the previous address and add one
            strAddresses(i) = Format(strAddresses(i), "00000")                              'Add leading zero if required
            strAddresses(i) = Left(strAddresses(i - 1), 1) & strAddresses(i)                'Add the function code
            strAddresses(i) = Replace(strAddresses(i), " ", "")                             'Remove extra spaces
        End If
        rs![DEVICE_ADDRESS] = strAddresses(i)
        
        '[DATA_TYPE]
        
        '[REVERSE]
        
        '[SCAN]
        
        '[COMMENT]
        If (Len(rs![COMMENT]) < 1) Then
            rs![COMMENT] = " "                                                              'Lazy programming, but to avoid a NULL error in strSQL
        End If
        
        '[DIOTAG]
        If (Len(rs![DIOTAG]) > 0) Then
            rs![DIOTAG] = Replace(rs![DIOTAG], "%%", "")                                    'Remove "%%" from Label so it can be used for JOIN in query
        Else
            rs![DIOTAG] = " "                                                               'Lazy programming, but to avoid a NULL error in strSQL
        End If
        
        '//---------------------------------------------------------------------//
        'Insert the record into the newly created "WW_Formatted" table
        If (Len(rs![Size]) > 0) Then                                                        'Don't worry about empty records
            strSQL = "INSERT INTO " & strWWXW & "_Formatted VALUES (" & _
                        "'" & rs![FCS] & "'," & _
                        "'" & rs![IOADDR/SYSTAG] & "'," & _
                        "'" & rs![BUFFER] & "'," & _
                        "'" & rs![PROGRAM_NAME] & "'," & _
                        "'" & rs![Size] & "'," & _
                        "'" & rs![Port] & "'," & _
                        "'" & rs![IP_ADDRESS] & "'," & _
                        "'" & rs![STATION] & "'," & _
                        "'" & rs![DEVICE_ADDRESS] & "'," & _
                        "'" & rs![DATA_TYPE] & "'," & _
                        "'" & rs![REVERSE] & "'," & _
                        "'" & rs![SCAN] & "'," & _
                        "'" & rs![COMMENT] & "'," & _
                        "'" & rs![DIOTAG] & "'" & _
                        ");"
                        'Debug.Print strSQL
                        
            db.Execute strSQL, dbFailOnError
        End If
     
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Debug.Print "Exporting Finished"


End Function

