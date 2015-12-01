Option Compare Database
Option Explicit
Option Base 1

Global strPjtPath As String
Global strProject As String
Global strHIS As String
Global strOPCHost As String


Public Sub RDBtoMDB()

    Dim lngRtn




End Sub




Public Sub SetPjtPath()

    Dim dbCur As DAO.Database
    Dim rstData As DAO.Recordset
    Dim strTest As String
    
    strPjtPath = ""
    
    Set dbCur = CurrentDb
    
    With dbCur
        Set rstData = .OpenRecordset("tblDBConfig", dbOpenDynaset)
        With rstData
            Call .FindFirst("[ID]=1")
            If (.NoMatch) Then
                Call MsgBox("Please set the project path in tblConfig.", vbCritical)
                GoTo ErrHdlr
            End If
            strTest = Dir(!ProjectPath)
            If (strTest = "") Then
                Call MsgBox("Error opening project path, please check connection.", vbCritical)
            End If
            strPjtPath = !ProjectPath
            strProject = !ProjectName
            strHIS = !MasterHIS
            strOPCHost = !OPCHost
        End With
        Call rstData.Close
        Set rstData = Nothing
        dbCur.Close
    End With
    Set dbCur = Nothing

    Exit Sub
    
ErrHdlr:

    Set rstData = Nothing
    dbCur.Close
    Set dbCur = Nothing
        
End Sub


Public Sub BuildMaster()

    'This subroutine creates a master Element table from each of the FCS element tables
    'Kinda useful
    
    Dim strFCS, strSQL As String
    Dim intFCS As Integer
    Dim i As Long
    Dim tbdTable As DAO.TableDef
    Dim dbCur As DAO.Database
    
    Set dbCur = CurrentDb
    
    With dbCur
    
        Form_frmOPC.lblStatus.Caption = "Linking UnitTbl.."
        Form_frmOPC.Refresh
        DoEvents
    
        On Error Resume Next
        Call .TableDefs.Delete("UnitTbl")                                   'Remove old UnitTbl
        On Error GoTo 0
        
        Set tbdTable = .CreateTableDef("UnitTbl")                           'Link in correct UnitTbl
        tbdTable.Connect = ";DATABASE=" & strPjtPath & "etc\PjtRef.accdb"
        tbdTable.SourceTableName = "UnitTbl"
        .TableDefs.Append tbdTable
        Set tbdTable = Nothing
        
        Form_frmOPC.lblStatus.Caption = "Clearing ElemMstr.."
        Form_frmOPC.Refresh
        DoEvents
        
        strSQL = "DELETE ElemMstr.* " & _
                "FROM ElemMstr;"
        Call .Execute(strSQL, dbFailOnError)                                  'Clear ElemMstr
        
        strFCS = Dir(strPjtPath & "FCS*", vbDirectory)
        Do While strFCS <> ""
        
            Form_frmOPC.lblStatus.Caption = "Importing " & strFCS
            Form_frmOPC.Refresh
            DoEvents
                
            intFCS = CInt(Right(strFCS, 4))
            Set tbdTable = .CreateTableDef("ElemTbl")
            tbdTable.Connect = ";DATABASE=" & strPjtPath & strFCS & "\etc\FcsRef.accdb"
            tbdTable.SourceTableName = "ElemTbl"
            .TableDefs.Append tbdTable

            Set tbdTable = Nothing

            strSQL = "INSERT INTO ElemMstr (FCS, ElemType, ElemNum, TagName, Comment, SL, SH, EngUnit, MSH, MSL, MVEngUnit, " & _
                "IOMActMode, HelpNum, AlarmLevel, WindowName, FileName) " & _
                "SELECT " & intFCS & " AS FCS, [ElemTbl].InstName, [ElemTbl].ElemNum, [ElemTbl].ElemName, [ElemTbl].Comment, " & _
                "[ElemTbl].SL , [ElemTbl].SH, [UnitTbl].Unit, [ElemTbl].RH, [ElemTbl].RL, [UnitTbl_1].Unit, " & _
                "[ElemTbl].IOMActMode , [ElemTbl].HelpNum, [ElemTbl].AlarmLevel, [ElemTbl].WindowName, [ElemTbl].FileName " & _
                "FROM (ElemTbl LEFT JOIN UnitTbl ON [ElemTbl].UnitID = [UnitTbl].UnitID) LEFT JOIN UnitTbl AS UnitTbl_1 " & _
                "ON [ElemTbl].MVUnitID = [UnitTbl_1].UnitID " & _
                "ORDER BY [ElemTbl].InstName, [ElemTbl].ElemNum;"
            
            Call .Execute(strSQL, dbFailOnError)
            Call .TableDefs.Delete("ElemTbl")
            strFCS = Dir(, vbDirectory)
            DoEvents
            
            
        Loop
    End With
    
    Set dbCur = Nothing


End Sub

Public Sub BuildConnect()

    'This subroutine creates a master Connect table from each of the FCS Connect Table
    
    Dim strFCS, strSQL As String
    Dim intFCS As Integer
    Dim tbdTable As DAO.TableDef
    Dim dbCur As DAO.Database
    
    Set dbCur = CurrentDb
    With dbCur
        
        Form_frmOPC.lblStatus.Caption = "Clearing Old ConnectMstr.."
        Form_frmOPC.Refresh
        DoEvents
        
        strSQL = "DELETE ConnectMstr.* " & _
                "FROM ConnectMstr;"
        Call .Execute(strSQL, dbFailOnError)                               'Clear old ConnectMstr
        
        strFCS = Dir(strPjtPath & "FCS*", vbDirectory)
        Do While strFCS <> ""
        
   
            Form_frmOPC.lblStatus.Caption = "Importing " & strFCS
            Form_frmOPC.Refresh
            DoEvents
                
            intFCS = CInt(Right(strFCS, 4))
            Set tbdTable = .CreateTableDef("ConnectTbl")
            tbdTable.Connect = ";DATABASE=" & strPjtPath & strFCS & "\etc\FcsRef.accdb"
            tbdTable.SourceTableName = "ConnectTbl"
            .TableDefs.Append tbdTable
            Set tbdTable = Nothing

            strSQL = "INSERT INTO ConnectMstr ( FCS, FromTagName, FromTagNum, FromItemName, ToTagName, ToTagNum, ToItemName ) " & _
                    "SELECT '" & intFCS & "' AS FCS, [ConnectTbl].FromTagName, [ConnectTbl].FromTagNum, [ConnectTbl].FromItemName, [ConnectTbl].ToTagName, [ConnectTbl].ToTagNum, [ConnectTbl].ToItemName " & _
                    "FROM ConnectTbl;"
            
            Call .Execute(strSQL, dbFailOnError)
            Call .TableDefs.Delete("ConnectTbl")
            strFCS = Dir(, vbDirectory)
        Loop
    
    End With
    Set dbCur = Nothing

  
End Sub

Public Sub BuildWinRef()

    'This subroutine creates a master WinRef table from each of the HIS WinRef tables
    'Kinda useful
    
    Dim strHIS, strSQL As String
    Dim intHIS As Integer
    Dim i As Long
    Dim tbdTable As DAO.TableDef
    Dim dbCur As DAO.Database
    
    Set dbCur = CurrentDb
    
    With dbCur
        
        Form_frmOPC.lblStatus.Caption = "Clearing WinRef.."
        Form_frmOPC.Refresh
        DoEvents
        
        strSQL = "DELETE WinRef.* " & _
                "FROM WinRef;"
        Call .Execute(strSQL, dbFailOnError)                                  'Clear ElemMstr
        
        strHIS = Dir(strPjtPath & "HIS*", vbDirectory)
        Do While strHIS <> ""
        
            Form_frmOPC.lblStatus.Caption = "Importing " & strHIS
            Form_frmOPC.Refresh
            DoEvents
                
            intHIS = CInt(Right(strHIS, 4))
            Set tbdTable = .CreateTableDef("WinRefTbl")
            tbdTable.Connect = ";DATABASE=" & strPjtPath & strHIS & "\etc\HisRef.accdb"
            tbdTable.SourceTableName = "WinRefTbl"
            .TableDefs.Append tbdTable

            Set tbdTable = Nothing

            strSQL = "INSERT INTO WinRef (HIS, WindowName, RefWinName) " & _
                    "SELECT '" & strHIS & "' AS HIS, WinRefTbl.WindowName AS WindowName, WinRefTbl.RefWinName AS RefWinName " & _
                    "FROM WinRefTbl;"
            
            Call .Execute(strSQL, dbFailOnError)
            Call .TableDefs.Delete("WinRefTbl")
            strHIS = Dir(, vbDirectory)
            DoEvents
            
        Loop
    End With
    
     
    Form_frmOPC.lblStatus.Caption = "Ready"
    Form_frmOPC.Refresh
    
    Set dbCur = Nothing


End Sub

Public Sub BuildDTLD()
    
    Dim strSrcFile As String
    Dim strPath As String
    Dim strTagName
    Dim strBlock As String
    Dim lngPos As Long          'Position of DTLD information in edf file
    Dim lngSize As Long         'Size of DTLD information in file
    
    Dim lngCount As Long
    
    Dim i, j, k As Long
    
    Dim bytData As Byte
    Dim strCH As String
    Dim strWord As String
    Dim strDTLD As String
    Dim strItem As String
    Dim strValue As String
    
    Dim strItemList As Variant
    
    Dim strItems() As String
    
    Dim strSQL As String
    Dim dbCur As DAO.Database
    Dim rstFNBK As DAO.Recordset
    Dim rstDTLD As DAO.Recordset
    Dim tdfDTLD As DAO.TableDef
    Dim fldITEM As DAO.field
    Dim dummy
    
    
    
    Set dbCur = CurrentDb
    With dbCur                      'Add DTLD table if it doesn't already exist
        
        On Error Resume Next
        Set tdfDTLD = .TableDefs("DTLD")
        If (Err.Description = "Item not found in this collection.") Then
            On Error GoTo 0
            
            Form_frmOPC.lblStatus.Caption = "Creating DTLD Table"
            Form_frmOPC.Refresh
            
            Set tdfDTLD = Nothing
            Set tdfDTLD = .CreateTableDef("DTLD")
            With tdfDTLD
                .Fields.Append .CreateField("TagName", dbText, 16)
            End With
            .TableDefs.Append tdfDTLD
            .TableDefs.Refresh
        End If
        On Error GoTo 0
        Set tdfDTLD = Nothing
        
        If (Form_frmOPC.optDTLDQRY.Value = 0) Then
        
            strSQL = "SELECT ElemMstr.FCS, ElemMstr.FileName, ElemMstr.TagName " & _
                "FROM ElemMstr " & _
                "WHERE (((ElemMstr.FileName) Like 'dr*')) " & _
                "ORDER BY ElemMstr.FCS, ElemMstr.FileName, ElemMstr.TagName;"
                
            Set rstFNBK = .OpenRecordset(strSQL, dbOpenDynaset)         'Query all function blocks
        Else
            Set rstFNBK = .OpenRecordset("qryDTLD", dbOpenDynaset)      'Use user defined Query
        End If

        Form_frmOPC.lblStatus.Caption = "Clearing old DTLD table"
        Form_frmOPC.Refresh
        
        strSQL = "DELETE DTLD.TagName " & _
                "FROM DTLD " & _
                "WHERE (((DTLD.TagName) Like '*'));"

        Call .Execute(strSQL, dbFailOnError)                            'Clear Old DTLD table
        Set rstDTLD = .OpenRecordset("DTLD", dbOpenDynaset)             'Open DTLD table
        
    End With
    
    
    With rstFNBK
        If Not .EOF Then
            .MoveLast
            .MoveFirst
        End If
        
        With Form_frmOPC
            .lblStatus.Caption = "Importing Function Block Data..."
            .pgbStatus.Min = 0
            .pgbStatus.Max = 100
            .pgbStatus.Value = 0
            .pgbStatus.Visible = True
            .Refresh
        End With

        
        Do While Not .EOF
            
            'Define file path and name
            strPath = strPjtPath & "FCS" & Format(!FCS, "0000") & _
                    "\Function_block\" & Left(!FileName, Len(!FileName) - 4)
            
            strTagName = !TagName
            strBlock = strTagName & ".edf"
            
            
            'Read file header to find location of block detailed info
            
            lngPos = 0
            lngSize = 0
            
            dummy = Dir(strPath & "\" & strBlock, vbNormal)
            If (dummy <> strBlock) Then
                dummy = dummy
                GoTo skip
            End If
            Open strPath & "\" & strBlock For Random Access Read Shared As #1 Len = 1
            Do While Not EOF(1)
            
                strWord = ""
                For i = 1 To 4
                    Get 1, , bytData
                    strWord = strWord & Chr(bytData)
                Next i
                
                If strWord = "DTLD" Then
                    'Read position of DTLD
                    Get 1, , bytData
                    lngPos = lngPos + bytData
                    Get 1, , bytData
                    lngPos = lngPos + bytData * &H100
                    Get 1, , bytData
                    lngPos = lngPos + bytData * &H10000
                    Get 1, , bytData
                    lngPos = lngPos + bytData * &H1000000
                    lngPos = lngPos + 1
                    
                    ' Read Length of DTLD string
                    Get 1, Seek(1) + 4, bytData
                    lngSize = lngSize + bytData
                    Get 1, , bytData
                    lngSize = lngSize + bytData * &H100
                    Get 1, , bytData
                    lngSize = lngSize + bytData * &H10000
                    Get 1, , bytData
                    lngSize = lngSize + bytData * &H1000000
                    
                    Exit Do
                End If
            Loop
            
            If (lngPos = 0 Or lngSize = 0) Then
                Close #1
                GoTo skip
            End If
            
            ' Read DTLD info into an array
            i = lngPos
            j = 0
            
            Do While i < (lngPos + lngSize - 1)
                j = j + 1
                strItem = ""
                strValue = ""
                
                Do While i < (lngPos + lngSize - 1)        'Read Item ID
                    Get 1, i, bytData
                    i = i + 1
                    strCH = Chr(bytData)
                    If (strCH = "!" Or strCH = ",") Then strCH = "_"
                    If (strCH = ":") Then Exit Do
                    strItem = strItem & strCH
                Loop
                
                
                If (strItem = "FEXP" Or strItem = "FPRM" Or strItem = "FZAS") Then  'Check for sub-items
                    k = i
                    strWord = ""
                    Do While i < (lngPos + lngSize - 1)         'Read sub-Item ID
                        Get 1, i, bytData
                        i = i + 1
                        strCH = Chr(bytData)
                        If (strCH = "!" Or strCH = ",") Then strCH = "_"
                        If (strCH = ":") Then
                            strItem = strItem & "_" & strWord
                            Exit Do
                        End If
                        If (strCH = ";") Then
                            i = k
                            Exit Do
                        End If
                        strWord = strWord & strCH
                    Loop
                End If
                
                Do While i < (lngPos + lngSize - 1)         'Read Value
                    Get 1, i, bytData
                    i = i + 1
                    strCH = Chr(bytData)
                    If strCH = ";" Then Exit Do
                    strValue = strValue & strCH
                Loop
                
                ReDim Preserve strItems(2, j)
                strItems(1, j) = strItem
                strItems(2, j) = strValue
                
            Loop
            Close #1
            
            'Add DTLD to DTLD table
Retry:
            With rstDTLD
                .AddNew
                !TagName = strTagName
                For i = 1 To j
                    On Error Resume Next
                    .Fields(strItems(1, i)) = strItems(2, i)
                    If (Err.Description = "Item not found in this collection.") Then
                        .Cancel
                        On Error GoTo 0
                        GoSub AddField
                        GoTo Retry
                    End If
                    On Error GoTo 0
                Next i
                .Update
            End With
skip:
            
            Erase strItems
            If (.PercentPosition * 10 Mod 10 = 0) Then
                Form_frmOPC.pgbStatus.Value = .PercentPosition
                DoEvents
            End If
            .MoveNext
        Loop
    End With
    
    With Form_frmOPC
        .lblStatus.Caption = "Ready"
        .pgbStatus.Value = 0
        .pgbStatus.Visible = False
        .Refresh
    End With
    
    dbCur.Close
    Set dbCur = Nothing
    
    Exit Sub
    
AddField:
    
    With dbCur
        rstDTLD.Close
        Set tdfDTLD = .TableDefs("DTLD")
        With tdfDTLD
            Set fldITEM = .CreateField(strItems(1, i), dbText, 32)
            fldITEM.AllowZeroLength = True
            .Fields.Append fldITEM
            Set fldITEM = Nothing
            .Fields.Refresh
            Set tdfDTLD = Nothing
        End With
        Set rstDTLD = .OpenRecordset("DTLD", dbOpenDynaset)
    End With
    
    Return

End Sub



Public Sub ListToFormat(strSTbl As String, strDTbl As String)
    
    Dim dbCur As DAO.Database
    Dim tbdDest As TableDef
    Dim rstTags As Recordset
    Dim rstItems As Recordset
    Dim rstSource As Recordset
    Dim rstDest As Recordset
    
    Dim strSQL As String
    
    Dim blnOk As Boolean
    
    Dim varSL
    Dim varSH
    Dim dummy
    
    Dim dblValue As Double
    
    On Error Resume Next
    
    Set dbCur = CurrentDb
    
    With dbCur
    
        On Error Resume Next
        Call .TableDefs.Delete(strDTbl)
        On Error GoTo 0
        
        strSQL = "SELECT DISTINCT " & strSTbl & ".Item " & _
                "FROM " & strSTbl & " " & _
                "WHERE (((" & strSTbl & ".Item) Not Like 'S[HL]'));"
        Set rstItems = .OpenRecordset(strSQL, dbOpenDynaset)
        
        Set tbdDest = .CreateTableDef(strDTbl)
        With tbdDest
            .Fields.Append .CreateField("TagName", dbText, 16)
            .Fields.Append .CreateField("Comment", dbText, 24)
            .Fields.Append .CreateField("EngUnit", dbText, 6)
            .Fields.Append .CreateField("SL", dbDouble)
            .Fields.Append .CreateField("SH", dbDouble)
            If Not rstItems.EOF Then
                Do
                    .Fields.Append .CreateField(rstItems!item, dbDouble)
                    rstItems.MoveNext
                Loop Until rstItems.EOF
            End If
        End With
        Call .TableDefs.Append(tbdDest)
        Set tbdDest = Nothing
        rstItems.Close
        
        strSQL = "SELECT " & strSTbl & ".* " & _
                "FROM " & strSTbl & " " & _
                "WHERE " & strSTbl & ".Item Not Like 'S[HL]' " & _
                "ORDER BY " & strSTbl & ".TagName, " & strSTbl & ".Item;"
        Set rstSource = .OpenRecordset(strSQL, dbOpenDynaset)
                
        strSQL = "SELECT DISTINCT ElemMstr.TagName, ElemMstr.Comment, ElemMstr.EngUnit, ElemMstr.SL, ElemMstr.SH " & _
                "FROM " & strSTbl & " INNER JOIN ElemMstr ON " & strSTbl & ".TagName = ElemMstr.TagName;"
                
        Set rstTags = .OpenRecordset(strSQL, dbOpenDynaset)
        
        Set rstDest = .OpenRecordset(strDTbl, dbOpenDynaset)
        Debug.Print strSQL
        With rstSource
            Do
                If !TagName <> rstTags!TagName Then rstTags.FindFirst ("TagName = '" & !TagName & "'")
                
                If Not .NoMatch Then
                
                    rstDest.AddNew
                    rstDest!TagName = rstTags!TagName
                    rstDest!Comment = rstTags!Comment
                    rstDest!EngUnit = rstTags!EngUnit
                    rstDest!SL = rstTags!SL
                    rstDest!SH = rstTags!SH
                    blnOk = False
                    Do
                        If !Value = "" Then
                            dblValue = 0
                        Else
                            On Error Resume Next
                            dblValue = !Value
                            On Error GoTo 0
                        End If
                        
                        Select Case !item
                            
                            Case "LL", "PL"
                                If ExcessDev(dblValue, rstDest!SL) Then
                                    rstDest.Fields(!item) = dblValue
                                    blnOk = True
                                    End If
                            Case "PH", "HH"
                                If ExcessDev(dblValue, rstDest!SH) Then
                                    rstDest.Fields(!item) = dblValue
                                    blnOk = True
                                End If
                            Case "DL", "VL", "DL2"
                                If ExcessDev(dblValue, (rstDest!SH - rstDest!SL)) Then
                                    rstDest.Fields(!item) = dblValue
                                    blnOk = True
                                End If
                            Case Else
                                rstDest.Fields(!item) = dblValue
                                blnOk = True
                        End Select
                        .MoveNext
                        If .EOF Then Exit Do
                    Loop Until (!TagName <> rstTags!TagName)
                    
                End If
                If (blnOk) Then rstDest.Update
                dummy = rstTags.AbsolutePosition
                
            Loop Until .EOF
        End With
        
        rstTags.Close
        rstSource.Close
        rstDest.Close
        
    End With
    Set dbCur = Nothing

    Exit Sub
hdlr:
    Set dbCur = Nothing
    Call MsgBox("Something went horribly wrong!!!", vbCritical)
End Sub

Public Function ExcessDev(dbl1 As Double, dbl2 As Double) As Boolean

    If (Abs(dbl1 + dbl2) > 0) Then
        ExcessDev = (200 * Abs(dbl1 - dbl2) / Abs(dbl1 + dbl2)) > 0.5
    Else
        ExcessDev = False
    End If
        
End Function

Public Sub FormatToList(strSTbl As String, strDTbl As String, lngCol As Long)
    
    Dim dbCur As DAO.Database
    Dim tbdDest As DAO.TableDef
    Dim rstSource As DAO.Recordset
    Dim rstDest As DAO.Recordset
    
    Dim strSQL As String
    Dim blnOk As Boolean
    Dim i, n As Long

    Dim dummy
    
    Set dbCur = CurrentDb
    With dbCur
        
        On Error Resume Next
        Call .TableDefs.Delete(strDTbl)
        On Error GoTo 0
    
        Set tbdDest = .CreateTableDef(strDTbl)
        With tbdDest
            .Fields.Append .CreateField("TagName", dbText, 16)
            .Fields.Append .CreateField("Item", dbText, 16)
            .Fields.Append .CreateField("Value", dbText, 255)
        End With
        Call .TableDefs.Append(tbdDest)
        Set tbdDest = Nothing
        
        Set rstSource = .OpenRecordset(strSTbl)
        Set rstDest = .OpenRecordset(strDTbl)
        
        
        With rstSource
        
            .MoveLast
            .MoveFirst
            With Form_frmOPC
                .lblStatus.Caption = "Transposing Table..."
                .pgbStatus.Min = 0
                .pgbStatus.Max = 100
                .pgbStatus.Value = 0
                .pgbStatus.Visible = True
                .Refresh
            End With
            
            n = rstSource.Fields.Count
            Do
                Form_frmOPC.pgbStatus.Value = .PercentPosition
                For i = lngCol To n - 1
                    If (Not IsNull(.Fields(i)) And .Fields(i) <> "") Then
                        rstDest.AddNew
                        rstDest!TagName = !TagName
                        rstDest!item = .Fields(i).Name
                        rstDest!Value = .Fields(i).Value
                        rstDest.Update
                    End If
                Next i
            DoEvents
            .MoveNext
            Loop Until .EOF
            .Close
            rstDest.Close
        End With
    End With
    Set dbCur = Nothing
    
    With Form_frmOPC
        .lblStatus.Caption = "Ready"

        .pgbStatus.Visible = False
        .Refresh
    End With

End Sub

