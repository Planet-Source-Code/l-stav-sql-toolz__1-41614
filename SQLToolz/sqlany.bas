Attribute VB_Name = "sqlany"
Option Explicit
Option Compare Text
Dim sReserved() As String
Dim bShowPlan As Boolean
Dim bSQLAny As Boolean

Const imgTABLE = 1
Const imgVIEW = 2
Const imgINDEX = 3
Const imgPROPERTY = 4
Const imgCOLUMN = 5
Const imgDB = 6
Const imgFOLDCLOSE = 7
Const imgFOLDOPEN = 8
Const imgPROC = 9
Const imgARG = 10
Const imgPRIMARY = 11
Const imgFOREIGN = 12
Const imgUDT = 13
Const imgTRIGGER = 14







Public Function DBInfo(iCase As Integer) As Boolean
Dim sSQL As String
Dim sDQUOTE As String
Dim sSQUOTE As String

sDQUOTE = Chr(34)
sSQUOTE = Chr(34)

If bDebug Then DebugWrite "DBInfo Starts"

Select Case iCase
    Case 0
        If bDebug Then DebugWrite "Case 0 - No Values"
        frmMain.sbStatus.Panels("Server").Text = "Server: "
        frmMain.sbStatus.Panels("DB").Text = "DB: "
        frmMain.sbStatus.Panels("User").Text = "User: "
        frmMain.sDBServer = ""
        frmMain.SetCaption
    Case 1
        
        ' find out if we are in sqlanywhere
        ' a real quick test that should weed out others
        If conSQL.Transactions = False Then
            bSQLAny = False
            frmMain.tvDB.Enabled = False
            If bDebug Then DebugWrite "Transactions = False"
        Else
            sSQL = "Select type " _
                & "From sysobjects " _
                & "where name = 'sysobjects'"
            If bDebug Then DebugWrite sSQL
            Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
            If rsSQL.rdoColumns(0).Value = "V" Then
                bSQLAny = True
                frmMain.tvDB.Enabled = True
                If bDebug Then DebugWrite "bSQLAny = True"
            Else
                bSQLAny = False
                frmMain.tvDB.Enabled = False
                If bDebug Then DebugWrite "bSQLAny = False"
            End If
            rsSQL.Close
            sSQL = "select @@servername " _
                & sDQUOTE & "Server" & sDQUOTE _
                & ", db_name() " & sDQUOTE & "Database" & sDQUOTE _
                & ", user_name() " & sDQUOTE & "User" & sDQUOTE
                If bDebug Then DebugWrite sSQL
            Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
            frmMain.sDBServer = rsSQL.rdoColumns(0).Value
            frmMain.sbStatus.Panels("Server").Text = "Server: " & rsSQL.rdoColumns(0).Value & "  "
            frmMain.sbStatus.Panels("DB").Text = "DB: " & rsSQL.rdoColumns(1).Value & "  "
            frmMain.sbStatus.Panels("User").Text = "User: " & rsSQL.rdoColumns(2).Value & "  "
            frmMain.SetCaption
            
            If bDebug Then
                DebugWrite "Server: " & rsSQL.rdoColumns(0).Value & " "
                DebugWrite "DB: " & rsSQL.rdoColumns(1).Value & " "
                DebugWrite "User: " & rsSQL.rdoColumns(2).Value & " "
            End If
            
            TreeInitialize
            frmMain.tvDB.Nodes(1).Text = rsSQL.rdoColumns(1).Value
            rsSQL.Close
        End If
End Select
If bDebug Then DebugWrite "DBInfo Ends"
DBInfo = True

End Function

Public Sub GetDirect()
Dim lDirect As Long
Dim sDirect As String

If bDebug Then DebugWrite "GetDirect Starts"
ReDim sReserved(0)

' Handle direct sql options
ReDim sReserved(13)
sReserved(0) = "INSERT "
sReserved(1) = "UPDATE "
sReserved(2) = "DELETE "
sReserved(3) = "CREATE "
sReserved(4) = "ALTER "
sReserved(5) = "DROP "
sReserved(6) = "USE "
sReserved(7) = "SET "
sReserved(8) = "SAVEPOINT"
sReserved(9) = "ROLLBACK TO SAVEPOINT"
sReserved(10) = "GRANT "
sReserved(11) = "REVOKE "
sReserved(12) = "TRUNCATE TABLE "

' Are there any additional ones
If Dir(App.Path & "\direct.txt") <> "" Then
    lDirect = FreeFile(0)
    Open App.Path & "\direct.txt" For Input As #lDirect
    Do While Not EOF(lDirect)
        Line Input #lDirect, sDirect
        If Left$(sDirect, 1) = "'" Then
            sDirect = Mid$(sDirect, 2, (InStr(2, sDirect, "'") - 2))
            If sDirect <> "" Then
                ReDim Preserve sReserved(UBound(sReserved) + 1)
                sReserved(UBound(sReserved) - 1) = sDirect
            End If
        End If
    Loop
    Close #lDirect
End If
        
'        sReserved(8) = "COMMENT "
'        sReserved(13) = "CHECKPOINT"

If bDebug Then DebugWrite "GetDirect Ends"
End Sub

Public Function GetTextChunks(rdoText As rdoColumn) As String
    Dim sChunkText As String, sText As String
    Dim bContinue As Boolean
    
    Dim iCount As Integer, iChunks As Integer
    Dim lColLength As Long
    Const ChunkSize As Integer = 16384
    
    ' since rdoText.ColumnSize doesn't always return a correct value
    ' we have to call getchunk() until we get a number less than
    ' ChunkSize or we get null
    bContinue = True
    Do
        sChunkText = rdoText.GetChunk(ChunkSize)
        If Len(sChunkText) <> 0 Then
            sText = sText & sChunkText
            If Len(sChunkText) < ChunkSize Then
                bContinue = False
            End If
        ElseIf IsNull(sChunkText) = True Then
            bContinue = False
        End If
        sChunkText = ""
    Loop Until bContinue = False
    If bDebug Then DebugWrite sText
    GetTextChunks = sText
End Function

Public Function ODBCTrim(sMsg As String) As String
Dim lTotalLength As Long, lCurrentPos As Long

If bDebug Then DebugWrite "ODBCTrim Starts"

lTotalLength = Len(sMsg)
If lTotalLength > 1 Then
    ' are there any end brackets
    Do
        lCurrentPos = 0
        lCurrentPos = InStr(sMsg, "]")
        If lCurrentPos = 0 Then
            Exit Do
        End If
        ' make sure that this is not the last character
        If lCurrentPos <= lTotalLength - 1 Then
            sMsg = Mid$(sMsg, lCurrentPos + 1)
        lTotalLength = Len(sMsg)
        End If
    Loop While lCurrentPos > 0
End If
    
ODBCTrim = sMsg

If bDebug Then DebugWrite "ODBCTrim Ends"

End Function

Public Function ShowPlanGet(sSQL As String) As Boolean
Dim sPlan As String, sMsg As String
Dim lTotalLength As Long, lCurrentPos As Long

If bDebug Then DebugWrite "ShowPlanGet Starts"

sPlan = sSQL
SQLEscape sPlan
On Error GoTo PlanError
If bSQLAny = True And bShowPlan Then

    ' since we use the plan function, we have to replace "'" with hex \x27
    lTotalLength = Len(sPlan)
    If lTotalLength > 1 Then
        ' are there any ' characters
        lCurrentPos = 0
        Do
            lCurrentPos = InStr(lCurrentPos + 1, sPlan, "'")
            If lCurrentPos = 0 Then
                Exit Do
            End If
            ' make sure that this is not the last character
            If lCurrentPos <= lTotalLength - 1 Then
                sPlan = Left$(sPlan, lCurrentPos - 1) & "\x27" _
                    & Mid$(sPlan, lCurrentPos + 1)
            Else
                sPlan = Left$(sPlan, lCurrentPos - 1) & "\x27"
            lTotalLength = Len(sPlan)
            End If
        Loop While lCurrentPos > 0
    End If


    ' get the plan
    sPlan = "Select plan('" & sPlan & "')"
    Set rsSQL = conSQL.OpenResultset(sPlan, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        If rsSQL.rdoColumns(0).ChunkRequired = True Then
            sPlan = GetTextChunks(rsSQL.rdoColumns(0))
        Else
            sPlan = rsSQL.rdoColumns(0).Value
        End If
    End If
    rsSQL.Close
    
    ' now we need to parse out the values
    lTotalLength = Len(sPlan)
    If lTotalLength > 1 Then
        sMsg = " "
        frmMain.StatsAdd sMsg
        ' are there any crlf
        Do
            lCurrentPos = 0
            lCurrentPos = InStr(sPlan, vbLf)
            If lCurrentPos = 0 Then
                sMsg = " "
                frmMain.StatsAdd sMsg
                Exit Do
            End If
            sMsg = Left$(sPlan, lCurrentPos - 1)
            ' make sure that this is not the last character
            If lCurrentPos <= lTotalLength - 1 Then
                sPlan = Mid$(sPlan, lCurrentPos + 1)
            Else
                sPlan = Left$(sPlan, lCurrentPos - 1)
            End If
            frmMain.StatsAdd sMsg
            lTotalLength = Len(sPlan)
        Loop While lCurrentPos > 0
    End If

End If
If bDebug Then DebugWrite "ShowPlanGet Ends"
ShowPlanGet = True
Exit Function

PlanError:
    rdoErrors.Clear
    On Error GoTo 0
    Exit Function
End Function

Public Function ShowPlanSet(iCase As Integer) As Boolean
Dim sSQL As String
Dim bRetVal As Boolean

Screen.MousePointer = vbHourglass

If bDebug Then DebugWrite "ShowPlanSet Starts"

On Error GoTo ConnectError
    If envSQL.rdoConnections.Count > 0 Then
        On Error GoTo SQLError
        Select Case iCase
            Case 0
                If bSQLAny = False Then
                    ' here we need to turn showplan off
                    sSQL = "set showplan off"
                    SQLDirect sSQL
                End If
                bShowPlan = False
                If bDebug Then DebugWrite "Case 0 - ShowPlan False"
            Case 1
                If bSQLAny = True Then
                    bShowPlan = True
                    If bDebug Then DebugWrite "Case 1 - ShowPlan True"
                Else
                    ' we are going to switch the value so that we don't execute Plan all the time
                    bShowPlan = False
                    ' here we need to turn showplan on
                    sSQL = "set showplan on"
                    SQLDirect sSQL
                    If bDebug Then DebugWrite "Case 1 - ShowPlan True/False"
                End If
                
        End Select
    End If
If bDebug Then DebugWrite "ShowPlanSet Ends"
ShowPlanSet = True
Exit Function

ConnectError:
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number = 91 Then
        bRetVal = frmMain.LoginShow
        If bRetVal = True Then
            Screen.MousePointer = vbHourglass
            Resume
        Else
            frmMain.StatsDisplay 7
            Screen.MousePointer = vbDefault
            ShowPlanSet = False
            Exit Function
        End If
    Else: GoTo SQLError
    End If
    frmMain.StatsDisplay 0
    frmMain.MenuSet 1
    ShowPlanSet = False
    Exit Function

SQLError:
    ' a cursor error may be triggered when we _
      run a stored procedure that doesn't return _
      and rows.  We will test for this and continue
    If Err.Number = 40088 Then
        DBActionRows
        Screen.MousePointer = vbDefault
        Resume Next
    End If
'    If Err.Number = 40086 Then
'        Resume Next
'    End If
    Dim iCount As Integer
    Dim sMsg As String
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    ElseIf rdoErrors.Count <> 0 Then
        For iCount = 0 To rdoErrors.Count - 1
            MsgBox rdoErrors(iCount).Description, vbExclamation + vbOKOnly
            Next
        rdoErrors.Clear
    End If
    Screen.MousePointer = vbDefault
    frmMain.MenuSet 1
    bCancel = False
    bProcessing = False
    ShowPlanSet = False
    Exit Function
End Function

Public Function SQLEscape(sSQL As String) As String
Dim lTotalLength As Long, lCurrentPos As Long

If bDebug Then DebugWrite "SQLEscape Starts"

' this should be handled by the db, but we need it for Plan()
lTotalLength = Len(sSQL)
If lTotalLength > 1 And bSQLAny = True Then
    ' are there any escape characters
    lCurrentPos = 0
    Do
        lCurrentPos = InStr(lCurrentPos + 1, sSQL, "\")
        If lCurrentPos = 0 Then
            Exit Do
        End If
        ' make sure that this is not the last character
        If lCurrentPos <= lTotalLength - 1 Then
            sSQL = Left$(sSQL, lCurrentPos) & "x" _
                & Trim(Hex(Asc(Mid$(sSQL, lCurrentPos + 1, 1)))) _
                & Mid$(sSQL, lCurrentPos + 2)
        lTotalLength = Len(sSQL)
        End If
    Loop While lCurrentPos > 0
End If
    
SQLEscape = sSQL

If bDebug Then DebugWrite "SQLEscape Ends"

End Function

Public Function SQLParse() As Boolean
Dim sSQL As String, sOriginalSQL As String
Dim sSingleSQL As String, sSeperator As String
Dim lCurrentPos As Long, lCursorPos As Long
Dim bRetVal As Boolean, bOrigRTF As Boolean

If bDebug Then DebugWrite "SQLParse Starts"

If Not bProcessing Then
    bProcessing = True
    Screen.MousePointer = vbHourglass
    frmMain.MenuSet 0
    frmMain.StatsDisplay 1
    
    ' cmax
    lCursorPos = frmMain.rtfSQL.GetSel(True)
    
    ' is it the whole control, or just highlighted text
    If Len(frmMain.rtfSQL.SelText) = 0 Then
        sSQL = frmMain.rtfSQL.Text
    Else
        sSQL = frmMain.rtfSQL.SelText
    End If

    ' ok, set the seperator
    sSeperator = vbCrLf & "go"
    
    ' add the sql to the history list
    With frmMain.grdHistory
        '.Col = 1
        If .Rows > 99 Then .RemoveItem 1
        If .Text <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
        .Text = sSQL
    End With
    
    ' trim the leading and trailing spaces
    sSQL = SQLTrim(sSQL)
    
    'Loop through sSQL and execute each statement
    If InStr(sSQL, sSeperator) > 0 Then     ' found at least once
        Do
           lCurrentPos = 0
           lCurrentPos = InStr(sSQL, sSeperator)
           If lCurrentPos = 0 Then
                ' nothing there
                If Len(sSQL) = 0 Then
                    Exit Do
                Else
                    ' last sql
                    sSingleSQL = sSQL
                End If
           Else
                ' get the current command, and everything that is left
                sSingleSQL = Left$(sSQL, lCurrentPos - 1)
                sSQL = Mid$(sSQL, lCurrentPos + Len(sSeperator))
           End If
           sSingleSQL = SQLTrim(sSingleSQL)
           If Len(sSingleSQL) > 0 Then
                ' display the current sql
                frmMain.cmRun.Text = ""
                frmMain.cmRun.Text = sSingleSQL
                frmMain.rtfSQL.Visible = False
                frmMain.cmRun.Visible = True
                DoEvents
                ' log it if necessary
                If bLogFile Then LogWrite (sSingleSQL)
                If bCancel Then Exit Do
                frmMain.StatsDisplay 1
                ShowPlanGet sSingleSQL
                bRetVal = SQLExecute(sSingleSQL)
                Screen.MousePointer = vbHourglass
                If bRetVal = False Then Exit Do
           Else
                ' nothing to do, reset the stats
                frmMain.StatsDisplay 3
           End If
        Loop While lCurrentPos > 0
    Else
        ' single statement, execute it
        sSQL = SQLTrim(sSQL)
        If Len(sSQL) > 0 Then
            frmMain.StatsDisplay 1
            If bLogFile Then LogWrite (sSQL)
            ShowPlanGet sSQL
            bRetVal = SQLExecute(sSQL)
            Screen.MousePointer = vbHourglass
        Else
            ' nothing to do, reset the stats
            frmMain.StatsDisplay 3
        End If
    End If
    
    ' cmax
    Call frmMain.rtfSQL.SetCaretPos(lCursorPos, 0)
    frmMain.rtfSQL.Visible = True
    frmMain.cmRun.Text = ""
    frmMain.cmRun.Visible = False
    frmMain.rtfSQL.SetFocus
    DoEvents
    
    Screen.MousePointer = vbDefault
    frmMain.MenuSet 1
    bCancel = False
    bProcessing = False
    If frmMain.grdSQL.Cols > 1 And frmMain.grdSQL.Rows > 1 Then
        frmMain.grdSQL.Col = 1
        frmMain.grdSQL.Row = 1
    End If
    DoEvents

End If
If bDebug Then DebugWrite "SQLParse Ends"
End Function


Public Function SQLTrim(sSQL As String) As String
Dim lTotalLength As Long, lCurrentPos As Long
Dim sCheckString As String

If bDebug Then DebugWrite "SQLTrim Starts"

lTotalLength = Len(sSQL)
If lTotalLength > 0 Then
    ' read leading characters
    lCurrentPos = 0
    Do
        lCurrentPos = lCurrentPos + 1
        sCheckString = Mid$(sSQL, lCurrentPos, 1)
        If Asc(sCheckString) > 32 And Asc(sCheckString) < 127 Then
            Exit Do
        End If
        sSQL = Left$(sSQL, lCurrentPos - 1) & " " & Mid$(sSQL, lCurrentPos + 1)
    Loop While lCurrentPos < lTotalLength
    sSQL = Trim(sSQL)
    ' read trailing characters
    lTotalLength = Len(sSQL)
    If lTotalLength > 0 Then
        lCurrentPos = lTotalLength + 1
        Do
            lCurrentPos = lCurrentPos - 1
            sCheckString = Mid$(sSQL, lCurrentPos, 1)
            If Asc(sCheckString) > 32 And Asc(sCheckString) < 127 Then
                Exit Do
            End If
            sSQL = Left$(sSQL, lCurrentPos - 1) & " " & Mid$(sSQL, lCurrentPos + 1)
        Loop While lCurrentPos > 0
        sSQL = Trim(sSQL)
    End If
    
    ' make sure that this isn't the seperator
    lTotalLength = Len(sSQL)
    If lTotalLength > 2 Then
        sCheckString = Mid$(sSQL, 3, 1)
        If Left$(sSQL, 2) = "go" And (Asc(sCheckString) <= 32 Or Asc(sCheckString) > 127) Then
            sSQL = SQLTrim(Mid$(sSQL, 3))
        End If
    ElseIf lTotalLength = 2 Then
        If sSQL = "go" Then sSQL = ""
    End If
End If
    
SQLTrim = sSQL

If bDebug Then DebugWrite "SQLTrim Ends"
End Function

Public Function SQLDirect(sSQL As String) As Boolean
    If bDebug Then DebugWrite "SQLDirect Starts"
    conSQL.Execute sSQL, rdExecDirect
    DBActionRows
    If bDebug Then DebugWrite "SQLDirect Ends"
    SQLDirect = True
End Function
Public Function DBActionRows() As Boolean
Dim lCount As Long
    lCount = conSQL.RowsAffected
    If lCount >= 0 Then
        frmMain.StatsDisplay 2, lCount
    Else
        frmMain.StatsDisplay 0
    End If
End Function
Public Function SQLSelect(sSQL As String) As Boolean
Dim sMsg As String

    If bDebug Then DebugWrite "SQLSelect Starts"
    If bDebug Then DebugWrite sSQL
    ' open the result set
    Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
    ' we just need to see if we can trigger a cursor error
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        'do nothing
        If bDebug Then DebugWrite "Not BOF and EOF"
    End If

    ' display the results
    frmMain.GridDisplay
    ' release the resources
    If bDebug Then DebugWrite "Closing ResultSet"
    rsSQL.Close
    If bDebug Then DebugWrite "SQLSelect Ends"
    SQLSelect = True
    
End Function
Public Function SQLExecute(sSQL As String) As Boolean
Dim bRetVal As Integer
Dim sTransaction() As String
Dim iCount As Integer
Dim sMsg As String
    
Execute:
    Screen.MousePointer = vbHourglass
    If bDebug Then DebugWrite "SQLExecute Starts"
    If bDebug Then DebugWrite sSQL
    On Error GoTo ConnectError
    If envSQL.rdoConnections.Count > 0 Then
        On Error GoTo SQLError
        ' Handle transaction options
        If InStr(sSQL, "BEGIN TRAN") = 1 Then
            If Not bCancel Then
                envSQL.CommitTrans
                envSQL.BeginTrans
            End If
            SQLExecute = True
            Screen.MousePointer = vbDefault
            frmMain.StatsDisplay 0
            If bDebug Then DebugWrite "SQLExecute Exit Function"
            Exit Function
        End If
        
        If InStr(sSQL, "COMMIT") = 1 Then
            If Not bCancel Then
                SQLDirect ("Commit")
                envSQL.CommitTrans
                If bSQLAny = True Then envSQL.BeginTrans
            End If
            SQLExecute = True
            Screen.MousePointer = vbDefault
            frmMain.StatsDisplay 0
            If bDebug Then DebugWrite "SQLExecute Exit Function"
            Exit Function
        End If
        
        If InStr(sSQL, "ROLLBACK") = 1 Then
            If InStr(sSQL, "SAVEPOINT") = 0 Then
                If Not bCancel Then
                    SQLDirect ("Rollback")
                    envSQL.RollbackTrans
                    If bSQLAny = True Then envSQL.BeginTrans
                End If
                SQLExecute = True
                Screen.MousePointer = vbDefault
                frmMain.StatsDisplay 0
                If bDebug Then DebugWrite "SQLExecute Exit Function"
                Exit Function
            End If
        End If
           
        ' Handle direct sql options
        For iCount = 0 To UBound(sReserved) - 1
            If InStr(sSQL, sReserved(iCount)) = 1 Then
                 'execute with direct
                bRetVal = SQLDirect(sSQL)
                SQLExecute = bRetVal
                If bRetVal And iCount = 6 Then DBInfo 1
                If bDebug Then DebugWrite "SQLExecute Exit Function"
                Exit Function
            End If
        Next
        
        'execute or select
         bRetVal = SQLSelect(sSQL)
        frmMain.grdSQL.Visible = True
    Else
        bRetVal = frmMain.LoginShow
        If bRetVal = True Then
            Screen.MousePointer = vbHourglass
            GoTo Execute
        Else
            frmMain.StatsDisplay 7
        End If
    End If
    
    ' are there messages to display
    If rdoErrors.Count <> 0 Then
        For iCount = (rdoErrors.Count - 1) To 0 Step -1
            sMsg = rdoErrors(iCount).Description
            sMsg = ODBCTrim(sMsg)
            sMsg = SQLTrim(sMsg)
            frmMain.StatsAdd sMsg
        Next
        rdoErrors.Clear
    End If
    
    If bDebug Then DebugWrite "SQLExecute Ends"
    SQLExecute = bRetVal
    Screen.MousePointer = vbDefault
    
Exit Function

ConnectError:
    If bDebug Then DebugWrite "Connect Error"
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number = 91 Then
        bRetVal = frmMain.LoginShow
        If bRetVal = True Then
            Screen.MousePointer = vbHourglass
            Resume
        Else
            frmMain.StatsDisplay 7
            Screen.MousePointer = vbDefault
            SQLExecute = False
            Exit Function
        End If
    Else: GoTo SQLError
    End If
    frmMain.StatsDisplay 0
    If frmMain.grdSQL.Rows > 1 Then
        frmMain.grdSQL.HighLight = flexHighlightAlways
    Else
        frmMain.grdSQL.HighLight = flexHighlightNever
    End If
    frmMain.grdSQL.Visible = True
    SQLExecute = False
    Exit Function

SQLError:
    If bDebug Then DebugWrite "SQL Error"
    ' a cursor error may be triggered when we _
      run a stored procedure that doesn't return _
      and rows.  We will test for this and continue
    If Err.Number = 40088 Then
        DBActionRows
        Screen.MousePointer = vbDefault
        bRetVal = True
        Resume Next
    End If
'    If Err.Number = 40086 Then
'        Resume Next
'    End If
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        If bDebug Then DebugWrite sMsg
        Err.Clear
        rdoErrors.Clear
    ElseIf rdoErrors.Count <> 0 Then
        For iCount = 0 To rdoErrors.Count - 1
            MsgBox rdoErrors(iCount).Description, vbExclamation + vbOKOnly
            If bDebug Then DebugWrite sMsg
        Next
        rdoErrors.Clear
    End If
    frmMain.StatsDisplay 0
    If frmMain.grdSQL.Rows > 1 Then
        frmMain.grdSQL.HighLight = flexHighlightAlways
    Else
        frmMain.grdSQL.HighLight = flexHighlightNever
    End If
    frmMain.grdSQL.Visible = True
    SQLExecute = False
    Exit Function

End Function

Public Function DBConnect(sDSN As String, Optional sUID As String, Optional sPWD As String) As Boolean
Dim iCount As Integer
Dim sMsg As String
   
If bDebug Then DebugWrite "DBConnect Starts"

    On Error GoTo ConnectError
    Screen.MousePointer = vbHourglass
    frmMain.StatsDisplay 6
    DBInfo 0
    DoEvents
   
    On Error Resume Next
    Err.Clear
    If envSQL.rdoConnections.Count > 0 Then
        If Err.Number <> 91 Then
            Call SQLExecute("COMMIT")
            conSQL.Close
        End If
    End If
    Err.Clear
    
    ' get the direct sql values
    GetDirect
    
    On Error GoTo ConnectError
    envSQL.UserName = sUID
    envSQL.Password = sPWD
    rdoErrors.Clear
    If GetSetting(App.Title, "Options", "CursorDriver", "Server") = "Server" Then
        envSQL.CursorDriver = rdUseServer
    Else
        envSQL.CursorDriver = rdUseOdbc
    End If
    Set conSQL = envSQL.OpenConnection(sDSN, rdDriverNoPrompt)
'    Call SQLExecute("BEGIN TRAN")
    Screen.MousePointer = vbHourglass
    frmMain.StatsDisplay 6
    DBInfo 1
    If bSQLAny = True Then Call SQLExecute("BEGIN TRAN")
    ' should we show the plan
    If GetSetting(App.Title, "Options", "ShowPLan", "0") = "0" Then
        bShowPlan = False
        ShowPlanSet 0
    Else
        bShowPlan = True
        ShowPlanSet 1
    End If
    frmMain.StatsDisplay 0
    Screen.MousePointer = vbDefault
    Err.Clear
    rdoErrors.Clear
    If bDebug Then DebugWrite "DBConnect Ends"
    DBConnect = True
    Exit Function
    
ConnectError:
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    ElseIf rdoErrors.Count <> 0 Then
        For iCount = 0 To rdoErrors.Count - 1
            MsgBox rdoErrors(iCount).Description, vbExclamation + vbOKOnly
            Next
        rdoErrors.Clear
    End If
    If bDebug Then DebugWrite "DBConnect Error"
    If bDebug Then DebugWrite sMsg
    frmMain.StatsDisplay 0
    frmMain.grdSQL.Visible = True
    DBConnect = False
    Exit Function

End Function


Public Function TreeGetTriggers(sParentKey As String) As Boolean
Dim sSQL As String, sKey As String, sText As String
Dim lObjectID As Long
Dim iCount As Integer

TreeGetTriggers = False
If bDebug Then DebugWrite "TreeGetTriggers Starts"
        
lObjectID = CLng(Mid$(sParentKey, 4))
    
sSQL = "Select isnull(trigger_name, 'REFACTION'), trigger_id, trigger_time, " _
    & "If event = 'C' then 'T' else event endif " & sDQUOTE & "MyEvent" _
    & sDQUOTE & ", isnull(trigger_order, 0) " _
    & "From SYS.SYSTRIGGER " _
    & "Where table_id = " & lObjectID
sSQL = sSQL & " and referential_action is null"
sSQL = sSQL & " Order by " & sDQUOTE & "MyEvent" _
    & sDQUOTE & ", trigger_time, trigger_order, trigger_name"
If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sText = rsSQL.rdoColumns(0).Value
        sKey = "tr_" & lObjectID & "_" & rsSQL.rdoColumns(1).Value
        ' add event
        If rsSQL.rdoColumns(3).Value = "I" Then
           sText = sText & "  Insert"
        ElseIf rsSQL.rdoColumns(3).Value = "U" Then
           sText = sText & "  Update"
        ElseIf rsSQL.rdoColumns(3).Value = "D" Then
           sText = sText & "  Delete"
        ElseIf rsSQL.rdoColumns(3).Value = "T" Then
           sText = sText & "  Update Column"
        End If
        ' add time
        If rsSQL.rdoColumns(2).Value = "A" Then
           sText = sText & "  After"
        ElseIf rsSQL.rdoColumns(2).Value = "B" Then
           sText = sText & "  Before"
        ElseIf rsSQL.rdoColumns(2).Value = "R" Then
           sText = sText & "  Resolve"
        End If
        ' add order
        sText = sText & "  " & CInt(rsSQL.rdoColumns(4).Value)
        
        frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgTRIGGER
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetTriggers Ends"
TreeGetTriggers = True
End Function

Public Function TreeGetViewDDL(sMyKey As String) As Boolean
Dim sSQL As String, sText As String, sAuth As String
' for permissions
Dim sSelect As String, sInsert As String
Dim sDelete As String, sUpdate As String
Dim sUpdateCols As String, sAlter As String
Dim sGrantee As String, sGrantor As String
Dim sOwner As String, sTableName As String
Dim sReference As String, sColumn As String
Dim lGrantee As Long, lGrantor As Long
Dim bAuth As Boolean
Dim rsSubSQL As rdoResultset

Dim lObjectID As Long


TreeGetViewDDL = False
If bDebug Then DebugWrite "TreeGetViewDDL Starts"

lObjectID = CLng(Mid$(sMyKey, 3))
bAuth = False
sText = ""
sAuth = ""
sColumn = ""

sSQL = "Select table_id, view_def, " _
    & "(If exists(select * from SYS.SYSTABLEPERM " _
    & "where table_id = " & lObjectID & ") then 'Y' " _
    & "else 'N' endif) Auth " _
    & "From SYS.SYSTABLE " _
    & "Where table_id = " & lObjectID

If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    If rsSQL.rdoColumns(1).ChunkRequired = True Then
        sText = GetTextChunks(rsSQL.rdoColumns(1))
    Else
        sText = rsSQL.rdoColumns(1).Value
    End If
    If rsSQL.rdoColumns(2).Value = "Y" Then bAuth = True
End If
rsSQL.Close
    
If sText <> "" Then
    If Not IsNull(sText) Then
        If UCase(Left(sText, 11)) = "CREATE VIEW" Then
            sText = "Alter View" & Mid(sText, 12)
        End If
        sText = sText & vbCrLf & "go" & vbCrLf
        
        ' are there any rights
        ' this section gets a little more complex, because we also
        ' need to check for individual columns
        If bAuth = True Then
            sSQL = "Select A.user_name " & sDQUOTE & "SGRANTEE" & sDQUOTE _
                & ", B.user_name " & sDQUOTE & "OWNER" & sDQUOTE _
                & ", table_name " _
                & ", C.user_name " & sDQUOTE & "SGRANTOR" & sDQUOTE _
                & ", SYSTABLEPERM.selectauth, SYSTABLEPERM.insertauth" _
                & ", SYSTABLEPERM.deleteauth, SYSTABLEPERM.updateauth" _
                & ", SYSTABLEPERM.updatecols, SYSTABLEPERM.alterauth" _
                & ", SYSTABLEPERM.referenceauth " _
                & ", SYSTABLEPERM.grantee " _
                & ", SYSTABLEPERM.grantor " _
                & "FROM SYS.SYSUSERPERMS A, SYS.SYSTABLEPERM, " _
                & "SYS.SYSTABLE, SYS.SYSUSERPERMS B, " _
                & "SYS.SYSUSERPERMS C " _
                & "Where SYSTABLEPERM.grantee = A.user_id " _
                & "and SYSTABLE.creator = B.user_id " _
                & "and SYSTABLE.table_id = SYSTABLEPERM.stable_id " _
                & "and SYSTABLEPERM.grantor = C.user_id " _
                & "and SYSTABLEPERM.stable_id = " & lObjectID

            If bDebug Then DebugWrite sSQL
            Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
            If Not rsSQL.BOF And Not rsSQL.EOF Then
                Do While Not rsSQL.EOF
                    ' lets get our values
                    sGrantee = sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE
                    sOwner = sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE
                    sTableName = sDQUOTE & rsSQL.rdoColumns(2).Value & sDQUOTE
                    sGrantor = sDQUOTE & rsSQL.rdoColumns(3).Value & sDQUOTE
                    sSelect = rsSQL.rdoColumns(4).Value
                    sInsert = rsSQL.rdoColumns(5).Value
                    sDelete = rsSQL.rdoColumns(6).Value
                    sUpdate = rsSQL.rdoColumns(7).Value
                    sUpdateCols = rsSQL.rdoColumns(8).Value
                    sAlter = rsSQL.rdoColumns(9).Value
                    sReference = rsSQL.rdoColumns(10).Value
                    lGrantee = rsSQL.rdoColumns(11).Value
                    lGrantor = rsSQL.rdoColumns(12).Value
                    
                    ' lets show the grantor for this set
                    sAuth = sAuth & "/* Grantor=" & sGrantor & " */" & vbCrLf
                    
                    ' select permissions
                    If sSelect <> "N" Then
                        sAuth = sAuth & "Grant Select on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sSelect = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    Else
                        ' they might have only permissions on columns
                        sSQL = "Select column_name, is_grantable " _
                            & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                            & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                            & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                            & "and SYSCOLPERM.table_id = " & lObjectID _
                            & " and SYSCOLPERM.grantee = " & lGrantee _
                            & " and SYSCOLPERM.grantor = " & lGrantor _
                            & " and SYSCOLPERM.privilege_type = 1 " _
                            & "order by SYSCOLUMN.column_id"
                        If bDebug Then DebugWrite sSQL
                        Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                        If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                            ' reset our 2 variables
                            sSelect = ""
                            sColumn = "Grant Select ("
                            Do While Not rsSubSQL.EOF
                                ' now to get the columns
                                sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                                    & sDQUOTE & ","
                                ' now to find out if it is grantable
                                sSelect = rsSubSQL.rdoColumns(1).Value
                                If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                                    DoEvents
                                    If bCancel Then Exit Do
                                End If
                                rsSubSQL.MoveNext
                            Loop
                            rsSubSQL.Close
                            If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                            sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                                & " to " & sGrantee
                            If sSelect = "Y" Then sColumn = sColumn & " with grant option"
                            sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                        End If
                    End If
                    
                    ' insert
                    If sInsert <> "N" Then
                        sAuth = sAuth & "Grant Insert on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sInsert = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    End If
                    
                    ' update
                    If sUpdate <> "N" Then
                        sAuth = sAuth & "Grant Update on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sUpdate = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    End If
                                        
                    ' update columns
                    If sUpdateCols <> "N" Then
                        ' they might have only permissions on columns
                        sSQL = "Select column_name, is_grantable " _
                            & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                            & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                            & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                            & "and SYSCOLPERM.table_id = " & lObjectID _
                            & " and SYSCOLPERM.grantee = " & lGrantee _
                            & " and SYSCOLPERM.grantor = " & lGrantor _
                            & " and SYSCOLPERM.privilege_type = 8 " _
                            & "order by SYSCOLUMN.column_id"
                        If bDebug Then DebugWrite sSQL
                        Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                        If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                            ' reset our 2 variables
                            sUpdateCols = ""
                            sColumn = "Grant Update ("
                            Do While Not rsSubSQL.EOF
                                ' now to get the columns
                                sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                                    & sDQUOTE & ","
                                ' now to find out if it is grantable
                                sUpdateCols = rsSubSQL.rdoColumns(1).Value
                                If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                                    DoEvents
                                    If bCancel Then Exit Do
                                End If
                                rsSubSQL.MoveNext
                            Loop
                            rsSubSQL.Close
                            If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                            sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                                & " to " & sGrantee
                            If sUpdateCols = "Y" Then sColumn = sColumn & " with grant option"
                            sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                        End If
                    End If
                    
                    ' delete
                    If sDelete <> "N" Then
                        sAuth = sAuth & "Grant Delete on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sDelete = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    End If
                    
                    'alter
                    If sAlter <> "N" Then
                        sAuth = sAuth & "Grant Alter on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sUpdate = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    End If
                    
                    ' reference
                    If sReference <> "N" Then
                        sAuth = sAuth & "Grant References on " & sOwner & "." & sTableName _
                            & " to " & sGrantee
                        If sReference = "G" Then sAuth = sAuth & " with grant option"
                        sAuth = sAuth & vbCrLf & "go" & vbCrLf
                    Else
                        ' they might have only permissions on columns
                        sSQL = "Select column_name, is_grantable " _
                            & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                            & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                            & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                            & "and SYSCOLPERM.table_id = " & lObjectID _
                            & " and SYSCOLPERM.grantee = " & lGrantee _
                            & " and SYSCOLPERM.grantor = " & lGrantor _
                            & " and SYSCOLPERM.privilege_type = 16 " _
                            & "order by SYSCOLUMN.column_id"
                        If bDebug Then DebugWrite sSQL
                        Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                        If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                            ' reset our 2 variables
                            sReference = ""
                            sColumn = "Grant References ("
                            Do While Not rsSubSQL.EOF
                                ' now to get the columns
                                sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                                    & sDQUOTE & ","
                                ' now to find out if it is grantable
                                sSelect = rsSubSQL.rdoColumns(1).Value
                                If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                                    DoEvents
                                    If bCancel Then Exit Do
                                End If
                                rsSubSQL.MoveNext
                            Loop
                            rsSubSQL.Close
                            If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                            sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                                & " to " & sGrantee
                            If sReference = "Y" Then sColumn = sColumn & " with grant option"
                            sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                        End If
                    End If
                    
                    If rsSQL.AbsolutePosition Mod 50 = 0 Then
                        DoEvents
                        If bCancel Then Exit Do
                    End If
                    rsSQL.MoveNext
                Loop
            End If
            rsSQL.Close
            sText = sText & sAuth
        End If
        Clipboard.Clear
        Clipboard.SetText sText
    End If
End If

If bDebug Then DebugWrite "TreeGetViewDDL Ends"
TreeGetViewDDL = True

End Function

Public Function TreeGetTableDDL(sMyKey As String) As Boolean
Dim sSQL As String, sTable As String, sText As String, sValue As String
Dim sPrimaryKey As String, sForeignKey As String, sUnique As String
Dim sIndex As String, sIndexName As String  ' added to find indexes
Dim sCols As String, sForeignCols As String
Dim sPrimaryCols As String, sRefAction As String
Dim bForeignKey As Boolean
Dim bIndex As Boolean  ' added to find indexes
Dim lCurrentPos As Long
Dim lObjectID As Long
Dim iCount As Integer
Dim vTypes(13, 1) As Variant

' for permissions
Dim sAuth As String
Dim sSelect As String, sInsert As String
Dim sDelete As String, sUpdate As String
Dim sUpdateCols As String, sAlter As String
Dim sGrantee As String, sGrantor As String
Dim sOwner As String, sTableName As String
Dim sReference As String, sColumn As String
Dim lGrantee As Long, lGrantor As Long
Dim bAuth As Boolean
Dim rsSubSQL As rdoResultset


TreeGetTableDDL = False
If bDebug Then DebugWrite "TreeGetTableDDL Starts"

lObjectID = CLng(Mid$(sMyKey, 3))
sPrimaryKey = ""
sForeignKey = ""
bForeignKey = False
bIndex = False
bAuth = False
sText = ""
sAuth = ""
sColumn = ""

' set up the array values
' type, scale
vTypes(0, 0) = "char"
vTypes(0, 1) = 0
vTypes(1, 0) = "nchar"
vTypes(1, 1) = 0
vTypes(2, 0) = "varchar"
vTypes(2, 1) = 0
vTypes(3, 0) = "nvarchar"
vTypes(3, 1) = 0
vTypes(4, 0) = "numeric"
vTypes(4, 1) = 1
vTypes(5, 0) = "num"
vTypes(5, 1) = 1
vTypes(6, 0) = "numericn"
vTypes(6, 1) = 1
vTypes(7, 0) = "decimal"
vTypes(7, 1) = 1
vTypes(8, 0) = "dec"
vTypes(8, 1) = 1
vTypes(9, 0) = "decimaln"
vTypes(9, 1) = 1
vTypes(10, 0) = "character"
vTypes(10, 1) = 0
vTypes(11, 0) = "binary"
vTypes(11, 1) = 1
vTypes(12, 0) = "varbinary"
vTypes(12, 1) = 1
    
' the sql statement
sSQL = ""
sSQL = "Select column_name, column_id, isnull((select type_name from" _
    & " SYS.SYSUSERTYPE where type_id = SYSCOLUMN.user_type)," _
    & "(select domain_name from SYS.SYSDOMAIN where domain_id = SYSCOLUMN.domain_id ))" _
    & sDQUOTE & "Type"
sSQL = sSQL _
    & sDQUOTE & ", width, scale, nulls, user_type" _
    & ", pkey"
sSQL = sSQL _
    & ", If exists (select table_id from SYS.SYSFKCOL" _
    & " where foreign_table_id = SYSCOLUMN.table_id" _
    & " and foreign_column_id = SYSCOLUMN.column_id)" _
    & " then 'Y' else 'N' endif " & sDQUOTE & "Foreign" _
    & sDQUOTE
sSQL = sSQL _
    & ", " & sDQUOTE & "default" & sDQUOTE _
    & ", " & sDQUOTE & "check" & sDQUOTE
sSQL = sSQL _
    & ", If exists (select table_id from SYS.SYSINDEX" _
    & " where table_id = SYSCOLUMN.table_id" _
    & " and " & sDQUOTE & "unique" & sDQUOTE _
    & " <> 'U' )" _
    & " then 'Y' else 'N' endif " & sDQUOTE & "Indx" _
    & sDQUOTE
sSQL = sSQL & " ,(If exists(select * from SYS.SYSTABLEPERM " _
    & "where table_id = " & lObjectID & ") then 'Y' " _
    & "else 'N' endif) Auth "
sSQL = sSQL _
    & "From SYS.SYSCOLUMN " _
    & "Where table_id = " & lObjectID _
    & " Order by column_id"

' set up the first line
sTable = Trim(Mid$(frmMain.tvDB.Nodes(sMyKey).Text, InStr(frmMain.tvDB.Nodes(sMyKey).Text, "  ")))
sTable = sDQUOTE & Mid$(sTable, InStr(sMyKey, "_"), Len(sTable) - InStr(sMyKey, "_")) & sDQUOTE & "." & sDQUOTE
sTable = sTable & Trim(Left$(frmMain.tvDB.Nodes(sMyKey).Text, InStr(frmMain.tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
sText = "Create Table " & sTable & "("

If bDebug Then DebugWrite sSQL
' open the recordset
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        ' set the values
        sText = sText & vbCrLf & "    " & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE
        sText = sText & " " & rsSQL.rdoColumns(2).Value
        ' do we need to append to the text
        ' if the value >= 100 then it is a user defined type
        If IsNull(rsSQL.rdoColumns(6).Value) Then
            ' go through the array
            For iCount = 0 To UBound(vTypes)
                ' does this type require the length
                If vTypes(iCount, 0) = rsSQL.rdoColumns(2).Value And rsSQL.rdoColumns(3).Value > 0 Then  ' yes
                    ' does it require the precision
                    If vTypes(iCount, 1) = 1 Then   ' yes
                        sText = sText & "(" & rsSQL.rdoColumns(3).Value _
                            & "," & rsSQL.rdoColumns(4).Value & ")"
                    Else    ' no
                        sText = sText & "(" & rsSQL.rdoColumns(3).Value & ")"
                    End If
                    Exit For
                End If
            Next
        End If
           
        ' add null/not null
         If rsSQL.rdoColumns(5).Value = "Y" Then
            sText = sText & " Null"
         Else
             sText = sText & " Not Null"
         End If
    
         ' add default
         On Error Resume Next
         sValue = ""
         sValue = rsSQL.rdoColumns(9).Value
         On Error GoTo 0
         If sValue <> "" Then
             sText = sText & " default " & sValue
         End If
         ' add check
         On Error Resume Next
         sValue = ""
         sValue = rsSQL.rdoColumns(10).Value
         On Error GoTo 0
         If sValue <> "" Then
             sText = sText & " " & sValue
         End If
            
        ' add comma
        sText = sText & ","
        
        ' is this in the primary key
        If rsSQL.rdoColumns(7).Value = "Y" Then
            sPrimaryKey = sPrimaryKey & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE & ","
        End If
        
        ' are there foreign keys
        If rsSQL.rdoColumns(8).Value = "Y" Then bForeignKey = True
        
        ' are there indexes
        If rsSQL.rdoColumns(11).Value = "Y" Then bIndex = True
        
        ' are there grants
        If rsSQL.rdoColumns(12).Value = "Y" Then bAuth = True
        
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

' do we need to add primary key constraint
If sPrimaryKey <> "" Then
    If Right$(sPrimaryKey, 1) = "," Then sPrimaryKey = Left$(sPrimaryKey, Len(sPrimaryKey) - 1)
    sText = sText & vbCrLf & "    Primary Key (" & sPrimaryKey & "),"
End If

' do we need to add unique indexes
sSQL = ""
sSQL = "Select index_name, index_id " _
    & " From SYS.SYSINDEX " _
    & "Where table_id = " & lObjectID _
    & " and " & sDQUOTE & "unique" & sDQUOTE & " = 'U'"
    
sSQL = sSQL & " Order by index_id"

If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sUnique = ""
        sUnique = "    Unique ("
        sCols = rsSQL.rdoColumns(0).Value
        sCols = Mid$(sCols, (InStr(sCols, "(") + 1), (InStr(sCols, ")") - (InStr(sCols, "(") + 1)))
        lCurrentPos = 0
            Do
                lCurrentPos = InStr(sCols, ",")
                If lCurrentPos = 0 Then
                    Exit Do
                End If
                ' make sure that this is not the last character
                If lCurrentPos <= Len(sCols) - 1 Then
                    sUnique = sUnique & sDQUOTE & _
                        Trim(Left$(sCols, lCurrentPos - 1)) & sDQUOTE & ","
                        sCols = Mid$(sCols, lCurrentPos + 1)
                End If
            Loop While lCurrentPos > 0
        sUnique = sUnique & sDQUOTE & sCols & sDQUOTE
        sText = sText & vbCrLf & sUnique & "),"
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

' do we need to add foreign key constraints
If bForeignKey Then
    sSQL = ""
    sSQL = "SELECT role, ( SELECT user_name " _
           & "FROM SYS.SYSUSERPERMS == SYS.SYSTABLE " _
           & "WHERE table_id = primary_table_id ), " _
           & "( SELECT table_name FROM SYS.SYSTABLE " _
           & "WHERE table_id = primary_table_id ), " _
           & "( SELECT list( string( FK.column_name, " _
           & "' IS ', PK.column_name ) ) " _
           & "FROM SYS.SYSFKCOL KEY JOIN " _
           & "SYS.SYSCOLUMN FK, SYS.SYSCOLUMN PK "
    sSQL = sSQL _
           & "WHERE foreign_table_id = SYSFOREIGNKEY.foreign_table_id " _
           & "AND foreign_key_id = SYSFOREIGNKEY.foreign_key_id " _
           & "AND PK.table_id = SYSFOREIGNKEY.primary_table_id " _
           & "AND PK.column_id = SYSFKCOL.primary_column_id ), " _
           & "nulls, check_on_commit, (SELECT referential_action " _
           & "FROM SYS.SYSTRIGGER WHERE foreign_table_id = " _
           & "SYSFOREIGNKEY.foreign_table_id and foreign_key_id = " _
           & "SYSFOREIGNKEY.foreign_key_id and event = 'D'), "
    sSQL = sSQL _
           & "(SELECT referential_action " _
           & "FROM SYS.SYSTRIGGER WHERE foreign_table_id = " _
           & "SYSFOREIGNKEY.foreign_table_id and foreign_key_id = " _
           & "SYSFOREIGNKEY.foreign_key_id and event = 'C') "
    sSQL = sSQL _
           & "From SYS.SYSFOREIGNKEY " _
           & "Where foreign_table_id = " & lObjectID

    If bDebug Then DebugWrite sSQL
    ' open the recordset
    Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        Do While Not rsSQL.EOF
            ' set the values
            sForeignKey = ""
            sCols = ""
            sPrimaryCols = ""
            sForeignCols = ""
            sCols = rsSQL.rdoColumns(3).Value
            ' check null
            If rsSQL.rdoColumns(4).Value = "N" Then
                sForeignKey = "Not Null"
            End If
            sForeignKey = "    Foreign Key "
            sForeignKey = sForeignKey & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE
            sForeignKey = sForeignKey & " ("
            
            lCurrentPos = 0
            Do
                lCurrentPos = InStr(sCols, " IS ")
                If lCurrentPos = 0 Then
                    Exit Do
                End If
                ' make sure that this is not the last character
                If lCurrentPos <= Len(sCols) - 1 Then
                    sForeignCols = sForeignCols & sDQUOTE & _
                        Trim(Left$(sCols, lCurrentPos - 1)) & sDQUOTE & ","
                        sCols = Mid$(sCols, lCurrentPos + 4)
                        If InStr(sCols, ",") > 0 Then
                            sPrimaryCols = sPrimaryCols & sDQUOTE & _
                                Trim(Left$(sCols, InStr(sCols, ",") - 1)) & sDQUOTE & ","
                            sCols = Trim(Mid$(sCols, InStr(sCols, ",") + 1))
                        Else
                            sPrimaryCols = sPrimaryCols & sDQUOTE & _
                                Trim(sCols) & sDQUOTE & ","
                        End If
                End If
            Loop While lCurrentPos > 0
            
            ' get rid of the trailing comma
            If Right$(sForeignCols, 1) = "," Then sForeignCols = Left$(sForeignCols, Len(sForeignCols) - 1)
            If Right$(sPrimaryCols, 1) = "," Then sPrimaryCols = Left$(sPrimaryCols, Len(sPrimaryCols) - 1)
            
            sForeignKey = sForeignKey & sForeignCols & ")" & vbCrLf _
                & "        References " & rsSQL.rdoColumns(1).Value _
                & "." & sDQUOTE & rsSQL.rdoColumns(2).Value & sDQUOTE _
                & " (" & sPrimaryCols & ")"
                
            ' get delete criteria
            sRefAction = ""
            On Error Resume Next
            sRefAction = rsSQL.rdoColumns(6).Value
            On Error GoTo 0
            If Not IsNull(sRefAction) Then
                Select Case sRefAction
                    Case "C"
                        sForeignKey = sForeignKey & " On Delete Cascade"
                    Case "N"
                        sForeignKey = sForeignKey & " On Delete Set Null"
                    Case "R"
                        sForeignKey = sForeignKey & " On Delete Restrict"
                    Case "D"
                        sForeignKey = sForeignKey & " On Delete Set Default"
                End Select
            End If
            
            ' get update criteria
            sRefAction = ""
            On Error Resume Next
            sRefAction = rsSQL.rdoColumns(7).Value
            On Error GoTo 0
            If Not IsNull(sRefAction) Then
                Select Case sRefAction
                    Case "C"
                        sForeignKey = sForeignKey & " On Update Cascade"
                    Case "N"
                        sForeignKey = sForeignKey & " On Update Set Null"
                    Case "R"
                        sForeignKey = sForeignKey & " On Update Restrict"
                    Case "D"
                        sForeignKey = sForeignKey & " On Update Set Default"
                End Select
            End If
            
            ' check on commit
            If rsSQL.rdoColumns(5).Value = "Y" Then
                sForeignKey = sForeignKey & " Check On Commit"
            End If
            
            ' add foreign keys
            sText = sText & vbCrLf & sForeignKey & ","
            rsSQL.MoveNext
        Loop
    End If
    rsSQL.Close
End If

If Right$(sText, 1) = "," Then sText = Left$(sText, Len(sText) - 1)
sText = sText & vbCrLf & ")" & vbCrLf & "go" & vbCrLf

' now we will add priveleges
If bAuth = True Then
    sSQL = "Select A.user_name " & sDQUOTE & "SGRANTEE" & sDQUOTE _
        & ", B.user_name " & sDQUOTE & "OWNER" & sDQUOTE _
        & ", table_name " _
        & ", C.user_name " & sDQUOTE & "SGRANTOR" & sDQUOTE _
        & ", SYSTABLEPERM.selectauth, SYSTABLEPERM.insertauth" _
        & ", SYSTABLEPERM.deleteauth, SYSTABLEPERM.updateauth" _
        & ", SYSTABLEPERM.updatecols, SYSTABLEPERM.alterauth" _
        & ", SYSTABLEPERM.referenceauth " _
        & ", SYSTABLEPERM.grantee " _
        & ", SYSTABLEPERM.grantor " _
        & "FROM SYS.SYSUSERPERMS A, SYS.SYSTABLEPERM, " _
        & "SYS.SYSTABLE, SYS.SYSUSERPERMS B, " _
        & "SYS.SYSUSERPERMS C " _
        & "Where SYSTABLEPERM.grantee = A.user_id " _
        & "and SYSTABLE.creator = B.user_id " _
        & "and SYSTABLE.table_id = SYSTABLEPERM.stable_id " _
        & "and SYSTABLEPERM.grantor = C.user_id " _
        & "and SYSTABLEPERM.stable_id = " & lObjectID

    If bDebug Then DebugWrite sSQL
    Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        Do While Not rsSQL.EOF
            ' lets get our values
            sGrantee = sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE
            sOwner = sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE
            sTableName = sDQUOTE & rsSQL.rdoColumns(2).Value & sDQUOTE
            sGrantor = sDQUOTE & rsSQL.rdoColumns(3).Value & sDQUOTE
            sSelect = rsSQL.rdoColumns(4).Value
            sInsert = rsSQL.rdoColumns(5).Value
            sDelete = rsSQL.rdoColumns(6).Value
            sUpdate = rsSQL.rdoColumns(7).Value
            sUpdateCols = rsSQL.rdoColumns(8).Value
            sAlter = rsSQL.rdoColumns(9).Value
            sReference = rsSQL.rdoColumns(10).Value
            lGrantee = rsSQL.rdoColumns(11).Value
            lGrantor = rsSQL.rdoColumns(12).Value
            
            ' lets show the grantor for this set
            sAuth = sAuth & "/* Grantor=" & sGrantor & " */" & vbCrLf
            
            ' select permissions
            If sSelect <> "N" Then
                sAuth = sAuth & "Grant Select on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sSelect = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            Else
                ' they might have only permissions on columns
                sSQL = "Select column_name, is_grantable " _
                    & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                    & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                    & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                    & "and SYSCOLPERM.table_id = " & lObjectID _
                    & " and SYSCOLPERM.grantee = " & lGrantee _
                    & " and SYSCOLPERM.grantor = " & lGrantor _
                    & " and SYSCOLPERM.privilege_type = 1 " _
                    & "order by SYSCOLUMN.column_id"
                If bDebug Then DebugWrite sSQL
                Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                    ' reset our 2 variables
                    sSelect = ""
                    sColumn = "Grant Select ("
                    Do While Not rsSubSQL.EOF
                        ' now to get the columns
                        sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                            & sDQUOTE & ","
                        ' now to find out if it is grantable
                        sSelect = rsSubSQL.rdoColumns(1).Value
                        If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                            DoEvents
                            If bCancel Then Exit Do
                        End If
                        rsSubSQL.MoveNext
                    Loop
                    rsSubSQL.Close
                    If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                    sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                        & " to " & sGrantee
                    If sSelect = "Y" Then sColumn = sColumn & " with grant option"
                    sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                End If
            End If
            
            ' insert
            If sInsert <> "N" Then
                sAuth = sAuth & "Grant Insert on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sInsert = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            End If
            
            ' update
            If sUpdate <> "N" Then
                sAuth = sAuth & "Grant Update on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sUpdate = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            End If
                                
            ' update columns
            If sUpdateCols <> "N" Then
                ' they might have only permissions on columns
                sSQL = "Select column_name, is_grantable " _
                    & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                    & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                    & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                    & "and SYSCOLPERM.table_id = " & lObjectID _
                    & " and SYSCOLPERM.grantee = " & lGrantee _
                    & " and SYSCOLPERM.grantor = " & lGrantor _
                    & " and SYSCOLPERM.privilege_type = 8 " _
                    & "order by SYSCOLUMN.column_id"
                If bDebug Then DebugWrite sSQL
                Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                    ' reset our 2 variables
                    sUpdateCols = ""
                    sColumn = "Grant Update ("
                    Do While Not rsSubSQL.EOF
                        ' now to get the columns
                        sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                            & sDQUOTE & ","
                        ' now to find out if it is grantable
                        sUpdateCols = rsSubSQL.rdoColumns(1).Value
                        If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                            DoEvents
                            If bCancel Then Exit Do
                        End If
                        rsSubSQL.MoveNext
                    Loop
                    rsSubSQL.Close
                    If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                    sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                        & " to " & sGrantee
                    If sUpdateCols = "Y" Then sColumn = sColumn & " with grant option"
                    sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                End If
            End If
            
            ' delete
            If sDelete <> "N" Then
                sAuth = sAuth & "Grant Delete on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sDelete = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            End If
            
            'alter
            If sAlter <> "N" Then
                sAuth = sAuth & "Grant Alter on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sUpdate = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            End If
            
            ' reference
            If sReference <> "N" Then
                sAuth = sAuth & "Grant References on " & sOwner & "." & sTableName _
                    & " to " & sGrantee
                If sReference = "G" Then sAuth = sAuth & " with grant option"
                sAuth = sAuth & vbCrLf & "go" & vbCrLf
            Else
                ' they might have only permissions on columns
                sSQL = "Select column_name, is_grantable " _
                    & "From SYS.SYSCOLPERM, SYS.SYSCOLUMN " _
                    & "Where SYSCOLPERM.table_id = SYSCOLUMN.table_id " _
                    & "and SYSCOLPERM.column_id = SYSCOLUMN.column_id " _
                    & "and SYSCOLPERM.table_id = " & lObjectID _
                    & " and SYSCOLPERM.grantee = " & lGrantee _
                    & " and SYSCOLPERM.grantor = " & lGrantor _
                    & " and SYSCOLPERM.privilege_type = 16 " _
                    & "order by SYSCOLUMN.column_id"
                If bDebug Then DebugWrite sSQL
                Set rsSubSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
                If Not rsSubSQL.BOF And Not rsSubSQL.EOF Then
                    ' reset our 2 variables
                    sReference = ""
                    sColumn = "Grant References ("
                    Do While Not rsSubSQL.EOF
                        ' now to get the columns
                        sColumn = sColumn & sDQUOTE & rsSubSQL.rdoColumns(0).Value _
                            & sDQUOTE & ","
                        ' now to find out if it is grantable
                        sSelect = rsSubSQL.rdoColumns(1).Value
                        If rsSubSQL.AbsolutePosition Mod 50 = 0 Then
                            DoEvents
                            If bCancel Then Exit Do
                        End If
                        rsSubSQL.MoveNext
                    Loop
                    rsSubSQL.Close
                    If Right$(sColumn, 1) = "," Then sColumn = Left$(sColumn, Len(sColumn) - 1)
                    sColumn = sColumn & ") on " & sOwner & "." & sTableName _
                        & " to " & sGrantee
                    If sReference = "Y" Then sColumn = sColumn & " with grant option"
                    sAuth = sAuth & sColumn & vbCrLf & "go" & vbCrLf
                End If
            End If
            
            If rsSQL.AbsolutePosition Mod 50 = 0 Then
                DoEvents
                If bCancel Then Exit Do
            End If
            rsSQL.MoveNext
        Loop
    End If
    rsSQL.Close
    sText = sText & sAuth
End If

' now we will add any indexes
If bIndex = True Then
    sSQL = "Select index_name, column_name, sequence, " _
        & sDQUOTE & "order" & sDQUOTE _
        & ", table_name, user_name, " _
        & sDQUOTE & "unique" & sDQUOTE _
        & " From SYS.SYSIXCOL, SYS.SYSCOLUMN, SYS.SYSTABLE, " _
        & "SYS.SYSUSERPERMS, SYS.SYSINDEX " _
        & "Where SYSIXCOL.table_id = SYSCOLUMN.table_id " _
        & "and SYSIXCOL.column_id = SYSCOLUMN.column_id " _
        & "and SYSIXCOL.table_id = SYSTABLE.table_id " _
        & "and SYSCOLUMN.table_id = SYSTABLE.table_id " _
        & "and SYSTABLE.creator = SYSUSERPERMS.user_id " _
        & "and SYSINDEX.index_id = SYSIXCOL.index_id " _
        & "and SYSINDEX.table_id = SYSIXCOL.table_id " _
        & "and " & sDQUOTE & "unique" & sDQUOTE & " <> 'U' " _
        & "and SYSIXCOL.table_id = " & lObjectID
    sSQL = sSQL & " Order by SYSIXCOL.index_id, sequence"
    If bDebug Then DebugWrite sSQL
    
    Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        sIndex = ""
        sIndexName = ""
        Do While Not rsSQL.EOF
            If sIndexName <> rsSQL.rdoColumns(0).Value Then
                If sIndex <> "" Then
                    If Right$(sIndex, 1) = "," Then sIndex = Left$(sIndex, Len(sIndex) - 1)
                    sIndex = sIndex & ")" & vbCrLf & "go" & vbCrLf
                End If
                ' we are not doing constraints here
                sIndexName = rsSQL.rdoColumns(0).Value
                sIndex = sIndex & "Create "
                If rsSQL.rdoColumns(6) = "Y" Then sIndex = sIndex & "Unique "
                sIndex = sIndex & "Index " & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE _
                    & vbCrLf & "on " & sDQUOTE & rsSQL.rdoColumns(5) & sDQUOTE & "." & sDQUOTE _
                    & rsSQL.rdoColumns(4) & sDQUOTE & "("
            End If
            ' get the column
            sIndex = sIndex & sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE
            ' add order
            If rsSQL.rdoColumns(3).Value = "D" Then
               sIndex = sIndex & "  Desc"
            End If
            sIndex = sIndex & ","
            If rsSQL.AbsolutePosition Mod 50 = 0 Then
                DoEvents
                If bCancel Then Exit Do
            End If
            rsSQL.MoveNext
        Loop
    End If
    rsSQL.Close
    
    If Right$(sIndex, 1) = "," Then sIndex = Left$(sIndex, Len(sIndex) - 1)
    sIndex = sIndex & ")" & vbCrLf & "go" & vbCrLf

    sText = sText & sIndex
    
End If

Clipboard.Clear
Clipboard.SetText sText

If bDebug Then DebugWrite "TreeGetTableDDL Ends"
TreeGetTableDDL = True
End Function

Public Function TreeGetProcDDL(sMyKey As String) As Boolean
Dim sSQL As String, sText As String, sAuth As String
Dim bAuth As Boolean
Dim lObjectID As Long

TreeGetProcDDL = False
If bDebug Then DebugWrite "TreeGetProcDDL Starts"

lObjectID = CLng(Mid$(sMyKey, 3))
bAuth = False
sText = ""
sAuth = ""

sSQL = "Select proc_id, proc_defn, " _
    & "(If exists(select * from SYS.SYSPROCPERM " _
    & "where proc_id = " & lObjectID & ") then 'Y' " _
    & "else 'N' endif) as Auth " _
    & "From SYS.SYSPROCEDURE " _
    & "Where proc_id = " & lObjectID

If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    If rsSQL.rdoColumns(1).ChunkRequired = True Then
        sText = GetTextChunks(rsSQL.rdoColumns(1))
    Else
        sText = rsSQL.rdoColumns(1).Value
    End If
    If rsSQL.rdoColumns(2).Value = "Y" Then bAuth = True
End If
rsSQL.Close
    
If sText <> "" Then
    If Not IsNull(sText) Then
        If UCase(Left(sText, 16)) = "CREATE PROCEDURE" Then
            sText = "Alter Procedure" & Mid(sText, 17)
        End If
        If UCase(Left(sText, 15)) = "CREATE FUNCTION" Then
            sText = "Alter Function" & Mid(sText, 16)
        End If
        sText = sText & vbCrLf & "go" & vbCrLf
        ' are there any rights
        If bAuth = True Then
            sSQL = "Select A.user_name " & sDQUOTE & "GRANTEE" & sDQUOTE _
                & ", B.user_name " & sDQUOTE & "OWNER" & sDQUOTE _
                & ", proc_name " _
                & "FROM SYS.SYSUSERPERMS A, SYS.SYSPROCPERM, " _
                & "SYS.SYSPROCEDURE, SYS.SYSUSERPERMS B " _
                & "Where SYSPROCPERM.grantee = A.user_id " _
                & "and SYSPROCEDURE.creator = B.user_id " _
                & "and SYSPROCEDURE.proc_id = SYSPROCPERM.proc_id " _
                & "and SYSPROCPERM.proc_id = " & lObjectID
            
            If bDebug Then DebugWrite sSQL
            Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
            If Not rsSQL.BOF And Not rsSQL.EOF Then
                Do While Not rsSQL.EOF
                    sAuth = sAuth & "/* Grantor=" & sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE & " */" & vbCrLf
                    sAuth = sAuth & "Grant Execute on "
                    sAuth = sAuth & sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE _
                        & "." & sDQUOTE & rsSQL.rdoColumns(2).Value & sDQUOTE _
                        & " to " & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE _
                        & vbCrLf & "go" & vbCrLf
                    If rsSQL.AbsolutePosition Mod 50 = 0 Then
                        DoEvents
                        If bCancel Then Exit Do
                    End If
                    rsSQL.MoveNext
                Loop
            End If
            rsSQL.Close
            sText = sText & sAuth
        End If
        
        Clipboard.Clear
        Clipboard.SetText sText
    End If
End If


If bDebug Then DebugWrite "TreeGetProcDDL Ends"
TreeGetProcDDL = True

End Function


Public Function TreeGetTriggerDDL(sMyKey As String) As Boolean
Dim sSQL As String, sText As String
Dim lObjectID As Long

TreeGetTriggerDDL = False
If bDebug Then DebugWrite "TreeGetTriggerDDL Starts"

lObjectID = CLng(Mid$(sMyKey, InStr(4, sMyKey, "_") + 1))

sSQL = "Select trigger_id, trigger_defn " _
    & "From SYS.SYSTRIGGER " _
    & "Where trigger_id = " & lObjectID

If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    If rsSQL.rdoColumns(1).ChunkRequired = True Then
        sText = GetTextChunks(rsSQL.rdoColumns(1))
    Else
        sText = rsSQL.rdoColumns(1).Value
    End If
    If sText <> "" Then
        If Not IsNull(sText) Then
            If UCase(Left(sText, 14)) = "CREATE TRIGGER" Then
                sText = "Alter Trigger" & Mid(sText, 15)
            End If
            sText = sText & vbCrLf & "go" & vbCrLf
            Clipboard.Clear
            Clipboard.SetText sText
        End If
    End If
End If
rsSQL.Close
If bDebug Then DebugWrite "TreeGetTriggerDDL Ends"
TreeGetTriggerDDL = True
End Function
Public Function TreeGetColumns(sParentKey As String) As Boolean
Dim sSQL As String, sKey As String, sText As String, sValue As String
Dim lObjectID As Long
Dim iCount As Integer
Dim vTypes(13, 1) As Variant

TreeGetColumns = False
If bDebug Then DebugWrite "TreeGetColumns Starts"
        
lObjectID = CLng(Mid$(sParentKey, 4))

' set up the array values
' type, scale
vTypes(0, 0) = "char"
vTypes(0, 1) = 0
vTypes(1, 0) = "nchar"
vTypes(1, 1) = 0
vTypes(2, 0) = "varchar"
vTypes(2, 1) = 0
vTypes(3, 0) = "nvarchar"
vTypes(3, 1) = 0
vTypes(4, 0) = "numeric"
vTypes(4, 1) = 1
vTypes(5, 0) = "num"
vTypes(5, 1) = 1
vTypes(6, 0) = "numericn"
vTypes(6, 1) = 1
vTypes(7, 0) = "decimal"
vTypes(7, 1) = 1
vTypes(8, 0) = "dec"
vTypes(8, 1) = 1
vTypes(9, 0) = "decimaln"
vTypes(9, 1) = 1
vTypes(10, 0) = "character"
vTypes(10, 1) = 0
vTypes(11, 0) = "binary"
vTypes(11, 1) = 1
vTypes(12, 0) = "varbinary"
vTypes(12, 1) = 1
    
' the sql statement
sSQL = "Select column_name, column_id, isnull((select type_name from" _
    & " SYS.SYSUSERTYPE where type_id = SYSCOLUMN.user_type)," _
    & "(select domain_name from SYS.SYSDOMAIN where domain_id = SYSCOLUMN.domain_id ))" _
    & sDQUOTE & "Type"
sSQL = sSQL _
    & sDQUOTE & ", width, scale, nulls, user_type" _
    & ", pkey"
sSQL = sSQL _
    & ", If exists (select table_id from SYS.SYSFKCOL" _
    & " where foreign_table_id = SYSCOLUMN.table_id" _
    & " and foreign_column_id = SYSCOLUMN.column_id)" _
    & " then 'Y' else 'N' endif " & sDQUOTE & "Foreign" _
    & sDQUOTE
sSQL = sSQL _
    & ", " & sDQUOTE & "default" & sDQUOTE _
    & ", " & sDQUOTE & "check" & sDQUOTE _
    & " From SYS.SYSCOLUMN " _
    & "Where table_id = " & lObjectID _
    & " Order by column_id"

If bDebug Then DebugWrite sSQL
' open the recordset
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        ' set the values
        sText = rsSQL.rdoColumns(0).Value
        sText = sText & "  " & rsSQL.rdoColumns(2).Value
        ' do we need to append to the text
        ' if the value >= 100 then it is a user defined type
        If IsNull(rsSQL.rdoColumns(6).Value) Then
            ' go through the array
            For iCount = 0 To UBound(vTypes)
                ' does this type require the length
                If vTypes(iCount, 0) = rsSQL.rdoColumns(2).Value And rsSQL.rdoColumns(3).Value > 0 Then  ' yes
                    ' does it require the precision
                    If vTypes(iCount, 1) = 1 Then   ' yes
                        sText = sText & "(" & rsSQL.rdoColumns(3).Value _
                            & "," & rsSQL.rdoColumns(4).Value & ")"
                    Else    ' no
                        sText = sText & "(" & rsSQL.rdoColumns(3).Value & ")"
                    End If
                    Exit For
                End If
            Next
        End If
        
        ' add null/not null
         If rsSQL.rdoColumns(5).Value = "Y" Then
            sText = sText & "  Null"
         Else
             sText = sText & "  Not Null"
         End If
    
         ' add default
         On Error Resume Next
         sValue = ""
         sValue = rsSQL.rdoColumns(9).Value
         On Error GoTo 0
         If sValue <> "" Then
             sText = sText & "  default " & sValue
         End If
         ' add check
         On Error Resume Next
         sValue = ""
         sValue = rsSQL.rdoColumns(10).Value
         On Error GoTo 0
         If sValue <> "" Then
             sText = sText & "  " & sValue
         End If
            
        sKey = "c_" & lObjectID & "_" & rsSQL.rdoColumns(1).Value
        If rsSQL.rdoColumns(7).Value = "Y" Then
            frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgPRIMARY
        ElseIf rsSQL.rdoColumns(8).Value = "Y" Then
            frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgFOREIGN
        Else
            frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgCOLUMN
        End If
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetColumns Ends"
TreeGetColumns = True
End Function
Public Function TreeGetIndexColumns(sParentKey As String) As Boolean
Dim sSQL As String, sKey As String, sText As String
Dim lObjectID As Long, lIndexID
Dim iCount As Integer

TreeGetIndexColumns = False
If bDebug Then DebugWrite "TreeGetIndexColumns Starts"
        
lObjectID = CLng(Mid$(sParentKey, 3, InStr(3, sParentKey, "_") - 3))
lIndexID = CLng(Mid$(sParentKey, InStr(3, sParentKey, "_") + 1))

sSQL = "Select column_name, sequence, " _
    & sDQUOTE & "order" & sDQUOTE _
    & " From SYS.SYSIXCOL, SYS.SYSCOLUMN " _
    & "Where SYSIXCOL.table_id = SYSCOLUMN.table_id " _
    & "and SYSIXCOL.column_id = SYSCOLUMN.column_id " _
    & "and SYSIXCOL.table_id = " & lObjectID _
    & " and SYSIXCOL.index_id = " & lIndexID
sSQL = sSQL & " Order by sequence"
If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sText = rsSQL.rdoColumns(0).Value
        sKey = "ic_" & lObjectID & "_" & lIndexID & "_" & rsSQL.rdoColumns(1).Value
        ' add order
        If rsSQL.rdoColumns(2).Value = "D" Then
           sText = sText & "  Desc"
        End If
        frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgCOLUMN
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetIndexes Ends"
TreeGetIndexColumns = True
End Function

Public Function TreeGetIndexes(sParentKey As String) As Boolean
Dim sSQL As String, sKey As String, sText As String
Dim lObjectID As Long
Dim iCount As Integer

TreeGetIndexes = False
If bDebug Then DebugWrite "TreeGetIndexes Starts"
        
lObjectID = CLng(Mid$(sParentKey, 4))
    
sSQL = "Select index_name, " & " index_id, " & sDQUOTE & "unique" & sDQUOTE _
    & " From SYS.SYSINDEX " _
    & "Where table_id = " & lObjectID

sSQL = sSQL & " Order by upper(index_name)"

If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        ' sText = rsSQL.rdoColumns(0).Value _
        '     & "  (" & rsSQL.rdoColumns(1).Value & ")"
        sText = rsSQL.rdoColumns(0).Value
        sKey = "i_" & lObjectID & "_" & rsSQL.rdoColumns(1).Value
        ' add unique
        If rsSQL.rdoColumns(2).Value = "Y" Then
           sText = sText & "  Unique"
        ElseIf rsSQL.rdoColumns(2).Value = "U" Then
           sText = sText & "  Constraint"
        End If
        frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgINDEX
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetIndexes Ends"
TreeGetIndexes = True

End Function
Public Function TreeGetIndexDDL(sMyKey As String) As Boolean
Dim sSQL As String, sIndex As String, sText As String
Dim lObjectID As Long, lIndexID
Dim iCount As Integer

TreeGetIndexDDL = False
If bDebug Then DebugWrite "TreeGetIndexDDL Starts"
        
lObjectID = CLng(Mid$(sMyKey, 3, InStr(3, sMyKey, "_") - 3))
lIndexID = CLng(Mid$(sMyKey, InStr(3, sMyKey, "_") + 1))

sSQL = "Select index_name, column_name, sequence, " _
    & sDQUOTE & "order" & sDQUOTE _
    & ", table_name, user_name, " _
    & sDQUOTE & "unique" & sDQUOTE _
    & " From SYS.SYSIXCOL, SYS.SYSCOLUMN, SYS.SYSTABLE, " _
    & "SYS.SYSUSERPERMS, SYS.SYSINDEX " _
    & "Where SYSIXCOL.table_id = SYSCOLUMN.table_id " _
    & "and SYSIXCOL.column_id = SYSCOLUMN.column_id " _
    & "and SYSIXCOL.table_id = SYSTABLE.table_id " _
    & "and SYSCOLUMN.table_id = SYSTABLE.table_id " _
    & "and SYSTABLE.creator = SYSUSERPERMS.user_id " _
    & "and SYSINDEX.index_id = SYSIXCOL.index_id " _
    & "and SYSINDEX.table_id = SYSIXCOL.table_id " _
    & "and SYSIXCOL.table_id = " & lObjectID _
    & " and SYSIXCOL.index_id = " & lIndexID
sSQL = sSQL & " Order by sequence"
If bDebug Then DebugWrite sSQL

Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    ' set up the first line
    ' we need to know if this is a unique constraint or a regular index
    
    If rsSQL.rdoColumns(6) = "U" Then
        sIndex = "Alter Table " _
            & rsSQL.rdoColumns(5) & "." & sDQUOTE _
            & rsSQL.rdoColumns(4) & sDQUOTE _
            & " Add Unique ("
    Else
        sIndex = "Create "
        If rsSQL.rdoColumns(6) = "Y" Then sIndex = sIndex & "Unique "
        sIndex = sIndex & "Index " & sDQUOTE & rsSQL.rdoColumns(0).Value & sDQUOTE _
            & vbCrLf & "on " & sDQUOTE & rsSQL.rdoColumns(5) & sDQUOTE & "." & sDQUOTE _
            & rsSQL.rdoColumns(4) & sDQUOTE & "("
    End If
    sText = ""
    Do While Not rsSQL.EOF
        sText = sText & sDQUOTE & rsSQL.rdoColumns(1).Value & sDQUOTE
        ' add order
        If rsSQL.rdoColumns(3).Value = "D" Then
           sText = sText & "  Desc"
        End If
        sText = sText & ","
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If Right$(sText, 1) = "," Then sText = Left$(sText, Len(sText) - 1)
sText = sIndex & sText & ")" & vbCrLf & "go" & vbCrLf

Clipboard.Clear
Clipboard.SetText sText

If bDebug Then DebugWrite "TreeGetIndexDDL Ends"
TreeGetIndexDDL = True
End Function

Public Function TreeGetParms(sParentKey As String) As Boolean
Dim sSQL As String, sKey As String, sText As String
Dim lObjectID As Long
Dim iCount As Integer
Dim vTypes(13, 1) As Variant
 
TreeGetParms = False
If bDebug Then DebugWrite "TreeGetParms Starts"
        
lObjectID = CLng(Mid$(sParentKey, 4))

' set up the array values
' type, scale
vTypes(0, 0) = "char"
vTypes(0, 1) = 0
vTypes(1, 0) = "nchar"
vTypes(1, 1) = 0
vTypes(2, 0) = "varchar"
vTypes(2, 1) = 0
vTypes(3, 0) = "nvarchar"
vTypes(3, 1) = 0
vTypes(4, 0) = "numeric"
vTypes(4, 1) = 1
vTypes(5, 0) = "num"
vTypes(5, 1) = 1
vTypes(6, 0) = "numericn"
vTypes(6, 1) = 1
vTypes(7, 0) = "decimal"
vTypes(7, 1) = 1
vTypes(8, 0) = "dec"
vTypes(8, 1) = 1
vTypes(9, 0) = "decimaln"
vTypes(9, 1) = 1
vTypes(10, 0) = "character"
vTypes(10, 1) = 0
vTypes(11, 0) = "binary"
vTypes(11, 1) = 1
vTypes(12, 0) = "varbinary"
vTypes(12, 1) = 1
        
' the sql statement
sSQL = "Select parm_name, parm_id" _
    & ", (select domain_name from SYS.SYSDOMAIN" _
    & " where domain_id = SYSPROCPARM.domain_id)" _
    & sDQUOTE & "Type" & sDQUOTE & ", width " _
    & ", scale, IF parm_mode_in = 'Y' AND" _
    & " parm_mode_out = 'N' THEN 'In'" _
    & " ELSE IF parm_mode_in = 'N'" _
    & " AND parm_mode_out = 'Y' THEN 'Out'" _
    & " Else 'InOut' ENDIF ENDIF " & sDQUOTE & "Direction" _
    & sDQUOTE
sSQL = sSQL _
    & " From SYS.SYSPROCPARM " _
    & "Where proc_id = " & lObjectID _
    & " and parm_type in (0,1) " _
    & "Order by parm_id"
If bDebug Then DebugWrite sSQL
' open the recordset
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        ' set the values
        sText = rsSQL.rdoColumns(0).Value
        sText = sText & "  " & rsSQL.rdoColumns(2).Value
        ' go through the array
        For iCount = 0 To UBound(vTypes)
            ' does this type require the length
            If vTypes(iCount, 0) = rsSQL.rdoColumns(2).Value And rsSQL.rdoColumns(3).Value > 0 Then  ' yes
                ' does it require the precision
                If vTypes(iCount, 1) = 1 Then   ' yes
                    sText = sText & "(" & rsSQL.rdoColumns(3).Value _
                        & "," & rsSQL.rdoColumns(4).Value & ")"
                Else    ' no
                    sText = sText & "(" & rsSQL.rdoColumns(3).Value & ")"
                End If
                Exit For
            End If
        Next
        ' add in/out
        sText = sText & "  " & rsSQL.rdoColumns(5).Value
        sKey = "a_" & lObjectID & "_" & rsSQL.rdoColumns(1).Value
        frmMain.tvDB.Nodes.Add sParentKey, tvwChild, sKey, sText, imgARG
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetParms Ends"
TreeGetParms = True
End Function


Public Function TreeGetProcs() As Boolean
Dim sSQL As String, sKey As String, sText As String

TreeGetProcs = False
If bDebug Then DebugWrite "TreeGetProcs Starts"
        
sSQL = "Select proc_name, (select user_name From SYS.SYSUSERPERMS " _
    & "where user_id = SYSPROCEDURE.creator) " & sDQUOTE & "Owner" _
    & sDQUOTE & ", proc_id From SYS.SYSPROCEDURE"
If GetSetting(App.Title, "Options", "ShowSystemObjects", "0") = "0" Then
    sSQL = sSQL _
        & " where " & sDQUOTE & "Owner" & sDQUOTE & " <> 'dbo'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'SYS'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'rs_systabgroup'"
End If
sSQL = sSQL & " Order by upper(proc_name), upper(" _
    & sDQUOTE & "Owner" & sDQUOTE & ")"
If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sText = rsSQL.rdoColumns(0).Value _
            & "  (" & rsSQL.rdoColumns(1).Value & ")"
        sKey = "p_" & rsSQL.rdoColumns(2).Value
        frmMain.tvDB.Nodes.Add "fProcs", tvwChild, sKey, sText, imgPROC
        frmMain.tvDB.Nodes.Add sKey, tvwChild, "fa_" & Mid$(sKey, 3), "Parameters", imgARG
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetProcs Ends"
TreeGetProcs = True
End Function

Public Function TreeGetTables() As Boolean
Dim sSQL As String, sKey As String, sText As String

TreeGetTables = False
If bDebug Then DebugWrite "TreeGetTables Starts"

sSQL = "Select table_name, (select user_name From SYS.SYSUSERPERMS " _
    & "where user_id = SYSTABLE.creator) " & sDQUOTE & "Owner" _
    & sDQUOTE & ", table_id From SYS.SYSTABLE " _
    & "Where table_type <> 'VIEW'"
If GetSetting(App.Title, "Options", "ShowSystemObjects", "0") = "0" Then
    sSQL = sSQL _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'dbo'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'SYS'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'rs_systabgroup'"
End If
sSQL = sSQL & " Order by upper(table_name), upper(" _
    & sDQUOTE & "Owner" & sDQUOTE & ")"
If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sText = rsSQL.rdoColumns(0).Value _
            & "  (" & rsSQL.rdoColumns(1).Value & ")"
        sKey = "t_" & rsSQL.rdoColumns(2).Value
        frmMain.tvDB.Nodes.Add "fTables", tvwChild, sKey, sText, imgTABLE
        frmMain.tvDB.Nodes.Add sKey, tvwChild, "fc_" & Mid$(sKey, 3), "Columns", imgCOLUMN
        frmMain.tvDB.Nodes.Add sKey, tvwChild, "fi_" & Mid$(sKey, 3), "Indexes", imgINDEX
        frmMain.tvDB.Nodes.Add sKey, tvwChild, "ft_" & Mid$(sKey, 3), "Triggers", imgTRIGGER
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetTables Ends"
TreeGetTables = True
End Function

Public Function TreeGetTypes() As Boolean
Dim sSQL As String, sKey As String, sText As String, sValue As String
Dim iCount As Integer
Dim vTypes(13, 1) As Variant

TreeGetTypes = False
If bDebug Then DebugWrite "TreeGetTypes Starts"
        
' set up the array values
' type, scale
vTypes(0, 0) = "char"
vTypes(0, 1) = 0
vTypes(1, 0) = "nchar"
vTypes(1, 1) = 0
vTypes(2, 0) = "varchar"
vTypes(2, 1) = 0
vTypes(3, 0) = "nvarchar"
vTypes(3, 1) = 0
vTypes(4, 0) = "numeric"
vTypes(4, 1) = 1
vTypes(5, 0) = "num"
vTypes(5, 1) = 1
vTypes(6, 0) = "numericn"
vTypes(6, 1) = 1
vTypes(7, 0) = "decimal"
vTypes(7, 1) = 1
vTypes(8, 0) = "dec"
vTypes(8, 1) = 1
vTypes(9, 0) = "decimaln"
vTypes(9, 1) = 1
vTypes(10, 0) = "character"
vTypes(10, 1) = 0
vTypes(11, 0) = "binary"
vTypes(11, 1) = 1
vTypes(12, 0) = "varbinary"
vTypes(12, 1) = 1
        
' the sql statement
sSQL = "Select type_name " _
    & ", type_id , (select domain_name from SYS.SYSDOMAIN " _
    & "where domain_id = SYSUSERTYPE.domain_id)"
sSQL = sSQL _
    & ", width, scale, nulls"
sSQL = sSQL _
    & ", " & sDQUOTE & "default" & sDQUOTE _
    & ", " & sDQUOTE & "check" & sDQUOTE _
    & "From SYS.SYSUSERTYPE "
If GetSetting(App.Title, "Options", "ShowSystemObjects", "0") = "0" Then
    sSQL = sSQL & " where not exists(select domain_id from SYS.SYSDOMAIN " _
        & "where domain_name like '%java.%' and domain_id = " _
        & "SYSUSERTYPE.domain_id) "
End If
sSQL = sSQL & "Order by upper(type_name)"
If bDebug Then DebugWrite sSQL
' open the recordset
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        ' set the values
        sText = rsSQL.rdoColumns(0).Value
        sText = sText & "  " & rsSQL.rdoColumns(2).Value
        ' go through the array
        For iCount = 0 To UBound(vTypes)
            ' does this type require the length
            If vTypes(iCount, 0) = rsSQL.rdoColumns(2).Value And rsSQL.rdoColumns(3).Value > 0 Then  ' yes
                ' does it require the precision
                If vTypes(iCount, 1) = 1 Then   ' yes
                    sText = sText & "(" & rsSQL.rdoColumns(3).Value _
                        & "," & rsSQL.rdoColumns(4).Value & ")"
                Else    ' no
                    sText = sText & "(" & rsSQL.rdoColumns(3).Value & ")"
                End If
                Exit For
            End If
        Next
        ' add null/not null
        If rsSQL.rdoColumns(5).Value = "N" Then
           sText = sText & "  Not Null"
        Else
            sText = sText & "  Null"
        End If
        
        ' add default
        On Error Resume Next
        sValue = ""
        sValue = rsSQL.rdoColumns(6).Value
        On Error GoTo 0
        If sValue <> "" Then
            sText = sText & "  default " & sValue
        End If
        ' add check
        On Error Resume Next
        sValue = ""
        sValue = rsSQL.rdoColumns(7).Value
        On Error GoTo 0
        If sValue <> "" Then
            sText = sText & "  " & sValue
        End If
         
        
        
        
        
        
        sKey = "u_" & rsSQL.rdoColumns(1).Value
        frmMain.tvDB.Nodes.Add "fTypes", tvwChild, sKey, sText, imgUDT
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetTypes Ends"
TreeGetTypes = True
End Function
Public Function TreeGetViews() As Boolean
Dim sSQL As String, sKey As String, sText As String

TreeGetViews = False
If bDebug Then DebugWrite "TreeGetViews Starts"
        
sSQL = "Select table_name, (select user_name From SYS.SYSUSERPERMS " _
    & "where user_id = SYSTABLE.creator) " & sDQUOTE & "Owner" _
    & sDQUOTE & ", table_id From SYS.SYSTABLE " _
    & "Where table_type = 'VIEW'"
If GetSetting(App.Title, "Options", "ShowSystemObjects", "0") = "0" Then
    sSQL = sSQL _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'dbo'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'SYS'" _
        & " and " & sDQUOTE & "Owner" & sDQUOTE & " <> 'rs_systabgroup'"
End If
sSQL = sSQL & " Order by upper(table_name), upper(" _
    & sDQUOTE & "Owner" & sDQUOTE & ")"
If bDebug Then DebugWrite sSQL
Set rsSQL = conSQL.OpenResultset(sSQL, rdOpenStatic, rdConcurReadOnly, rdExecDirect)
If Not rsSQL.BOF And Not rsSQL.EOF Then
    Do While Not rsSQL.EOF
        sText = rsSQL.rdoColumns(0).Value _
            & "  (" & rsSQL.rdoColumns(1).Value & ")"
        sKey = "v_" & rsSQL.rdoColumns(2).Value
        frmMain.tvDB.Nodes.Add "fViews", tvwChild, sKey, sText, imgVIEW
        frmMain.tvDB.Nodes.Add sKey, tvwChild, "fc_" & Mid$(sKey, 3), "Columns", imgCOLUMN
        If rsSQL.AbsolutePosition Mod 50 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
        rsSQL.MoveNext
    Loop
End If
rsSQL.Close

If bDebug Then DebugWrite "TreeGetViews Ends"
TreeGetViews = True

End Function

Public Function TreeInitialize() As Boolean
Screen.MousePointer = vbHourglass
If bDebug Then DebugWrite "TreeInitialize Starts"
    
    With frmMain.tvDB.Nodes
        .Clear
        .Add , , "fDatabase", "Database", imgDB
        .Add "fDatabase", tvwChild, "fTables", "Tables", imgTABLE
        .Add "fDatabase", tvwChild, "fViews", "Views", imgVIEW
        .Add "fDatabase", tvwChild, "fProcs", "Procedures", imgPROC
        .Add "fDatabase", tvwChild, "fTypes", "User-Defined Types", imgUDT
    End With
    frmMain.tvDB.Nodes(1).Expanded = True
TreeInitialize = True
If bDebug Then DebugWrite "TreeInitialize Ends"
Screen.MousePointer = vbDefault
End Function



