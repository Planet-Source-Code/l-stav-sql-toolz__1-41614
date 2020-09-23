Attribute VB_Name = "sqltool"
Option Explicit

Global iDebugFile As Integer
Global iLogFile As Integer

Global conSQL As rdoConnection
Global rsSQL As rdoResultset
Global envSQL As rdoEnvironment

Global bCancel As Boolean
Global bDebug As Boolean
Global bLogFile As Boolean
Global bGlobalRetVal As Boolean
Global bProcessing As Boolean
Global bLogTime As Boolean

Global sDQUOTE As String
Global sLogFile As String
Global sDateFormat As String
Global sTimeFormat As String
Global sComment As String
Global sFile As String

' Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)

Public Function LogWrite(ByVal sText As String) As Boolean
Dim sMsg As String
Dim iCount As Integer
Dim iMaxCount As Integer
Dim bContinue As Boolean
Dim iFilePool As Integer
Dim sLogTime As String

On Error GoTo AppError
LogWrite = False
    
If bDebug Then DebugWrite "LogWrite Starts"
sText = sText & vbCrLf & "go"
iCount = 0
iMaxCount = 25
bContinue = True
iFilePool = 0
If bLogTime Then
    sLogTime = "/* " & Format(Now, sDateFormat & " " & sTimeFormat) & " */" & vbCrLf
    sText = sLogTime & sText
End If

If sLogFile = "" Then
    ' no log file in memory, lets check the registry
    sLogFile = GetSetting(App.Title, "Options", "LogFile", App.Path & "\" & App.Title & ".log")
End If

For iCount = 0 To iMaxCount
    If bContinue = True Then
        ' file not open, lets get a new number
        If iLogFile <= 0 Then
            On Error Resume Next
            iLogFile = FreeFile(iFilePool)
            Open sLogFile For Append Lock Write As #iLogFile
            If Err.Number = 70 Or Err.Number = 52 Or Err.Number = 20477 Or Err.Number = 75 Or Err.Number = 55 Then
                bContinue = True
                iLogFile = 0
                If iFilePool = 0 Then
                    iFilePool = 1
                Else
                    iFilePool = 0
                End If
                If iCount >= iMaxCount Then GoTo AppError
            ElseIf Err.Number > 0 Then
                bContinue = False
                iLogFile = 0
            End If
        End If
        'print the text
        If iLogFile > 0 Then
            On Error GoTo 0
            Print #iLogFile, sText
            Close #iLogFile
            iLogFile = 0
            Exit For
        End If
    End If
Next

On Error GoTo 0

If bDebug Then DebugWrite "LogWrite Ends"
LogWrite = True
Exit Function

AppError:
    frmMain.StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf
        If Err.Number = 52 Or Err.Number = 20477 Or Err.Number = 70 Or Err.Number = 75 Or Err.Number = 55 Then
            sMsg = sMsg & "Unable to write to log file. "
            sMsg = sMsg & "Please ensure that : " & vbCrLf _
                & "1) A valid filename is entered on the " _
                & "Database Tab of the Options Window." & vbCrLf _
                & "2) The drive which holds the log file is available." & vbCrLf _
                & "3) You have rights to the directory which holds the log file." & vbCrLf _
                & "4) The log file is not write protected or in use by another process."
        Else
            sMsg = sMsg & Err.Description
        End If
        sMsg = sMsg & vbCrLf & "Logging of commands has been disabled temporarily."
        If MsgBox(sMsg, vbExclamation + vbOKCancel) = vbCancel Then
            bCancel = True
        End If
        Err.Clear
    End If
    If iLogFile > 0 Then
        Close #iLogFile
        iLogFile = 0
    End If
    bLogFile = False
    Screen.MousePointer = vbDefault
    LogWrite = False
    Exit Function

End Function

Public Sub Main()
Dim aCommand() As String, sCommand As String
Dim iCount As Integer
    
    sDQUOTE = Chr(34)
    sFile = ""
    sComment = "-- "
    iCount = 0
    
    bGlobalRetVal = False
    
    ' bDebug = True
    
    ' get the command line
    If Command$ <> "" Then
        aCommand = SplitAroundQuotes(Command$, " ")
        ' only 2 allowable, -z and a filename
        For iCount = LBound(aCommand) To UBound(aCommand)
            ' get the value
            sCommand = aCommand(iCount)
            If aCommand(iCount) = "-z" Then
                bDebug = True
            Else
                If aCommand(iCount) <> "" Then
                    sFile = aCommand(iCount)
                    ' strip the quotes
                    sFile = Replace(sFile, """", "")
                End If
            End If
        Next
    End If
            
    If bDebug Then
        iDebugFile = FreeFile(0)
        Open App.Path & "\debug.log" For Output As #iDebugFile
        DebugWrite "App Starts"
        DebugWrite "Command Line:  " & Trim(Command())
    End If

    rdoEngine.rdoDefaultCursorDriver = rdUseIfNeeded
    
    Set envSQL = rdoEnvironments(0)
    If bDebug Then DebugWrite "frmMain.Show"
    frmMain.Show

End Sub


Function SplitAroundQuotes(TextToSplit As String, _
    Optional Delimiter As String = ",") As String()
Dim QuoteDelimited() As String
Dim WorkingArray() As String
Dim iCount As Integer

QuoteDelimited = Split(TextToSplit, """")

For iCount = 1 To UBound(QuoteDelimited) Step 2
    QuoteDelimited(iCount) = Replace$(QuoteDelimited(iCount), Delimiter, Chr$(0))
Next

TextToSplit = Join(QuoteDelimited, """")

TextToSplit = Replace$(TextToSplit, Delimiter, Chr$(1))

TextToSplit = Replace$(TextToSplit, Chr$(0), Delimiter)

WorkingArray = Split(TextToSplit, Chr$(1))

SplitAroundQuotes = WorkingArray

End Function

Public Function DebugWrite(ByVal sText As String) As Boolean
Dim sTime As String
    sTime = Str(Now)
    sTime = sTime & Space(24 - Len(sTime))
    sText = sTime & sText
    Print #iDebugFile, sText
    DebugWrite = True
End Function





