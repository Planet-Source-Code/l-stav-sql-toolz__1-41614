VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4425
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6120
   ControlBox      =   0   'False
   HelpContextID   =   20
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3315
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3315
      ScaleWidth      =   5685
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.CheckBox ckPrintColor 
         Alignment       =   1  'Right Justify
         Caption         =   "Print in Color"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   2700
         Width           =   2500
      End
      Begin VB.TextBox txtCommandFontName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtCommandFontSize 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4740
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCommandFont 
         Caption         =   "..."
         Height          =   285
         Left            =   5205
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox ckRestrictCursor 
         Alignment       =   1  'Right Justify
         Caption         =   "Restrict Cursor to Text"
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   1500
         Width           =   2500
      End
      Begin VB.TextBox txtTabSize 
         Height          =   285
         Left            =   2460
         TabIndex        =   23
         Top             =   1860
         Width           =   375
      End
      Begin VB.CheckBox ckExpandTabs 
         Alignment       =   1  'Right Justify
         Caption         =   "Expand Tabs to Spaces"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   2280
         Width           =   2500
      End
      Begin VB.CheckBox ckHighlightLine 
         Alignment       =   1  'Right Justify
         Caption         =   "Highlight Current Line"
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   1080
         Width           =   2500
      End
      Begin VB.ComboBox cboIndent 
         Height          =   315
         ItemData        =   "frmOptions.frx":0442
         Left            =   2460
         List            =   "frmOptions.frx":0444
         TabIndex        =   20
         Text            =   "cboIndent"
         Top             =   630
         Width           =   3010
      End
      Begin VB.Label Label1 
         Caption         =   "Command Window Font"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   42
         Top             =   285
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Tab Size"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   41
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Auto Indent Style"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   40
         Top             =   660
         Width           =   1695
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3315
      Index           =   0
      Left            =   210
      ScaleHeight     =   3315
      ScaleWidth      =   5685
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.CheckBox ckLogDateTime 
         Alignment       =   1  'Right Justify
         Caption         =   "Log Date/Time in File"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   3000
         Width           =   2500
      End
      Begin VB.TextBox txtLogToFile 
         Height          =   285
         Left            =   2460
         TabIndex        =   7
         Top             =   2580
         Width           =   2655
      End
      Begin VB.CommandButton cmdLogToFile 
         Caption         =   "..."
         Height          =   285
         Left            =   5220
         TabIndex        =   8
         Top             =   2580
         Width           =   255
      End
      Begin VB.CheckBox ckLogToFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Log Commands To File"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   2220
         Width           =   2500
      End
      Begin VB.ComboBox cboCursorDriver 
         Height          =   315
         ItemData        =   "frmOptions.frx":0446
         Left            =   2460
         List            =   "frmOptions.frx":0448
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1765
         Width           =   3015
      End
      Begin VB.ComboBox cboDSNList 
         Height          =   315
         ItemData        =   "frmOptions.frx":044A
         Left            =   2460
         List            =   "frmOptions.frx":044C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3010
      End
      Begin VB.CheckBox ckShowSystemObjects 
         Alignment       =   1  'Right Justify
         Caption         =   "Show System Objects"
         Height          =   255
         Left            =   150
         TabIndex        =   3
         Top             =   1050
         Width           =   2500
      End
      Begin VB.CheckBox ckShowPlan 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Plan"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   1440
         Width           =   2500
      End
      Begin VB.CheckBox ckForceDefaultDSN 
         Alignment       =   1  'Right Justify
         Caption         =   "Force Default DSN"
         Height          =   255
         Left            =   150
         TabIndex        =   2
         Top             =   660
         Width           =   2500
      End
      Begin VB.Label Label1 
         Caption         =   "Log File"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   36
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Cursor Driver"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   34
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Default DSN"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   33
         Top             =   285
         Width           =   900
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3315
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3315
      ScaleWidth      =   5685
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.CheckBox ckMakeBakFileOnSave 
         Alignment       =   1  'Right Justify
         Caption         =   "Make .BAK File On Save"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   1020
         Width           =   2500
      End
      Begin VB.ComboBox cboTimeFormat 
         Height          =   315
         ItemData        =   "frmOptions.frx":044E
         Left            =   2460
         List            =   "frmOptions.frx":0450
         TabIndex        =   18
         Text            =   "cboTimeFormat"
         Top             =   2880
         Width           =   3010
      End
      Begin VB.ComboBox cboDateFormat 
         Height          =   315
         ItemData        =   "frmOptions.frx":0452
         Left            =   2460
         List            =   "frmOptions.frx":0454
         TabIndex        =   17
         Text            =   "cboDateFormat"
         Top             =   2460
         Width           =   3010
      End
      Begin VB.CheckBox ckShowSplash 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Splash Screen"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   2160
         Width           =   2500
      End
      Begin VB.CheckBox ckSaveNewOnExit 
         Alignment       =   1  'Right Justify
         Caption         =   "Save New On Exit"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1800
         Width           =   2500
      End
      Begin VB.TextBox txtGridFontSize 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4740
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox txtGridFontName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2460
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   630
         Width           =   2175
      End
      Begin VB.CommandButton cmdGridFont 
         Caption         =   "..."
         Height          =   285
         Left            =   5200
         TabIndex        =   12
         Top             =   630
         Width           =   255
      End
      Begin VB.ComboBox cboListLocation 
         Height          =   315
         ItemData        =   "frmOptions.frx":0456
         Left            =   2460
         List            =   "frmOptions.frx":0458
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1380
         Width           =   3010
      End
      Begin VB.CommandButton cmdEditorPath 
         Caption         =   "..."
         Height          =   285
         Left            =   5200
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   285
         Left            =   2460
         TabIndex        =   10
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Time Format"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   38
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Date Format"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   37
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Result Grid Font"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   35
         Top             =   660
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "QuickList Location"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Default Editor"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   30
         Top             =   285
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   345
      Left            =   4740
      TabIndex        =   28
      Top             =   3990
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3360
      TabIndex        =   27
      Top             =   3990
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1980
      TabIndex        =   26
      Top             =   3990
      Width           =   1260
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3765
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6641
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Database"
            Key             =   "Database"
            Object.ToolTipText     =   "Set Database Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            Object.ToolTipText     =   "Set General Options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Editor"
            Key             =   "Editor"
            Object.ToolTipText     =   "Set Editor Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim sCommandFontStyle As String
Dim sGridFontStyle As String

Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1


Private Sub GetOptions()
Screen.MousePointer = vbHourglass
On Error Resume Next
  
    cboListLocation.AddItem "Left"
    cboListLocation.AddItem "Right"
    cboCursorDriver.AddItem "ODBC"
    cboCursorDriver.AddItem "Server"
    
    cboDateFormat.AddItem "yyyy-mm-dd"
    
    cboDateFormat.AddItem "mm-dd-yyyy"
    cboDateFormat.AddItem "mm/dd/yyyy"
    cboDateFormat.AddItem "m-d-yyyy"
    cboDateFormat.AddItem "m/d/yyyy"
    
    cboDateFormat.AddItem "dd-mm-yyyy"
    cboDateFormat.AddItem "dd/mm/yyyy"
    cboDateFormat.AddItem "d-m-yyyy"
    cboDateFormat.AddItem "d/m/yyyy"
    
    cboTimeFormat.AddItem "hh:nn:ss"
    cboTimeFormat.AddItem "hh:nn"
    cboTimeFormat.AddItem "h:n:s"
    cboTimeFormat.AddItem "h:n:s AMPM"
    cboTimeFormat.AddItem "h:n AMPM"
    cboTimeFormat.AddItem "h:n:s a/p"
    cboTimeFormat.AddItem "h:n a/p"
    
    cboIndent.AddItem "None"
    cboIndent.AddItem "Previous Line"
    cboIndent.AddItem "SQL Scope"
    
    GetDSNsAndDrivers
    
    cboDSNList.Text = GetSetting(App.Title, "Options", "DefaultDSN", "")
    ckForceDefaultDSN = GetSetting(App.Title, "Options", "ForceDefaultDSN", "0")
    txtEdit.Text = GetSetting(App.Title, "Options", "DefaultEditor", "Notepad")
    txtCommandFontName.Text = GetSetting(App.Title, "Options", "CommandFontName", "Courier New")
    txtCommandFontSize.Text = GetSetting(App.Title, "Options", "CommandFontSize", "10")
    sCommandFontStyle = GetSetting(App.Title, "Options", "CommandFontStyle", "Regular")
    txtGridFontName.Text = GetSetting(App.Title, "Options", "GridFontName", "MS Sans Serif")
    txtGridFontSize.Text = GetSetting(App.Title, "Options", "GridFontSize", "10")
    sGridFontStyle = GetSetting(App.Title, "Options", "GridFontStyle", "Regular")
    cboListLocation.Text = GetSetting(App.Title, "Options", "ListLocation", "Left")
    ckSaveNewOnExit = GetSetting(App.Title, "Options", "SaveNewOnExit", "1")
    ckShowSystemObjects = GetSetting(App.Title, "Options", "ShowSystemObjects", "0")
    ckShowPlan = GetSetting(App.Title, "Options", "ShowPlan", "0")
    cboCursorDriver.Text = GetSetting(App.Title, "Options", "CursorDriver", "Server")
    ckShowSplash = GetSetting(App.Title, "Options", "ShowSplash", "1")
    txtLogToFile.Text = GetSetting(App.Title, "Options", "LogFile", sLogFile)
    If bLogFile = True Then
        ckLogToFile = 1
    Else
        ckLogToFile = 0
    End If
    
    If GetSetting(App.Title, "Options", "LogDateTime", "0") = "0" Then
        ckLogDateTime = 0
    Else
        ckLogDateTime = 1
    End If
    
    If GetSetting(App.Title, "Options", "MakeBakFileOnSave", "0") = "0" Then
        ckMakeBakFileOnSave = 0
    Else
        ckMakeBakFileOnSave = 1
    End If
    
    cboDateFormat.Text = GetSetting(App.Title, "Options", "DateFormat", "yyyy-mm-dd")
    cboTimeFormat.Text = GetSetting(App.Title, "Options", "TimeFormat", "hh:nn:ss")
   
    cboIndent = GetSetting(App.Title, "Options", "IndentMode", "Previous")
    ckHighlightLine = GetSetting(App.Title, "Options", "HighlightLine", "0")
    ckRestrictCursor = GetSetting(App.Title, "Options", "RestrictCursor", "1")
    ckExpandTabs = GetSetting(App.Title, "Options", "ExpandTabs", "1")
    txtTabSize = CInt(GetSetting(App.Title, "Options", "TabSize", "4"))
    ckPrintColor = GetSetting(App.Title, "Options", "PrintColor", "0")
    
Screen.MousePointer = vbDefault
End Sub
Public Sub SaveOptions()
Dim lVerticalPos As Long
On Error Resume Next
    Screen.MousePointer = vbHourglass
        
    ' do we need to move the list
    If frmMain.sListLocation <> cboListLocation.Text Then
        If cboListLocation.Text = "Left" Then
            lVerticalPos = (frmMain.Width / 5) * 1 - 104
            If lVerticalPos < 2420 Then
                lVerticalPos = 2420
            End If
            frmMain.frSplitVertical.Left = lVerticalPos
         Else
            lVerticalPos = (frmMain.Width / 5) * 4 - 100
            If lVerticalPos > 8084 Then
                lVerticalPos = 8084
            End If
            frmMain.frSplitVertical.Left = lVerticalPos
        End If
    End If
    
    ' do we need to execute showplan
    If CInt(GetSetting(App.Title, "Options", "ShowPlan", "0")) <> ckShowPlan.Value Then
        ' execute showplan function
        ShowPlanSet ckShowPlan.Value
    End If
    
    SaveSetting App.Title, "Options", "DefaultDSN", cboDSNList.Text
    SaveSetting App.Title, "Options", "ForceDefaultDSN", ckForceDefaultDSN.Value
    SaveSetting App.Title, "Options", "DefaultEditor", txtEdit.Text
    SaveSetting App.Title, "Options", "CommandFontName", txtCommandFontName.Text
    SaveSetting App.Title, "Options", "CommandFontSize", txtCommandFontSize.Text
    SaveSetting App.Title, "Options", "CommandFontStyle", sCommandFontStyle
    SaveSetting App.Title, "Options", "GridFontName", txtGridFontName.Text
    SaveSetting App.Title, "Options", "GridFontSize", txtGridFontSize.Text
    SaveSetting App.Title, "Options", "GridFontStyle", sGridFontStyle
    SaveSetting App.Title, "Options", "ListLocation", cboListLocation.Text
    SaveSetting App.Title, "Options", "SaveNewOnExit", ckSaveNewOnExit.Value
    SaveSetting App.Title, "Options", "ShowSystemObjects", ckShowSystemObjects.Value
    SaveSetting App.Title, "Options", "ShowPlan", ckShowPlan.Value
    SaveSetting App.Title, "Options", "ShowSplash", ckShowSplash.Value
    SaveSetting App.Title, "Options", "DateFormat", cboDateFormat.Text
    SaveSetting App.Title, "Options", "TimeFormat", cboTimeFormat.Text
    SaveSetting App.Title, "Options", "LogDateTime", ckLogDateTime.Value
    SaveSetting App.Title, "Options", "MakeBakFileOnSave", ckMakeBakFileOnSave.Value
        
    sDateFormat = cboDateFormat.Text
    sTimeFormat = cboTimeFormat.Text
        
    'log file
    SaveSetting App.Title, "Options", "LogFile", txtLogToFile.Text
    SaveSetting App.Title, "Options", "LogToFile", ckLogToFile.Value
    If ckLogToFile.Value = 0 Then
        bLogFile = False
    Else
        bLogFile = True
    End If
    sLogFile = txtLogToFile.Text
    ' since we might have a new name, go ahead and close the file
    If iLogFile > 0 Then
        Close #iLogFile
    End If
    SaveSetting App.Title, "Options", "LogDateTime", ckLogDateTime.Value
    If ckLogDateTime.Value = 0 Then
        bLogTime = False
    Else
        bLogTime = True
    End If
    
    frmMain.sListLocation = GetSetting(App.Title, "Options", "ListLocation", "Left")
    
    If cboCursorDriver.Text = "Server" Then
        SaveSetting App.Title, "Options", "CursorDriver", "Server"
    Else
        SaveSetting App.Title, "Options", "CursorDriver", "ODBC"
    End If
        
    frmMain.rtfSQL.Font.Name = txtCommandFontName.Text
    frmMain.rtfSQL.Font.Size = CInt(txtCommandFontSize.Text)
    Select Case sCommandFontStyle
        Case Is = "Italic"
            frmMain.rtfSQL.Font.Italic = True
            frmMain.rtfSQL.Font.Bold = False
        Case Is = "Bold"
            frmMain.rtfSQL.Font.Italic = False
            frmMain.rtfSQL.Font.Bold = True
        Case Is = "Bold Italic"
            frmMain.rtfSQL.Font.Italic = True
            frmMain.rtfSQL.Font.Bold = True
        Case Else
            frmMain.rtfSQL.Font.Bold = False
            frmMain.rtfSQL.Font.Italic = False
    End Select

    frmMain.cmRun.Font.Name = txtCommandFontName.Text
    frmMain.cmRun.Font.Size = CInt(txtCommandFontSize.Text)
    Select Case sCommandFontStyle
        Case Is = "Italic"
            frmMain.cmRun.Font.Italic = True
            frmMain.cmRun.Font.Bold = False
        Case Is = "Bold"
            frmMain.cmRun.Font.Italic = False
            frmMain.cmRun.Font.Bold = True
        Case Is = "Bold Italic"
            frmMain.cmRun.Font.Italic = True
            frmMain.cmRun.Font.Bold = True
        Case Else
            frmMain.cmRun.Font.Bold = False
            frmMain.cmRun.Font.Italic = False
    End Select

    frmMain.grdSQL.Font.Name = txtGridFontName.Text
    frmMain.grdSQL.Font.Size = CInt(txtGridFontSize.Text)
    Select Case sGridFontStyle
        Case Is = "Italic"
            frmMain.grdSQL.Font.Italic = True
            frmMain.grdSQL.Font.Bold = False
        Case Is = "Bold"
            frmMain.grdSQL.Font.Italic = False
            frmMain.grdSQL.Font.Bold = True
        Case Is = "Bold Italic"
            frmMain.grdSQL.Font.Italic = True
            frmMain.grdSQL.Font.Bold = True
        Case Else
            frmMain.grdSQL.Font.Bold = False
            frmMain.grdSQL.Font.Italic = False
    End Select
   
    SaveSetting App.Title, "Options", "HighlightLine", ckHighlightLine.Value
    If ckHighlightLine.Value = 1 Then
        frmMain.bHighlight = True
        frmMain.rtfSQL.HighlightedLine = frmMain.rtfSQL.GetSel(True).EndLineNo
    Else
        frmMain.bHighlight = False
        frmMain.rtfSQL.HighlightedLine = -1
    End If

    
    ' restrict cursor to text
    SaveSetting App.Title, "Options", "RestrictCursor", ckRestrictCursor.Value
    If ckRestrictCursor.Value = 1 Then
        frmMain.rtfSQL.SelBounds = True
    Else
        frmMain.rtfSQL.SelBounds = False
    End If
    
    ' expand tabs
    SaveSetting App.Title, "Options", "ExpandTabs", ckExpandTabs.Value
    If ckExpandTabs.Value = 1 Then
        frmMain.rtfSQL.ExpandTabs = True
    Else
        frmMain.rtfSQL.ExpandTabs = False
    End If
    
    ' tabsize
    SaveSetting App.Title, "Options", "TabSize", CInt(txtTabSize.Text)
    frmMain.rtfSQL.TabSize = CInt(txtTabSize.Text)

    ' indent mode
    SaveSetting App.Title, "Options", "IndentMode", cboIndent.Text
    If GetSetting(App.Title, "Options", "IndentMode", "Previous Line") = "Previous Line" Then
        frmMain.rtfSQL.AutoIndentMode = cmIndentPrevLine
    Else
        If GetSetting(App.Title, "Options", "IndentMode", "Previous Line") = "SQL Scope" Then
            frmMain.rtfSQL.AutoIndentMode = cmIndentScope
        Else
            frmMain.rtfSQL.AutoIndentMode = cmIndentOff
        End If
    End If
    
    ' print color
    SaveSetting App.Title, "Options", "PrintColor", ckPrintColor.Value
    
    frmMain.ControlSize
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub cmdApply_Click()
    SaveOptions
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommandFont_Click()
        With frmMain.dlgCommonDialog
            .DialogTitle = "Fonts"
            .Flags = cdlCFANSIOnly Or cdlCFBoth Or cdlCFForceFontExist Or cdlCFFixedPitchOnly
            .FontName = GetSetting(App.Title, "Options", "CommandFontName", "Courier New")
            .FontSize = GetSetting(App.Title, "Options", "CommandFontSize", "10")
            Select Case sCommandFontStyle
                Case Is = "Italic"
                    .FontItalic = True
                    .FontBold = False
                Case Is = "Bold"
                    .FontItalic = False
                    .FontBold = True
                Case Is = "Bold Italic"
                    .FontItalic = True
                    .FontBold = True
                Case Else
                    .FontBold = False
                    .FontItalic = False
            End Select
            
            .CancelError = True
            On Error GoTo ErrHandler
            .ShowFont
            txtCommandFontName.Text = .FontName
            txtCommandFontSize.Text = .FontSize
            If .FontBold And .FontItalic Then
                sCommandFontStyle = "Bold Italic"
            ElseIf .FontBold Then
                sCommandFontStyle = "Bold"
            ElseIf .FontItalic Then
                sCommandFontStyle = "Italic"
            Else
                sCommandFontStyle = "Regular"
            End If
        End With
        Exit Sub

ErrHandler:
    Exit Sub
End Sub

Private Sub cmdEditorPath_Click()
Dim sText As String

        With frmMain.dlgCommonDialog
            .DialogTitle = "Default Editor"
            .Flags = cdlOFNHideReadOnly
            sText = .Filter
            .Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
            .CancelError = True
            On Error GoTo ErrHandler
            .ShowSave
            If Len(Trim(.FileName)) <> 0 Then
                txtEdit.Text = .FileName
            End If
            .Filter = sText
        End With
        Exit Sub

ErrHandler:
    'User pressed the Cancel button
    frmMain.dlgCommonDialog.Filter = sText
    Exit Sub
End Sub

Private Sub cmdGridFont_Click()
Dim sText As String

        With frmMain.dlgCommonDialog
            .DialogTitle = "Fonts"
            .Flags = cdlCFANSIOnly Or cdlCFBoth Or cdlCFForceFontExist
            .FontName = GetSetting(App.Title, "Options", "GridFontName", "MS Sans Serif")
            .FontSize = GetSetting(App.Title, "Options", "GridFontSize", "10")
            Select Case sGridFontStyle
                Case Is = "Italic"
                    .FontItalic = True
                    .FontBold = False
                Case Is = "Bold"
                    .FontItalic = False
                    .FontBold = True
                Case Is = "Bold Italic"
                    .FontItalic = True
                    .FontBold = True
                Case Else
                    .FontBold = False
                    .FontItalic = False
            End Select

            .CancelError = True
            On Error GoTo ErrHandler
            .ShowFont
            txtGridFontName.Text = .FontName
            txtGridFontSize.Text = .FontSize
            If .FontBold And .FontItalic Then
                sGridFontStyle = "Bold Italic"
            ElseIf .FontBold Then
                sGridFontStyle = "Bold"
            ElseIf .FontItalic Then
                sGridFontStyle = "Italic"
            Else
                sGridFontStyle = "Regular"
            End If
        End With
        Exit Sub

ErrHandler:
    Exit Sub
End Sub


Private Sub cmdLogToFile_Click()
Dim sText As String
Dim sMsg As String
        With frmMain.dlgCommonDialog
            .FileName = sLogFile
            .DialogTitle = "Log File"
            .Flags = cdlOFNHideReadOnly
            sText = .Filter
            .Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
            .CancelError = True
            On Error GoTo ErrHandler
            .ShowSave
            If Len(Trim(.FileName)) <> 0 Then
                txtLogToFile.Text = .FileName
            End If
            .Filter = sText
        End With
        Exit Sub

ErrHandler:
If Err.Number = 32755 Then
    'User pressed the Cancel button
    frmMain.dlgCommonDialog.Filter = sText
    Exit Sub
ElseIf Err.Number = 20477 Then
    'file name is bad
    sMsg = sMsg & "An invalid filename has prevented the dialog window from opening. "
    sMsg = sMsg & "Please ensure that : " & vbCrLf _
        & "1) No invalid characters are entered for the filename." & vbCrLf _
        & "2) The drive which holds the log file is available." & vbCrLf _
        & "3) You have rights to the directory which holds the log file." & vbCrLf
    sMsg = sMsg & "If you delete the filename in the Log File text box, the dialog window will open."
Else
    'something else happened
    sMsg = Err.Description
End If
    
sMsg = "Error:  " & Str(Err.Number) & vbCrLf & sMsg
Err.Clear
MsgBox sMsg, vbExclamation + vbOKOnly
Exit Sub

End Sub

Private Sub cmdOK_Click()
    SaveOptions
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim i As Integer
'    'handle ctrl+tab to move to the next tab
'    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
'        i = tbsOptions.SelectedItem.Index
'        If i = tbsOptions.Tabs.Count Then
'            'last tab so we need to wrap to tab 1
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
'        Else
'            'increment the tab
'            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
'        End If
'    End If
End Sub
Sub GetDSNsAndDrivers()
If bDebug Then DebugWrite "GetDSNsAndDrivers Starts"
    On Error Resume Next
    Dim i As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSN As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long         'handle to the environment
    Dim sDefaultDSN As String
    'cboDSNList.AddItem "(None)"

    
    'get the DSNs
    If bDebug Then DebugWrite "Alloc Starts"
    If SQLAllocEnv(lHenv) <> -1 Then
        If bDebug Then DebugWrite "Loop Starts"
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSN = Left$(sDSNItem, iDSNLen)
            sDRV = Left$(sDRVItem, iDRVLen)
            If sDSN <> Space(iDSNLen) Then
                cboDSNList.AddItem sDSN
                If bDebug Then DebugWrite "sDSN Added " & sDSN
            End If
        Loop
        If bDebug Then DebugWrite "Loop Ends"
    End If
    If bDebug Then DebugWrite "Alloc Ends"
    sDefaultDSN = GetSetting(App.Title, "Options", "DefaultDSN")
    If sDefaultDSN <> "" Then
        cboDSNList.Text = sDefaultDSN
    Else
        cboDSNList.ListIndex = 0
    End If
If bDebug Then DebugWrite "GetDSNsAndDrivers Ends"
End Sub


Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    cmdCancel.Left = cmdApply.Left - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    GetOptions
End Sub

Private Sub tbsOptions_Click()

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next

End Sub
