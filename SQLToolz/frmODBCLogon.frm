VERSION 5.00
Begin VB.Form frmODBCLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ODBC Logon"
   ClientHeight    =   2130
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4680
   ControlBox      =   0   'False
   HelpContextID   =   10
   Icon            =   "frmODBCLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdODBC 
      Caption         =   "ODBC &Admin"
      Height          =   345
      Left            =   3360
      TabIndex        =   5
      Top             =   1575
      Width           =   1260
   End
   Begin VB.ComboBox cboDSNList 
      Height          =   315
      ItemData        =   "frmODBCLogon.frx":0442
      Left            =   1290
      List            =   "frmODBCLogon.frx":0444
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   930
      Width           =   3015
   End
   Begin VB.TextBox txtUID 
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2220
      TabIndex        =   4
      Top             =   1575
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   615
      TabIndex        =   3
      Top             =   1575
      Width           =   1260
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "DSN"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   8
      Top             =   300
      UseMnemonic     =   0   'False
      Width           =   825
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   7
      Top             =   653
      UseMnemonic     =   0   'False
      Width           =   825
   End
   Begin VB.Label lblStep3 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   6
      Top             =   983
      UseMnemonic     =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmODBCLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv&, ByVal fDirection%, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN%, ByVal szDescription$, ByVal cbDescriptionMax%, pcbDescription%) As Integer
Private Declare Function SQLAllocEnv% Lib "ODBC32.DLL" (env&)
Private Declare Function SQLCreateDataSource Lib "odbccp32.dll" (ByVal hWnd As Long, ByVal szDSN As String) As Boolean
Private Declare Function SQLManageDataSources Lib "odbccp32.dll" (ByVal hWnd As Long) As Boolean
Const SQL_SUCCESS As Long = 0
Const SQL_FETCH_NEXT As Long = 1




Private Sub cmdCancel_Click()
    bGlobalRetVal = False
    Unload Me
End Sub

Private Sub cmdODBC_Click()
    SQLManageDataSources frmODBCLogon.hWnd
    GetDSNsAndDrivers
End Sub

Private Sub cmdOK_Click()
Dim bRetVal As Boolean
Dim sDSN As String
Dim sUID As String
Dim sPWD As String
Dim sConnect As String
    If bDebug Then DebugWrite "cmdOK_Click Starts"
    Me.Caption = App.Title + " Connecting..."
    sDSN = cboDSNList.Text
    sUID = txtUID.Text
    sPWD = txtPWD.Text
    bRetVal = DBConnect(sDSN, sUID, sPWD)
    If bRetVal Then
        bGlobalRetVal = True
        If GetSetting(App.Title, "Options", "ForceDefaultDSN", "0") = 0 Then
            SaveSetting App.Title, "Options", "DefaultDSN", cboDSNList.Text
        End If
        Unload Me
    Else
        Me.Caption = App.Title
    End If
    If bDebug Then DebugWrite "cmdOK_Click Ends"
End Sub




Private Sub Form_Load()
    If bDebug Then DebugWrite "Form_Load Starts"
    Me.Caption = App.Title & " Login"
    cmdOK.Left = Me.Width / 3 - cmdOK.Width - 50
    cmdCancel.Left = cmdOK.Left + cmdOK.Width + 200
    cmdODBC.Left = cmdCancel.Left + cmdCancel.Width + 200
    GetDSNsAndDrivers
    If bDebug Then DebugWrite "Form_Load Ends"
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
    cboDSNList.Clear
    'get the DSNs
    If SQLAllocEnv(lHenv) <> -1 Then
        If bDebug Then DebugWrite "Alloc Starts"
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
        If bDebug Then DebugWrite "Alloc Ends"
    End If
    
    sDefaultDSN = GetSetting(App.Title, "Options", "DefaultDSN")
    If sDefaultDSN <> "" Then
        cboDSNList.Text = sDefaultDSN
    Else
        cboDSNList.ListIndex = 0
    End If
If bDebug Then DebugWrite "GetDSNsAndDrivers Ends"
End Sub

