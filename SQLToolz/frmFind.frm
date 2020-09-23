VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   HelpContextID   =   500
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ckWholeWord 
      Caption         =   "Whole Word"
      Enabled         =   0   'False
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame frDirection 
      Caption         =   "Direction"
      Enabled         =   0   'False
      Height          =   675
      Left            =   2100
      TabIndex        =   8
      Top             =   540
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton optDown 
         Caption         =   "&Down"
         Enabled         =   0   'False
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optUp 
         Caption         =   "&Up"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.CheckBox ckCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Top             =   480
      Width           =   1260
   End
   Begin VB.ComboBox cboFindText 
      Height          =   315
      ItemData        =   "frmFind.frx":0442
      Left            =   1020
      List            =   "frmFind.frx":0444
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Top             =   60
      Width           =   1260
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "Find What:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   180
      UseMnemonic     =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFindText_Change()
    If cboFindText.Text <> "" Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
End Sub

Private Sub ckCase_Click()
    If ckCase.Value = 1 Then
        frmMain.iFindCase = 1
    Else
        frmMain.iFindCase = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim sText As String
sText = cboFindText.Text
If sText <> "" Then
    frmMain.sFindText = cboFindText.Text
    frmMain.FileFind
End If


'Dim sText As String
'Dim lFound As Long
'Dim lStart As Long
'
'    bProcessing = True
'    Screen.MousePointer = vbHourglass
'    If cboFindText.Text <> cboFindText.List(0) Then
'        cboFindText.AddItem cboFindText.Text, 0
'    End If
'
'    frmMain.rtfSQL.HideSelection = False
'    lStart = frmMain.rtfSQL.SelStart
'    If bFound Then
'        lStart = lStart + 2
'        frmMain.rtfSQL.SelLength = 0
'    End If
'    If lStart < 1 Then
'        lStart = 1
'    End If
'
'    'search for the text
'    sText = cboFindText.Text
'    If sText <> "" Then
'        frmMain.sFindText = cboFindText.Text
'        If ckCase = 1 Then
'            lFound = InStr(lStart, frmMain.rtfSQL.Text, sText, vbBinaryCompare)
'        Else
'            lFound = InStr(lStart, frmMain.rtfSQL.Text, sText, vbTextCompare)
'        End If
'
'        If lFound < 1 Then
'            MsgBox "Text '" & sText & "' Not Found", vbInformation
'            bFound = False
'        Else
'            bFound = True
'            frmMain.rtfSQL.SelStart = lFound - 1
'            frmMain.rtfSQL.SelLength = Len(sText)
'        End If
'    End If
'    Screen.MousePointer = vbDefault
'    bProcessing = False
End Sub

Private Sub Form_Load()
Dim sText As String

Me.Left = frmMain.Left + frmMain.Width - Me.Width - 120
Me.Top = frmMain.Top + 655
sText = frmMain.sFindText

If sText <> "" Then
    If cboFindText.Text <> sText Then
        cboFindText.AddItem sText, 0
        cboFindText.Text = sText
    End If
End If
If cboFindText.Text <> "" Then
    cmdFind.Enabled = True
Else
    cmdFind.Enabled = False
End If
ckCase.Value = frmMain.iFindCase

End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub


