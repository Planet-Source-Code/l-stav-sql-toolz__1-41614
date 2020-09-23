VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   ScaleHeight     =   4710
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1800
      Left            =   60
      Top             =   4620
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   600
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame fraMainFrame 
      Height          =   4650
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   1020
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblCMax 
         Alignment       =   1  'Right Justify
         Caption         =   "Portions Copyright© 1997-2000 Barry Allyn.  All rights reserved."
         Height          =   255
         Left            =   510
         TabIndex        =   7
         Tag             =   "Copyright"
         Top             =   2820
         Width           =   6495
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2700
         TabIndex        =   6
         Tag             =   "Version"
         Top             =   1920
         Width           =   930
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2700
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5985
         TabIndex        =   4
         Tag             =   "Platform"
         Top             =   3180
         Width           =   1020
      End
      Begin VB.Label lblWarning 
         Caption         =   $"frmSplash.frx":0442
         Height          =   855
         Left            =   300
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3600
         Width           =   6855
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright© 1998 Larry W. Stavinoha"
         Height          =   255
         Left            =   510
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   2460
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim sPlatform As String
    Select Case SysInfo1.OSPlatform
        Case 1
        sPlatform = ""
        lblWarning.Caption = ""
        Case 2
            If CStr(SysInfo1.OSVersion) = "4" Then
                sPlatform = "Windows NT " & CStr(SysInfo1.OSVersion) _
                    & " Build " & CStr(SysInfo1.OSBuild)
            ElseIf CStr(SysInfo1.OSVersion) = "5" Then
                    sPlatform = "Windows 2000 Build " & CStr(SysInfo1.OSBuild)
            ElseIf CStr(SysInfo1.OSVersion) > "5" Then
                    sPlatform = "Windows XP Build " & CStr(SysInfo1.OSBuild)
            End If
            lblWarning.Caption = App.Comments
        Case Else
            sPlatform = ""
            lblWarning.Caption = ""
    End Select
        
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
    lblProductName.Caption = App.Title
    lblCopyright.Caption = App.LegalCopyright
    lblPlatform.Caption = sPlatform
    DoEvents
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub


