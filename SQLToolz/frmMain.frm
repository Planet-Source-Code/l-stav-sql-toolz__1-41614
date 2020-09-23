VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{ECEDB943-AC41-11D2-AB20-000000000000}#2.0#0"; "cmax20.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   10560
   FillStyle       =   7  'Diagonal Cross
   HelpContextID   =   30
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   39
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print Results"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Description     =   "Export"
            Object.ToolTipText     =   "Export"
            ImageKey        =   "Export"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ExportView"
            Description     =   "Export View"
            Object.ToolTipText     =   "Export View"
            ImageKey        =   "ExportView"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Description     =   "Redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Description     =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageKey        =   "Connect"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Description     =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageKey        =   "Disconnect"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Replace"
            Description     =   "Replace"
            Object.ToolTipText     =   "Replace"
            ImageKey        =   "Replace"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ExecuteSQL"
            Description     =   "Execute SQL"
            Object.ToolTipText     =   "Execute SQL"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Description     =   "Cancel"
            Object.ToolTipText     =   "Cancel"
            ImageKey        =   "Cancel"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SortAscending"
            Description     =   "Sort Ascending"
            Object.ToolTipText     =   "Sort Ascending"
            ImageKey        =   "SortAscending"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SortDescending"
            Description     =   "Sort Descending"
            Object.ToolTipText     =   "Sort Descending"
            ImageKey        =   "SortDescending"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indent"
            Description     =   "Increase Indent"
            Object.ToolTipText     =   "Increase Indent"
            ImageKey        =   "Indent"
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outdent"
            Description     =   "Decrease Indent"
            Object.ToolTipText     =   "Decrease Indent"
            ImageKey        =   "Outdent"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Comment"
            Description     =   "Comment Block"
            Object.ToolTipText     =   "Comment Block"
            ImageKey        =   "Comment"
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Uncomment"
            Description     =   "Uncomment Block"
            Object.ToolTipText     =   "Uncomment Block"
            ImageKey        =   "Uncomment"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BookToggle"
            Description     =   "Toggle Bookmark"
            Object.ToolTipText     =   "Toggle Bookmark"
            ImageKey        =   "BookToggle"
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BookNext"
            Description     =   "Next Bookmark"
            Object.ToolTipText     =   "Next Bookmark"
            ImageKey        =   "BookNext"
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BookPrevious"
            Description     =   "Previous Bookmark"
            Object.ToolTipText     =   "Previous Bookmark"
            ImageKey        =   "BookPrevious"
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BookDel"
            Description     =   "Clear All Bookmarks"
            Object.ToolTipText     =   "Clear All Bookmarks"
            ImageKey        =   "BookDel"
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Log"
            Description     =   "Log"
            Object.ToolTipText     =   "Log"
            ImageKey        =   "Log"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrMenu 
      Interval        =   2500
      Left            =   9540
      Top             =   2520
   End
   Begin VB.Frame frSplitVertical 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5580
      Left            =   6420
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   440
      Width           =   50
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8880
      Top             =   540
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6945
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10478
            Key             =   "Main"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Ln: 1"
            TextSave        =   "Ln: 1"
            Key             =   "Line"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1147
            Text            =   "Col: 1"
            TextSave        =   "Col: 1"
            Key             =   "Col"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "Server:  "
            TextSave        =   "Server:  "
            Key             =   "Server"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "DB:  "
            TextSave        =   "DB:  "
            Key             =   "DB"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "User:  "
            TextSave        =   "User:  "
            Key             =   "User"
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".sql"
      DialogTitle     =   "Save As"
      Filter          =   "SQL Files (*.sql)|*.sql|DDL Files (*.ddl)|*.ddl|All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
   End
   Begin VB.Frame frSQL 
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   6135
      Begin CodeMaxCtl.CodeMax rtfSQL 
         Height          =   2655
         Left            =   0
         OleObjectBlob   =   "frmMain.frx":0442
         TabIndex        =   15
         Top             =   3840
         Width           =   2175
      End
      Begin VB.PictureBox picSQL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         ScaleHeight     =   210
         ScaleWidth      =   4095
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3500
         Width           =   4095
         Begin VB.Label lblSQL 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Command"
            ForeColor       =   &H80000009&
            Height          =   195
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   705
         End
      End
      Begin VB.PictureBox picGrd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   60
         ScaleHeight     =   210
         ScaleWidth      =   4095
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   4095
         Begin VB.Label lblGrd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Results"
            ForeColor       =   &H80000009&
            Height          =   195
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.Frame frSplitHorizontal 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   50
         Left            =   30
         MousePointer    =   7  'Size N S
         TabIndex        =   7
         Top             =   2460
         Width           =   4155
      End
      Begin MSFlexGridLib.MSFlexGrid grdSQL 
         Height          =   2955
         Left            =   0
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   5212
         _Version        =   393216
         BackColorBkg    =   -2147483636
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         AllowUserResizing=   3
      End
      Begin CodeMaxCtl.CodeMax cmRun 
         Height          =   2655
         Left            =   3480
         OleObjectBlob   =   "frmMain.frx":05B6
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3840
         Width           =   2175
      End
   End
   Begin VB.Frame frList 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   6660
      TabIndex        =   9
      Top             =   420
      Width           =   2235
      Begin MSComctlLib.TreeView tvDB 
         Height          =   1995
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   3519
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlTreePics"
         Appearance      =   1
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
      Begin MSComctlLib.TabStrip tbsList 
         Height          =   1755
         Left            =   60
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4500
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   3096
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Schema"
               Key             =   "Schema"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Messages"
               Key             =   "Messages"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "History"
               Key             =   "History"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstStatus 
         BackColor       =   &H8000000F&
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid grdHistory 
         Height          =   1515
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2672
         _Version        =   393216
         Rows            =   1
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   -2147483633
         ForeColorFixed  =   -2147483640
         BackColorBkg    =   -2147483633
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   0
         GridLinesFixed  =   0
         AllowUserResizing=   3
      End
   End
   Begin MSComctlLib.ImageList imlTreePics 
      Left            =   9360
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0728
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":083A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":094C
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E9E
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B8
            Key             =   "Column"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12CA
            Key             =   "Database"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DC
            Key             =   "FoldClose"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14EE
            Key             =   "FoldOpen"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1600
            Key             =   "Procedure"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1712
            Key             =   "Argument"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1824
            Key             =   "Primary"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1936
            Key             =   "Foreign"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A48
            Key             =   "UDT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F9A
            Key             =   "Trigger"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9360
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20AC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21BE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22D0
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23E2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24F4
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2606
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2718
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":282A
            Key             =   "PrintOld"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":293C
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A4E
            Key             =   "SortAscending"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B60
            Key             =   "SortDescending"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C72
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D84
            Key             =   "ExportView"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E96
            Key             =   "Disconnect"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FA8
            Key             =   "Connect"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30BA
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31CC
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32DE
            Key             =   "Log"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33F0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3942
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A54
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B66
            Key             =   "Indent"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C78
            Key             =   "Outdent"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D8A
            Key             =   "Case"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E9C
            Key             =   "BookToggle"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FAE
            Key             =   "BookNext"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40C0
            Key             =   "BookPrevious"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41D2
            Key             =   "BookDel"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42E4
            Key             =   "Replace"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43F6
            Key             =   "Comment"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4508
            Key             =   "Uncomment"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   410
      Begin VB.Menu mnuFileChild 
         Caption         =   "&New"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "&Open..."
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "&Save..."
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "&Print Results"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "Print &Command Window"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "Print Setup ..."
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "&Export..."
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "Export &View..."
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFileChild 
         Caption         =   "E&xit"
         Enabled         =   0   'False
         Index           =   13
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   420
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Execute SQL	Ctrl+E"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Undo	Ctrl+Z"
         Index           =   2
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Redo"
         Index           =   3
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "Cu&t	Ctrl+X"
         Index           =   5
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Copy	Ctrl+C"
         Index           =   6
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Paste	Ctrl+V"
         Index           =   7
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Delete	Del"
         Index           =   8
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "Select &All"
         Index           =   10
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "F&ormat"
         Index           =   12
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Tabs to Spaces"
            Index           =   0
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Spaces to Tabs"
            Index           =   1
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "Sh&ow Tabs and Spaces"
            Index           =   2
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Lower Case"
            Index           =   4
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Upper Case"
            Index           =   5
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "Ca&pitalize"
            Index           =   6
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Increase Indent	Tab"
            Index           =   8
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Decrease Indent	Shift+Tab"
            Index           =   9
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "Transpose &Characters"
            Index           =   11
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "Transpose &Words"
            Index           =   12
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "Transpose L&ines"
            Index           =   13
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "-"
            Index           =   14
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Comment Block"
            Index           =   15
         End
         Begin VB.Menu mnuEditChildFormat 
            Caption         =   "&Uncomment Block"
            Index           =   16
         End
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Find..."
         Index           =   14
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "Find &Next"
         Index           =   15
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "R&eplace..."
         Index           =   16
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuEditChild 
         Caption         =   "&Go To"
         Index           =   18
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "&Go To Line...	Ctrl+G"
            Index           =   0
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "Matching &Brace	Ctrl+]"
            Index           =   2
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "&Toggle Bookmark"
            Index           =   4
            Shortcut        =   {F2}
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "&Next Bookmark"
            Index           =   5
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "&Previous Bookmark"
            Index           =   6
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuEditChildGoTo 
            Caption         =   "&Clear All Bookmarks"
            Index           =   7
         End
      End
   End
   Begin VB.Menu mnuPopGrid 
      Caption         =   "PopGrid"
      Enabled         =   0   'False
      HelpContextID   =   440
      Visible         =   0   'False
      Begin VB.Menu mnuPopGridCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopGridSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGridPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopGridSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGridClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuPopGridSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGridSortAsc 
         Caption         =   "Sort &Ascending"
      End
      Begin VB.Menu mnuPopGridSortDesc 
         Caption         =   "Sort &Descending"
      End
   End
   Begin VB.Menu mnuPopTreeObject 
      Caption         =   "PopTreeObject"
      Enabled         =   0   'False
      HelpContextID   =   450
      Visible         =   0   'False
      Begin VB.Menu mnuPopTreeObjectCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPopTreeObjectCopyColumnSelect 
         Caption         =   "Copy Select Statement"
      End
      Begin VB.Menu mnuPopTreeObjectCopyColumnInsert 
         Caption         =   "Copy Insert Statement"
      End
      Begin VB.Menu mnuPopTreeObjectCopyColumnUpdate 
         Caption         =   "Copy Update Statement"
      End
      Begin VB.Menu mnuPopTreeObjectCopyColumnDelete 
         Caption         =   "Copy Delete Statement"
      End
      Begin VB.Menu mnuPopTreeObjectCopyParameterExecute 
         Caption         =   "Copy Execute Statement"
      End
      Begin VB.Menu mnuPopTreeObjectCopyColumn 
         Caption         =   "Copy with Columns"
      End
      Begin VB.Menu mnuPopTreeObjectCopyParameter 
         Caption         =   "Copy with Parameters"
      End
      Begin VB.Menu mnuPopTreeObjectCopyDDL 
         Caption         =   "Copy DDL"
      End
      Begin VB.Menu mnuPopObjectTableSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopTreeObjectPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuPopTreeObjectPasteColumnSelect 
         Caption         =   "Paste Select Statement"
      End
      Begin VB.Menu mnuPopTreeObjectPasteColumnInsert 
         Caption         =   "Paste Insert Statement"
      End
      Begin VB.Menu mnuPopTreeObjectPasteColumnUpdate 
         Caption         =   "Paste Update Statement"
      End
      Begin VB.Menu mnuPopTreeObjectPasteColumnDelete 
         Caption         =   "Paste Delete Statement"
      End
      Begin VB.Menu mnuPopTreeObjectPasteParameterExecute 
         Caption         =   "Paste Execute Statement"
      End
      Begin VB.Menu mnuPopTreeObjectPasteColumn 
         Caption         =   "Paste with Columns"
      End
      Begin VB.Menu mnuPopTreeObjectPasteParameter 
         Caption         =   "Paste with Parameters"
      End
      Begin VB.Menu mnuPopTreeObjectPasteDDL 
         Caption         =   "Paste DDL"
      End
   End
   Begin VB.Menu mnuPopGridHistory 
      Caption         =   "PopGridHistory"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopGridHistoryCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopGridHistorySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGridHistoryPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopGridHistorySep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopGridHistoryDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuDB 
      Caption         =   "&Database"
      HelpContextID   =   460
      Begin VB.Menu mnuDBChild 
         Caption         =   "&Execute SQL"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDBChild 
         Caption         =   "Cancel E&xecution"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuDBChild 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuDBChild 
         Caption         =   "&Connect..."
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuDBChild 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
         Index           =   4
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   470
      Begin VB.Menu mnuViewChild 
         Caption         =   "Sort &Ascending"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuViewChild 
         Caption         =   "Sort &Descending"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuViewChild 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuViewChild 
         Caption         =   "&Options..."
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuViewChild 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuViewChild 
         Caption         =   "&Refresh"
         Enabled         =   0   'False
         Index           =   5
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      HelpContextID   =   480
      Begin VB.Menu mnuWindowCommand 
         Caption         =   "&Command"
      End
      Begin VB.Menu mnuWindowResult 
         Caption         =   "&Results"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuWindowHistory 
         Caption         =   "H&istory"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuWindowStatus 
         Caption         =   "&Messages"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuWindowSchema 
         Caption         =   "&Schema"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   490
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public sListLocation As String
Public sDBServer As String
Public bHighlight As Boolean

Dim mbMoving As Boolean
Dim bPictureAdded As Boolean
Dim bPictureCleared As Boolean

Dim iGridCol() As Integer

Const sglSplitLimit = 500
Const MAXDESIGNCELLS = 350000




Private Function FilePrint() As Boolean
Dim PrintString As String
Dim sMsg As String
        
On Error GoTo AppError
FilePrint = False
        
If bDebug Then DebugWrite "FilePrint Starts"
 
    PrintString = rtfSQL.Text
    ' now to wrap the text
    PrintString = TextWrap(PrintString, 70, False)
    ' now to set the printing fonts
    cmRun.Font.Name = "Courier New"
    cmRun.Font.Size = 10
    ' now to print
    cmRun.Text = PrintString
    If GetSetting(App.Title, "Options", "PrintColor", "0") = "0" Then
        Call cmRun.PrintContents(0, cmPrnRichFonts + cmPrnDefaultPrn + cmPrnPageNums + cmPrnDateTime)
    Else
        Call cmRun.PrintContents(0, cmPrnRichFonts + cmPrnColor + cmPrnDefaultPrn + cmPrnPageNums + cmPrnDateTime)
    End If
        
        '   cmPrnPromptDlg - Prompt the user to configure a printer.
        '   cmPrnDefaultPrn - No prompting; use the default printer.
        '   cmPrnHDC - The hDC parameter should be used.
        '   cmPrnRichFonts - Use bold, italics, and underline in the output.
        '   cmPrnColor - Print in color.
        '   cmPrnPageNums - Print page numbers in the footer.
        '   cmPrnDateTime - Print the date and time in the header.
        '   cmPrnBorderThin - Print a thick border around the text.
        '   cmPrnBorderThick - Print a thick border around the page.
        '   cmPrnBorderDouble - Print a double border (thick or thin) around the page.
        '   CmPrnSelection - Prints the selection if non-empty.
        
    ' now to reset
    cmRun.Text = ""
    cmRun.Font.Name = rtfSQL.Font.Name
    cmRun.Font.Size = rtfSQL.Font.Size

If bDebug Then DebugWrite "FilePrint Ends"
FilePrint = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    FilePrint = False
    Exit Function

End Function

Public Function MenuSet(iCase As Integer) As Boolean
Dim sControl As String
Dim iCount As Integer

MenuSet = False
On Error Resume Next
DoEvents
Select Case iCase
    Case 0
        ' disable
        ' file menu
        For iCount = mnuFileChild.LBound To mnuFileChild.UBound
            mnuFileChild(iCount).Enabled = False
        Next
        
        ' Edit menu
        For iCount = mnuEditChild.LBound To mnuEditChild.UBound
            mnuEditChild(iCount).Enabled = False
        Next
        
        ' db menu
        For iCount = mnuDBChild.LBound To mnuDBChild.UBound
            mnuDBChild(iCount).Enabled = False
        Next
        mnuDBChild(1).Enabled = True
        ' view menu
        For iCount = mnuViewChild.LBound To mnuViewChild.UBound
            mnuViewChild(iCount).Enabled = False
        Next
        
        ' help menu
        mnuHelpAbout.Enabled = False
        
        ' toolbar
        ' For iCount = 1 To tbToolbar.Buttons.Count - 1
        For iCount = 1 To tbToolbar.Buttons.Count
            tbToolbar.Buttons(iCount).Enabled = False
        Next iCount

        tbToolbar.Buttons("Cancel").Enabled = True
    Case 1
        ' enable
        ' file menu
        For iCount = mnuFileChild.LBound To mnuFileChild.UBound
            mnuFileChild(iCount).Enabled = True
        Next
        
        ' Edit menu
        For iCount = mnuEditChild.LBound To mnuEditChild.UBound
            mnuEditChild(iCount).Enabled = True
        Next
        
        ' db menu
        For iCount = mnuDBChild.LBound To mnuDBChild.UBound
            mnuDBChild(iCount).Enabled = True
        Next
        ' cancel menu
        mnuDBChild(1).Enabled = False
        
        ' view menu
        For iCount = mnuViewChild.LBound To mnuViewChild.UBound
            mnuViewChild(iCount).Enabled = True
        Next
        
        ' help menu
        mnuHelpAbout.Enabled = True
        
        ' toolbar
        ' For iCount = 1 To tbToolbar.Buttons.Count - 1
        For iCount = 1 To tbToolbar.Buttons.Count
            tbToolbar.Buttons(iCount).Enabled = True
        Next iCount
        tbToolbar.Buttons("Cancel").Enabled = False
        If bLogFile = True Then
            tbToolbar.Buttons("Log").Value = tbrPressed
        Else
            tbToolbar.Buttons("Log").Value = tbrUnpressed
        End If
              
        ' now we need to find out which items we should disable
        sControl = Me.ActiveControl.Name
        Select Case sControl
            Case "rtfSQL"
                ' if there is no data we can paste, then just set paste = false
                If Clipboard.GetFormat(vbCFText) = False And Clipboard.GetFormat(vbCFRTF) = False Then
                    mnuEditChild(7).Enabled = False
                    tbToolbar.Buttons("Paste").Enabled = False
                End If
                ' disable find next if necessary
                If rtfSQL.FindText = "" Then
                    mnuEditChild(15).Enabled = False
                End If
                lblGrd.ForeColor = vbInactiveCaptionText
                lblSQL.ForeColor = vbTitleBarText
                mnuViewChild(5).Enabled = False
                picGrd.BackColor = vbInactiveTitleBar
                picSQL.BackColor = vbActiveTitleBar
                
                
            Case "grdSQL"
                For iCount = mnuEditChild.LBound To mnuEditChild.UBound
                    mnuEditChild(iCount).Enabled = False
                Next
                lblGrd.ForeColor = vbTitleBarText
                lblSQL.ForeColor = vbInactiveCaptionText
                mnuEditChild(6).Enabled = True
                mnuEditChild(7).Enabled = True
                mnuViewChild(5).Enabled = False
                picGrd.BackColor = vbActiveTitleBar
                picSQL.BackColor = vbInactiveTitleBar
                tbToolbar.Buttons("Find").Enabled = False
                tbToolbar.Buttons("Replace").Enabled = False
                tbToolbar.Buttons("Cut").Enabled = False
                tbToolbar.Buttons("Undo").Enabled = False
                tbToolbar.Buttons("Redo").Enabled = False
                tbToolbar.Buttons("Indent").Enabled = False
                tbToolbar.Buttons("Outdent").Enabled = False
                tbToolbar.Buttons("BookToggle").Enabled = False
                tbToolbar.Buttons("BookNext").Enabled = False
                tbToolbar.Buttons("BookPrevious").Enabled = False
                tbToolbar.Buttons("BookDel").Enabled = False
                tbToolbar.Buttons("Comment").Enabled = False
                tbToolbar.Buttons("Uncomment").Enabled = False
                
            Case "tvDB"
                For iCount = mnuEditChild.LBound To mnuEditChild.UBound
                    mnuEditChild(iCount).Enabled = False
                Next
                lblGrd.ForeColor = vbInactiveCaptionText
                lblSQL.ForeColor = vbInactiveCaptionText
                mnuEditChild(6).Enabled = True
                mnuEditChild(7).Enabled = True
                mnuViewChild(0).Enabled = False
                mnuViewChild(1).Enabled = False
                picGrd.BackColor = vbInactiveTitleBar
                picSQL.BackColor = vbInactiveTitleBar
                tbToolbar.Buttons("Find").Enabled = False
                tbToolbar.Buttons("Replace").Enabled = False
                tbToolbar.Buttons("SortAscending").Enabled = False
                tbToolbar.Buttons("SortDescending").Enabled = False
                tbToolbar.Buttons("Cut").Enabled = False
                tbToolbar.Buttons("Undo").Enabled = False
                tbToolbar.Buttons("Redo").Enabled = False
                tbToolbar.Buttons("Indent").Enabled = False
                tbToolbar.Buttons("Outdent").Enabled = False
                tbToolbar.Buttons("BookToggle").Enabled = False
                tbToolbar.Buttons("BookNext").Enabled = False
                tbToolbar.Buttons("BookPrevious").Enabled = False
                tbToolbar.Buttons("BookDel").Enabled = False
                tbToolbar.Buttons("Comment").Enabled = False
                tbToolbar.Buttons("Uncomment").Enabled = False
                
            Case "grdHistory"
                For iCount = mnuEditChild.LBound To mnuEditChild.UBound
                    mnuEditChild(iCount).Enabled = False
                Next
                lblGrd.ForeColor = vbInactiveCaptionText
                lblSQL.ForeColor = vbInactiveCaptionText
                mnuEditChild(6).Enabled = True
                mnuEditChild(7).Enabled = True
                mnuEditChild(8).Enabled = True
                mnuViewChild(0).Enabled = False
                mnuViewChild(1).Enabled = False
                mnuViewChild(5).Enabled = False
                picGrd.BackColor = vbInactiveTitleBar
                picSQL.BackColor = vbInactiveTitleBar
                tbToolbar.Buttons("Find").Enabled = False
                tbToolbar.Buttons("Replace").Enabled = False
                tbToolbar.Buttons("SortAscending").Enabled = False
                tbToolbar.Buttons("SortDescending").Enabled = False
                tbToolbar.Buttons("Cut").Enabled = False
                tbToolbar.Buttons("Undo").Enabled = False
                tbToolbar.Buttons("Redo").Enabled = False
                tbToolbar.Buttons("Indent").Enabled = False
                tbToolbar.Buttons("Outdent").Enabled = False
                tbToolbar.Buttons("BookToggle").Enabled = False
                tbToolbar.Buttons("BookNext").Enabled = False
                tbToolbar.Buttons("BookPrevious").Enabled = False
                tbToolbar.Buttons("BookDel").Enabled = False
                tbToolbar.Buttons("Comment").Enabled = False
                tbToolbar.Buttons("Uncomment").Enabled = False
                
            Case "lstStatus"
                For iCount = mnuEditChild.LBound To mnuEditChild.UBound
                    mnuEditChild(iCount).Enabled = False
                Next
                lblGrd.ForeColor = vbInactiveCaptionText
                lblSQL.ForeColor = vbInactiveCaptionText
                mnuViewChild(0).Enabled = False
                mnuViewChild(1).Enabled = False
                mnuViewChild(5).Enabled = False
                picGrd.BackColor = vbInactiveTitleBar
                picSQL.BackColor = vbInactiveTitleBar
                tbToolbar.Buttons("Copy").Enabled = False
                tbToolbar.Buttons("Paste").Enabled = False
                tbToolbar.Buttons("Find").Enabled = False
                tbToolbar.Buttons("Replace").Enabled = False
                tbToolbar.Buttons("SortAscending").Enabled = False
                tbToolbar.Buttons("SortDescending").Enabled = False
                tbToolbar.Buttons("Cut").Enabled = False
                tbToolbar.Buttons("Undo").Enabled = False
                tbToolbar.Buttons("Redo").Enabled = False
                tbToolbar.Buttons("Indent").Enabled = False
                tbToolbar.Buttons("Outdent").Enabled = False
                tbToolbar.Buttons("BookToggle").Enabled = False
                tbToolbar.Buttons("BookNext").Enabled = False
                tbToolbar.Buttons("BookPrevious").Enabled = False
                tbToolbar.Buttons("BookDel").Enabled = False
                tbToolbar.Buttons("Comment").Enabled = False
                tbToolbar.Buttons("Uncomment").Enabled = False
            
        End Select
        
        ' if the file or command window has not changed, we don't need to show save
        If Not rtfSQL.Modified Then
            mnuFileChild(3).Enabled = False
            'mnuFileChild(4).Enabled = False
            tbToolbar.Buttons(3).Enabled = False
        End If
        If bHighlight = True Then
            rtfSQL.HighlightedLine = rtfSQL.GetSel(True).EndLineNo
        Else
            rtfSQL.HighlightedLine = -1
        End If
        ' just to keep it in sync
        sbStatus.Panels("Line").Text = "Ln: " & rtfSQL.GetSel(True).EndLineNo + 1 & "     "
        sbStatus.Panels("Col").Text = "Col: " & rtfSQL.GetSel(True).EndColNo + 1 & "     "

End Select
DoEvents
On Error GoTo 0
MenuSet = True
End Function

Public Sub ControlSize()
On Error Resume Next
Dim sglWidth As Single
Dim sglLimit As Single
Dim sglHeight As Single
Dim sglLeft As Single

' some initial values
sglHeight = Me.ScaleHeight - 720  'so we also have borders
sglWidth = Me.ScaleWidth - 10

    If Me.WindowState <> vbMinimized Then
        If Me.Height < 2295 Then Me.Height = 2295
        'set up the frSplitVertical splitter
        frSplitVertical.Height = sglHeight - 80
        sglLimit = sglWidth - sglSplitLimit
        If frSplitVertical.Left > sglLimit Then
           frSplitVertical.Left = sglLimit
        End If
        If frSplitVertical.Left < 100 Then frSplitVertical.Left = 100
    End If
    ' set up our control frame heights and widths
    Select Case UCase(sListLocation)
        Case "LEFT"
            frList.Left = 0
            frList.Width = frSplitVertical.Left - 20
            sglLeft = frList.Width + 30
            sglWidth = sglWidth - frList.Width - 20
        Case Else
            frList.Left = frSplitVertical.Left + 50
            frList.Width = sglWidth - frList.Left - 21
            sglWidth = frSplitVertical.Left + 20
            sglLeft = 0
    End Select
            
    frList.Height = sglHeight
    
    ' set up the controls in frList
    With tbsList
        .Top = 8
        .Left = 10
        .Width = frList.Width - 10
        .Height = frList.Height - 69
    End With
    
    With lstStatus
'        .Top = 400
        .Top = 100
        .Left = 100
        .Width = frList.Width - 200
        .Height = frList.Height - 570
    End With
    
    With tvDB
'        .Top = 400
        .Top = 100
        .Left = 100
        .Width = frList.Width - 200
        .Height = frList.Height - 570
    End With
    
            
            
    With grdHistory
'        .Top = 400
        .Top = 100
        .Left = 100
        .Width = frList.Width - 200
        .Height = frList.Height - 570
        .ColWidth(0) = .Width
    End With
    
    
    ' set up the controls in frSQL
    frSQL.Left = sglLeft
    frSQL.Height = sglHeight
    frSQL.Width = sglWidth
    
    'set up the splitter in frSQL
    sglLimit = frSQL.Height - sglSplitLimit
    If frSplitHorizontal.Top > sglLimit Then
       frSplitHorizontal.Top = sglLimit
    End If
    If frSplitHorizontal.Top < sglSplitLimit Then frSplitHorizontal.Top = sglSplitLimit
    
    ' set up the controls in frSQL
    sglWidth = frSQL.Width - 10
    grdSQL.Width = sglWidth
    rtfSQL.Width = sglWidth
    frSplitHorizontal.Width = sglWidth - 80
    grdSQL.Height = frSplitHorizontal.Top - grdSQL.Top
    picSQL.Top = frSplitHorizontal.Top + frSplitHorizontal.Height - 30
    rtfSQL.Top = picSQL.Top + picSQL.Height + 64
    rtfSQL.Height = frSQL.Height - frSplitHorizontal.Top - 60 - picSQL.Height - 64
    picGrd.Left = 25
    picGrd.Width = sglWidth - (picGrd.Left * 2) - 12
    picSQL.Left = 25
    picSQL.Width = sglWidth - (picSQL.Left * 2) - 12
    
    ' set up the run sql window
    cmRun.Top = rtfSQL.Top
    cmRun.Left = rtfSQL.Left
    cmRun.Height = rtfSQL.Height
    cmRun.Width = rtfSQL.Width
    
End Sub

Private Function FileNew() As Boolean
Dim iRetVal As Integer
Dim sMsg As String

On Error GoTo AppError
FileNew = False
If bDebug Then DebugWrite "FileNew Starts"

If rtfSQL.Modified = True Then
    If sFile <> "" Or GetSetting(App.Title, "Options", "SaveNewOnExit", "1") = "1" Then
        'do we want to save
        iRetVal = MsgBox("File Changed.  Save Changes?", vbDefaultButton1 + vbYesNoCancel + vbQuestion)
        'if yes, save
        If iRetVal = vbYes Then
            If FileSaveAs <> True Then Exit Function
        ElseIf iRetVal = vbCancel Then
            FileNew = False
            Exit Function
        End If
        If bDebug Then DebugWrite "Save Changes = " & Str(iRetVal)
    End If
End If

sFile = ""
rtfSQL.Text = ""
rtfSQL.Modified = False
SetCaption
MenuSet 1

If bDebug Then DebugWrite "FileNew Ends"
FileNew = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    FileNew = False
    Exit Function
End Function
Private Function GridCopy() As Boolean
Dim sText As String, sMsg As String
Dim lCount As Long

On Error GoTo AppError
GridCopy = False
If bDebug Then DebugWrite "GridCopy Starts"
If Not bProcessing Then
    bProcessing = True
    Screen.MousePointer = vbHourglass
    MenuSet 0
    StatsDisplay 1

    sText = grdSQL.Clip
    'replace carriagereturn with carriagereturnlinefeed
    lCount = 1
    Do
        lCount = InStr(lCount, sText, vbCr)
        If lCount > 0 Then
            ' make sure that this is not a crlf already
            If Mid$(sText, lCount + 1) <> vbLf Then
                sText = Left$(sText, lCount - 1) & vbCrLf & Mid$(sText, lCount + 1)
            End If
            lCount = lCount + 1
        End If
        If lCount Mod 10 = 0 Then
            DoEvents
            If bCancel Then Exit Do
        End If
    Loop While lCount > 0
    If Not bCancel Then
        Clipboard.Clear ' Clear Clipboard.
        Clipboard.SetText sText    ' Put text on Clipboard.
    End If

    Screen.MousePointer = vbDefault
    MenuSet 1
    StatsDisplay 0
    bCancel = False
    bProcessing = False

End If

GridCopy = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    MenuSet 1
    bCancel = False
    bProcessing = False
    GridCopy = False
    Exit Function

End Function

Private Function GridPicture(iCase) As Boolean
Dim lThisColumn As Long
Dim sMsg As String

On Error GoTo AppError
GridPicture = False

Select Case iCase
    Case 0
        If Not bProcessing And Not bPictureCleared And grdSQL.Col <> 0 Then
            bPictureCleared = True
            grdSQL.Col = 0
            Set grdSQL.CellPicture = Nothing
            bPictureCleared = False
        End If
    Case 1
        If Not bProcessing And Not bPictureAdded And grdSQL.Col <> 0 Then
            bPictureAdded = True
            lThisColumn = grdSQL.Col
            grdSQL.Col = 0
            Set grdSQL.CellPicture = imlToolbarIcons.ListImages(12).Picture
            grdSQL.Col = lThisColumn
            bPictureAdded = False
        End If
End Select
GridPicture = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    GridPicture = False
    Exit Function

End Function

Private Function SetTreeMenu() As Boolean
Dim sMyKey As String
Dim sMyPrefix As String
Dim sMyChild As String
Dim sMsg As String

If bDebug Then DebugWrite "SetTreeMenu Starts"

SetTreeMenu = False

On Error GoTo AppError

sMyKey = tvDB.SelectedItem.Key
sMyPrefix = Left$(sMyKey, InStr(sMyKey, "_"))

Select Case sMyPrefix
    
    Case "t_", "v_"
        ' set up the child
        sMyChild = "fc_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
        ' does the table or view have columns available
        If tvDB.Nodes(sMyChild).Children > 0 Then
            mnuPopTreeObjectCopyColumn.Enabled = True
            mnuPopTreeObjectCopyColumn.Visible = True
            mnuPopTreeObjectPasteColumn.Enabled = True
            mnuPopTreeObjectPasteColumn.Visible = True
        Else
            ' no, only copy
            mnuPopTreeObjectCopyColumn.Enabled = False
            mnuPopTreeObjectCopyColumn.Visible = False
            mnuPopTreeObjectCopyParameter.Enabled = False
            mnuPopTreeObjectCopyParameter.Visible = False
            mnuPopTreeObjectPasteColumn.Enabled = False
            mnuPopTreeObjectPasteColumn.Visible = False
        End If
        
        mnuPopTreeObjectCopyColumnSelect.Enabled = True
        mnuPopTreeObjectCopyColumnSelect.Visible = True
        mnuPopTreeObjectCopyColumnInsert.Enabled = True
        mnuPopTreeObjectCopyColumnInsert.Visible = True
        mnuPopTreeObjectPasteColumnSelect.Enabled = True
        mnuPopTreeObjectPasteColumnSelect.Visible = True
        mnuPopTreeObjectPasteColumnInsert.Enabled = True
        mnuPopTreeObjectPasteColumnInsert.Visible = True
        mnuPopTreeObjectCopyColumnUpdate.Enabled = True
        mnuPopTreeObjectCopyColumnUpdate.Visible = True
        mnuPopTreeObjectCopyColumnDelete.Enabled = True
        mnuPopTreeObjectCopyColumnDelete.Visible = True
        mnuPopTreeObjectPasteColumnUpdate.Enabled = True
        mnuPopTreeObjectPasteColumnUpdate.Visible = True
        mnuPopTreeObjectPasteColumnDelete.Enabled = True
        mnuPopTreeObjectPasteColumnDelete.Visible = True
        mnuPopTreeObjectCopyParameter.Enabled = False
        mnuPopTreeObjectCopyParameter.Visible = False
        mnuPopTreeObjectCopyParameterExecute.Enabled = False
        mnuPopTreeObjectCopyParameterExecute.Visible = False
        mnuPopTreeObjectPasteParameterExecute.Enabled = False
        mnuPopTreeObjectPasteParameterExecute.Visible = False
        mnuPopTreeObjectPasteParameter.Enabled = False
        mnuPopTreeObjectPasteParameter.Visible = False
        mnuPopTreeObjectCopyDDL.Enabled = True
        mnuPopTreeObjectCopyDDL.Visible = True
        mnuPopTreeObjectPasteDDL.Enabled = True
        mnuPopTreeObjectPasteDDL.Visible = True
        
    Case "p_"
        ' set up the child
        sMyChild = "fa_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
        
        ' does the procedure have parameters available
        If tvDB.Nodes(sMyChild).Children > 0 Then
            mnuPopTreeObjectCopyParameter.Enabled = True
            mnuPopTreeObjectCopyParameter.Visible = True
            mnuPopTreeObjectPasteParameter.Enabled = True
            mnuPopTreeObjectPasteParameter.Visible = True
        Else
            ' no, only copy
            mnuPopTreeObjectCopyParameter.Enabled = False
            mnuPopTreeObjectCopyParameter.Visible = False
            mnuPopTreeObjectPasteParameter.Enabled = False
            mnuPopTreeObjectPasteParameter.Visible = False
        End If
        
        mnuPopTreeObjectCopyParameterExecute.Enabled = True
        mnuPopTreeObjectCopyParameterExecute.Visible = True
        mnuPopTreeObjectPasteParameterExecute.Enabled = True
        mnuPopTreeObjectPasteParameterExecute.Visible = True
        mnuPopTreeObjectCopyColumn.Enabled = False
        mnuPopTreeObjectCopyColumn.Visible = False
        mnuPopTreeObjectPasteColumn.Enabled = False
        mnuPopTreeObjectPasteColumn.Visible = False
        mnuPopTreeObjectCopyColumnSelect.Enabled = False
        mnuPopTreeObjectCopyColumnSelect.Visible = False
        mnuPopTreeObjectCopyColumnInsert.Enabled = False
        mnuPopTreeObjectCopyColumnInsert.Visible = False
        mnuPopTreeObjectPasteColumnSelect.Enabled = False
        mnuPopTreeObjectPasteColumnSelect.Visible = False
        mnuPopTreeObjectPasteColumnInsert.Enabled = False
        mnuPopTreeObjectPasteColumnInsert.Visible = False
        mnuPopTreeObjectCopyColumnUpdate.Enabled = False
        mnuPopTreeObjectCopyColumnUpdate.Visible = False
        mnuPopTreeObjectCopyColumnDelete.Enabled = False
        mnuPopTreeObjectCopyColumnDelete.Visible = False
        mnuPopTreeObjectPasteColumnUpdate.Enabled = False
        mnuPopTreeObjectPasteColumnUpdate.Visible = False
        mnuPopTreeObjectPasteColumnDelete.Enabled = False
        mnuPopTreeObjectPasteColumnDelete.Visible = False
        mnuPopTreeObjectCopyDDL.Enabled = True
        mnuPopTreeObjectCopyDDL.Visible = True
        mnuPopTreeObjectPasteDDL.Enabled = True
        mnuPopTreeObjectPasteDDL.Visible = True
        
    Case "i_"
        ' does the index have columns available
        If tvDB.Nodes(sMyKey).Children > 0 Then
            mnuPopTreeObjectCopyColumn.Enabled = True
            mnuPopTreeObjectCopyColumn.Visible = True
            mnuPopTreeObjectPasteColumn.Enabled = True
            mnuPopTreeObjectPasteColumn.Visible = True
        Else
            ' no, only copy
            mnuPopTreeObjectCopyColumn.Enabled = False
            mnuPopTreeObjectCopyColumn.Visible = False
            mnuPopTreeObjectPasteColumn.Enabled = False
            mnuPopTreeObjectPasteColumn.Visible = False
        End If
        
        mnuPopTreeObjectCopyColumnSelect.Enabled = False
        mnuPopTreeObjectCopyColumnSelect.Visible = False
        mnuPopTreeObjectCopyColumnInsert.Enabled = False
        mnuPopTreeObjectCopyColumnInsert.Visible = False
        mnuPopTreeObjectCopyParameter.Enabled = False
        mnuPopTreeObjectCopyParameter.Visible = False
        mnuPopTreeObjectPasteColumnSelect.Enabled = False
        mnuPopTreeObjectPasteColumnSelect.Visible = False
        mnuPopTreeObjectPasteColumnInsert.Enabled = False
        mnuPopTreeObjectPasteColumnInsert.Visible = False
        mnuPopTreeObjectCopyColumnUpdate.Enabled = False
        mnuPopTreeObjectCopyColumnUpdate.Visible = False
        mnuPopTreeObjectCopyColumnDelete.Enabled = False
        mnuPopTreeObjectCopyColumnDelete.Visible = False
        mnuPopTreeObjectPasteColumnUpdate.Enabled = False
        mnuPopTreeObjectPasteColumnUpdate.Visible = False
        mnuPopTreeObjectPasteColumnDelete.Enabled = False
        mnuPopTreeObjectPasteColumnDelete.Visible = False
        mnuPopTreeObjectCopyParameterExecute.Enabled = False
        mnuPopTreeObjectCopyParameterExecute.Visible = False
        mnuPopTreeObjectPasteParameterExecute.Enabled = False
        mnuPopTreeObjectPasteParameterExecute.Visible = False
        mnuPopTreeObjectPasteParameter.Enabled = False
        mnuPopTreeObjectPasteParameter.Visible = False
        mnuPopTreeObjectCopyDDL.Enabled = True
        mnuPopTreeObjectCopyDDL.Visible = True
        mnuPopTreeObjectPasteDDL.Enabled = True
        mnuPopTreeObjectPasteDDL.Visible = True
        
    Case Else
'        Case "c_", "a_", "u_", "ic_", "tr_"
        ' just a column, only copy
            mnuPopTreeObjectCopyColumn.Enabled = False
            mnuPopTreeObjectCopyColumn.Visible = False
            mnuPopTreeObjectCopyColumnSelect.Enabled = False
            mnuPopTreeObjectCopyColumnSelect.Visible = False
            mnuPopTreeObjectCopyColumnInsert.Enabled = False
            mnuPopTreeObjectCopyColumnInsert.Visible = False
            mnuPopTreeObjectCopyColumnUpdate.Enabled = False
            mnuPopTreeObjectCopyColumnUpdate.Visible = False
            mnuPopTreeObjectCopyColumnDelete.Enabled = False
            mnuPopTreeObjectCopyColumnDelete.Visible = False
            mnuPopTreeObjectPasteColumnUpdate.Enabled = False
            mnuPopTreeObjectPasteColumnUpdate.Visible = False
            mnuPopTreeObjectPasteColumnDelete.Enabled = False
            mnuPopTreeObjectPasteColumnDelete.Visible = False
            mnuPopTreeObjectCopyParameter.Enabled = False
            mnuPopTreeObjectCopyParameter.Visible = False
            mnuPopTreeObjectPasteColumn.Enabled = False
            mnuPopTreeObjectPasteColumn.Visible = False
            mnuPopTreeObjectPasteParameter.Enabled = False
            mnuPopTreeObjectPasteColumnSelect.Enabled = False
            mnuPopTreeObjectPasteColumnSelect.Visible = False
            mnuPopTreeObjectPasteColumnInsert.Enabled = False
            mnuPopTreeObjectCopyParameterExecute.Enabled = False
            mnuPopTreeObjectCopyParameterExecute.Visible = False
            mnuPopTreeObjectPasteParameterExecute.Enabled = False
            mnuPopTreeObjectPasteParameterExecute.Visible = False
            mnuPopTreeObjectPasteColumnInsert.Visible = False
            mnuPopTreeObjectPasteParameter.Visible = False
        
        If sMyPrefix = "tr_" Then
            mnuPopTreeObjectCopyDDL.Enabled = True
            mnuPopTreeObjectCopyDDL.Visible = True
            mnuPopTreeObjectPasteDDL.Enabled = True
            mnuPopTreeObjectPasteDDL.Visible = True
        Else
            mnuPopTreeObjectCopyDDL.Enabled = False
            mnuPopTreeObjectCopyDDL.Visible = False
            mnuPopTreeObjectPasteDDL.Enabled = False
            mnuPopTreeObjectPasteDDL.Visible = False
        End If
End Select

If bDebug Then DebugWrite "SetTreeMenu Ends"

SetTreeMenu = True

Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    Exit Function

End Function

'Private Function RTFCopy() As Boolean
'Dim sText As String
'    sText = rtfSQL.SelText
'    If sText <> "" Then
'        Clipboard.Clear
'        Clipboard.SetText sText
'    End If
'    RTFCopy = True
'End Function
Public Function StatsAdd(sMsg As String) As Boolean
    lstStatus.AddItem sMsg
    DoEvents
    If lstStatus.ListCount > 2000 Then lstStatus.RemoveItem 0
    lstStatus.ListIndex = lstStatus.ListCount - 1
    DoEvents
End Function
Public Function StatsDisplay(iCase As Integer, Optional lCount As Long) As Boolean
Dim sMsg As String
    Select Case iCase
        Case 0
            sMsg = "Ready"
        Case 1
            sMsg = "Please Wait..."
        Case 2
            sMsg = "Rows Affected: " & Str(lCount)
        Case 3
            sMsg = "Nothing To Do"
        Case 4
            sMsg = "Cancelling...Please Wait..."
        Case 5
            sMsg = "Error..."
        Case 6
            sMsg = "Connecting..."
        Case 7
            sMsg = "Not Connected"
        Case 8
            sMsg = "WARNING:  Max # Grid Cells Reached.  Rows Affected: " & Str(lCount)
    End Select

    sbStatus.Panels("Main").Text = sMsg
    StatsAdd sMsg
    StatsDisplay = True
    DoEvents
End Function
Private Function FileOpen(bHaveFileName As Boolean) As Boolean
Dim iRetVal As Integer
Dim sMsg As String

On Error GoTo AppError
FileOpen = False
If bDebug Then DebugWrite "FileOpen Starts"

If rtfSQL.Modified Then
    If sFile <> "" Or GetSetting(App.Title, "Options", "SaveNewOnExit", "1") = "1" Then
        'do we want to save
        iRetVal = MsgBox("File Changed.  Save Changes?", vbDefaultButton1 + vbYesNoCancel + vbQuestion)
        'if yes, save
        If iRetVal = vbYes Then
            If FileSaveAs <> True Then Exit Function
        ElseIf iRetVal = vbCancel Then
            FileOpen = False
            Exit Function
        End If
        If bDebug Then DebugWrite "Save Changes = " & Str(iRetVal)
    End If
End If

' do we have a filename already.  only used at startup
If bHaveFileName = False Then
    If bDebug Then DebugWrite "Get FileName"
    With dlgCommonDialog
        .DialogTitle = "Open"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .FileName = ""
        .CancelError = True
        On Error GoTo ErrHandler
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Function
        sFile = .FileName
    End With
End If

On Error GoTo AppError

Screen.MousePointer = vbHourglass
If bDebug Then DebugWrite "Filename = " & sFile

Call rtfSQL.OpenFile(sFile)

SetCaption

Call rtfSQL.SetCaretPos(0, 0)

Screen.MousePointer = vbDefault
rtfSQL.Modified = False
If bDebug Then DebugWrite "FileOpen Ends"
FileOpen = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        If sFile <> "" And bHaveFileName = True Then
            sMsg = sMsg & sFile
            sFile = ""
        End If
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    FileOpen = False
    Exit Function

ErrHandler:
    'User pressed the Cancel button
    If bDebug Then DebugWrite "FileOpen Cancel"
    FileOpen = False
    Exit Function

End Function

Private Function FileSaveAs(Optional bForce As Boolean) As Boolean
Dim sMsg As String
Dim sBakFile As String
Dim iLength As Integer, iDotPos As Integer, iLastPos As Integer

On Error GoTo AppError
FileSaveAs = False
    
If bDebug Then DebugWrite "FileSaveAs Starts"
'If rtfSQL.Modified = True Then
    If sFile = "" Or bForce = True Then
        With dlgCommonDialog
            .DialogTitle = "Save As"
            .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
            .FileName = sFile
            .CancelError = True
            On Error GoTo ErrHandler
            .ShowSave
            sFile = .FileName
            If Len(.FileName) = 0 Then Exit Function
        End With
        On Error GoTo AppError
        If bDebug Then DebugWrite "Filename = " & sFile
    End If
'End If

' do we need to make a backup file
If GetSetting(App.Title, "Options", "MakeBakFileOnSave", "0") = "1" Then
    iDotPos = 0
    iLastPos = 0
    iLength = 0
    ' parse the extension from the file, and change to .bak
    For iLength = 1 To Len(sFile)
        iDotPos = InStr(iLength, sFile, ".", 1)
        If iDotPos > iLastPos And iDotPos < Len(sFile) Then
            iLastPos = iDotPos
        Else
            iDotPos = iLastPos
            Exit For
        End If
        iLength = iDotPos
    Next iLength
    If iDotPos = 0 Then
        ' no extension
        sBakFile = sFile & ".bak"
    Else
        sBakFile = Left(sFile, iDotPos) & "bak"
    End If
    ' does the original file exist
    If Dir(sFile) <> "" Then
        ' ok, the original file exists, does the backup
        If Dir(sBakFile) <> "" Then
            ' ok, the backup exists, we need to delete it
            Kill (sBakFile)
        End If
        ' now we need to rename the original
        Name sFile As sBakFile
    End If
End If

Call rtfSQL.SaveFile(sFile, False)

SetCaption

rtfSQL.Modified = False
If bDebug Then DebugWrite "FileSaveAs Ends"
FileSaveAs = True
Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    FileSaveAs = False
    Exit Function

ErrHandler:
    'User pressed the Cancel button
    If bDebug Then DebugWrite "FileSaveAs Cancel"
    FileSaveAs = False
    Exit Function

End Function

Private Function GridSave(Optional bView As Boolean) As Boolean
Dim iColCount As Integer
Dim lCount As Long
Dim lCurrentRow As Long, lOrigRow As Long
Dim lRowCount As Long, lOrigCol As Long
Dim lOutFile As Long
Dim sText As String, sOutFile As String
Dim sProgram As String

GridSave = False

sProgram = GetSetting(App.Title, "Options", "DefaultEditor", "Notepad")
If bDebug Then
    DebugWrite "GridSave Starts"
    DebugWrite "bView = " & bView
    DebugWrite "Editor = " & sProgram
End If

    With dlgCommonDialog
        .DialogTitle = "Export Results"
        .Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .FileName = ""
        .FilterIndex = 4
        sText = .DefaultExt
        .DefaultExt = "*.txt"
        .CancelError = True
        On Error GoTo NoFileError
        .ShowSave
        .DefaultExt = sText
        .FilterIndex = 1
        sOutFile = .FileName
        If Len(.FileName) = 0 Then Exit Function
    End With

    If bDebug Then DebugWrite "Filename = " & sOutFile

    On Error GoTo WriteFileError
    
    Screen.MousePointer = vbHourglass
    bProcessing = True
    MenuSet 0
    DoEvents
    rtfSQL.SetFocus
    DoEvents
    DoEvents
    grdSQL.Redraw = False
    
    StatsDisplay 1
    sText = ""
    With grdSQL
        lOrigRow = .Row
        lOrigCol = .Col
        iColCount = .Cols
        lRowCount = .Rows
        .Col = 1
        .Row = 0
    End With


    If (iColCount >= 2) And grdSQL.Text <> "" Then
        ' make sure we don't append to an existing file
        If Dir(sOutFile) <> "" Then
            Kill (sOutFile)
        End If

        ' get the file numbers we will use
        lOutFile = FreeFile(0)

        ' open the output file
        Open sOutFile For Output As #lOutFile

        ' print the headers
        For lCount = 0 To iColCount - 2
            With grdSQL
                .Col = lCount + 1
                .Row = 0
                sText = sText & "  " & .Text & _
                Space(iGridCol(lCount, 0) - Len(.Text))
            End With
        Next
        If Trim(sText) <> "" Then
            Print #lOutFile, sText
        End If
        sText = ""

        ' print the seperator line
        For lCount = 0 To iColCount - 2
            With grdSQL
                .Col = lCount + 1
                sText = sText & "  " & String(iGridCol(lCount, 0), 61)
            End With
        Next
        If Trim(sText) <> "" Then
            Print #lOutFile, sText
        End If

        'Print the grid
        If lRowCount > 1 Then
            grdSQL.Row = 1
            grdSQL.Col = 1
            For lCurrentRow = 1 To lRowCount - 1
                On Error Resume Next
                sText = ""
                For lCount = 0 To iColCount - 1
                    With grdSQL
                        grdSQL.Row = lCurrentRow
                        .Col = lCount + 1
                        If (iGridCol(lCount, 1) = rdTypeVARCHAR) _
                        Or (iGridCol(lCount, 1) = rdTypeLONGVARCHAR) _
                        Or (iGridCol(lCount, 1) = rdTypeTIMESTAMP) _
                        Or (iGridCol(lCount, 1) = rdTypeTIME) _
                        Or (iGridCol(lCount, 1) = rdTypeDATE) _
                        Or (iGridCol(lCount, 1) = rdTypeCHAR) Then
                            If iGridCol(lCount, 0) >= Len(.Text) Then
                                sText = sText & "  " & .Text & _
                                    Space(iGridCol(lCount, 0) - Len(.Text))
                            Else
                                sText = sText & "  " & Left$(.Text, iGridCol(lCount, 0))
                            End If
                        Else
                            If iGridCol(lCount, 0) >= Len(.Text) Then
                                sText = sText & "  " & _
                                    Space(iGridCol(lCount, 0) - Len(.Text)) & _
                                    .Text
                            Else
                                sText = sText & "  " & Left$(.Text, iGridCol(lCount, 0))
                            End If
                        End If
                    End With
                Next
                If Trim(sText) <> "" Then
                    Print #lOutFile, sText
                End If
                If lCurrentRow Mod 100 = 0 Then
                    DoEvents
                    If bCancel Then
                        bCancel = False
                        Exit For
                    End If
                End If
            Next
        End If

        ' close output file
        Close #lOutFile

        If bView = True Then Shell sProgram & " " & sOutFile, vbNormalFocus

    End If

    bProcessing = False
    grdSQL.Redraw = True
    With grdSQL
        .Row = lOrigRow
        .Col = lOrigCol
    End With
    MenuSet 1
    StatsDisplay 0
    Screen.MousePointer = vbDefault
    GridSave = True
If bDebug Then DebugWrite "GridSave Ends"
Exit Function

NoFileError:
    'User pressed the Cancel button
    dlgCommonDialog.DefaultExt = sText
    dlgCommonDialog.FilterIndex = 4
    GridSave = False
    bProcessing = False
    MenuSet 1
    grdSQL.Redraw = True
    If bDebug Then DebugWrite "GridSave Cancel"
    Exit Function

WriteFileError:
    Dim sMsg As String
    StatsDisplay 5
    ' close output file
    Close #lOutFile
    bProcessing = False
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    StatsDisplay 5
    MenuSet 1
    grdSQL.Visible = True
    grdSQL.Redraw = True
    With grdSQL
        .Row = lOrigRow
        .Col = lOrigCol
    End With
    Screen.MousePointer = vbDefault
    GridSave = False
    If bDebug Then DebugWrite "GridSave Error"
    Exit Function

End Function

Private Sub PrintSetup()
Dim iPrintOrient As Integer

iPrintOrient = CInt(GetSetting(App.Title, "Options", "PrintOrient", "1"))

With dlgCommonDialog
    '.DialogTitle = "Print Setup"
    If iPrintOrient = 1 Then
        .Orientation = cdlPortrait
    Else
        .Orientation = cdlLandscape
    End If
    .Flags = cdlPDPrintSetup
    .CancelError = True
    On Error GoTo ErrHandler
    .ShowPrinter
    If .Orientation = cdlPortrait Then
        SaveSetting App.Title, "Options", "PrintOrient", "1"
    Else
        SaveSetting App.Title, "Options", "PrintOrient", "2"
    End If
End With

Exit Sub

ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub

Public Function GridDisplay() As Boolean
Dim lCount As Long
Dim iColCount As Integer
Dim lRowCount As Long
Dim iDefaultRows As Integer
Dim iIncrementRows As Integer
Dim lMaxRows As Long
Dim lCurrentRow As Long

If bDebug Then DebugWrite "GridDisplay Starts"

    grdSQL.Visible = False
    DoEvents
    If Not rsSQL.BOF And Not rsSQL.EOF Then
        If bCancel = True Then
            StatsDisplay 0
            GridDisplay = True
            grdSQL.Visible = True
            Exit Function
        End If
    End If

    iColCount = rsSQL.rdoColumns.Count
    iDefaultRows = 999
    iIncrementRows = 1000
    lMaxRows = (MAXDESIGNCELLS / (iColCount + 1)) - 1
    lCurrentRow = 0
    
    lRowCount = iDefaultRows

    ReDim iGridCol(iColCount, 2)

    ' size the grid
    With grdSQL
        .FixedCols = 0
        .FixedRows = 0
        .Cols = 1
        .Rows = 1
        .Cols = iColCount + 1
        If .Cols < 2 Then .Cols = 2
        .Rows = lRowCount + 1
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = .RowHeight(0)
        .Row = 0
        .HighLight = flexHighlightNever
    End With

    ' set the column headers
    For lCount = 0 To iColCount - 1
        With grdSQL
            .Row = 0
            .Col = lCount + 1
            .ColWidth(lCount + 1) = .Font.Size * 175
            .Text = rsSQL.rdoColumns(lCount).Name
        End With
        ' get the size
        If rsSQL.rdoColumns(lCount).Size <= 256 Then
            iGridCol(lCount, 0) = rsSQL.rdoColumns(lCount).Size
        Else
            iGridCol(lCount, 0) = 256
        End If
        ' increase if name is longer than data value
        If iGridCol(lCount, 0) < Len(rsSQL.rdoColumns(lCount).Name) Then
            iGridCol(lCount, 0) = Len(rsSQL.rdoColumns(lCount).Name)
        End If
        ' get the type
        iGridCol(lCount, 1) = rsSQL.rdoColumns(lCount).Type
        ' if this is a timestamp, or int type we need to set up enough room to print
        If iGridCol(lCount, 1) = rdTypeTIMESTAMP And iGridCol(lCount, 0) < (Len(sDateFormat & " " & sTimeFormat) + 2) Then
            iGridCol(lCount, 0) = (Len(sDateFormat & " " & sTimeFormat) + 2)
        ElseIf (iGridCol(lCount, 1) = rdTypeINTEGER Or iGridCol(lCount, 1) = rdTypeSMALLINT _
            Or iGridCol(lCount, 1) = rdTypeTINYINT) And iGridCol(lCount, 0) < 11 Then
            iGridCol(lCount, 0) = 11
        End If
    Next

    ' populate the grid
    grdSQL.Row = 1
    lRowCount = 0
    If bCancel = False Then
        Do While Not rsSQL.EOF
            Screen.MousePointer = vbArrowHourglass
            On Error Resume Next
            If grdSQL.Row = lMaxRows - 1 Then Exit Do
            lCurrentRow = lCurrentRow + 1
            ' populate the grid
            For lCount = 0 To iColCount - 1
                With grdSQL
                    .Row = lCurrentRow
                    .Col = lCount + 1
                    If (iGridCol(lCount, 1) = rdTypeVARCHAR) _
                        Or (iGridCol(lCount, 1) = rdTypeTIMESTAMP) _
                        Or (iGridCol(lCount, 1) = rdTypeTIME) _
                        Or (iGridCol(lCount, 1) = rdTypeDATE) _
                        Or (iGridCol(lCount, 1) = rdTypeCHAR) Then
                        .CellAlignment = flexAlignLeftCenter
                    End If
                    If (iGridCol(lCount, 1) = rdTypeTIMESTAMP) Then
                        .Text = Format(rsSQL.rdoColumns(lCount).Value, sDateFormat & " " & sTimeFormat)
                    ElseIf (iGridCol(lCount, 1) = rdTypeTIME) Then
                        .Text = Format(rsSQL.rdoColumns(lCount).Value, sTimeFormat)
                    ElseIf (iGridCol(lCount, 1) = rdTypeDATE) Then
                        .Text = Format(rsSQL.rdoColumns(lCount).Value, sDateFormat)
                    ElseIf rsSQL.rdoColumns(lCount).ChunkRequired = True Then
                        .Text = GetTextChunks(rsSQL.rdoColumns(lCount))
                    Else
                        .Text = rsSQL.rdoColumns(lCount).Value
                    End If
                End With
            Next
            On Error GoTo 0
            ' check to see if we need more rows
            rsSQL.MoveNext
            lRowCount = lRowCount + 1
            If (lRowCount + 1) = iDefaultRows Then
                If grdSQL.Rows + iIncrementRows <= lMaxRows Then
                    grdSQL.Rows = grdSQL.Rows + iIncrementRows
                    iDefaultRows = iDefaultRows + iIncrementRows
                Else
                    grdSQL.Rows = lMaxRows
                End If
            End If
            ' check to see if we should update stats
            If (rsSQL.AbsolutePosition - 1) Mod 50 = 0 Then
                grdSQL.Visible = True
                StatsDisplay 2, lRowCount
                DoEvents
                If bCancel = True Then Exit Do
            End If
        Loop
    End If
    
    grdSQL.Rows = lRowCount + 1
    If grdSQL.Rows < lMaxRows Then
        StatsDisplay 2, lRowCount
    Else
        StatsDisplay 8, lRowCount
    End If
    If grdSQL.Rows > 1 Then
        grdSQL.HighLight = flexHighlightAlways
    Else
        grdSQL.HighLight = flexHighlightNever
    End If
    grdSQL.Visible = True
    GridDisplay = True
    If bDebug Then DebugWrite "GridDisplay Ends"

End Function

Public Function LoginShow() As Boolean
    Dim iCount As Integer
    If bDebug Then DebugWrite "LoginShow Starts"
    Screen.MousePointer = vbHourglass
    'If bDebug Then DebugWrite "frmODBCLogon.Hide Starts"
    'frmODBCLogon.Hide
'    If bDebug Then DebugWrite "FindSplash Starts"
'    For iCount = 0 To Forms.Count - 1
'        If Forms(iCount).Name = "frmSplash" Then
'            Unload frmSplash
'            Exit For
'        End If
'    Next
'    If bDebug Then DebugWrite "FindSplash Ends"
    Screen.MousePointer = vbDefault
    If bDebug Then DebugWrite "frmODBCLogon.Show vbModal Starts"
    frmODBCLogon.Show vbModal
    If bDebug Then DebugWrite "frmODBCLogon.Show vbModal Ends"
    LoginShow = bGlobalRetVal
    If LoginShow = False Then
        On Error Resume Next
        If envSQL.rdoConnections.Count < 1 Then
            StatsDisplay 7
            DBInfo 0
        End If
    End If
    bGlobalRetVal = False
    If bDebug Then DebugWrite "LoginShow Ends"
End Function

Private Function GridSort(iCase As Integer) As Boolean
Screen.MousePointer = vbHourglass
Dim lThisColumn As Long

    lThisColumn = grdSQL.Col

    GridPicture (0)
    bPictureAdded = True
    grdSQL.Col = lThisColumn
    Select Case iCase

        Case 0
            grdSQL.Sort = flexSortGenericAscending
        Case 1
            grdSQL.Sort = flexSortGenericDescending
    End Select
    bPictureAdded = False
    GridPicture (1)
    GridSort = True

Screen.MousePointer = vbDefault
End Function
Public Function TreeCopy(ByVal sMyKey As String, iCase As Integer) As Boolean
Dim sMsg As String
Dim iErrorCount As Integer
iErrorCount = 0

On Error GoTo AppError
    
TreeCopy = False
If bDebug Then DebugWrite "TreeCopy Starts"
Screen.MousePointer = vbHourglass

Select Case iCase
    Case Is = 1
        ' copy item
        TreeCopyItem sMyKey
        
    Case Is = 2, 3, 4, 5, 7, 8
        ' copy with children
        TreeCopyColumns sMyKey, iCase
            
    Case Is = 6
        ' copy ddl
        TreeGetDDL sMyKey
End Select
    
If bDebug Then DebugWrite "TreeCopy Ends"
Screen.MousePointer = vbDefault
TreeCopy = True
Exit Function

AppError:
    If Err.Number = 521 And iErrorCount < 5 Then
        iErrorCount = iErrorCount + 1
        Err.Clear
        Resume
    End If
    
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    TreeCopy = False
    Exit Function

End Function

Private Function FixHotKeys() As Boolean
Dim HotKey As CodeMaxCtl.HotKey, HotKeyIndex As Integer
Dim HotKeyNumber As Long, cmGlobals As CodeMaxCtl.Globals
Dim Command(31) As CodeMaxCtl.cmCommand, CommandIndex As Integer

If bDebug Then DebugWrite "FixHotKeys Starts"

FixHotKeys = False

Set cmGlobals = New CodeMaxCtl.Globals
Set HotKey = New CodeMaxCtl.HotKey


'Unregister a few of the hotkeys in the codemax control
Command(1) = cmCmdFind
Command(2) = cmCmdRedo
Command(3) = cmCmdLineCut
Command(4) = cmCmdLowercaseSelection
Command(5) = cmCmdTabifySelection
Command(6) = cmCmdUntabifySelection
Command(7) = cmCmdUppercaseSelection
Command(8) = cmCmdCutSelection
Command(9) = cmCmdFindNextWord
Command(10) = cmCmdFindPrevWord
Command(11) = cmCmdFindReplace
Command(12) = cmCmdLineCut
Command(13) = cmCmdLineOpenAbove
Command(14) = cmCmdProperties
Command(15) = cmCmdRecordMacro
Command(16) = cmCmdSelectLine
Command(17) = cmCmdSelectSwapAnchor
Command(18) = cmCmdSentenceCut
Command(19) = cmCmdSetRepeatCount
Command(20) = cmCmdTabifySelection
Command(21) = cmCmdToggleWhitespaceDisplay
Command(22) = cmCmdUntabifySelection
Command(23) = cmCmdWindowScrollDown
Command(24) = cmCmdWindowScrollLeft
Command(25) = cmCmdWindowScrollRight
Command(26) = cmCmdWindowScrollUp
Command(27) = cmCmdWordDeleteToEnd
Command(28) = cmCmdWordDeleteToStart
Command(29) = cmCmdBookmarkNext
Command(30) = cmCmdBookmarkPrev
Command(31) = cmCmdBookmarkToggle

' not going to take these out
'SentenceLeft Control + Alt + Left
'SentenceRight Control + Alt + Right

For CommandIndex = 1 To 31
    HotKeyNumber = cmGlobals.GetNumHotKeysForCmd(Command(CommandIndex))
    For HotKeyIndex = HotKeyNumber - 1 To 0 Step -1
        Set HotKey = cmGlobals.GetHotKeyForCmd(Command(CommandIndex), HotKeyIndex)
        Call cmGlobals.UnregisterHotKey(HotKey)
    Next HotKeyIndex
Next CommandIndex

If bDebug Then DebugWrite "Setting Variables to Nothing."

Set HotKey = Nothing
Set cmGlobals = Nothing

If bDebug Then DebugWrite "FixHotKeys Ends"

FixHotKeys = True

End Function

Public Function CommentBlock() As Boolean
Dim sString As String
Dim iCount As Integer
Dim aString() As String

If bDebug Then DebugWrite "CommentBlock Starts"

CommentBlock = False

sString = ""

' we need to go through the text, and insert the comment
aString = Split(rtfSQL.SelText, Chr$(10))
For iCount = 0 To UBound(aString)
    sString = sString & sComment & aString(iCount)
Next

InsertText sString

If bDebug Then DebugWrite "Setting Variables to Nothing."
Erase aString

If bDebug Then DebugWrite "CommentBlock Ends"

CommentBlock = True

End Function
Public Function UnCommentBlock() As Boolean
If bDebug Then DebugWrite "UnCommentBlock Starts"

UnCommentBlock = False

InsertText Replace(rtfSQL.SelText, sComment, "")

If bDebug Then DebugWrite "UnCommentBlock Ends"

UnCommentBlock = True

End Function

Private Sub Form_Load()
Dim lVerticalPos As Long
Dim sFontStyle As String

    Screen.MousePointer = vbArrowHourglass
    ' SetAppHelp Me.hWnd
    sDBServer = ""
    
    grdSQL.ColWidth(0) = grdSQL.RowHeight(0)
    
    ' DragAcceptFiles ByVal rtfSQL.hWnd, ByVal 0&
    bCancel = False
    bProcessing = False
    bPictureAdded = False
    bPictureCleared = False
            
    ' set up the cmax control
    rtfSQL.Font.Name = GetSetting(App.Title, "Options", "CommandFontName", "Courier New")
    rtfSQL.Font.Size = CInt(GetSetting(App.Title, "Options", "CommandFontSize", "10"))
    sFontStyle = GetSetting(App.Title, "Options", "CommandFontStyle", "Regular")
    Select Case sFontStyle
        Case Is = "Italic"
            rtfSQL.Font.Italic = True
            rtfSQL.Font.Bold = False
        Case Is = "Bold"
            rtfSQL.Font.Italic = False
            rtfSQL.Font.Bold = True
        Case Is = "Bold Italic"
            rtfSQL.Font.Italic = True
            rtfSQL.Font.Bold = True
        Case Else
            rtfSQL.Font.Bold = False
            rtfSQL.Font.Italic = False
    End Select
    ' cmax added required these
    Call rtfSQL.SetFontStyle(cmStyKeyword, cmFontNormal)
    Call rtfSQL.SetColor(cmClrLineNumberBk, 12632256)
    rtfSQL.DisplayLeftMargin = False
    
    rtfSQL.TabSize = CInt(GetSetting(App.Title, "Options", "TabSize", "4"))
    If GetSetting(App.Title, "Options", "ExpandTabs", "1") = "1" Then
        rtfSQL.ExpandTabs = True
    Else
        rtfSQL.ExpandTabs = False
    End If
        
    ' restrict cursor to text
    If GetSetting(App.Title, "Options", "RestrictCursor", "1") = "1" Then
        rtfSQL.SelBounds = True
    Else
        rtfSQL.SelBounds = False
    End If
    
    ' indent mode
    If GetSetting(App.Title, "Options", "IndentMode", "Previous Line") = "Previous Line" Then
        rtfSQL.AutoIndentMode = cmIndentPrevLine
    Else
        If GetSetting(App.Title, "Options", "IndentMode", "Previous Line") = "SQL Scope" Then
                rtfSQL.AutoIndentMode = cmIndentScope
        Else
            rtfSQL.AutoIndentMode = cmIndentOff
        End If
    End If
    ' highlight line
    If GetSetting(App.Title, "Options", "HighlightLine", "0") = "0" Then
        bHighlight = False
        rtfSQL.HighlightedLine = -1
    Else
        bHighlight = True
        rtfSQL.HighlightedLine = rtfSQL.GetSel(True).EndLineNo
    End If

    rtfSQL.LineNumberStart = 1
    
    cmRun.Font.Name = GetSetting(App.Title, "Options", "CommandFontName", "Courier New")
    cmRun.Font.Size = CInt(GetSetting(App.Title, "Options", "CommandFontSize", "10"))
    sFontStyle = GetSetting(App.Title, "Options", "CommandFontStyle", "Regular")
    Select Case sFontStyle
        Case Is = "Italic"
            cmRun.Font.Italic = True
            cmRun.Font.Bold = False
        Case Is = "Bold"
            cmRun.Font.Italic = False
            cmRun.Font.Bold = True
        Case Is = "Bold Italic"
            cmRun.Font.Italic = True
            cmRun.Font.Bold = True
        Case Else
            cmRun.Font.Bold = False
            cmRun.Font.Italic = False
    End Select
    Call cmRun.SetFontStyle(cmStyKeyword, cmFontNormal)
    Call cmRun.SetColor(cmClrLineNumberBk, 12632256)
    cmRun.DisplayLeftMargin = False
    cmRun.SelBounds = True  ' restrict cursor to text
    cmRun.ExpandTabs = True
    cmRun.LineNumberStart = 1
    
    ' fix the hotkeys from the cmax control
    FixHotKeys
    
    ' set up the grid
    grdSQL.Font.Name = GetSetting(App.Title, "Options", "GridFontName", "MS Sans Serif")
    grdSQL.Font.Size = CInt(GetSetting(App.Title, "Options", "GridFontSize", "10"))
    sFontStyle = GetSetting(App.Title, "Options", "GridFontStyle", "Regular")
    Select Case sFontStyle
        Case Is = "Italic"
            grdSQL.Font.Italic = True
            grdSQL.Font.Bold = False
        Case Is = "Bold"
            grdSQL.Font.Italic = False
            grdSQL.Font.Bold = True
        Case Is = "Bold Italic"
            grdSQL.Font.Italic = True
            grdSQL.Font.Bold = True
        Case Else
            grdSQL.Font.Bold = False
            grdSQL.Font.Italic = False
    End Select
       
    ' set up the log file
    sLogFile = GetSetting(App.Title, "Options", "LogFile", App.Path & "\" & App.Title & ".log")
    bLogFile = GetSetting(App.Title, "Options", "LogToFile", "0")
    iLogFile = 0
       
    If GetSetting(App.Title, "Options", "LogDateTime", "0") = "0" Then
        bLogTime = False
    Else
        bLogTime = True
    End If
        
    ' set up the date format
    sDateFormat = GetSetting(App.Title, "Options", "DateFormat", "yyyy-mm-dd")
    sTimeFormat = GetSetting(App.Title, "Options", "TimeFormat", "hh:nn:ss")
       
    ' set up the toolbar
    tbToolbar.Buttons("Cancel").Enabled = False
            
    ' set up the quicklist
    ' set up the set fonts
    lstStatus.Font.Name = "MS Sans Serif"
    lstStatus.Font.Size = 9
    tvDB.Font.Name = "MS Sans Serif"
    tvDB.Font.Size = 9
    sbStatus.Font.Name = "MS Sans Serif"
    sbStatus.Font.Size = 9
    tbsList.Font.Name = "MS Sans Serif"
    tbsList.Font.Size = 9
    
    With grdHistory
        .Font.Name = "MS Sans Serif"
        .Font.Size = 9
        .ColWidth(0) = .RowHeight(0)
    End With
    
    sListLocation = GetSetting(App.Title, "Options", "ListLocation", "Left")
    If UCase(sListLocation) = "LEFT" Then
        lVerticalPos = (Me.Width / 5) * 1 - 104
        If lVerticalPos < 2420 Then
            lVerticalPos = 2420
        End If
        frSplitVertical.Left = lVerticalPos
    Else
        lVerticalPos = (Me.Width / 5) * 4 - 100
        If lVerticalPos > 8084 Then
            lVerticalPos = 8084
        End If
        frSplitVertical.Left = lVerticalPos
    End If
    
    frSplitHorizontal.Top = (Me.Height / 2) - 1200
    TreeInitialize
    
    ' make sure that this is false
    rtfSQL.Modified = False
    
    ' show the form, and maybe splash screen
    DoEvents
    DoEvents
    DoEvents
    Me.Show
    DoEvents
    DoEvents
    DoEvents
    If GetSetting(App.Title, "Options", "ShowSplash", "1") = 1 Then
        frmSplash.Show vbModal
    End If
    DoEvents
    DoEvents
    Screen.MousePointer = vbDefault
    
    ' do we need to show license agreement
    If GetSetting(App.Title, "Options", "License", "") <> "L1C6U" & App.Major & "5A8C2U" & App.Minor & App.Revision Then
        frmAgree.Show vbModal, Me
        If bGlobalRetVal = False Then End
    End If
    
    ' on the very first run, we will show help
    ' this is moved to the frmagree area
'    If GetSetting(App.Title, "Options", "ShowHelp", "1") = 1 Then
'        mnuHelpContents_Click
'        SaveSetting App.Title, "Options", "ShowHelp", "0"
'    End If
    
    ' if sfile is passed, open it
    If bDebug Then DebugWrite "sFile:  " & sFile
    If sFile <> "" Then
        FileOpen True
    Else
        FileNew
    End If
    ' start the timer
    If bDebug Then DebugWrite "Starting Timer"
    tmrConnect.Enabled = True
End Sub
Public Function SetCaption() As Boolean
Dim sFileName As String, sCaption As String

If bDebug Then DebugWrite "SetCaption Starts"

SetCaption = False

sCaption = ""
sFileName = ""

' do we need the server name
If sDBServer <> "" Then sCaption = sDBServer & " - "

' add an * if the file is modified
If sFile <> "" Then
    sFileName = "["
    If rtfSQL.Modified = True Then
        If bDebug Then DebugWrite "File is Modified"
        sFileName = sFileName & sFile & " *"
    Else
        If bDebug Then DebugWrite "File is not Modified"
        sFileName = sFileName & sFile
    End If
    sFileName = sFileName & "] - "
End If

' do we need the file name
If sDBServer <> "" And sFileName <> "" Then
    sCaption = sCaption & sFileName
Else
    If sFileName <> "" Then
        sCaption = sFileName
    End If
End If

' set the caption
If sCaption = "" Then
    sCaption = App.Title
Else
    sCaption = sCaption & App.Title
End If

Me.Caption = sCaption

If bDebug Then DebugWrite "SetCaption Ends"

SetCaption = True

End Function


Public Sub GridPrintPage(lStartCol As Long, lEndCol As Long, lStartRow As Long, lEndRow As Long, _
sTopMargin As Single, sBottomMargin As Single, sLeftMargin As Single, sRightMargin As Single)
Dim lRowCount As Long
Dim lColCount As Long
Dim lColWidths As Long
Dim sText As String
Dim iColWidth As Integer, iTextWidth As Integer
Dim ColLeftX As Single
Dim LeftX As Single, RightX As Single
Dim LeftY As Single, RightY As Single
Dim LineHeight As Single
Dim LastLine As Single

If bDebug Then DebugWrite "GridPrintPage Starts"

    ' Some Setup
    With grdSQL
        Printer.Print ""
        'Set the font to the grid font
        Printer.FontName = .Font.Name
        Printer.FontSize = .Font.Size
        Printer.FontBold = False
        .HighLight = flexHighlightNever
        ' get the total col widths
        For lColCount = lStartCol To lEndCol
            lColWidths = lColWidths + .ColWidth(lColCount)
        Next
    End With
    
    ' make sure the printer is positioned
    Printer.CurrentX = sLeftMargin
    Printer.CurrentY = sTopMargin
    ColLeftX = sLeftMargin
    
    LineHeight = Printer.TextHeight("AAA") + 50
    LastLine = sTopMargin
    
    'Draw the box for the heading
'    Printer.Print " "
    LeftX = sLeftMargin
    LeftY = sTopMargin
    RightX = lColWidths + LeftX
    RightY = sTopMargin + LineHeight
    Printer.FillColor = RGB(192, 192, 192)
    Printer.FillStyle = vbFSSolid
    Printer.Line (LeftX, LeftY)-(RightX, RightY), , B
    Printer.FillColor = vbBlack
    Printer.FillStyle = vbFSTransparent
    Printer.CurrentX = sLeftMargin
    Printer.CurrentY = LastLine + 25
    
    ' print the top line again to get rid of any red
    LeftX = sLeftMargin
    LeftY = sTopMargin
    RightX = lColWidths + LeftX
    RightY = sTopMargin
    Printer.Line (LeftX, LeftY)-(RightX, RightY)
    Printer.CurrentX = sLeftMargin
    Printer.CurrentY = LastLine + 25
    
    'print the header
    grdSQL.Row = 0
    For lColCount = lStartCol To lEndCol
        ' set the column
        grdSQL.Col = lColCount
        ' get the column width
        iColWidth = grdSQL.ColWidth(lColCount) - 25
        ' get the text
        sText = Trim$(grdSQL.Text)
        ' do we print the column
        If iColWidth > 0 Then
            iTextWidth = Printer.TextWidth(sText)
            If iTextWidth > iColWidth And iTextWidth <> 0 Then
                'This trims the text to fit in the grid
                Do While iTextWidth > iColWidth
                    sText = Left$(sText, (Len(sText) - 1))
                    iTextWidth = Printer.TextWidth(sText)
                Loop
            End If
            Printer.CurrentX = ColLeftX + 25
            Printer.Print sText;
        End If
        'Move the start pointer to the next column position
        ColLeftX = ColLeftX + grdSQL.ColWidth(lColCount)
        Printer.CurrentX = ColLeftX
    Next lColCount
    
    'Draw the Horizontal Line for this line of the Grid
    Printer.Print " "
    LastLine = LastLine + LineHeight
    LeftX = sLeftMargin
    LeftY = LastLine
    RightX = lColWidths + LeftX
    RightY = LastLine
    Printer.Line (LeftX, LeftY)-(RightX, RightY)
    Printer.CurrentX = LeftX
    Printer.CurrentY = LastLine + 25
    ColLeftX = sLeftMargin
    
    ' print the remainder of the rows
    For lRowCount = lStartRow To lEndRow
        'Point to the current Row to Print
        grdSQL.Row = lRowCount
        For lColCount = lStartCol To lEndCol
            ' set the column
            grdSQL.Col = lColCount
            ' get the column width
            iColWidth = grdSQL.ColWidth(lColCount) - 25
            ' get the text
            sText = Trim$(grdSQL.Text)
            ' do we print the column
            If iColWidth > 0 Then
                iTextWidth = Printer.TextWidth(sText)
                If iTextWidth > iColWidth And iTextWidth <> 0 Then
                    'This trims the text to fit in the grid
                    Do While iTextWidth > iColWidth
                        sText = Left$(sText, (Len(sText) - 1))
                        iTextWidth = Printer.TextWidth(sText)
                    Loop
                End If
                Printer.CurrentX = ColLeftX + 25
                Printer.Print sText;
            End If
            'Move the start pointer to the next column position
            ColLeftX = ColLeftX + grdSQL.ColWidth(lColCount)
            Printer.CurrentX = ColLeftX
        Next lColCount
            
        'Draw the Horizontal Line for this line of the Grid
        Printer.Print ""
        LastLine = LastLine + LineHeight
        LeftX = sLeftMargin
        LeftY = LastLine
        RightX = lColWidths + LeftX
        RightY = LastLine
        Printer.Line (LeftX, LeftY)-(RightX, RightY)
        Printer.CurrentX = LeftX
        Printer.CurrentY = LastLine + 25
        ColLeftX = sLeftMargin
        
        If lRowCount Mod 50 = 0 Then
            DoEvents
            If bCancel = True Then Exit For
        End If
    Next lRowCount
    
    ' print the first vertical line
    LeftX = sLeftMargin
    LeftY = sTopMargin
    RightX = sLeftMargin
    RightY = LastLine
    Printer.Line (LeftX, LeftY)-(RightX, RightY)
    
    ' Print the rest of the Vertical Lines
    For lColCount = lStartCol To lEndCol
        iColWidth = grdSQL.ColWidth(lColCount)
        If iColWidth > 0 Then
            LeftX = LeftX + iColWidth
            RightX = LeftX
            Printer.Line (LeftX, LeftY)-(RightX, RightY)
        End If
    Next lColCount
    
If bDebug Then DebugWrite "GridPrintPage Ends"

End Sub
Public Sub GridPrint()
Dim lRowCount As Long, lCurrentRow As Long, lStartRow As Long, lEndRow As Long
Dim lColCount As Long, lColPage As Long
Dim lStartCol() As Long, lEndCol() As Long
Dim lCol As Long, lRow As Long
Dim lColWidths As Long
Dim sMsg As String
Const TOP_MARGIN = 480
Const LEFT_MARGIN = 480
Dim RIGHT_MARGIN As Single
Dim BOTTOM_MARGIN As Single
Dim lRowsPerPage As Long
Dim iPrintOrient As Integer
Dim MAX_WIDTH As Single

If bDebug Then DebugWrite "GridPrint Starts"

On Error GoTo AppError
If Not bProcessing Then
    bProcessing = True
    Screen.MousePointer = vbHourglass
    MenuSet 0
    StatsDisplay 1

    ' printer setup
    iPrintOrient = CInt(GetSetting(App.Title, "Options", "PrintOrient", "1"))
    If iPrintOrient = 1 Then
            Printer.Orientation = vbPRORPortrait
        Else
            Printer.Orientation = vbPRORLandscape
    End If
        
    ' find out how many rows
    lRowCount = grdSQL.Rows

    If lRowCount > 1 Then
        ' Some Setup
        With grdSQL
            Printer.Print ""
            'Set the font to the grid font
            Printer.FontName = .Font.Name
            Printer.FontSize = .Font.Size
            Printer.FontBold = False
            'we need this
            lCol = .Col
            lRow = .Row
            .HighLight = flexHighlightNever
            .Enabled = False
        End With

        BOTTOM_MARGIN = Printer.ScaleHeight - 480
        RIGHT_MARGIN = Printer.ScaleWidth - 480
        If iPrintOrient = 1 Then
                MAX_WIDTH = RIGHT_MARGIN - LEFT_MARGIN - 100
            Else
                MAX_WIDTH = RIGHT_MARGIN - LEFT_MARGIN - 1175
        End If
        
        ' find out how many rows per page
        lRowsPerPage = (BOTTOM_MARGIN - TOP_MARGIN) \ (Printer.TextHeight("AAA") + 50) - 1

        ' set up the array to hold the start and end col numbers
        lColWidths = 0
        lColPage = 0
        lStartRow = 1
        For lColCount = 1 To grdSQL.Cols - 1
            If (lColWidths + grdSQL.ColWidth(lColCount)) < MAX_WIDTH Then
                lColWidths = lColWidths + grdSQL.ColWidth(lColCount)
                If lColCount = grdSQL.Cols - 1 Then
                    ReDim Preserve lStartCol(lColPage + 1)
                    ReDim Preserve lEndCol(lColPage + 1)
                    lStartCol(lColPage) = lStartRow
                    lEndCol(lColPage) = lColCount
                End If
            Else
                ReDim Preserve lStartCol(lColPage + 1)
                ReDim Preserve lEndCol(lColPage + 1)
                lStartCol(lColPage) = lStartRow
                If lColCount = 1 Then
                    lEndCol(lColPage) = lColCount
                    Exit For
                ElseIf lColCount = grdSQL.Cols - 1 Then
                    lEndCol(lColPage) = lColCount
                    lStartRow = lColCount
                    Exit For
                Else
                    lEndCol(lColPage) = lColCount - 1
                    lStartRow = lColCount
                End If
                lColPage = lColPage + 1
                lColWidths = 0
            End If
        Next

        ' loop through the rows until none are left
        lCurrentRow = 1
        lStartRow = 1
        lEndRow = lRowCount - 1
        Do
            lEndRow = lStartRow + lRowsPerPage
            If lEndRow > lRowCount - 1 Then lEndRow = lRowCount - 1
            For lColCount = 0 To lColPage
                ' print
                GridPrintPage lStartCol(lColCount), lEndCol(lColCount), lStartRow, lEndRow, TOP_MARGIN, BOTTOM_MARGIN, LEFT_MARGIN, RIGHT_MARGIN
                'send a new page if needed
                If lColCount < lColPage Then Printer.NewPage
            Next
            lStartRow = lEndRow + 1
            If lStartRow > lRowCount - 1 Then Exit Do
            If bCancel Then Exit Do

        Loop Until lStartRow > lRowCount - 1
        ' force the print
         Printer.EndDoc
    End If

    ' reset the grid
    With grdSQL
        .Col = lCol
        .Row = lRow
        .HighLight = flexHighlightAlways
        .Enabled = True
    End With

    Screen.MousePointer = vbDefault
    MenuSet 1
    StatsDisplay 0
    bCancel = False
    bProcessing = False

End If
If bDebug Then DebugWrite "GridPrint Ends"

Exit Sub

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    ' reset the grid
    With grdSQL
        .Col = lCol
        .Row = lRow
        .HighLight = flexHighlightAlways
        .Enabled = True
    End With
    Screen.MousePointer = vbDefault
    MenuSet 1
    bCancel = False
    bProcessing = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim bRetVal As Boolean
Dim iRetVal As Integer
Dim sMsg As String

On Error GoTo AppError

If bDebug Then DebugWrite "Query Unload Starts"
If bProcessing Then
    If bDebug Then DebugWrite "Query Unload bProcessing Exit Sub"
    Cancel = True
    Exit Sub
End If

On Error Resume Next
Err.Clear
If envSQL.rdoConnections.Count > 0 Then
    If Err.Number <> 91 Then
        bRetVal = SQLExecute("COMMIT")
        If bRetVal = False Then
            sMsg = "An error has occured.  Do you still want to Exit?"
            If MsgBox(sMsg, vbYesNo + vbQuestion) = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
        If bDebug Then DebugWrite "Closing Connection"
        conSQL.Close
        If bDebug Then DebugWrite "Closed Connection"
    End If
End If
On Error GoTo AppError
' the file has been changed
If rtfSQL.Modified = True Then
   If sFile <> "" Or GetSetting(App.Title, "Options", "SaveNewOnExit", "1") = "1" Then
        'do we want to save
        iRetVal = MsgBox("File Changed.  Save Changes?", vbDefaultButton1 + vbYesNoCancel + vbQuestion)
        'if yes, save
        If iRetVal = vbYes Then
            If FileSaveAs <> True Then
                Cancel = True
                Exit Sub
            End If
        Else
            If iRetVal = vbCancel Then
                Cancel = True
                Exit Sub
            End If
        End If
        If bDebug Then DebugWrite "Save Changes = " & Str(iRetVal)
    End If
End If

If bDebug Then DebugWrite "Query Unload Ends"
Exit Sub

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'*** Code added by HelpWriter ***
'*** Subroutine added by HelpWriter ***
'    QuitHelp
'***********************************
    If bDebug Then DebugWrite "Closing Environment"
    envSQL.Close
    If bDebug Then DebugWrite "App Ends"
    If bDebug Then Close #iDebugFile
    If iLogFile > 0 Then
        Close #iLogFile
        iLogFile = 0
    End If
End Sub

Private Sub frSplitVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMoving = True
    frSplitVertical.BackColor = &H80000006
End Sub

Private Sub frSplitVertical_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sglPos As Single
    If mbMoving Then
        sglPos = X + frSplitVertical.Left
        If sglPos < sglSplitLimit Then
            frSplitVertical.Left = sglSplitLimit
        Else
            frSplitVertical.Left = sglPos
        End If
    End If

End Sub

Private Sub frSplitVertical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMoving = False
    frSplitVertical.BackColor = &H8000000F
    ControlSize
End Sub

Private Sub grdHistory_DblClick()
    ' reset the rtf window
    Clipboard.Clear
    Clipboard.SetText grdHistory.Text
    Call rtfSQL.ExecuteCmd(cmCmdSelectAll)
    Call rtfSQL.ExecuteCmd(cmCmdDelete)
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
    rtfSQL.SetFocus
End Sub


Private Sub grdHistory_GotFocus()
If Not bProcessing Then MenuSet 1
End Sub


Private Sub grdHistory_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then
    If Chr(KeyCode) = "C" Or Chr(KeyCode) = "c" Then
        mnuEditChild_Click 6
    ElseIf Chr(KeyCode) = "V" Or Chr(KeyCode) = "v" Then
        mnuEditChild_Click 7
    End If
End If
If KeyCode = vbKeyDelete Then
    mnuEditChild_Click 8
End If
If Shift = 0 Then
    If Not bProcessing And KeyCode = 93 Then
        grdHistory.SetFocus
        PopupMenu mnuPopGridHistory, vbPopupMenuLeftAlign, _
            frList.Left + grdHistory.CellLeft + grdHistory.CellWidth, _
            frList.Top + grdHistory.CellTop + grdHistory.CellHeight
    End If
End If

End Sub

Private Sub grdHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not bProcessing Then
            grdHistory.SetFocus
            PopupMenu mnuPopGridHistory
        End If
    End If
End Sub

Private Sub grdHistory_SelChange()
    grdHistory.Row = grdHistory.RowSel
End Sub


Private Sub grdSQL_EnterCell()
    GridPicture (1)
End Sub

Private Sub grdSQL_GotFocus()
If Not bProcessing Then MenuSet 1
End Sub


Private Sub grdSQL_KeyDown(KeyCode As Integer, Shift As Integer)
If (Shift And vbCtrlMask) > 0 Then
    If Chr(KeyCode) = "C" Or Chr(KeyCode) = "c" Then
        mnuEditChild_Click 6
    ElseIf Chr(KeyCode) = "V" Or Chr(KeyCode) = "v" Then
        mnuEditChild_Click 7
    End If
End If

If Shift = 0 Then
    If Not bProcessing And KeyCode = 93 Then
        grdSQL.SetFocus
        PopupMenu mnuPopGrid, vbPopupMenuLeftAlign, _
            frSQL.Left + grdSQL.CellLeft + grdSQL.CellWidth, _
            frSQL.Top + grdSQL.CellTop + grdSQL.CellHeight
    End If
End If
End Sub

Private Sub grdSQL_LeaveCell()
    GridPicture (0)
End Sub

Private Sub grdSQL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not bProcessing Then
            grdSQL.SetFocus
            PopupMenu mnuPopGrid
        End If
    End If
End Sub


Private Sub lstStatus_GotFocus()
If Not bProcessing Then MenuSet 1
End Sub


Private Sub mnuDBChild_Click(Index As Integer)
Select Case Index
    Case Is = 0
        SQLParse
    Case Is = 1
        bCancel = True
    Case Is = 3
        LoginShow
    Case Is = 4
        On Error Resume Next
        If bDebug Then DebugWrite "Disconnect Starts"
        Err.Clear
        If envSQL.rdoConnections.Count > 0 Then
            If Err.Number <> 91 Then
                If bDebug Then DebugWrite "Commit Starts"
                Call SQLExecute("COMMIT")
                If bDebug Then DebugWrite "Commit Ends"
                conSQL.Close
                If bDebug Then DebugWrite "Connection Closed"
            End If
        End If
        If bDebug Then DebugWrite "Call StatsDisplay 7"
        StatsDisplay 7
        If bDebug Then DebugWrite "Call DBInfo 0"
        DBInfo 0
        If bDebug Then DebugWrite "Disconnect Ends"
End Select
End Sub

Private Sub mnuEdit_Click()
    MenuSet 1
End Sub


Private Sub mnuEditChild_Click(Index As Integer)
Select Case Index
    Case Is = 0
        SQLParse
    Case Is = 2
        Call rtfSQL.ExecuteCmd(cmCmdUndo)
    Case Is = 3
        Call rtfSQL.ExecuteCmd(cmCmdRedo)
    Case Is = 5
        Call rtfSQL.ExecuteCmd(cmCmdCut)
    Case Is = 6
        Select Case Me.ActiveControl.Name
            Case "rtfSQL"
                Call rtfSQL.ExecuteCmd(cmCmdCopy)
            Case "grdSQL"
                GridCopy
            Case "tvDB"
                TreeCopy tvDB.SelectedItem.Key, 1
            Case "grdHistory"
                Clipboard.Clear
                Clipboard.SetText grdHistory.Text
        End Select
    Case Is = 7
        Select Case Me.ActiveControl.Name
            Case "rtfSQL"
                Call rtfSQL.ExecuteCmd(cmCmdPaste)
            Case "grdSQL"
                GridCopy
                rtfSQL.SetFocus
                Call rtfSQL.ExecuteCmd(cmCmdPaste)
            Case "tvDB"
                TreeCopy tvDB.SelectedItem.Key, 1
                rtfSQL.SetFocus
                Call rtfSQL.ExecuteCmd(cmCmdPaste)
            Case "grdHistory"
                Clipboard.Clear
                Clipboard.SetText grdHistory.Text
                rtfSQL.SetFocus
                Call rtfSQL.ExecuteCmd(cmCmdPaste)
        End Select
        
    Case Is = 8
        Select Case Me.ActiveControl.Name
            Case "rtfSQL"
                rtfSQL.ExecuteCmd (cmCmdDelete)
            Case "grdHistory"
                With grdHistory
                    If .Rows > 1 Then
                        .RemoveItem .Row
                    Else
                        .Text = ""
                    End If
                End With
        End Select
    Case Is = 10
         Call rtfSQL.ExecuteCmd(cmCmdSelectAll)
         
    Case Is = 14
        Call rtfSQL.ExecuteCmd(cmCmdFind)
    Case Is = 15
        Call rtfSQL.ExecuteCmd(cmCmdFindNext)
    Case Is = 16
        Call rtfSQL.ExecuteCmd(cmCmdFindReplace)
        
End Select
MenuSet 1
End Sub

Private Sub mnuEditChildFormat_Click(Index As Integer)
Select Case Index
    Case Is = 0
        Call rtfSQL.ExecuteCmd(cmCmdUntabifySelection)
    Case Is = 1
        Call rtfSQL.ExecuteCmd(cmCmdTabifySelection)
    Case Is = 2
        rtfSQL.DisplayWhitespace = Not rtfSQL.DisplayWhitespace
        mnuEditChildFormat(2).Checked = Not mnuEditChildFormat(2).Checked
    Case Is = 4
        Call rtfSQL.ExecuteCmd(cmCmdLowercaseSelection)
    Case Is = 5
        Call rtfSQL.ExecuteCmd(cmCmdUppercaseSelection)
    Case Is = 6
        Call rtfSQL.ExecuteCmd(cmCmdWordCapitalize)
    Case Is = 8
        Call rtfSQL.ExecuteCmd(cmCmdIndentSelection)
    Case Is = 9
        Call rtfSQL.ExecuteCmd(cmCmdUnindentSelection)
    Case Is = 11
        Call rtfSQL.ExecuteCmd(cmCmdCharTranspose)
    Case Is = 12
        Call rtfSQL.ExecuteCmd(cmCmdWordTranspose)
    Case Is = 13
        Call rtfSQL.ExecuteCmd(cmCmdLineTranspose)
    Case Is = 15
        CommentBlock
    Case Is = 16
        UnCommentBlock
    
End Select
MenuSet 1

End Sub
Public Function InsertText(sString As String) As Boolean
Dim cmRange As CodeMaxCtl.Range

If bDebug Then DebugWrite "InsertText Starts"

InsertText = False

Set cmRange = New CodeMaxCtl.Range
    
If bDebug Then DebugWrite "sString:  " & sString
rtfSQL.SelText = sString

Set cmRange = rtfSQL.GetSel(False)
rtfSQL.SetCaretPos cmRange.EndLineNo, cmRange.StartColNo + Len(sString)

If bDebug Then DebugWrite "Setting Variables to Nothing."

Set cmRange = Nothing
    
If bDebug Then DebugWrite "InsertText Ends"

InsertText = True

End Function


Private Sub mnuEditChildGoTo_Click(Index As Integer)
Select Case Index
    Case Is = 0
        Call rtfSQL.ExecuteCmd(cmCmdGotoLine, -1)
    Case Is = 2
        Call rtfSQL.ExecuteCmd(cmCmdGotoMatchBrace)
    Case Is = 4
        Call rtfSQL.ExecuteCmd(cmCmdBookmarkToggle)
    Case Is = 5
        Call rtfSQL.ExecuteCmd(cmCmdBookmarkNext)
    Case Is = 6
        Call rtfSQL.ExecuteCmd(cmCmdBookmarkPrev)
    Case Is = 7
        Call rtfSQL.ExecuteCmd(cmCmdBookmarkClearAll)
    
End Select
MenuSet 1
        
End Sub

Private Sub mnuFileChild_Click(Index As Integer)
Select Case Index
    Case Is = 0
        FileNew
    Case Is = 1
        FileOpen False
    Case Is = 3
        FileSaveAs
    Case Is = 4
        FileSaveAs True
    Case Is = 6
        GridPrint
    Case Is = 7
        FilePrint
    Case Is = 8
        PrintSetup
    Case Is = 10
        GridSave
    Case Is = 11
        GridSave True
    Case Is = 13
        'unload the form and end
        Unload Me
End Select
End Sub

Public Function TextWrap(sWholeText As String, iMaxLenOfLine As Integer, Optional bForceSplit As Boolean = True) As String
Dim sForcedLines() As String
Dim lCounter As Long
Dim sTempString As String
Dim sMsg As String
    
On Error GoTo AppError
        
If bDebug Then DebugWrite "TextWrap Starts"
    
    'fragment for newlines
    sForcedLines = Split(sWholeText, vbCrLf)
    'for each line (element of array)
    'decide if line needs to be divided
    'in smallest pieces

    For lCounter = LBound(sForcedLines) To UBound(sForcedLines)
        If Len(sForcedLines(lCounter)) > iMaxLenOfLine Then
            sForcedLines(lCounter) = LineWrap(sForcedLines(lCounter), iMaxLenOfLine, bForceSplit)
        End If
    Next lCounter
    'rebuild final string
    TextWrap = Join(sForcedLines, vbCrLf)

If bDebug Then DebugWrite "TextWrap Ends"

Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    TextWrap = ""
    Exit Function

End Function


Private Function LineWrap(sInput As String, iMaxLen As Integer, bForceSplit As Boolean) As String
Dim iFoundPos As Integer
Dim sSplitChar As String
Dim iSplitLen As Integer
Dim sMsg As String

On Error GoTo AppError
        
If bDebug Then DebugWrite "LineWrap Starts"
    
    'strip unnecessary right spaces
    sInput = RTrim(sInput)
    'the bForceSplit is to determine if you want
    'fixed len lines where division is made with "-" char
    '(when breaking a word) and vbcrlf
    If bForceSplit Then
        Do While Len(sInput) > iMaxLen
            If Mid(sInput, iMaxLen, 1) = " " Then
                sSplitChar = ""
            Else
                sSplitChar = "-"
            End If
            LineWrap = LineWrap & Left(sInput, iMaxLen) & sSplitChar & vbCrLf
            
            If Len(sInput) - iMaxLen > 0 Then
                sInput = Right(sInput, Len(sInput) - iMaxLen)
                'trim non significant left spaces
                sInput = LTrim(sInput)
            End If
        Loop
            
        'add last piece of string
        LineWrap = LineWrap & sInput
           
    Else
        'If the bForceSplit is false, then you want this code
        'to divide lines where a break is found (a space), and
        'only if a space is not found, you want it to divide lines
        'like before, with a "-" char when breaking a word
        Do While Len(sInput) > iMaxLen
        
            ' lws fix to break on the larger of , or space
            If InStrRev(sInput, ",", iMaxLen) > InStrRev(sInput, " ", iMaxLen) Then
                sSplitChar = ","
            Else
                sSplitChar = " "
            End If
            
            iFoundPos = InStrRev(sInput, sSplitChar, iMaxLen)
            
            If iFoundPos > 0 Then
                If sSplitChar = "," Then
                    LineWrap = LineWrap & Left(sInput, iFoundPos) & vbCrLf
                    iSplitLen = iFoundPos + 1
                    
                Else
                    LineWrap = LineWrap & Left(sInput, iFoundPos - 1) & vbCrLf
                    iSplitLen = iFoundPos
                End If
                
            Else
                LineWrap = LineWrap & Left(sInput, iMaxLen - 1) & "-" & vbCrLf
                iSplitLen = iMaxLen
            End If

            If Len(sInput) - iSplitLen > 0 Then
                sInput = Right(sInput, Len(sInput) + 1 - iSplitLen)
                'trim non significant left spaces
                sInput = LTrim(sInput)
            End If
        Loop
        'add last piece of string
        LineWrap = LineWrap & sInput
    End If


If bDebug Then DebugWrite "LineWrap Ends"

Exit Function

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    Screen.MousePointer = vbDefault
    LineWrap = ""
    Exit Function
    
End Function

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub



Private Sub mnuHelpContents_Click()
'ShowHelpContents
End Sub



Private Sub mnuPopGridClear_Click()
Dim iRetVal As Integer
Dim sMsg As String

On Error GoTo AppError

If bDebug Then DebugWrite "mnuPopGridClear Starts"
If Not bProcessing Then
    bProcessing = True
    bCancel = False
    iRetVal = MsgBox("Clear Results Grid Contents?" _
        , vbDefaultButton1 + vbYesNo + vbQuestion)
    'if yes, clear
    If iRetVal = vbYes Then
        Screen.MousePointer = vbHourglass
        MenuSet 0
        StatsDisplay 1
        With grdSQL
            .Clear
            .Rows = 2
            .Cols = 2
        End With
        Screen.MousePointer = vbDefault
        MenuSet 1
        StatsDisplay 0
    End If
    bCancel = False
    bProcessing = False
End If
If bDebug Then DebugWrite "mnuPopGridClear Ends"
Exit Sub

AppError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
        MsgBox sMsg, vbExclamation + vbOKOnly
        Err.Clear
    End If
    MenuSet 1
    StatsDisplay 0
    Screen.MousePointer = vbDefault
    bProcessing = False
    Exit Sub
End Sub

Private Sub mnuPopGridCopy_Click()
    grdSQL.SetFocus
    mnuEditChild_Click 6
End Sub

Private Sub mnuPopGridHistoryCopy_Click()
    mnuEditChild_Click 6
End Sub

Private Sub mnuPopGridHistoryDelete_Click()
    mnuEditChild_Click 8
End Sub


Private Sub mnuPopGridHistoryPaste_Click()
    mnuEditChild_Click 7
End Sub


Private Sub mnuPopGridPaste_Click()
    mnuPopGridCopy_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopGridSortAsc_Click()
    GridSort 0
End Sub

Private Sub mnuPopGridSortDesc_Click()
    GridSort 1
End Sub
Public Function TreeCopyItem(sMyKey As String) As Boolean
Dim sText As String
Dim sMyPrefix As String

TreeCopyItem = False
If bDebug Then DebugWrite "TreeCopyItem Starts"

sMyPrefix = Left$(sMyKey, InStr(sMyKey, "_"))
Select Case sMyPrefix
    Case "t_", "v_", "p_"
        ' include the owner
        sText = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
        sText = sDQUOTE & Mid$(sText, InStr(sMyKey, "_"), Len(sText) - InStr(sMyKey, "_")) & sDQUOTE & "."
        sText = sText & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
    Case "c_", "a_", "tr_"
        ' only the item
        If InStr(tvDB.Nodes(sMyKey).Text, "  ") > 0 Then
            sText = sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
        Else
            sText = sDQUOTE & tvDB.Nodes(sMyKey).Text & sDQUOTE
        End If
    Case "i_"
        ' include the owner
        sText = sDQUOTE & Trim(tvDB.Nodes(sMyKey).Text) & sDQUOTE
        ' do we need to remove the "  Constraint" or the "  Unique"
        If Right(sText, 8) = "  Unique" Then
            sText = Left$(sText, Len(sText) - 8)
            sText = "Unique " & sText
        ElseIf Right(sText, 12) = "  Constraint" Then
            sText = Left$(sText, Len(sText) - 12)
            sText = "Constraint " & sText
        End If
    Case Else
        ' all of the item
        sText = tvDB.Nodes(sMyKey).Text
End Select
Clipboard.Clear
Clipboard.SetText sText
    
If bDebug Then DebugWrite "TreeCopyItem Ends"
TreeCopyItem = True
End Function

Public Function TreeCopyColumns(sMyKey As String, Optional iCase As Integer) As Boolean
Dim sText As String, sOwner As String
Dim sMyPrefix As String
Dim sMyChild As String
Dim sGrandChildKey As String
                   
TreeCopyColumns = False
If bDebug Then DebugWrite "TreeCopyColumns Starts"

sMyPrefix = Left$(sMyKey, InStr(sMyKey, "_"))
Select Case sMyPrefix
    Case "t_", "v_"
        Select Case iCase
            Case 3
            'select statement
                sText = "SELECT "
                ' now we have to get each column
                ' set up the child
                sMyChild = "fc_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) & sDQUOTE
                    Else
                        Do
                            sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) _
                                & sDQUOTE & "," & vbCrLf & "    "
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        sText = sText _
                            & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyChild).Child.LastSibling.Text, (InStr(tvDB.Nodes(sMyChild).Child.LastSibling.Text, "  ")))) & sDQUOTE
                        End If
                Else
                    sText = sText & "*"
                End If
                ' rest of the select statement
                sText = sText & vbCrLf & "FROM "
                ' include the owner
                sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
                sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
                sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
                sText = sText & sOwner & vbCrLf
            
            
            Case 4
            'insert statement
                ' include the owner
                sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
                sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
                sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
                sText = "INSERT INTO " & sOwner
                ' now we have to get each column
                ' set up the child
                sMyChild = "fc_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sText = sText & "("
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) & sDQUOTE & ")"
                    Else
                        Do
                            sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) _
                                & sDQUOTE & "," & vbCrLf & "    "
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        sText = sText _
                            & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyChild).Child.LastSibling.Text, (InStr(tvDB.Nodes(sMyChild).Child.LastSibling.Text, "  ")))) _
                            & sDQUOTE & ")"
                    End If
                End If
                sText = sText & vbCrLf & "VALUES("
                
            Case 7
            ' update statement
                ' include the owner
                sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
                sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
                sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
                sText = "UPDATE " & sOwner & vbCrLf & "SET "
                ' now we have to get each column
                ' set up the child
                sMyChild = "fc_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) & sDQUOTE
                    Else
                        Do
                            sText = sText _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) _
                                & sDQUOTE & " = " _
                                & sDQUOTE & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  ")))) _
                                & sDQUOTE & "," & vbCrLf & "    "
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        sText = sText _
                            & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyChild).Child.LastSibling.Text, (InStr(tvDB.Nodes(sMyChild).Child.LastSibling.Text, "  ")))) _
                            & sDQUOTE & " = " _
                            & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyChild).Child.LastSibling.Text, (InStr(tvDB.Nodes(sMyChild).Child.LastSibling.Text, "  ")))) _
                            & sDQUOTE
                    End If
                Else
                    sText = sText & " = "
                End If
                sText = sText & vbCrLf & "WHERE "
                
            
            Case 8
            'delete statement
                sText = "DELETE "
                sText = sText & vbCrLf & "FROM "
                ' include the owner
                sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
                sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
                sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
                sText = sText & sOwner & vbCrLf
                sText = sText & "WHERE "
                
            
            Case Else
                ' include the owner
                sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
                sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
                sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
                sText = sOwner
                ' now we have to get each column
                ' set up the child
                sMyChild = "fc_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sText = sText & "("
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          sText = sText _
                                & Trim(tvDB.Nodes(sGrandChildKey).Text)
                    Else
                        Do
                            sText = sText _
                                & Trim(tvDB.Nodes(sGrandChildKey).Text) & "," _
                                & vbCrLf & "    "
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        sText = sText _
                            & Trim(tvDB.Nodes(sMyChild).Child.LastSibling.Text) & ")"
                    End If
                End If
                sText = sText & vbCrLf
        End Select
        
   Case "p_"
        ' include the owner
        sOwner = Trim(Mid$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  ")))
        sOwner = sDQUOTE & Mid$(sOwner, InStr(sMyKey, "_"), Len(sOwner) - InStr(sMyKey, "_")) & sDQUOTE & "."
        sOwner = sOwner & sDQUOTE & Trim(Left$(tvDB.Nodes(sMyKey).Text, InStr(tvDB.Nodes(sMyKey).Text, "  "))) & sDQUOTE
        
        Select Case iCase
            Case 5
                sText = "EXECUTE " & sOwner & " "
                ' now we have to get each column
                ' set up the child
                sMyChild = "fa_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          If (Right$(tvDB.Nodes(sGrandChildKey).Text, 4) = "  In") Or (Right$(tvDB.Nodes(sGrandChildKey).Text, 7) = "  InOut") Then
                            sText = sText _
                                & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  "))))
                          End If
                    Else
                        Do
                            If (Right$(tvDB.Nodes(sGrandChildKey).Text, 4) = "  In") Or (Right$(tvDB.Nodes(sGrandChildKey).Text, 7) = "  InOut") Then
                                If Right$(sText, 1) <> " " And Right$(sText, 1) <> "," Then
                                    sText = sText & ","
                                End If
                                sText = sText _
                                    & Trim(Left$(tvDB.Nodes(sGrandChildKey).Text, (InStr(tvDB.Nodes(sGrandChildKey).Text, "  "))))
                                    
                            End If
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        If (Right$(tvDB.Nodes(sGrandChildKey).Text, 4) = "  In") Or (Right$(tvDB.Nodes(sGrandChildKey).Text, 7) = "  InOut") Then
                            If Right$(sText, 1) <> " " Then
                                sText = sText & ","
                            End If
                            sText = sText _
                                & Trim(Left$(tvDB.Nodes(sMyChild).Child.LastSibling.Text, (InStr(tvDB.Nodes(sMyChild).Child.LastSibling.Text, "  "))))
                        End If
                    End If
                End If
                sText = sText & vbCrLf
            Case Else
                sText = sOwner & " "
                ' now we have to get each column
                ' set up the child
                sMyChild = "fa_" & Mid$(sMyKey, InStr(sMyKey, "_") + 1)
                If tvDB.Nodes(sMyChild).Children > 0 Then
                    sGrandChildKey = tvDB.Nodes(sMyChild).Child.FirstSibling.Key
                    If tvDB.Nodes(sMyChild).Children = 1 Then
                          sText = sText _
                                & Trim(tvDB.Nodes(sGrandChildKey).Text)
                    Else
                        Do
                            sText = sText _
                                & Trim(tvDB.Nodes(sGrandChildKey).Text) & "," _
                                & vbCrLf & "    "
                            sGrandChildKey = tvDB.Nodes(sGrandChildKey).Next.Key
                        Loop Until sGrandChildKey = tvDB.Nodes(sMyChild).Child.LastSibling.Key
                        sText = sText _
                            & Trim(tvDB.Nodes(sMyChild).Child.LastSibling.Text)
                    End If
                End If
                sText = sText & vbCrLf
        End Select
    Case "i_"
        sText = Trim(tvDB.Nodes(sMyKey).Text)
        ' do we need to remove the "  Constraint" or the "  Unique"
        If Right(sText, 8) = "  Unique" Then
            sText = Left$(sText, Len(sText) - 8)
            sText = "Unique " & sText
        ElseIf Right(sText, 12) = "  Constraint" Then
            sText = Left$(sText, Len(sText) - 12)
            sText = "Constraint " & sText
        Else
            sText = sDQUOTE & sText & sDQUOTE
        End If
        If Left(sText, 11) <> "Constraint " Then
            ' now we have to get each column
            If tvDB.Nodes(sMyKey).Children > 0 Then
                sText = sText & "("
                sMyChild = tvDB.Nodes(sMyKey).Child.FirstSibling.Key
                If tvDB.Nodes(sMyKey).Children = 1 Then
                      sText = sText & sDQUOTE & tvDB.Nodes(sMyChild).Text & sDQUOTE & ")"
                Else
                    Do
                        sText = sText & sDQUOTE & tvDB.Nodes(sMyChild).Text & sDQUOTE & ","
                        sMyChild = tvDB.Nodes(sMyChild).Next.Key
                    Loop Until sMyChild = tvDB.Nodes(sMyKey).Child.LastSibling.Key
                    sText = sText & sDQUOTE & tvDB.Nodes(sMyKey).Child.LastSibling.Text & sDQUOTE & ")" & vbCrLf
                End If
            End If
        End If
    Case Else
        TreeCopyColumns = False
        Exit Function
End Select

If sText <> "" Then
    Clipboard.Clear
    Clipboard.SetText sText
End If
If bDebug Then DebugWrite "TreeCopyColumns Ends"
TreeCopyColumns = True
End Function

Public Function TreeNodeClick(ByVal sMyKey As String) As Boolean
Dim sMsg As String
Dim bRetVal As Boolean
Dim iCount As Integer

TreeNodeClick = False
If bDebug Then DebugWrite "TreeNodeClick Starts"

If Not bProcessing Then
    bProcessing = True
    Screen.MousePointer = vbHourglass
    MenuSet 0
    StatsDisplay 1
    
    On Error GoTo ConnectError
    If envSQL.rdoConnections.Count > 0 Then
        On Error GoTo SQLError

        Select Case sMyKey
            Case "fDatabase"
                ' do nothing
            Case "fTables"
                ' do tables
                If tvDB.Nodes(2).Children = 0 Then
                    TreeGetTables
                End If
            Case "fViews"
                ' do views
                If tvDB.Nodes(3).Children = 0 Then
                    TreeGetViews
                End If
            Case "fProcs"
                ' do procs
                If tvDB.Nodes(4).Children = 0 Then
                    TreeGetProcs
                End If
            Case "fTypes"
                ' do procs
                If tvDB.Nodes(5).Children = 0 Then
                    TreeGetTypes
                End If
            Case Else
                ' get the detail info
                If tvDB.SelectedItem.Children = 0 Then
                    If Left$(sMyKey, 3) = "fc_" Then
                        TreeGetColumns sMyKey
                    ElseIf Left$(sMyKey, 3) = "fa_" Then
                        TreeGetParms sMyKey
                    ElseIf Left$(sMyKey, 3) = "fi_" Then
                        TreeGetIndexes sMyKey
                    ElseIf Left$(sMyKey, 3) = "ft_" Then
                        TreeGetTriggers sMyKey
                    ElseIf Left$(sMyKey, 2) = "i_" Then
                        TreeGetIndexColumns sMyKey
                    End If
                End If
        End Select
    End If
    
    Screen.MousePointer = vbDefault
    MenuSet 1
    StatsDisplay 0
    bCancel = False
    bProcessing = False
End If

If bDebug Then DebugWrite "TreeNodeClick Ends"
TreeNodeClick = True
Exit Function

ConnectError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number = 91 Then
        bRetVal = LoginShow
        If bRetVal = True Then
            Screen.MousePointer = vbHourglass
            Resume
        Else
            StatsDisplay 7
            Screen.MousePointer = vbDefault
            TreeNodeClick = False
            Exit Function
        End If
    Else: GoTo SQLError
    End If
    StatsDisplay 0
    MenuSet 1
    TreeNodeClick = False
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
    StatsDisplay 5
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
    MenuSet 1
    bCancel = False
    bProcessing = False
    TreeNodeClick = False
    Exit Function

End Function

Public Function TreeGetDDL(sMyKey As String) As Boolean
Dim sMsg As String
Dim bRetVal As Boolean
Dim iCount As Integer
Dim iErrorCount As Integer
iErrorCount = 0

If Not bProcessing Then
    bProcessing = True
    Screen.MousePointer = vbHourglass
    MenuSet 0
    StatsDisplay 1

    On Error GoTo ConnectError
    If envSQL.rdoConnections.Count > 0 Then
        On Error GoTo SQLError
        
        TreeGetDDL = False
        If bDebug Then DebugWrite "TreeGetDDL Starts"
        Screen.MousePointer = vbHourglass
            
            
            If Left(sMyKey, 2) = "t_" Then
                TreeGetTableDDL sMyKey
            End If
            If Left(sMyKey, 2) = "v_" Then
                TreeGetViewDDL sMyKey
            End If
            If Left(sMyKey, 2) = "p_" Then
                TreeGetProcDDL sMyKey
            End If
            If Left(sMyKey, 3) = "tr_" Then
                TreeGetTriggerDDL sMyKey
            End If
            If Left(sMyKey, 2) = "i_" Then
                TreeGetIndexDDL sMyKey
            End If
            
        End If
    TreeGetDDL = True

    Screen.MousePointer = vbDefault
    MenuSet 1
    StatsDisplay 0
    bCancel = False
    bProcessing = False
End If


If bDebug Then DebugWrite "TreeGetDDL Ends"
Screen.MousePointer = vbDefault
Exit Function

ConnectError:
    StatsDisplay 5
    Screen.MousePointer = vbDefault
    If Err.Number = 91 Then
        bRetVal = LoginShow
        If bRetVal = True Then
            Screen.MousePointer = vbHourglass
            Resume
        Else
            StatsDisplay 7
            Screen.MousePointer = vbDefault
            TreeGetDDL = False
            Exit Function
        End If
    Else: GoTo SQLError
    End If
    StatsDisplay 0
    MenuSet 1
    TreeGetDDL = False
    Exit Function

SQLError:
    ' a cursor error may be triggered when we _
      run a stored procedure that doesn't return _
      and rows.  We will test for this and continue
    If Err.Number = 521 And iErrorCount < 5 Then
        iErrorCount = iErrorCount + 1
        Err.Clear
        Resume
    End If
          
    If Err.Number = 40088 Then
        DBActionRows
        Screen.MousePointer = vbDefault
        Resume Next
    End If
'    If Err.Number = 40086 Then
'        Resume Next
'    End If
    StatsDisplay 5
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
    MenuSet 1
    bCancel = False
    bProcessing = False
    TreeGetDDL = False
    Exit Function

End Function


Private Sub mnuPopTreeObjectCopy_Click()
    TreeCopy tvDB.SelectedItem.Key, 1
End Sub

Private Sub mnuPopTreeObjectCopyColumn_Click()
    TreeCopy tvDB.SelectedItem.Key, 2
End Sub

Private Sub mnuPopTreeObjectCopyColumnDelete_Click()
    TreeCopy tvDB.SelectedItem.Key, 8
End Sub

Private Sub mnuPopTreeObjectCopyColumnInsert_Click()
    TreeCopy tvDB.SelectedItem.Key, 4
End Sub

Private Sub mnuPopTreeObjectCopyColumnSelect_Click()
    TreeCopy tvDB.SelectedItem.Key, 3
End Sub

Private Sub mnuPopTreeObjectCopyColumnUpdate_Click()
    TreeCopy tvDB.SelectedItem.Key, 7
End Sub

Private Sub mnuPopTreeObjectCopyDDL_Click()
    TreeCopy tvDB.SelectedItem.Key, 6
End Sub

Private Sub mnuPopTreeObjectCopyParameter_Click()
    TreeCopy tvDB.SelectedItem.Key, 2
End Sub

Private Sub mnuPopTreeObjectCopyParameterExecute_Click()
    TreeCopy tvDB.SelectedItem.Key, 5
End Sub

Private Sub mnuPopTreeObjectPaste_Click()
    mnuPopTreeObjectCopy_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteColumn_Click()
    mnuPopTreeObjectCopyColumn_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteColumnDelete_Click()
    mnuPopTreeObjectCopyColumnDelete_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteColumnInsert_Click()
    mnuPopTreeObjectCopyColumnInsert_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteColumnSelect_Click()
    mnuPopTreeObjectCopyColumnSelect_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteColumnUpdate_Click()
    mnuPopTreeObjectCopyColumnUpdate_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuPopTreeObjectPasteDDL_Click()
Dim lPos As Long
    lPos = rtfSQL.GetSel(True)
    mnuPopTreeObjectCopyDDL_Click
    rtfSQL.SetFocus
    
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
    
    Call rtfSQL.SetCaretPos(lPos, 0)
End Sub

Private Sub mnuPopTreeObjectPasteParameter_Click()
    mnuPopTreeObjectCopyParameter_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub
Private Sub Form_Resize()
On Error Resume Next

    If Me.Height < 2000 Then Me.Height = 2000
    If Me.Width < 2000 Then Me.Width = 2000
    ControlSize
End Sub


Private Sub frSplitHorizontal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMoving = True
    frSplitHorizontal.BackColor = &H80000006
End Sub

Private Sub frSplitHorizontal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sglPos As Single
    If mbMoving Then
        sglPos = Y + frSplitHorizontal.Top
        If sglPos < sglSplitLimit Then
            frSplitHorizontal.Top = sglSplitLimit
        Else
            frSplitHorizontal.Top = sglPos
        End If
    End If
End Sub


Private Sub frSplitHorizontal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mbMoving = False
    frSplitHorizontal.BackColor = &H8000000F
    ControlSize
End Sub


Private Sub mnuPopTreeObjectPasteParameterExecute_Click()
    mnuPopTreeObjectCopyParameterExecute_Click
    rtfSQL.SetFocus
    Call rtfSQL.ExecuteCmd(cmCmdPaste)
End Sub

Private Sub mnuViewChild_Click(Index As Integer)
Dim sText As String

Select Case Index
    Case Is = 0
        GridSort 0
    Case Is = 1
        GridSort 1
    Case Is = 3
        frmOptions.Show vbModal
        MenuSet 1
    Case Is = 5
        ' sText = Trim(Mid$(sbStatus.Panels(3), 5))
        sText = Trim(Mid$(sbStatus.Panels("DB"), 5))
        If sText = "" Then sText = "Database"
        TreeInitialize
        tvDB.Nodes(1).Text = sText
End Select
End Sub

Private Sub mnuWindowCommand_Click()
    rtfSQL.SetFocus

End Sub

Private Sub mnuWindowHistory_Click()
    tbsList.Tabs(3).Selected = True
End Sub

Private Sub mnuWindowResult_Click()
    grdSQL.SetFocus

End Sub

Private Sub mnuWindowSchema_Click()
    tbsList.Tabs(1).Selected = True
End Sub

Private Sub mnuWindowStatus_Click()
    tbsList.Tabs(2).Selected = True
End Sub

Private Sub rtfSQL_GotFocus()
If Not bProcessing Then MenuSet 1
End Sub

Private Function rtfSQL_RClick(ByVal Control As CodeMaxCtl.ICodeMax) As Boolean
    If Not bProcessing Then
        rtfSQL.SetFocus
        mnuEditChild(0).Visible = True
        mnuEditChild(1).Visible = True
        PopupMenu mnuEdit
        mnuEditChild(0).Visible = False
        mnuEditChild(1).Visible = False
    End If
    ' now we have to return false to keep the built in from popping up
    rtfSQL_RClick = True
    
End Function

Private Sub rtfSQL_SelChange(ByVal Control As CodeMaxCtl.ICodeMax)
If bHighlight = True Then
    rtfSQL.HighlightedLine = rtfSQL.GetSel(True).EndLineNo
Else
    rtfSQL.HighlightedLine = -1
End If
sbStatus.Panels("Line").Text = "Ln: " & rtfSQL.GetSel(True).EndLineNo + 1 & "     "
sbStatus.Panels("Col").Text = "Col: " & rtfSQL.GetSel(True).EndColNo + 1 & "     "
End Sub


'Private Sub rtfSQL_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim iErrorCount As Integer
'Dim sMsg As String
'iErrorCount = 0
'On Error GoTo AppError
'If (Shift And vbCtrlMask) > 0 Then
'    If Chr(KeyCode) = "V" Or Chr(KeyCode) = "v" And Not bAmPasting Then
'        bAmPasting = True
'        Dim sText As String
'        sText = Clipboard.GetText
'        Clipboard.Clear
'        Clipboard.SetText sText
'        bAmPasting = False
'    End If
'End If
'Exit Sub
'AppError:
'    If Err.Number <> 0 Then
'        If Err.Number = 521 And iErrorCount < 5 Then
'            iErrorCount = iErrorCount + 1
'            Err.Clear
'            Resume
'        End If
'        StatsDisplay 5
'        Screen.MousePointer = vbDefault
'        sMsg = "Error:  " & Str(Err.Number) & vbCrLf & Err.Description
'        MsgBox sMsg, vbExclamation + vbOKOnly
'        Err.Clear
'    End If
'    MenuSet 1
'Exit Sub
'End Sub

Private Sub tbsList_Click()
On Error Resume Next
Select Case tbsList.SelectedItem.Key
    Case "History"
        grdHistory.ZOrder 0
        grdHistory.SetFocus
    Case "Schema"
        tvDB.ZOrder 0
        tvDB.SetFocus
    Case "Messages"
        lstStatus.ZOrder 0
        lstStatus.SetFocus
End Select
On Error GoTo 0
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"
            FileNew
        Case "Open"
            FileOpen False
        Case "Save"
            FileSaveAs
        Case "Print"
            mnuFileChild_Click 6
        Case "Export"
            GridSave
        Case "ExportView"
            GridSave True
        Case "Cut"
            mnuEditChild_Click 5
        Case "Copy"
            mnuEditChild_Click 6
        Case "Paste"
            mnuEditChild_Click 7
        Case "Undo"
            mnuEditChild_Click 2
        Case "Redo"
            mnuEditChild_Click 3
        Case "Find"
            mnuEditChild_Click 14
        Case "Replace"
            mnuEditChild_Click 16
        Case "Connect"
            LoginShow
        Case "Disconnect"
            mnuDBChild_Click 4
        Case "ExecuteSQL"
            SQLParse
        Case "Cancel"
            bCancel = True
        Case "SortAscending"
            GridSort 0
        Case "SortDescending"
            GridSort 1
        Case "Indent"
            mnuEditChildFormat_Click 8
        Case "Outdent"
            mnuEditChildFormat_Click 9
        Case "Comment"
            CommentBlock
        Case "Uncomment"
            UnCommentBlock
        Case "BookToggle"
            mnuEditChildGoTo_Click 4
        Case "BookNext"
            mnuEditChildGoTo_Click 5
        Case "BookPrevious"
            mnuEditChildGoTo_Click 6
        Case "BookDel"
            mnuEditChildGoTo_Click 7
        Case "Log"
            If bLogFile = False Then
                bLogFile = True
                SaveSetting App.Title, "Options", "LogToFile", "1"
                tbToolbar.Buttons("Log").Value = tbrPressed
            Else
                bLogFile = False
                SaveSetting App.Title, "Options", "LogToFile", "0"
                tbToolbar.Buttons("Log").Value = tbrUnpressed
            End If
    End Select
End Sub

Private Sub tmrConnect_Timer()
    ' we need this timer, otherwise the form will not show before the login
    tmrConnect.Enabled = False
    If bDebug Then
        DebugWrite "Timer Disabled"
        DebugWrite "Showing Login"
    End If
    LoginShow
End Sub


Private Sub tmrMenu_Timer()
    If Not bProcessing Then
        MenuSet 1
        SetCaption
    Else
        DoEvents
    End If
End Sub

Private Sub tvDB_GotFocus()
If Not bProcessing Then MenuSet 1
End Sub

Private Sub tvDB_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Asc(KeyCode) = 57 Then TreeContext 1, 1
If (Shift And vbCtrlMask) > 0 Then
    If Chr(KeyCode) = "C" Or Chr(KeyCode) = "c" Then
        mnuEditChild_Click 6
    ElseIf Chr(KeyCode) = "V" Or Chr(KeyCode) = "v" Then
        mnuEditChild_Click 7
    End If
End If
If Shift = 0 Then
    If Not bProcessing And KeyCode = 93 Then
        tvDB.SetFocus
        SetTreeMenu
        PopupMenu mnuPopTreeObject, vbPopupMenuLeftAlign, _
            frList.Left + frList.Width - 100, _
            frList.Top + 100
    End If
End If

End Sub


Private Sub tvDB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Not bProcessing Then
        SetTreeMenu
        PopupMenu mnuPopTreeObject
    End If
End If
End Sub

Private Sub tvDB_NodeClick(ByVal Node As MSComctlLib.Node)
    TreeNodeClick Node.Key
End Sub


