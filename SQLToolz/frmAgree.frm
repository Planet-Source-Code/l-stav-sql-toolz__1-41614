VERSION 5.00
Begin VB.Form frmAgree 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "frmAgree"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ControlBox      =   0   'False
   Icon            =   "frmAgree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "I Agree"
      Height          =   345
      Left            =   1980
      TabIndex        =   0
      Top             =   5040
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "I Do NOT Agree"
      Height          =   345
      Left            =   3480
      TabIndex        =   1
      Top             =   5040
      Width           =   1440
   End
   Begin VB.TextBox txtAgree 
      BackColor       =   &H8000000F&
      Height          =   4875
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "frmAgree.frx":000C
      Top             =   60
      Width           =   6915
   End
End
Attribute VB_Name = "frmAgree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sText As String
Dim bLoading As Boolean

Private Sub cmdCancel_Click()
bGlobalRetVal = False
End
End Sub


Private Sub cmdOK_Click()
SaveSetting App.Title, "Options", "License", "L1C6U" & App.Major & "5A8C2U" & App.Minor & App.Revision
bGlobalRetVal = True
    
    ' on the very first run, we will show help
'    If GetSetting(App.Title, "Options", "ShowHelp", "1") = 1 Then
'    ShowHelpContents
    SaveSetting App.Title, "Options", "ShowHelp", "0"
'    End If

Unload Me
End Sub


Private Sub Form_Load()
    
Me.Caption = "License Agreement - " + App.Title

bLoading = True

sText = "By using SQL Toolz you give your express agreement to the"
sText = sText & vbCrLf & "below copyright, license and disclaimer notice."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "SQL Toolz is not " & Chr(34) & "public domain" & Chr(34) & " software."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "SQL Toolz is distributed as " & Chr(34) & "freeware" & Chr(34) & ", whose meaning is"
sText = sText & vbCrLf & "that although no charge is made for this present version,"
sText = sText & vbCrLf & "sole and exclusive copyright and ownership of SQL Toolz is"
sText = sText & vbCrLf & "retained by the author, Larry W. Stavinoha."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "Larry W. Stavinoha grants you, without charge, the right to"
sText = sText & vbCrLf & "reproduce, distribute and use this version of SQL Toolz on"
sText = sText & vbCrLf & "the express condition that you do not receive any payment,"
sText = sText & vbCrLf & "commercial benefit or other consideration for any such act,"
sText = sText & vbCrLf & "other than a nominal media charge, and that the wording of "
sText = sText & vbCrLf & "this copyright notice and disclaimer is not changed in any "
sText = sText & vbCrLf & "way within the documentation, software or other media."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "Larry W. Stavinoha grants you, without charge, the right to"
sText = sText & vbCrLf & "use the code contained within this software for educational "
sText = sText & vbCrLf & "or commercial purposes.  Larry W. Stavinoha grants you, "
sText = sText & vbCrLf & "without charge, the right to create derivative works based "
sText = sText & vbCrLf & "on SQL Toolz, as long as credit is given to the original "
sText = sText & vbCrLf & "author and as long as the name of the derivative is not SQL Toolz."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "You also acknowledge that SQL Toolz remains the"
sText = sText & vbCrLf & "intellectual property of Larry W. Stavinoha."
sText = sText & vbCrLf & " "
sText = sText & vbCrLf & "THIS SOFTWARE IS PROVIDED " & Chr(34) & "AS IS" & Chr(34) & " WITHOUT WARRANTY OF ANY"
sText = sText & vbCrLf & "KIND, EITHER EXPRESS OR IMPLIED, INCLUDING, WITHOUT"
sText = sText & vbCrLf & "LIMITATION, ANY WARRANTY OF MERCHANTABILITY AND FITNESS FOR"
sText = sText & vbCrLf & "A PARTICULAR PURPOSE. IN NO EVENT SHALL LARRY W. STAVINOHA"
sText = sText & vbCrLf & "BE LIABLE FOR ANY DAMAGES ARISING OUT OF THE USE OR"
sText = sText & vbCrLf & "INABILITY TO USE THE SOFTWARE, EVEN IF LARRY W. STAVINOHA"
sText = sText & vbCrLf & "HAS BEEN ADVISED OF THE LIKELIHOOD OF SUCH DAMAGES"
sText = sText & vbCrLf & "OCCURRING. LARRY W. STAVINOHA SHALL NOT BE LIABLE FOR ANY"
sText = sText & vbCrLf & "LOSS, DAMAGES OR COSTS, ARISING OUT OF, BUT NOT LIMITED TO,"
sText = sText & vbCrLf & "LOST PROFITS OR REVENUE, LOSS OF USE OF THE SOFTWARE, LOSS"
sText = sText & vbCrLf & "OF DATA OR EQUIPMENT, THE COSTS OF RECOVERING SOFTWARE, DATA"
sText = sText & vbCrLf & "OR EQUIPMENT, THE COST OF SUBSTITUTE SOFTWARE, DATA OR"
sText = sText & vbCrLf & "EQUIPMENT OR CLAIMS BY THIRD PARTIES, OR OTHER SIMILAR"
sText = sText & vbCrLf & "COSTS."
sText = sText & vbCrLf & " "


txtAgree.Text = sText
bLoading = False
cmdOK.Left = Me.Width / 2 - cmdOK.Width - 250
cmdCancel.Left = cmdOK.Left + cmdOK.Width + 300

End Sub


Private Sub txtAgree_Change()
    If Not bLoading Then
        bLoading = True
        txtAgree.Text = sText
        bLoading = False
    End If
End Sub

Private Sub txtAgree_GotFocus()
    cmdOK.SetFocus
End Sub


