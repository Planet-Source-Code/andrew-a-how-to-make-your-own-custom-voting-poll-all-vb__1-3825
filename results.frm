VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Results"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2520
      Top             =   3900
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4380
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Ready..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar nobar 
      Height          =   2835
      Left            =   3540
      TabIndex        =   5
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   5001
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar yesbar 
      Height          =   2835
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   5001
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
      Scrolling       =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4500
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   2940
      TabIndex        =   1
      Top             =   3720
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label VoteNo 
      Alignment       =   2  'Center
      Caption         =   "Refreshing..."
      Height          =   255
      Left            =   3300
      TabIndex        =   8
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label VoteYes 
      Alignment       =   2  'Center
      Caption         =   "Refreshing..."
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   2580
      X2              =   2580
      Y1              =   3300
      Y2              =   120
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   60
      X2              =   5100
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   3000
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   3195
      Index           =   0
      Left            =   60
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Unload the current form
Unload Me
End Sub

Private Sub Command2_Click()
Me.MousePointer = vbHourglass
'Change the captions
VoteYes.Caption = "Refreshing..."
VoteNo.Caption = "Refreshing..."
'Disable the timer
Timer1.Enabled = False
'Set up the error handler
On Error GoTo refresherror
'Display the current status
StatusBar1.SimpleText = "Refreshing..."
'Disable the close and refresh buttons, so when
'Refreshing, no errors will occure if one of the buttons
'are clicked
Command1.Enabled = False
Command2.Enabled = False
'Get the current results from the web site
yesdat = Inet1.OpenURL("# Example #: http://www.myserver.com/myaccount/my voting folder/the yes votes.dat")
nodat = Inet1.OpenURL("# Example #: http://www.myserver.com/myaccount/my voting folder/the no votes.dat")
'Set the progress bar's value's to the number of votes
yesbar.Value = yesdat
nobar.Value = nodat
'Set the caption to the number of votes
VoteYes.Caption = yesdat
VoteNo.Caption = nodat
'Enable the command buttons
Command1.Enabled = True
Command2.Enabled = True
'Set the current status
StatusBar1.SimpleText = "Done."
'Change the mouse pointer
Me.MousePointer = vbNormal
Exit Sub
refresherror:
'Re-enable the command buttons
Command1.Enabled = True
Command2.Enabled = True
'If the error number is 13 then
If Err = 13 Then
'Set the captions
VoteYes.Caption = "Error"
VoteNo.Caption = "Error"
StatusBar1.SimpleText = "Error: Please check your internet connection."
Else
'Set the captions
VoteNo.Caption = "Error"
VoteYes.Caption = "Error"
StatusBar1.SimpleText = "Error " & Err.Number & ": " & Err.Description & ". Please try refreshing again."
End If
'Change the mouse pointer
Me.MousePointer = vbNormal
End Sub


Private Sub Timer1_Timer()
'Set up the error handler
On Error GoTo timererror
'Disable the timer
Timer1.Enabled = False

Me.MousePointer = vbHourglass
'Display the current status
StatusBar1.SimpleText = "Refreshing..."
'Disable the close and refresh buttons, so when
'Refreshing, no errors will occure if one of the buttons
'are clicked
Command1.Enabled = False
Command2.Enabled = False
'Get the current results from the web site
yesdat = Inet1.OpenURL("# Example #: http://www.myserver.com/myaccount/my voting folder/the yes votes.dat")
nodat = Inet1.OpenURL("# Example #: http://www.myserver.com/myaccount/my voting folder/the no votes.dat")
'Set the progress bar's value's to the number of votes
yesbar.Value = yesdat
nobar.Value = nodat
'Set the caption to the number of votes
VoteYes.Caption = yesdat & " votes say yes"
VoteNo.Caption = nodat & " votes say no"
'Re-enable the command buttons
Command1.Enabled = True
Command2.Enabled = True
'Set the current status
StatusBar1.SimpleText = "Done."
'Change the mouse pointer
Me.MousePointer = vbNormal
Exit Sub
timererror:
'Re-enable the command buttons
Command1.Enabled = True
Command2.Enabled = True
'If error 13 then
If Err = 13 Then
StatusBar1.SimpleText = "Error: Please check your internet connection."
'Set the captions
yesnumber.Caption = "Error"
nonumber.Caption = "Error"
Else
'Set the captions
VoteYes.Caption = "Error"
VoteNo.Caption = "Error"
StatusBar1.SimpleText = "Error " & Err.Number & ": " & Err.Description & ". Please try refreshing again."
End If
Me.MousePointer = vbNormal
End Sub

