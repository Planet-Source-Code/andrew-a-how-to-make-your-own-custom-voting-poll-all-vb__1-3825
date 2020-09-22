VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4260
      Top             =   3180
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   3180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ICQ # 14344635"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "If you use this code please let me know and add me in the credits."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   4560
      Top             =   1980
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   4680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enough talk, this program was made by -={ Zer0 Cool }=- with thanks to Kristian Trenskow for his code."
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   2160
      Width           =   4515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"about2.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"about2.frx":00B8
      ForeColor       =   &H0000FF00&
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Unload the current form
Unload Me
End Sub

Private Sub Timer1_Timer()
'If the yellow box's position is <= 0 then
If Shape1.Left <= 0 Then
'Enable timer2
Timer2.Enabled = True
'Disable timer1
Timer1.Enabled = False
'Exit the sub
Exit Sub
End If

'If the yellow box's position is <= 4560 then
If Shape1.Left <= 4560 Then
'Move the box's position - 100
Shape1.Left = Shape1.Left - 100
Exit Sub
End If

End Sub

Private Sub Timer2_Timer()

'If the yellow box's position is >= 4560 then
If Shape1.Left >= 4560 Then
'Disable timer1
Timer1.Enabled = True
'Disable timer2
Timer2.Enabled = False
Exit Sub
End If

'If the yellow box's position is <> 0 then
If Shape1.Left <> 0 Then
'Move the box's position + 100
Shape1.Left = Shape1.Left + 100
Exit Sub
End If


End Sub
