VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Web Voting Poll"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   6360
      _ExtentX        =   11218
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
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5100
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4620
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   6075
      Begin VB.CommandButton Command5 
         Caption         =   "About..."
         Height          =   315
         Left            =   4140
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   315
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "View current web poll results"
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vote"
      Height          =   735
      Left            =   180
      TabIndex        =   6
      Top             =   480
      Width           =   6075
      Begin VB.CommandButton Command2 
         Caption         =   "No"
         Height          =   315
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Yes"
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
   End
   Begin VB.Label Label1 
      Caption         =   "Do you like PlanetSourceCode.com's web page layout?"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Commun(5) As Com
Dim CommunState As Integer
Dim Site As String
Dim Username As String
Dim Password As String
Dim Remotefile As String
Dim Localfile As String
Dim Buffersize As Long
Dim CloseAfterSend As Boolean


Private Sub Command1_Click()
        '### Vote once protection ###
'   SaveSetting "Voting", "Voted", "Voted", "True"
        '### Vote once protection ###

On Error GoTo voteyeserror
StatusBar1.SimpleText = "Adding vote to current results..."
Command1.Enabled = False
Command2.Enabled = False

yesnum = Inet1.OpenURL("# Example # http://www.myserver.com/myaccount/myvotefolder/my yes votes.dat")

yesnumt = yesnum + 1

Open "C:\voteyes.dat" For Output As #99
Print #99, yesnumt
Close #99

'------------------------------------------------
    Site = "" '# Private # This field is where u put your FTP site. I.E. ftp.myserver.com
    Username = "" '# Private # Put your account/username in this field
    Password = "" '# Private # Put your account/username password in this field
    Localfile = "C:\voteyes.dat" 'This field is the program's temp file, and the file to be uploaded. (Calculation results)
'**** NOTE!!!
'You MUST have the remote files ( /web poll/yes.dat
'and /web poll/no.dat already in the FTP site
'and have their content ( In the file, the text ) zero (0)
    
    Remotefile = "/web poll/yes.dat" 'This is where the temp file is to be uploaded. (The names dont have to match)
    Commun(0).Reply = "220"
    Commun(0).BackCommand = "USER " + Username
    Commun(1).Reply = "331"
    Commun(1).BackCommand = "PASS " + Password
    Commun(2).Reply = "230"
    Commun(2).BackCommand = "TYPE I"
    Commun(3).Reply = "200"
    Commun(3).BackCommand = "PORT"
    Commun(4).Reply = "200"
    Commun(4).BackCommand = "STOR " + Remotefile
    Commun(5).Reply = ""
    Commun(5).BackCommand = ""
    Buffersize = 2920
    Dim Nr1 As Integer
    Dim Nr2 As Integer
    Dim LocalIP As String
    LocalIP = Winsock1.LocalIP


    Do Until InStr(LocalIP, ".") = 0
        LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
    Loop

    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1
    Commun(3).BackCommand = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
    Winsock2.Close


    Do Until Winsock2.State = 0


        DoEvents
        Loop

        Winsock2.LocalPort = (Nr1 * 256) + Nr2
        Winsock2.Listen
        Winsock1.Close


        Do Until Winsock1.State = 0


            DoEvents
            Loop

            Winsock1.RemoteHost = Site
            Winsock1.RemotePort = 21
            Winsock1.Connect
            CommunState = 0


            Do Until Winsock1.State = 7 Or Winsock1.State = 9


                DoEvents
                Loop



                Select Case Winsock1.State
                    Case 9
                    
                    StatusBar1.SimpleText = "Error connecting to server."
                    Case 7
                    Open Localfile For Binary As #1
                End Select
                Exit Sub
voteyeserror:
Command1.Enabled = True
Command2.Enabled = True
If Err = 13 Then
StatusBar1.SimpleText = "Error: Please check your internet connection."
Else
StatusBar1.SimpleText = "Error " & Err.Number & ": " & Err.Description & ". Please try voting again."
 End If
        End Sub



Private Sub Command2_Click()
        '### Vote once protection ###
'   SaveSetting "Voting", "Voted", "Voted", "True"
        '### Vote once protection ###

On Error GoTo votenoerror
StatusBar1.SimpleText = "Adding vote to current results..."
Command1.Enabled = False
Command2.Enabled = False

nonum = Inet1.OpenURL("# Example # http://www.myserver.com/myaccount/myvotefolder/my no votes.dat")

nonumt = nonum + 1
Open "C:\voteno.dat" For Output As #1
Print #1, nonumt
Close #1


'---------------------------------------------------

Site = "" '# Private # This field is where u put your FTP site. I.E. ftp.myserver.com
    Username = "" '# Private # Put your account/username in this field
    Password = "" '# Private # Put your account/username password in this field
    Localfile = "C:\voteyes.dat" 'This field is the program's temp file, and the file to be uploaded. (Calculation results)
    Remotefile = "/web poll/yes.dat" 'This is where the temp file is to be uploaded. (The names dont have to match)
    Commun(0).Reply = "220"
    Commun(0).BackCommand = "USER " + Username
    Commun(1).Reply = "331"
    Commun(1).BackCommand = "PASS " + Password
    Commun(2).Reply = "230"
    Commun(2).BackCommand = "TYPE I"
    Commun(3).Reply = "200"
    Commun(3).BackCommand = "PORT"
    Commun(4).Reply = "200"
    Commun(4).BackCommand = "STOR " + Remotefile
    Commun(5).Reply = ""
    Commun(5).BackCommand = ""
    Buffersize = 2920
    Dim Nr1 As Integer
    Dim Nr2 As Integer
    Dim LocalIP As String
    LocalIP = Winsock1.LocalIP


    Do Until InStr(LocalIP, ".") = 0
        LocalIP = Left(LocalIP, InStr(LocalIP, ".") - 1) + "," + Right(LocalIP, Len(LocalIP) - InStr(LocalIP, "."))
    Loop

    Randomize Timer
    Nr1 = Int(Rnd * 12) + 5
    Nr2 = Int(Rnd * 254) + 1
    Commun(3).BackCommand = "PORT " + LocalIP + "," + Trim(Str(Nr1)) + "," + Trim(Str(Nr2))
    Winsock2.Close


    Do Until Winsock2.State = 0


        DoEvents
        Loop

        Winsock2.LocalPort = (Nr1 * 256) + Nr2
        Winsock2.Listen
        Winsock1.Close


        Do Until Winsock1.State = 0


            DoEvents
            Loop

            Winsock1.RemoteHost = Site
            Winsock1.RemotePort = 21
            Winsock1.Connect
            CommunState = 0


            Do Until Winsock1.State = 7 Or Winsock1.State = 9


                DoEvents
                Loop



                Select Case Winsock1.State
                    Case 9
                    StatusBar1.SimpleText = "Error connecting to server."
                    
                    Case 7
                    Open Localfile For Binary As #1
                End Select
Exit Sub
votenoerror:
Command1.Enabled = True
Command2.Enabled = True
If Err = 13 Then
StatusBar1.SimpleText = "Error: Please check your internet connection."
Else
StatusBar1.SimpleText = "Error " & Err.Number & ": " & Err.Description & ". Please try voting again."
End If
End Sub

Private Sub Command3_Click()
'Load and show the results window
Load Form2
Form2.Show
End Sub

Private Sub Command5_Click()
'Load and show the about box
Load Form3
Form3.Show
End Sub

Private Sub Form_Load()

        '### Vote once protection ###
'v = GetSetting("Voting", "Voted", "Voted", "False")
'
'If v = True Then
'Command1.Enabled = False
'Command2.Enabled = False
'MsgBox "You have already voted, and can not vote again.", vbExclamation, "Already voted"
'End If
        '### Vote once protection ###

MsgBox "Please note: I have removed my server passwords and usernames from this project, so if you want to try it out, please run 'web poll.exe' in the zip file." & vbNewLine _
& " -={ Zer0 Cool }=-", vbInformation, "Notice:"

MsgBox "I have also added an extra feature in the program, you can have the program only allow the user to vote once. Which can be useful for surveys. In the EXE and this demo, it is disabled, to enable it, search the entire project for:" & vbNewLine & "### Vote once protection ###" & vbNewLine & "And clear the ' from all of the statements in between the two ### Vote once protection ### remarks. Have fun!" & vbNewLine & "-={ Zer0 Cool }=-", vbInformation, "New feature"

'Make the temp files
Open "C:\voteyes.dat" For Output As #1
Print #1, "1"
Close #1

Open "C:\voteno.dat" For Output As #2
Print #2, "1"
Close #2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Delete the temp files
Kill "C:\voteyes.dat"
Kill "C:\voteno.dat"
'Terminate the application
End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'Declare variables
    Dim tmpS As String
    Winsock1.GetData tmpS, , bytesTotal
    Debug.Print tmpS;


'Select what command to send
Select Case Left(tmpS, 3)
        Case Commun(CommunState).Reply
        Winsock1.SendData Commun(CommunState).BackCommand + Chr(13) + Chr(10)
        Debug.Print Commun(CommunState).BackCommand
        CommunState = CommunState + 1
        Case "150"


        Do Until Winsock2.State = 7


            DoEvents
            Loop

'Send the required Data
            SendNextData
            Case "226"
            Winsock1.Close


            Do Until Winsock1.State = 0


                DoEvents
                Loop

'Set the current status
StatusBar1.SimpleText = "Done."
'Display a message to show the vote was a success

        '### Vote once protection ###
'         Command1.Enabled = False
'         Command2.Enabled = False
        '### Vote once protection ###


                MsgBox "Thank you for voting", vbInformation, "Vote done"
'Enable the command buttons
Command1.Enabled = True
Command2.Enabled = True

Case Else
               
            End Select

    End Sub



Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)

'Close winsock2's socket
    Winsock2.Close


    Do Until Winsock2.State = 0


        DoEvents
        Loop

'Accept the connection if their is one
        Winsock2.Accept requestID


        Do Until Winsock2.State = 7


            DoEvents
            Loop

        End Sub



Sub SendNextData()

    Dim Take As Long
    Dim Buffer As String
    If LOF(1) - Seek(1) < Buffersize Then Take = LOF(1) - Seek(1) + 1 Else Take = Buffersize
    Buffer = Input(Take, 1)
    Winsock2.SendData Buffer


    If Take < Buffersize Then
        Close #1
        CloseAfterSend = True
    End If

    On Error Resume Next
    Label1 = Trim(Str(Seek(1))) + "/" + Trim(Str(LOF(1)))
    On Error GoTo 0
End Sub



Private Sub Winsock2_SendComplete()


'Variable arguments

    If CloseAfterSend = True Then
        Winsock2.Close


        Do Until Winsock2.State = 0


            DoEvents
            Loop

            CloseAfterSend = False
        Else
'Send data
            SendNextData
        End If

    End Sub

Private Sub Command4_Click()
'Delete the temp files
Kill "C:\voteyes.dat"
Kill "C:\voteno.dat"
'Terminate the application
End
End Sub
