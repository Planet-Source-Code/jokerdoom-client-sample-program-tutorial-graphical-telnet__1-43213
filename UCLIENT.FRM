VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cory Dambach's Universal Client"
   ClientHeight    =   6765
   ClientLeft      =   3000
   ClientTop       =   3885
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10590
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   10050
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Disconnect 
      Caption         =   "&Disconnect"
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton bSEND 
      Caption         =   "&Send"
      Height          =   255
      Left            =   8640
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox MyData 
      Height          =   5535
      Left            =   7800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox TheirData 
      Height          =   5535
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7815
   End
   Begin VB.CommandButton Connect 
      Caption         =   "&Connect"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox OutData 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   840
      Width           =   8535
   End
   Begin VB.TextBox Port 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Port"
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox Host 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Host"
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6795
      TabIndex        =   8
      Top             =   270
      Width           =   3720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----~~~~~***%%%Cory's Client Tutorial / Sample Application ***~~~~~----'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'I hope you have some good use for this simple program and remember VOTE FOR ME!!!

Private Sub bSEND_Click()
On Error GoTo Ignore:
'I include this so nothing really ugly goes wrong
'You might think its stupid, but it will help you if you have windows compatibility problems trust me
Winsock.SendData OutData.Text + vbCrLf 'Send the data
MyData.Text = MyData.Text + OutData.Text + vbCrLf 'Put the data you sent in the appropriate box
Ignore:
End Sub

Private Sub Connect_Click()
On Error GoTo Ah:
Label1.Caption = "Connecting"
Winsock.Connect Host.Text, Port.Text 'Connect to the supplied host and port
'Then We Update the buttons
Disconnect.Enabled = True
Connect.Enabled = False
Connect.Default = False
TheirData.Text = ""
MyData.Text = ""
Exit Sub
Ah:
Label1.Caption = "Disconnected"
MsgBox (Err.Description)
End Sub

Private Sub Disconnect_Click()
On Error GoTo Ignore:
Winsock.Close 'Close the socket
'Update the User Interface again
Connect.Enabled = True
Connect.Default = True
bSEND.Enabled = False
bSEND.Default = False
Disconnect.Enabled = False
Label1.Caption = "Disconnected"
Ignore:
End Sub

Private Sub Form_Load()
On Error GoTo Help:
'This load gets the command line stuff if any
'it should be passed like this uclient host port
'ie: uclient 127.0.0.1 80
bSEND.Enabled = False
OutData.Text = ""
MyData.Text = ""
TheirData.Text = ""
Connect.Default = True
Disconnect.Enabled = False
OutData.TabStop = False
OutData.Enabled = False
Dim ipandport() As String
'Parsing the command line
If Len(Command) > 4 Then
    ipandport = Split(Command, " ")
    Host.Text = ipandport(0)
    Port.Text = ipandport(1)
    Connect_Click
    Exit Sub
End If
Host.SelLength = Len(Host.Text) 'Selecting the text lets the user get on with life
Exit Sub
Help:
    MsgBox ("Command Line Contains invalid data") 'Notify the user but still allow them to do stuff the normal way
End Sub

'After voting for me take this code out
Private Sub Form_Unload(Cancel As Integer)
MsgBox ("Thank you, Please vote for me")
End Sub

Private Sub MyData_Change()
MyData.SelStart = Len(MyData.Text) 'scroll always to the newest data
End Sub

Private Sub Port_GotFocus()
Port.SelLength = Len(Port.Text) 'select all the text so they can keep on with their life
End Sub

Private Sub TheirData_Change()
TheirData.SelStart = Len(TheirData.Text) 'Scroll always to the newest data
End Sub



Private Sub Winsock_Close()
On Error GoTo Ignore:
MsgBox ("Connection Closed")
Winsock.Close 'Make sure our socket is closed
'Then update the Buttons and labels and stuff AGAIN!
Connect.Enabled = True
Connect.Default = True
bSEND.Enabled = False
bSEND.Default = False
Disconnect.Enabled = False
OutData.TabStop = False
Label1.Caption = "Disconnected"
Ignore:
End Sub

Private Sub Winsock_Connect()
On Error GoTo Ignore:
'WE CONNECTED! GREAT! now what?
'Tell the user STUPID!
Label1.Caption = "Connected"
bSEND.Enabled = True 'Enable the send button
bSEND.Default = True 'Make the send button default so the user just has to press enter to send stuff
OutData.Enabled = True
OutData.TabStop = True
OutData.SetFocus
Ignore:
End Sub

'Wanna hear a joke?
'Two guys walk into a bar, one ducks the other says OW!
'You shouldn't be laughing

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
On Error GoTo Ignore:
If Winsock.State = 7 Then '7 Is the connected state
    Dim inbuff As String
    'The GetData function places the data it recieves into the variable you specify
    'which in this case is the inbuff variable
    Winsock.GetData inbuff
    TheirData.Text = TheirData.Text + inbuff 'Show the user what we recieved
End If
Ignore:
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Ignore:
    Label1.Caption = "Disconnected"
    MsgBox ("Winsock Error: " + Description)
    Disconnect_Click
Ignore:
End Sub

'VOTE FOR ME
