VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "HTTP Subnet Scanner"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTrue 
      Height          =   3735
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5280
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Log.txt"
      Filter          =   "*.txt"
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Save Results"
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   855
   End
   Begin VB.Timer TimWait 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   2000
      Left            =   240
      Top             =   1320
   End
   Begin VB.TextBox TxtTimeOut 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6480
      TabIndex        =   8
      Text            =   "2"
      Top             =   1200
      Width           =   615
   End
   Begin InetCtlsObjects.Inet Inet1 
      Index           =   0
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdStop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Stop"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1500
      Left            =   600
      Top             =   360
   End
   Begin VB.TextBox TxtInfo 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton CmdGo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scan"
      Height          =   615
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Text            =   "163"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Text            =   "46"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox TxtIP 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Text            =   "80"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   17
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Open Port 80 Subnet Scanner"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Open Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Closed Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Time Out (In Seconds)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGo_Click()


If TxtIP(3).Text = 255 Then CmdStop_Click 'make sure it does not get past 255



lngnextip = TxtIP(3).Text 'set the next ip variable
   
   
For intI = 1 To Val(15) 'do this 15 times, for the 15 winsock connections (threads)
   
 
    Load Me.Winsock1(intI) 'load all the bits
    Load Me.Inet1(intI)
    Load Me.TimeOut(intI)
    Load Me.TimWait(intI)
        
      lngnextip = lngnextip + 1 'increase the ip number
        
    host = TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & lngnextip 'set the host
  
    Me.Winsock1(intI).Connect host, 80 'connect to host on port 80
   

    TxtIP(3).Text = lngnextip 'update the last ip block

    TimeOut(intI).Interval = TxtTimeOut.Text * 1000 'set the time out

    TimeOut(intI).Enabled = True 'enable the timeout timer


Next intI 'do it all over again
   


End Sub

Private Sub CmdSave_Click()
CD1.ShowSave 'show the commen dialog box

Open CD1.FileName For Output As #1 'save the file
    Print #1, TxtTrue.Text
Close #1
End Sub

Private Sub CmdStop_Click() 'stop the scan



For a = 1 To 15 'uhhhh my head god so confused doing this for some wierd reason


    Winsock1(a).Close 'close winsock
    TimeOut(a).Enabled = False 'stop the time out
    TimWait(a).Enabled = False 'stop the wait timer

Next a 'do it again


For intI = 1 To Val(15) 'now unload all the arrays so it don't bugger up if user clicks "start"  again
        Unload Me.Winsock1(intI)
        Unload Me.Inet1(intI)
        Unload Me.TimeOut(intI)
        Unload Me.TimWait(intI)
Next intI

End Sub




Private Sub TimeOut_Timer(Index As Integer) 'if connection times out (just quicker than the winsock timeout

TxtInfo.Text = TxtInfo.Text + Winsock1(Index).RemoteHost + "     " + "NO Server" + vbNewLine 'obvisuly there was no server
TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down
TimeOut(Index).Enabled = False 'stop time out


Winsock1(Index).Close 'close winsock (if not already done)


If TxtIP(3).Text >= 255 Then 'make sure it don't try and scan somthing out of range

     CmdStop_Click 'make sure it does not get past 255
     Exit Sub
     
End If


TimWait(Index).Enabled = True 'enable the wait timer (YOU NEED TO WAIT A SECOND OR TWO BEFORE WINSOCK RECONNECTS, OTHERWISE IT BUGGERS UP! (yes this did take me quite some time to work out!)
End Sub



Private Sub TimWait_Timer(Index As Integer) 'this stupid thign to wait before winsock reconnects

TimWait(Index).Enabled = False 'stop the timer

If TxtIP(3).Text >= 255 Then
    CmdStop_Click 'make sure it does not get past 255
    Exit Sub
     
End If


lngnextip = TxtIP(3).Text 'set hte variable
   
lngnextip = lngnextip + 1
        
host = TxtIP(0).Text & "." & TxtIP(1).Text & "." & TxtIP(2).Text & "." & lngnextip 'set the host
  
Me.Winsock1(Index).Connect host, 80 'connect to the host

TxtIP(3).Text = lngnextip 'update txtip(3)

TimeOut(Index).Enabled = False 'stop the time out (should already be stopped, but lets just make sure!)
TimeOut(Index).Interval = TxtTimeOut.Text * 1000 'set the time out

TimeOut(Index).Enabled = True 'enable the timeout timer


End Sub



Private Sub Winsock1_Connect(Index As Integer)
On Error Resume Next 'if winsock connects...

TimeOut(intI).Enabled = False 'disable the time out
site = Inet1(Index).OpenURL("http://" + Winsock1(Index).RemoteHost, icByteArray)  'set the site


servers = Inet1(Index).GetHeader("server") 'grab the server header (if there is one)

'you can grab any header you want, as long as it's there!, normal headers include, Content-type, Content-length, and Expires
'if you want to grab all the headers, then just use servers = Inet1.GetHeader()
'the headers you can get from kazaa are,
'X-Kazaa-Username
'X-Kazaa-Network
'X-Kazaa-IP
'X-Kazaa-SupernodeIP



If servers = "" Then 'if there isn't a server header, chances are it's going to be kazaa

    site = Inet1(Index).OpenURL("http://" + Winsock1(Index).RemoteHost, icByteArray) 'reconnect to grab the kaazaa username
    
    user = Inet1(Index).GetHeader("X-Kazaa-Username") 'get the kazaa username
    'so we try and get the kazaa username, (nothing important, just somthing i thought could be fun?!)
        If user = "" Then 'there was no user, maybe they arn't running a webserver, or kazaa...
            TxtTrue.Text = TxtTrue.Text + Winsock1(Index).RemoteHost + "     " + "Unkown Server" + vbNewLine
        Else
            'print out hte username
            TxtTrue.Text = TxtTrue.Text + Winsock1(Index).RemoteHost + "     " + "Kazaa Username: " + user + vbNewLine
        End If
        

Else
    'show the http server type
    TxtTrue.Text = TxtTrue.Text + Winsock1(Index).RemoteHost + "     " + servers + vbNewLine
End If


TxtTrue.SelStart = Len(TxtInfo.Text) 'scroll down the txtinfo box

Winsock1(Index).Close 'close winsock

Inet1(Index).Cancel 'close the inet, if not already closed.

TimWait(Index).Enabled = True 'wait a second or two before winsock reconnects

End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)


TxtInfo.Text = TxtInfo.Text + Winsock1(Index).RemoteHost + "     " + "NO Server" + vbNewLine 'winsock couldn't connect


TxtInfo.SelStart = Len(TxtInfo.Text) 'scroll down
TimeOut(Index).Enabled = False 'stop time out



Winsock1(Index).Close 'close the winsock (should already be closed, but sometimes isn't)

'close winsock (if not already done)

'CmdGo_Click 'start again
If TxtIP(3).Text >= 255 Then

     CmdStop_Click 'make sure it does not get past 255
     Exit Sub
     
End If

TimWait(Index).Enabled = True 'wait a while before reconnecting


End Sub




