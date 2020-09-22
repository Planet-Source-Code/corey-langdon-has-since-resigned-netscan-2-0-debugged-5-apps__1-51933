VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTelnetXS 
   Caption         =   "NetScan 2.0 Telnet"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   4455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   6255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   6240
      Top             =   2760
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Command Line"
      Top             =   4440
      Width           =   6255
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSvLg 
         Caption         =   "Save Log   . . . . . . "
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuTelnet 
      Caption         =   "Telnet"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect     . . . . . "
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuListen 
         Caption         =   "Listen         . . . . . "
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuClAll 
         Caption         =   "Close All     . . . . . "
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuClrLg 
         Caption         =   "Clear Log"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuTut 
         Caption         =   "Tutorials"
         Begin VB.Menu mnuPOP 
            Caption         =   "POP Tutorial"
         End
         Begin VB.Menu mnuSMTP 
            Caption         =   "SMTP Tutorial"
         End
         Begin VB.Menu mnuPrtUse 
            Caption         =   "Ports and Uses"
         End
         Begin VB.Menu mnuHaxxor 
            Caption         =   "Haxx0r"
            Enabled         =   0   'False
            Shortcut        =   ^B
            Visible         =   0   'False
         End
         Begin VB.Menu mnuHTTP 
            Caption         =   "HTTP Tutorial"
         End
      End
   End
End
Attribute VB_Name = "frmTelnetXS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSCKHOST
Dim WSCKPRT

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Dim SVDTA
SVDTA = InputBox("Log name:")
If SVDTA <> "" Then
    Open App.Path & "\logs\" & SVDTA & ".txt" For Output As #1
    Print #1, Text2.Text
    Close #1
End If
End Sub

Private Sub Form_Resize()
Text3.Width = Me.Width - 100
Text3.Top = Me.Height - 1095
Text2.Height = Me.Height - 1095
Text2.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me

End Sub

Private Sub mnuClAll_Click()
Winsock1.Close
Text2.Text = Text2.Text + vbNewLine + "All Sockets Closed" + vbNewLine
End Sub

Private Sub mnuClrLg_Click()
Text2.Text = ""
End Sub

Private Sub mnuConnect_Click()
WSCKHOST = InputBox("Host")
WSCKPRT = InputBox("Port")

Winsock1.Connect WSCKHOST, WSCKPRT
DoEvents:
DoEvents:
DoEvents:
DoEvents
End Sub

Private Sub mnuLgEd_Click()
Form3.Show
End Sub

Private Sub mnuHaxxor_Click()
HaXxOr.Show
End Sub

Private Sub mnuHTTP_Click()
Shell ("c:\WINDOWS\explorer.exe " & App.Path & "\HTTP.html")
End Sub

Private Sub mnuListen_Click()
Winsock1.Close
WSCKPRT = InputBox("Port:")
Winsock1.LocalPort = WSCKPRT
Winsock1.Listen
Text2.Text = Text2.Text & "Listening on port " & WSCKPRT & vbNewLine & Time & vbNewLine & Date & vbNewLine
End Sub

Private Sub mnuPOP_Click()
Shell ("c:\WINDOWS\explorer.exe " & App.Path & "\Tutorials\POP.html")
End Sub

Private Sub mnuPrtUse_Click()
Shell ("c:\WINDOWS\explorer.exe " & App.Path & "\Tutorials\Prt.html")
End Sub

Private Sub mnuSMTP_Click()
Shell ("c:\WINDOWS\explorer.exe " & App.Path & "\Tutorials\SMTP.html")
End Sub

Private Sub mnuSvLg_Click()
Dim HTDOC As Variant
Dim CurLin
Dim FinalHTML As String
CommonDialog1.Filter = "HTML Files (*.html)|*.html"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    
    HTDOC = Text2.Text
    HTDOC = Split(HTDOC, vbNewLine)
    DoEvents
    FinalHTML = "<html><body><font size=3 face=terminal><p><center>Telnet XS<br>Date: " & Date & "<br>Time: " & Time & "<p><br> _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _<p><br></center>"
    For i = 0 To UBound(HTDOC)
        CurLin = HTDOC(i)
        FinalHTML = FinalHTML & CurLin & "<br>"
    Next i
    DoEvents
    FinalHTML = FinalHTML & "</body></html>"
    Print #1, FinalHTML
    Close #1
End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Text2.Text = Text2.Text + Text3.Text & vbNewLine
    Winsock1.SendData Text3.Text & vbCrLf
    If Text3.Text = "ENDCONNECT" Then
        Winsock1.Close
        Text2.Text = Text2.Text + vbCr & "Socket Killed on port " & Winsock1.LocalPort & vbNewLine & Time & vbNewLine & Date & vbNewLine
    End If
    
    Text3.Text = ""
End If

End Sub

Private Sub Winsock1_Close()
Text2.Text = Text2.Text & "Socket Closed" & vbNewLine & Time & vbNewLine & Date
End Sub

Private Sub Winsock1_Connect()
Text2.Text = Text2.Text + vbNewLine + "Connected to " & Winsock1.RemoteHost & " on port " & Winsock1.RemotePort & vbNewLine & Time & vbNewLine & Date & vbNewLine
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
Text2.Text = Text2.Text & vbNewLine & Winsock1.RemoteHostIP & "  Connected to you on port  " & Winsock1.LocalPort & vbNewLine & Time & vbNewLine & Date & vbNewLine
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Texter, vbString, bytesTotal
Text2.Text = Text2.Text + vbNewLine + Texter
End Sub

