VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmPingPong 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   1800
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Port"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1200
      Top             =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pong"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Host"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "NetScan Ping/Pong"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      X1              =   0
      X2              =   2640
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   2520
      Y2              =   120
   End
End
Attribute VB_Name = "frmPingPong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FX
Dim FY
Dim IY
Dim IX
Dim FileQuality
Dim DragFLag As Integer
Dim MsResponse As String
Dim WskDat As String
Dim PPmode As String
Private Sub Command1_Click()
If Winsock1.State = sckConnected Then Exit Sub
Winsock1.Connect Text1.Text, Text2.Text
Timer2.Enabled = True
DoEvents
PPmode = "ping"
Timer1.Enabled = True
WskDat = ""
End Sub

Private Sub Command2_Click()
If Winsock1.State = sckConnected Then Exit Sub
Winsock1.Connect Text1.Text, Text2.Text
Timer2.Enabled = True
DoEvents
PPmode = "pong"
Timer1.Enabled = True
WskDat = ""
End Sub

Private Sub Form_Load()
MsResponse = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragFLag = 0 Then
        IX = X: IY = Y
        FX = Me.Left: FY = Me.Top
        DragFLag = 1
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DragFLag = 1 Then
        Me.Move FX + (X - IX), FY + (Y - IY)
        FX = Me.Left: FY = Me.Top
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragFLag = 0
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Me.WindowState = 1
End Sub

Private Sub Timer1_Timer()
If WskDat = "" Then MsgBox "The remote host has not responded in 10 seconds.", vbExclamation, "NetScan Ping/Pong"
End Sub

Private Sub Timer2_Timer()
MsResponse = MsResponse + 1
End Sub

Private Sub Winsock1_Connect()
If PPmode = "ping" Then Winsock1.SendData Ping
If PPmode = "pong" Then Winsock1.SendData Pong
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData WskDat, vbString, bytesTotal
Timer1.Enabled = False
Timer2.Enabled = False
MsgBox Text1.Text & " responded in " & MsResponse & " milliseconds.", vbInformation, "NetScan Ping/Pong"
MsResponse = 0
Winsock1.Close
End Sub

