VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmIPStealer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "IP thing"
   ClientHeight    =   6480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6480
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Text            =   "80"
      Top             =   1200
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deactivate"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "__"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
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
      Left            =   2760
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "NetScan IP Grabber"
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
      TabIndex        =   9
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "255.255.255.255"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No Connection"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Off"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   3000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCelcrapIP.frx":0000
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
End
Attribute VB_Name = "frmIPStealer"
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
Private Sub Command1_Click()
If Text1.Text = "" Then
    MsgBox "Please enter a port", vbExclamation
    Exit Sub
End If
If Winsock1.State = sckListening Then GoTo Finish
Winsock1.LocalPort = Text1.Text
Winsock1.Listen
Finish:
Label3.Caption = "On"
Label4.Caption = "No Connection"
End Sub

Private Sub Command2_Click()
Winsock1.Close
Label3.Caption = "Off"
End Sub

Private Sub Form_Load()
Label5.Caption = Winsock1.LocalIP
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragFLag = 0
End Sub

Private Sub Label7_Click()
Unload Me
End Sub

Private Sub Label8_Click()
Me.WindowState = 1
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

If Winsock1.State <> sckClosed Then
    Winsock1.Close
    DoEvents
    Winsock1.Accept requestID
    Label4.Caption = "Connected"
End If

    
    Label4.Caption = "No Connection"
    List1.AddItem Winsock1.RemoteHostIP + " Connected to you"
    List1.AddItem " "
    Winsock1.Close


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Check1.Value = 1 Then
    Winsock1.SendData "<html><center><font size=8>HTTP 404<p><br><h1><font size=2>Apache/1.32 Internal server error</font></html>"
End If
End Sub

Private Sub Winsock1_SendComplete()
 List1.AddItem "And HTTP Emulated"
 Label4.Caption = "No Connection"
 
    List1.AddItem " "
 Winsock1.Close
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
