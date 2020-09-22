VERSION 5.00
Begin VB.Form frmNetScan 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "NetScan 2.0"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IP Grabber"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Telnet"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Port Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Host Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ping Pong"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "__"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
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
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "NetScan 2.0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmNetScan"
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
    Label4.ForeColor = &HFF00&
    Label5.ForeColor = &HFF00&
    Label6.ForeColor = &HFF00&
    Label7.ForeColor = &HFF00&
    Label8.ForeColor = &HFF00&

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

Private Sub Label4_Click()
frmPingPong.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF&
End Sub

Private Sub Label5_Click()
frmHostScan.Show
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFF&
End Sub

Private Sub Label6_Click()
frmPortScan.Show
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFF&
End Sub

Private Sub Label7_Click()
frmTelnetXS.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFF&
End Sub

Private Sub Label8_Click()
frmIPStealer.Show
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFF&
End Sub


