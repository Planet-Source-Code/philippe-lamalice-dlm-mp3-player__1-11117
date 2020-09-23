VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "DLM MP3 Player"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   LinkTopic       =   "Form5"
   Picture         =   "Playerv2-6.frx":0000
   ScaleHeight     =   495
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   720
      Top             =   480
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   2400
      Picture         =   "Playerv2-6.frx":476A
      ScaleHeight     =   210
      ScaleWidth      =   240
      TabIndex        =   2
      ToolTipText     =   "Return to Full Mode"
      Top             =   140
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "Playerv2-6.frx":4A4C
      ScaleHeight     =   255
      ScaleWidth      =   1710
      TabIndex        =   0
      Top             =   120
      Width           =   1710
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         ToolTipText     =   "Fwd Track"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "Pause"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   360
         TabIndex        =   4
         ToolTipText     =   "Play"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Rwd Track"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Play Time"
      Top             =   150
      Width           =   375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form5.Top = 10
Form5.Left = Screen.Width - Form5.Width + 10
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Label2_Click()
Form1.SecSlider.Value = 0
Form1.MinSlider.Value = 0
Form1.ActiveMovie1.CurrentPosition = 0
End Sub

Private Sub Label3_Click()
If Form1.ActiveMovie1.PlayState = mpPaused Then
    Form1.ActiveMovie1.Play
Else
    If Form1.ActiveMovie1.PlayState = mpPlaying Then
    Else
        If Form1.ActiveMovie1.FileName <> "" Then Form1.ActiveMovie1.Play
    End If
End If
End Sub

Private Sub Label4_Click()
If Form1.ActiveMovie1.PlayState = mpPaused Then
    Form1.ActiveMovie1.Play
Else
    If Form1.ActiveMovie1.PlayState = mpPlaying Then
        Form1.ActiveMovie1.Pause
    End If
End If
End Sub

Private Sub Label5_Click()
Form1.ActiveMovie1.Stop
End Sub

Private Sub Label6_Click()
Form1.SecSlider.Value = 0
Form1.MinSlider.Value = 0
Form1.ActiveMovie1.CurrentPosition = Form1.ActiveMovie1.Duration
End Sub

Private Sub Picture2_Click()
Form5.Visible = False
Form1.Visible = True
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Form1.Label13.Caption
End Sub
