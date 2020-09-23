VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Open Streaming File..."
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Playerv2-3.frx":0000
   ScaleHeight     =   1320
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-3.frx":30A2
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   16
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-3.frx":6BAC
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   18
         ToolTipText     =   "Close Streaming Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open Streaming File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   17
         Top             =   10
         Width           =   1560
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-3.frx":6E8E
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   13
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-3.frx":A758
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   15
         ToolTipText     =   "Close Streaming Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Streaming File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   10
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-3.frx":AA3A
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   10
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-3.frx":E664
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   12
         ToolTipText     =   "Close Streaming Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Streaming File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   10
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-3.frx":E946
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   6
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-3.frx":127F8
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   8
         ToolTipText     =   "Close Streaming Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Streaming File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   15
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-3.frx":12ADA
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   4
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-3.frx":166BC
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   9
         ToolTipText     =   "Close Streaming Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Open Streaming File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   15
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Type location :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.ActiveMovie1.FileName = Text1.Text
Form1.Label1.Caption = "Streaming "
Form1.Command4.Enabled = False
Form1.Command5.Enabled = False
Form3.Enabled = False
Form3.Visible = False
Form1.Visible = True
Form1.Enabled = True
Form1.Visible = True
Form4.Enabled = True

Unload Form3
End Sub

Private Sub Command2_Click()
Form3.Enabled = False
Form3.Visible = False
Form2.Visible = False
Form1.Enabled = True
Form1.Visible = True
Form4.Enabled = True
Unload Form3
End Sub

Private Sub Form_Load()
If Form1.Picture5.Visible = True Then Form3.Picture2.Visible = True
If Form1.Picture7.Visible = True Then Form3.Picture1.Visible = True
If Form1.Picture11.Visible = True Then Form3.Picture5.Visible = True
If Form1.Picture15.Visible = True Then Form3.Picture7.Visible = True
Form3.Top = Screen.Height / 2 - Form3.Height / 2
Form3.Left = Screen.Width / 2 - Form3.Width / 2
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture10_Click()
Command2_Click

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture3_Click()
Command2_Click
End Sub

Private Sub Picture4_Click()
Command2_Click
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture6_Click()
Command2_Click
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture8_Click()
Command2_Click

End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub
