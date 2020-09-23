VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Playerv2-5.frx":0000
   ScaleHeight     =   1740
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-5.frx":3FF2
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   22
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-5.frx":7AFC
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   23
         ToolTipText     =   "Close Options Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Top             =   10
         Width           =   540
      End
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Orange Scheme"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "Orange Color Scheme"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   360
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-5.frx":7DDE
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   16
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-5.frx":B6A8
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   20
         ToolTipText     =   "Close Options Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   10
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-5.frx":B98A
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   15
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-5.frx":F5B4
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   19
         ToolTipText     =   "Close Options Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   10
         Width           =   735
      End
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Silver Scheme"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      ToolTipText     =   "Silver Color Scheme"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Gold Scheme"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      ToolTipText     =   "Gold Color Scheme"
      Top             =   840
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-5.frx":F896
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   9
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-5.frx":13478
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   11
         ToolTipText     =   "Close Options Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   10
         Width           =   540
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Blue Scheme"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Blue Color Scheme"
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Green Scheme"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      ToolTipText     =   "Green Color Scheme"
      Top             =   360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      ToolTipText     =   "Reset Balance"
      Top             =   1320
      Width           =   735
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Balance"
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -9640
      Max             =   9640
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Auto-Attach"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Rewind when done playing"
      Top             =   600
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto-Play "
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Play on load/open"
      Top             =   360
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-5.frx":1375A
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-5.frx":1760C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   12
         ToolTipText     =   "Close Options Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   1
         Top             =   10
         Width           =   1335
      End
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   480
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Left     Balance     Right"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Balance"
      Top             =   870
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check3_Click()

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)


End Sub

Private Sub Option5_Click()
    Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, "4"
    Close #1
    Form1.Label1.ForeColor = &H80FF&
    Form1.Label15.ForeColor = &H80FF&
    Form1.Label7.ForeColor = &H80FF&
    Form1.Label4.ForeColor = &H80FF&
    Form1.Label13.ForeColor = &H80FF&
    Form1.Label6.ForeColor = &H80FF&
    Form1.Label5.ForeColor = &H4080&
    Form1.Label3.ForeColor = &H4080&
    'GREEN
    Form1.Picture5.Visible = False
    Form2.Picture2.Visible = False
    Form3.Picture2.Visible = False
    Form4.Picture5.Visible = False
    Form7.Picture5.Visible = False
    'BLUE
    Form1.Picture7.Visible = False
    Form2.Picture1.Visible = False
    Form3.Picture1.Visible = False
    Form4.Picture1.Visible = False
    Form7.Picture4.Visible = False
    'GOLD
    Form1.Picture11.Visible = False
    Form2.Picture5.Visible = False
    Form3.Picture5.Visible = False
    Form4.Picture4.Visible = False
    Form7.Picture7.Visible = False
    'GREY
    Form1.Picture15.Visible = False
    Form2.Picture7.Visible = False
    Form3.Picture7.Visible = False
    Form4.Picture6.Visible = False
    Form7.Picture6.Visible = False
    'ORANGE
    Form1.Picture19.Visible = True
    Form2.Picture9.Visible = True
    Form3.Picture9.Visible = True
    Form4.Picture9.Visible = True
    Form7.Picture1.Visible = True

End Sub

Private Sub Picture10_Click()
    Form1.Check2.Value = 0
    Form4.Visible = False

End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Timer1_Timer()
If Check2.Value = 1 Then
    Form4.Top = Form1.Top + Form1.Height + 20
    Form4.Left = Form1.Left
    Form7.Top = Form1.Top
    Form7.Left = Form1.Left + Form1.Width + 20
    If Form7.Visible = False Then
        Form6.Top = Form1.Top
        Form6.Left = Form1.Left + Form1.Width + 20
    Else
        Form6.Top = Form7.Top + Form7.Height + 20
        Form6.Left = Form7.Left
    End If
End If
End Sub

Private Sub Command1_Click()
Slider1.Value = 0
Form1.ActiveMovie1.Balance = 0
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

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

Private Sub Option1_Click()
    Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, "0"
    Close #1
    
    Form1.Label1.ForeColor = &HFF00&
    Form1.Label15.ForeColor = &HFF00&
    Form1.Label7.ForeColor = &HFF00&
    Form1.Label4.ForeColor = &HFF00&
    Form1.Label13.ForeColor = &HFF00&
    Form1.Label6.ForeColor = &HFF00&
    Form1.Label5.ForeColor = &H8000&
    Form1.Label3.ForeColor = &H8000&
    
    'GREEN
    Form1.Picture5.Visible = True
    Form2.Picture2.Visible = True
    Form3.Picture2.Visible = True
    Form4.Picture5.Visible = True
    Form7.Picture5.Visible = True
    'BLUE
    Form1.Picture7.Visible = False
    Form2.Picture1.Visible = False
    Form3.Picture1.Visible = False
    Form4.Picture1.Visible = False
    Form7.Picture4.Visible = False
    'GOLD
    Form1.Picture11.Visible = False
    Form2.Picture5.Visible = False
    Form3.Picture5.Visible = False
    Form4.Picture4.Visible = False
    Form7.Picture7.Visible = False
    'GREY
    Form1.Picture15.Visible = False
    Form2.Picture7.Visible = False
    Form3.Picture7.Visible = False
    Form4.Picture6.Visible = False
    Form7.Picture6.Visible = False
    'ORANGE
    Form1.Picture19.Visible = False
    Form2.Picture9.Visible = False
    Form3.Picture9.Visible = False
    Form4.Picture9.Visible = False
    Form7.Picture1.Visible = False
End Sub

Private Sub Option2_Click()
    Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, "1"
    Close #1
    Form1.Label1.ForeColor = &HFF0000
    Form1.Label15.ForeColor = &HFF0000
    Form1.Label7.ForeColor = &HFF0000
    Form1.Label4.ForeColor = &HFF0000
    Form1.Label13.ForeColor = &HFF0000
    Form1.Label6.ForeColor = &HFF0000
    Form1.Label5.ForeColor = &H800000
    Form1.Label3.ForeColor = &H800000
    'GREEN
    Form1.Picture5.Visible = False
    Form2.Picture2.Visible = False
    Form3.Picture2.Visible = False
    Form4.Picture5.Visible = False
    Form7.Picture5.Visible = False
    'BLUE
    Form1.Picture7.Visible = True
    Form2.Picture1.Visible = True
    Form3.Picture1.Visible = True
    Form4.Picture1.Visible = True
    Form7.Picture4.Visible = True
    'GOLD
    Form1.Picture11.Visible = False
    Form2.Picture5.Visible = False
    Form3.Picture5.Visible = False
    Form4.Picture4.Visible = False
    Form7.Picture7.Visible = False
    'GREY
    Form1.Picture15.Visible = False
    Form2.Picture7.Visible = False
    Form3.Picture7.Visible = False
    Form4.Picture6.Visible = False
    Form7.Picture6.Visible = False
    'ORANGE
    Form1.Picture19.Visible = False
    Form2.Picture9.Visible = False
    Form3.Picture9.Visible = False
    Form4.Picture9.Visible = False
    Form7.Picture1.Visible = False
End Sub

Private Sub Option3_Click()
        Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, "2"
    Close #1

    Form1.Label1.ForeColor = &HC0C0&
    Form1.Label15.ForeColor = &HC0C0&
    Form1.Label7.ForeColor = &HC0C0&
    Form1.Label4.ForeColor = &HC0C0&
    Form1.Label13.ForeColor = &HC0C0&
    Form1.Label6.ForeColor = &HC0C0&
    Form1.Label5.ForeColor = &H8080&
    Form1.Label3.ForeColor = &H8080&
    'GREEN
    Form1.Picture5.Visible = False
    Form2.Picture2.Visible = False
    Form3.Picture2.Visible = False
    Form4.Picture5.Visible = False
    Form7.Picture5.Visible = False
    'BLUE
    Form1.Picture7.Visible = False
    Form2.Picture1.Visible = False
    Form3.Picture1.Visible = False
    Form4.Picture1.Visible = False
    Form7.Picture4.Visible = False
    'GOLD
    Form1.Picture11.Visible = True
    Form2.Picture5.Visible = True
    Form3.Picture5.Visible = True
    Form4.Picture4.Visible = True
    Form7.Picture7.Visible = True
    'GREY
    Form1.Picture15.Visible = False
    Form2.Picture7.Visible = False
    Form3.Picture7.Visible = False
    Form4.Picture6.Visible = False
    Form7.Picture6.Visible = False
    'ORANGE
    Form1.Picture19.Visible = False
    Form2.Picture9.Visible = False
    Form3.Picture9.Visible = False
    Form4.Picture9.Visible = False
    Form7.Picture1.Visible = False
End Sub

Private Sub Option4_Click()
'&H00C0C0C0&
    Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, ""
    Write #1, "3"
    Close #1
    Form1.Label1.ForeColor = &HC0C0C0
    Form1.Label15.ForeColor = &HC0C0C0
    Form1.Label7.ForeColor = &HC0C0C0
    Form1.Label4.ForeColor = &HC0C0C0
    Form1.Label13.ForeColor = &HC0C0C0
    Form1.Label6.ForeColor = &HC0C0C0
    Form1.Label5.ForeColor = &H808080
    Form1.Label3.ForeColor = &H808080
    'GREEN
    Form1.Picture5.Visible = False
    Form2.Picture2.Visible = False
    Form3.Picture2.Visible = False
    Form4.Picture5.Visible = False
    Form7.Picture5.Visible = False
    'BLUE
    Form1.Picture7.Visible = False
    Form2.Picture1.Visible = False
    Form3.Picture1.Visible = False
    Form4.Picture1.Visible = False
    Form7.Picture5.Visible = False
    'GOLD
    Form1.Picture11.Visible = False
    Form2.Picture5.Visible = False
    Form3.Picture5.Visible = False
    Form4.Picture4.Visible = False
    Form7.Picture7.Visible = False
    'GREY
    Form1.Picture15.Visible = True
    Form2.Picture7.Visible = True
    Form3.Picture7.Visible = True
    Form4.Picture6.Visible = True
    Form7.Picture6.Visible = True
    'ORANGE
    Form1.Picture19.Visible = False
    Form2.Picture9.Visible = False
    Form3.Picture9.Visible = False
    Form4.Picture9.Visible = False
    Form7.Picture1.Visible = False
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture2_Click()
    Form1.Check2.Value = 0
    Form4.Visible = False

End Sub

Private Sub Picture3_Click()
    Form1.Check2.Value = 0
    Form4.Visible = False

End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture7_Click()
    Form1.Check2.Value = 0
    Form4.Visible = False

End Sub

Private Sub Picture8_Click()
    Form1.Check2.Value = 0
    Form4.Visible = False

End Sub

Private Sub Slider1_Change()
Form1.ActiveMovie1.Balance = Slider1.Value
End Sub

