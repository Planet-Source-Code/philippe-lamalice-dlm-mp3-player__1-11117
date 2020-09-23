VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   LinkTopic       =   "Form7"
   Picture         =   "Playerv2-8.frx":0000
   ScaleHeight     =   3495
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   120
      Left            =   4080
      Top             =   1920
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-8.frx":7FEE
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   13
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture11 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-8.frx":BC18
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   21
         ToolTipText     =   "Close PlayList Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   10
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-8.frx":BEFA
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   12
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-8.frx":F7C4
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   20
         ToolTipText     =   "Close PlayList Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   20
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-8.frx":FAA6
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   11
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-8.frx":13958
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   19
         ToolTipText     =   "Close PlayList Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   10
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-8.frx":13C3A
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   10
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-8.frx":1781C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   18
         ToolTipText     =   "Close PlayList Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   10
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      ScaleHeight     =   315
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0:00 / 0:00"
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   -120
         TabIndex        =   9
         ToolTipText     =   "Play Time / Total Time"
         Top             =   30
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Save Playlist"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "Load Playlist"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Remove File"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Add file"
      Top             =   3000
      Width           =   495
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000080FF&
      Height          =   2595
      ItemData        =   "Playerv2-8.frx":17AFE
      Left            =   120
      List            =   "Playerv2-8.frx":17B00
      TabIndex        =   3
      ToolTipText     =   "Double-click file to play"
      Top             =   360
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-8.frx":17B02
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-8.frx":1B60C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   1
         Top             =   10
         Width           =   540
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.CommonDialog1.FileName = ""
Form1.CommonDialog1.ShowOpen
List1.AddItem Form1.CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command3_Click()
Dim InputPlayList$(0 To 255)
Form1.CommonDialog1.DefaultExt = "DLM MP3 Playlist (*.DPL)"
Form1.CommonDialog1.DialogTitle = "Open PlayList"
Form1.CommonDialog1.Filter = "*.DPL"
Form1.CommonDialog1.ShowOpen
If Form1.CommonDialog1.FileName <> "" Then
    Open Form1.CommonDialog1.FileName For Input As #1
    On Error GoTo CloseIt
Go:
    Input #1, InputPlayList$(X)
    X = X + 1
    GoTo Go
CloseIt:
TotalFiles = X - 1
For z = 0 To TotalFiles
List1.AddItem (InputPlayList$(z))
Next z
    Close #1
End If
Form1.CommonDialog1.DefaultExt = "MP3 Files"
Form1.CommonDialog1.DialogTitle = "Add MP3 File..."
Form1.CommonDialog1.Filter = "*.MP3;*.MP2"

End Sub

Private Sub Command4_Click()
Form1.CommonDialog1.DefaultExt = "DLM MP3 Playlist (*.DPL)"
Form1.CommonDialog1.DialogTitle = "Save PlayList"
Form1.CommonDialog1.Filter = "*.DPL"
Form1.CommonDialog1.ShowSave
If Form1.CommonDialog1.FileName <> "" Then
    Open Form1.CommonDialog1.FileName For Output As #1
    OldIndex = List1.ListIndex
    List1.ListIndex = 0
    For X = 0 To List1.ListCount - 1
    List1.ListIndex = List1.ListIndex + 1
    Write #1, List1.Text
    Next X
    Close #1
    List1.ListIndex = OldIndex
End If
Form1.CommonDialog1.DefaultExt = "MP3 Files"
Form1.CommonDialog1.DialogTitle = "Add MP3 File..."
Form1.CommonDialog1.Filter = "*.MP3;*.MP2"
End Sub
Private Sub Form_Load()
Picture1.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
If Form1.Picture5.Visible = True Then Picture5.Visible = True
If Form1.Picture7.Visible = True Then Picture4.Visible = True
If Form1.Picture11.Visible = True Then Picture7.Visible = True
If Form1.Picture15.Visible = True Then Picture6.Visible = True
If Form1.Picture19.Visible = True Then Picture1.Visible = True
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub List1_DblClick()
Form1.ActiveMovie1.FileName = List1.Text
Form1.Label1.Caption = LCase$(Left$(List1.Text, Len(List1.Text) - 4))
Form1.ActiveMovie1.Play
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture10_Click()
Form7.Visible = False
Form1.Check4.Value = 0

End Sub

Private Sub Picture11_Click()
Form7.Visible = False
Form1.Check4.Value = 0

End Sub

Private Sub Picture2_Click()
Form7.Visible = False
Form1.Check4.Value = 0
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

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture8_Click()
Form7.Visible = False
Form1.Check4.Value = 0

End Sub

Private Sub Picture9_Click()
Form7.Visible = False
Form1.Check4.Value = 0

End Sub

Private Sub Timer1_Timer()
List1.ForeColor = Form1.Label1.ForeColor
Label2.ForeColor = Form1.Label1.ForeColor
End Sub
