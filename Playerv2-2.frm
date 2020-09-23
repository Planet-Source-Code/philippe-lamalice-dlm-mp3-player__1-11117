VERSION 5.00
Object = "{05B9F8C4-05D2-11D1-A081-444553540000}#1.0#0"; "NEWEX.OCX"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Open File..."
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Playerv2-2.frx":0000
   ScaleHeight     =   4335
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-2.frx":1066A
      ScaleHeight     =   225
      ScaleWidth      =   6615
      TabIndex        =   17
      Top             =   100
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6370
         Picture         =   "Playerv2-2.frx":17264
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   19
         ToolTipText     =   "Close Open Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Open File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-2.frx":17546
      ScaleHeight     =   225
      ScaleWidth      =   6615
      TabIndex        =   14
      Top             =   100
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6370
         Picture         =   "Playerv2-2.frx":1D6D8
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   16
         ToolTipText     =   "Close Open Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Open File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   10
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-2.frx":1D9BA
      ScaleHeight     =   225
      ScaleWidth      =   6615
      TabIndex        =   11
      Top             =   100
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6370
         Picture         =   "Playerv2-2.frx":23B4C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   13
         ToolTipText     =   "Close Open Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Open File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   10
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-2.frx":23E2E
      ScaleHeight     =   225
      ScaleWidth      =   6615
      TabIndex        =   7
      Top             =   100
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6370
         Picture         =   "Playerv2-2.frx":2A128
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   9
         ToolTipText     =   "Close Open Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Open File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   10
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-2.frx":2A40A
      ScaleHeight     =   225
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   100
      Visible         =   0   'False
      Width           =   6615
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6370
         Picture         =   "Playerv2-2.frx":3344C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   10
         ToolTipText     =   "Close Open Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Open File..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   10
         Width           =   1095
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   2400
      Pattern         =   "*.MP3;*.MP2;*.M3U"
      TabIndex        =   4
      Top             =   375
      Width           =   4335
   End
   Begin NEWEXLib.ExplorerTree ExplorerTree1 
      Height          =   3405
      Left            =   120
      TabIndex        =   3
      Top             =   375
      Width           =   2295
      _Version        =   65536
      _ExtentX        =   4048
      _ExtentY        =   6006
      _StockProps     =   161
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Streaming"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   3910
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   3910
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   3910
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Open "C:\WINDOWS\DLMPLAY2.INI" For Output As #1
Write #1, Form2.ExplorerTree1.Path
Close #1
Form1.ActiveMovie1.FileName = ExplorerTree1.Path + "\" + File1.FileName
  '  For x = 1 To Len(Text1.Text)
  '      Tmp1$ = Left$(Text1.Text, x)
  '      Tmp2$ = Right$(Tmp1$, 1)
  '      If Tmp2$ = "\" Then
  '          SongNameTmp$ = Right$(Text1.Text, x + 9)
  '      End If
  '  Next x
  ' SongName$ = Left$(SongNameTmp$, Len(SongNameTmp$) - 4)
If UCase$(Right$(File1.FileName, 3)) = "M3U" Then
    Form1.Label1.Caption = "MP3 Playlist"
    Form1.Command4.Enabled = True
    Form1.Command5.Enabled = True
    
Else
    Form1.Command4.Enabled = False
    Form1.Command5.Enabled = False
    Form1.Label1.Caption = LCase$(Left$(File1.FileName, Len(File1.FileName) - 4))
End If
Form1.Timer1.Enabled = True
Form2.Enabled = False
Form2.Visible = False
Form1.Visible = True
Form1.Enabled = True
Form1.Visible = True
Form4.Enabled = True

Unload Form2
End Sub

Private Sub Command2_Click()
    Form2.Enabled = False
    Form2.Visible = False
    Form1.Visible = True
    Form4.Enabled = True
    
    Form1.Enabled = True
    Form1.Visible = True
Unload Form2
End Sub

Private Sub Command3_Click()
Form2.Visible = False
Form2.Enabled = False
Form3.Visible = True
Form3.Enabled = True
Form3.Visible = True
Unload Form2
End Sub

Private Sub Dir1_Change()

End Sub

Private Sub Drive1_Change()

End Sub

Private Sub ExplorerList1_FolderClick()
End Sub

Private Sub ExplorerTree1_OnDirChanged()
File1.Path = ExplorerTree1.Path
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

Private Sub Form_Load()
'Makes all Title-Bar Unvisible
Picture1.Visible = False
Picture2.Visible = False
Picture5.Visible = False
Picture7.Visible = False
Picture9.Visible = False
'Select the Visible Title Bar (According to main window)
If Form1.Picture5.Visible = True Then Form2.Picture2.Visible = True
If Form1.Picture7.Visible = True Then Form2.Picture1.Visible = True
If Form1.Picture11.Visible = True Then Form2.Picture5.Visible = True
If Form1.Picture15.Visible = True Then Form2.Picture7.Visible = True
If Form1.Picture19.Visible = True Then Form2.Picture9.Visible = True
'Centers the Open Window
Form2.Top = Screen.Height / 2 - Form2.Height / 2
Form2.Left = Screen.Width / 2 - Form2.Width / 2
'Close Any Unclosed file using #1
Close #1
'Open Path Stored Data from file DLMPLAY2.INI
Open "C:\WINDOWS\DLMPLAY2.INI" For Input As #1
Input #1, CurrentPathTree
Close #1
'Code for File Selection and Three Display
ExplorerTree1.Path = CurrentPathTree
ExplorerTree1.InitialDir = CurrentPathTree
File1.Path = ExplorerTree1.Path
'Error recovery Code
On Error Resume Next
End Sub

Private Sub Timer1_Timer()
Label1.Caption = Dir1.Path + "\" + File1.FileName
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
