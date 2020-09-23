VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "DLM MP3 Player"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Playerv2-7.frx":0000
   ScaleHeight     =   3495
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture9 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-7.frx":7FEE
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   15
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-7.frx":BAF8
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   17
         ToolTipText     =   "Close Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MiniBrowser"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   10
         Width           =   855
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1800
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3360
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-7.frx":BDDA
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   10
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-7.frx":F9BC
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   11
         ToolTipText     =   "Close Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MiniBrowser"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-7.frx":FC9E
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   7
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-7.frx":13B50
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   8
         ToolTipText     =   "Close Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "MiniBrowser"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-7.frx":13E32
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   4
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-7.frx":17A5C
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   5
         ToolTipText     =   "Close Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MiniBrowser"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-7.frx":17D3E
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   1
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-7.frx":1B608
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   2
         ToolTipText     =   "Close Window"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "MiniBrowser"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   380
      Width           =   3880
      ExtentX         =   6844
      ExtentY         =   5292
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1B8EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1BBCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1BEAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1C190
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1C472
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-7.frx":1C754
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Tag             =   "&Address:"
         Top             =   720
         Width           =   3915
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
If Form4.Option1.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = True
    Picture3.Visible = False
    Picture4.Visible = False
End If
If Form4.Option2.Value = True Then
    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
End If
If Form4.Option3.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = True
    Picture4.Visible = False
End If
If Form4.Option4.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = True
End If
 
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
'    Form_Resize

    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
    StartingAddress = "http://www3.sympatico.ca/michel.lamalice/dlmplay.htm"

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

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
Form6.Visible = False
Form1.Check3.Value = False

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)


End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)


End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)


End Sub

Private Sub Picture5_Click()
Form6.Visible = False
Form1.Check3.Value = False

End Sub

Private Sub Picture6_Click()
Form6.Visible = False
Form1.Check3.Value = False

End Sub

Private Sub Picture7_Click()
Form6.Visible = False
Form1.Check3.Value = False

End Sub

Private Sub Picture8_Click()
Form6.Visible = False
Form1.Check3.Value = False

End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Timer1_Timer()
If Form4.Option1.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = True
    Picture3.Visible = False
    Picture4.Visible = False
    Picture9.Visible = False
End If
If Form4.Option2.Value = True Then
    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    Picture9.Visible = False
End If
If Form4.Option3.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = True
    Picture4.Visible = False
    Picture9.Visible = False
End If
If Form4.Option4.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = True
    Picture9.Visible = False
End If
If Form4.Option5.Value = True Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    Picture9.Visible = True
End If
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If

End Sub
