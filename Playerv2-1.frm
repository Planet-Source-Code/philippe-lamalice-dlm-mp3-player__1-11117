VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DLM MP3 Player"
   ClientHeight    =   1740
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Playerv2-1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Playerv2-1.frx":0442
   ScaleHeight     =   1740
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check4 
      Height          =   195
      Left            =   2880
      TabIndex        =   57
      ToolTipText     =   "Playlist"
      Top             =   1350
      Width           =   255
   End
   Begin VB.PictureBox Picture19 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-1.frx":4434
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   52
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture22 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3070
         Picture         =   "Playerv2-1.frx":7F3E
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   56
         ToolTipText     =   "About"
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture21 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3340
         Picture         =   "Playerv2-1.frx":8220
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   55
         ToolTipText     =   "Minimize"
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture20 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-1.frx":8502
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   54
         ToolTipText     =   "Exit DLM MP3 Player"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLM MP3 Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   53
         Top             =   10
         Width           =   1200
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mp3;*.mp2"
      DialogTitle     =   "Add MP3 File..."
      Filter          =   "MP3 Files|*.mp3;*.mp2"
   End
   Begin VB.CheckBox Check3 
      Height          =   255
      Left            =   3160
      TabIndex        =   51
      ToolTipText     =   "MiniBrowser"
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox Picture15 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-1.frx":87E4
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   45
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture18 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3070
         Picture         =   "Playerv2-1.frx":C0AE
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   48
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture17 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3340
         Picture         =   "Playerv2-1.frx":C390
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   47
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture16 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-1.frx":C672
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   46
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLM MP3 Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   49
         Top             =   10
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture11 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-1.frx":C954
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   40
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture14 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3070
         Picture         =   "Playerv2-1.frx":1057E
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   43
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture13 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3340
         Picture         =   "Playerv2-1.frx":10860
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   42
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture12 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-1.frx":10B42
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   41
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLM MP3 Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   44
         Top             =   10
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-1.frx":10E24
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   35
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3080
         Picture         =   "Playerv2-1.frx":14A06
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   39
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3350
         Picture         =   "Playerv2-1.frx":14CE8
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   38
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-1.frx":14FCA
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   37
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DLM MP3 Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   36
         Top             =   10
         Width           =   1200
      End
   End
   Begin VB.CheckBox Check2 
      Height          =   255
      Left            =   3440
      TabIndex        =   34
      ToolTipText     =   "Options"
      Top             =   1320
      Width           =   220
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   2040
      Top             =   2040
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-1.frx":152AC
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   29
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3080
         Picture         =   "Playerv2-1.frx":1915E
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   33
         ToolTipText     =   "About"
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3350
         Picture         =   "Playerv2-1.frx":19440
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   32
         ToolTipText     =   "Minimize"
         Top             =   10
         Width           =   240
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-1.frx":19722
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   31
         ToolTipText     =   "Exit DLM MP3 Player"
         Top             =   10
         Width           =   240
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "DLM MP3 Player"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   0
         TabIndex        =   30
         Top             =   10
         Width           =   1335
      End
   End
   Begin ComctlLib.Slider DurationMin 
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   327682
      Max             =   100000
   End
   Begin ComctlLib.Slider DurationSec 
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      _Version        =   327682
      Max             =   100000
   End
   Begin ComctlLib.Slider MinSlider 
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100000
   End
   Begin ComctlLib.Slider SecSlider 
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100000
   End
   Begin VB.Timer Timer5 
      Interval        =   50
      Left            =   3120
      Top             =   2520
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3680
      Picture         =   "Playerv2-1.frx":19A04
      ScaleHeight     =   255
      ScaleWidth      =   315
      TabIndex        =   20
      ToolTipText     =   "Open File / Location"
      Top             =   1320
      Width           =   320
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "Playerv2-1.frx":19B52
      ScaleHeight     =   255
      ScaleWidth      =   1710
      TabIndex        =   14
      Top             =   1320
      Width           =   1710
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         ToolTipText     =   "Fwd Track"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   720
         TabIndex        =   17
         ToolTipText     =   "Pause"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   360
         TabIndex        =   16
         ToolTipText     =   "Play"
         Top             =   0
         Width           =   330
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Back Track"
         Top             =   0
         Width           =   330
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   3600
      Top             =   2400
   End
   Begin VB.CheckBox Check1 
      Height          =   195
      Left            =   1880
      TabIndex        =   12
      ToolTipText     =   "Repeat"
      Top             =   1350
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   0
      Top             =   2280
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2100
      TabIndex        =   7
      ToolTipText     =   "Volume"
      Top             =   1365
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -4000
      Max             =   0
      SelectRange     =   -1  'True
      TickStyle       =   3
   End
   Begin VB.CommandButton Command6 
      Caption         =   "O"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      ToolTipText     =   "Open File / Location"
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Fwd Track"
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Back Track"
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Stop"
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "| |"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Pause"
      Top             =   2550
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Play"
      Top             =   2550
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   840
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   2520
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   160
      TabIndex        =   28
      ToolTipText     =   "Song Position (Click to step)"
      Top             =   670
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
      MousePointer    =   2
      MouseIcon       =   "Playerv2-1.frx":19FD0
   End
   Begin MediaPlayerCtl.MediaPlayer ActiveMovie1 
      Height          =   615
      Left            =   840
      TabIndex        =   50
      Top             =   5640
      Width           =   1215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   2040
      TabIndex        =   25
      ToolTipText     =   "Total Time"
      Top             =   890
      Width           =   645
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Vol"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   390
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1440
      TabIndex        =   21
      ToolTipText     =   "Play Time"
      Top             =   885
      Width           =   525
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3360
      TabIndex        =   13
      ToolTipText     =   "Volume (%)"
      Top             =   390
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mono"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      ToolTipText     =   "Stereo / Mono"
      Top             =   890
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   165
      TabIndex        =   10
      ToolTipText     =   "Current State"
      Top             =   885
      Width           =   810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Stereo"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Stereo / Mono"
      Top             =   890
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- no file opened -"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   165
      TabIndex        =   8
      ToolTipText     =   "Song File / Title"
      Top             =   390
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Display Area"
      Top             =   360
      Width           =   3885
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
'DLM MP3 Player
'Written by Philippe Lamalice
'
'I wrote this MP3 Program for no particular reasons...
'Just wanted to do something... It ended that my player was
'very popular along my friends who had seen it. And they
'wanted it, so i went on with it... I've been developping my
'player for about 1 month now... (about 1 hour a day) and
'it's getting pretty good!
'
'If you have any questions, comments, or improvements, feel
'free to write me (dlm@biosys.net). I will try to answer your
'questions the best i can! And if you modify it, please leave
'me credit for it... and i would also like to see what
'improvements are made by others... so if you could send
'me back the sources after modifications... would be really
'appreciated!
'
'08/30/2000 Added Playlist. and cleaned up some code..
'           (still messy tough). Playlist is working
'           except for an unresolved bug when saving the
'           playlist from times to times (File Not Found
'           Error!!)  Note : M3U Playlist are not supported.
'           DLM MP3 Players uses its own playlist format...
'
'Thankx
'Philippe Lamalice
'DLM@Biosys.net
'
'
'


Private Sub ActiveMovie1_EndOfStream(ByVal Result As Long)
If Form7.List1.ListCount = 0 Then
Else
    If Form7.List1.ListIndex + 1 = Form7.List1.ListCount Then
        Form7.List1.ListIndex = 0
    Else
        Form7.List1.ListIndex = Form7.List1.ListIndex + 1
    End If
    If Form7.List1.Text <> "" Then
        Form1.ActiveMovie1.FileName = Form7.List1.Text
        Form1.ActiveMovie1.Play
        Form1.Label1.Caption = LCase$(Left$(Form7.List1.Text, Len(Form7.List1.Text) - 4))
    End If
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Form4.Show
    Form4.Top = Form1.Top + Form1.Height + 10
    Form4.Left = Form1.Left
Else
    Form4.Visible = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Form6.Show
    Form6.Top = Form1.Top
    Form6.Left = Form1.Left + Form1.Width + 20
Else
    Form6.Visible = False
End If
End Sub

Private Sub Check4_Click()
If Form1.Check4.Value = 1 Then
    Form7.Visible = True
Else
    Form7.Visible = False
End If
End Sub

Private Sub Command1_Click()
ActiveMovie1.Play
End Sub

Private Sub Command2_Click()
If ActiveMovie1.PlayState = mpPaused Then
    ActiveMovie1.Play
Else
    If ActiveMovie1.PlayState = mpPlaying Then
        ActiveMovie1.Pause
    End If
End If

End Sub

Private Sub Command3_Click()
ActiveMovie1.Stop

End Sub

Private Sub Command4_Click()
ActiveMovie1.CurrentPosition = 0
End Sub

Private Sub Command5_Click()
ActiveMovie1.CurrentPosition = ActiveMovie1.Duration
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 2 Then
    Form1.Enabled = False
    Form1.Visible = False
    Form3.Enabled = True
    Form3.Visible = True
    Form3.Enabled = True
Else
    Form1.Enabled = False
    Form1.Visible = False
    Form2.Visible = True
    Form2.Enabled = True
    Form2.Visible = True
End If

End Sub

Private Sub Form_Load()


'   DLM MP3 Player
'   Written by Philippe Lamalice
'
'   See General_Declarations for comments, notes
'   and known bugs...
'
'
'   Note : Code has been trowhen... Only matter was to make it work..
'          I haven't had the time yet to simplify it... Sorry!



'==============================================================
'Beginning of code --->
'Not many comments, sorry! I didn't planned on distributing it!
'==============================================================


'Settings Load Code
Form1.Check2.Value = 1
Form1.Check3.Value = 0
Form1.Check4.Value = 1
Me.Top = Screen.Height / 2 - Me.Height - 20
Me.Left = Screen.Width / 2 - Me.Width - 20
On Error GoTo CreateConfig
LoadConfig:
Dim Setting$(1 To 5)
Dim ColorScheme$, CurrentPath$
Open "C:\WINDOWS\DLMPLAY.INI" For Input As #1
Input #1, Setting$(1)
Input #1, Setting$(2)
Input #1, Setting$(3)
Input #1, Setting$(4)
Input #1, Setting$(5)
Input #1, ColorScheme$
Close #1
Open "C:\WINDOWS\DLMPLAY2.INI" For Input As #1
Input #1, CurrentPath$
Close #1
GoTo ContinueLoad
CreateConfig: 'Create Configuration Files if not found or if incomplete!
Close #1
Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1  ' General Configuration File
Write #1, "1"
Write #1, "1"
Write #1, Form1.ActiveMovie1.Balance
Write #1, Form1.ActiveMovie1.Volume
Write #1, ""
Write #1, "1"
Close #1
Open "C:\WINDOWS\DLMPLAY2.INI" For Output As #1  ' Latest Path Access File
Write #1, "C:\"
Close #1
GoTo LoadConfig
ContinueLoad:
On Error Resume Next
'Color Scheme Code
Select Case ColorScheme$
    Case "0"
        Form4.Option1.Value = True
    Case "1"
        Form4.Option2.Value = True
    Case "2"
        Form4.Option3.Value = True
    Case "3"
        Form4.Option4.Value = True
    Case "4"
        Form4.Option5.Value = True
    Case Else
        Form4.Option2.Value = True
End Select
'Current Path Code
Form2.ExplorerTree1.InitialDir = CurrentPath$
Form2.ExplorerTree1.Path = CurrentPath$
Form2.File1.Path = Form2.ExplorerTree1.Path
'AutoPlay Code
If Setting$(1) = 1 Then
    Form4.Check1.Value = 1
    Form1.ActiveMovie1.AutoStart = True
Else
    Form4.Check1.Value = 0
    Form1.ActiveMovie1.AutoStart = False
End If
'Auto-Attach Code
If Setting$(2) = 1 Then
    Form4.Check2.Value = 1
Else
    Form4.Check2.Value = 0
End If
'Balance Code
Form4.Slider1.Value = Setting$(3)
'Volume Code
Form1.Slider1.Value = Setting$(4)
ActiveMovie1.Volume = Form1.Slider1.Value
'Re-Load Code
FileLoad$ = Setting$(5)
'END of settings load code

On Error Resume Next
If Command$ <> "" Then FileLoad$ = Command$
'--- Code to remove path and keep only filename ... without extension
  '  For x = 4 To Len(Command$)
  '      Tmp1$ = Left$(Command$, x)
  '      Tmp2$ = Right$(Tmp1$, 1)
  '      If Tmp2$ = "\" Then
  '          SongNameTmp$ = Right$(Command$, x + 9)
  '      End If
  '  Next x
If FileLoad$ <> "" Then
    Form1.ActiveMovie1.FileName = FileLoad$
  '  SongName$ = Left$(SongNameTmp$, Len(SongNameTmp$) - 4)
    If UCase$(Right$(FileLoad$, 3)) = "M3U" Then
        Form1.Label1.Caption = "MP3 Playlist"
        Form1.Command4.Enabled = True
        Form1.Command5.Enabled = True
    Else
        Form1.Label1.Caption = LCase$(Left$(FileLoad$, Len(FileLoad$) - 4))
        Form1.Command4.Enabled = False
        Form1.Command5.Enabled = False
    End If
Else
End If
Timer1.Enabled = True
GoTo EndAll
ErrorHandler:
Select Case ErrorNum
    Case 6
        ProgressTmp = 0
        Progress = 0
        ProgressBar1.Value = 0
    Case Else
        MsgBox "Error " + ErrorNum + " happened. Quitting...", vbCritical, "ERROR"
        End
End Select
EndAll:
Picture10.ToolTipText = Picture6.ToolTipText
Picture14.ToolTipText = Picture6.ToolTipText
Picture18.ToolTipText = Picture6.ToolTipText
Picture9.ToolTipText = Picture4.ToolTipText
Picture13.ToolTipText = Picture4.ToolTipText
Picture17.ToolTipText = Picture4.ToolTipText
Picture8.ToolTipText = Picture3.ToolTipText
Picture12.ToolTipText = Picture3.ToolTipText
Picture16.ToolTipText = Picture3.ToolTipText
Form6.Visible = False
On Error Resume Next
End Sub

Private Sub Label10_Click()
Command2_Click
End Sub

Private Sub Label11_Click()
SecSlider.Value = 0
MinSlider.Value = 0
Timer5_Timer
Command3_Click
End Sub

Private Sub Label12_Click()
SecSlider.Value = 0
MinSlider.Value = 0
Timer5_Timer
Command5_Click
End Sub

Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Label8_Click()
SecSlider.Value = 0
MinSlider.Value = 0
Timer5_Timer
Command4_Click
End Sub

Private Sub Label9_Click()
If ActiveMovie1.PlayState = mpPaused Then
    Command2_Click
Else
    If ActiveMovie1.PlayState = mpPlaying Then
    Else
        If ActiveMovie1.FileName <> "" Then Command1_Click
    End If
End If
End Sub

Private Sub Picture10_Click()
Picture6_Click
Picture10.ToolTipText = Picture6.ToolTipText
End Sub

Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture12_Click()
Picture3_Click
End Sub

Private Sub Picture13_Click()
Picture4_Click
End Sub

Private Sub Picture14_Click()
Picture6_Click
End Sub

Private Sub Picture15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture16_Click()
Picture3_Click
End Sub

Private Sub Picture17_Click()
Picture4_Click
End Sub

Private Sub Picture18_Click()
Picture6_Click
End Sub

Private Sub Picture19_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 2 Then
    Form1.Enabled = False
    Form4.Enabled = False
    Form1.Visible = True
    Form3.Enabled = True
    Form3.Visible = True
    Form3.Enabled = True
Else
    Form1.Enabled = False
    Form1.Visible = True
    Form2.Visible = True
    Form2.Enabled = True
    Form2.Visible = True
End If

End Sub

Private Sub Picture20_Click()
Picture3_Click

End Sub

Private Sub Picture21_Click()
Picture4_Click

End Sub

Private Sub Picture22_Click()
Picture6_Click

End Sub

Private Sub Picture3_Click()
Open "C:\WINDOWS\DLMPLAY.INI" For Output As #1
If Form4.Check1.Value = 1 Then
    Write #1, "1"
Else
    Write #1, "0"
End If
If Form4.Check2.Value = 1 Then
    Write #1, "1"
Else
    Write #1, "0"
End If
On Error Resume Next
Write #1, Form4.Slider1.Value
Write #1, Form1.Slider1.Value
Write #1, Form1.ActiveMovie1.FileName
If Form4.Option1.Value = True Then Write #1, "0"
If Form4.Option2.Value = True Then Write #1, "1"
If Form4.Option3.Value = True Then Write #1, "2"
If Form4.Option4.Value = True Then Write #1, "3"
If Form4.Option5.Value = True Then Write #1, "4"
Close #1
End
End Sub

Private Sub Picture4_Click()
If Check2.Value = 1 Then Form4.Visible = False
Form1.Visible = False
Form5.Visible = True
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture6_Click()
frmAbout.Show
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub Picture8_Click()
Picture3_Click
End Sub

Private Sub Picture9_Click()
Picture4_Click
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And 1 Then
    Step1 = ActiveMovie1.Duration / 3572
    Step2 = X * Step1
    ActiveMovie1.CurrentPosition = Step2
End If
End Sub

Private Sub Slider1_Change()
ActiveMovie1.Volume = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
Label15.Visible = True
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Left$(Form1.Label1.Caption, 9) = "Streaming" Then
    If ActiveMovie1.BufferingProgress = 0 Or ActiveMovie1.BufferingProgress > 99 Then
        LostPack$ = ActiveMovie1.LostPackets
        BandWith$ = ActiveMovie1.Bandwidth
        Form1.Label1.Caption = "Streaming [LP:" + LostPack$ + " / BW:" + BandWith$ + "]"
    Else
        BufferPrg$ = ActiveMovie1.BufferingProgress
        Form1.Label1.Caption = "Streaming [Buffering : " + BufferPrg$ + "%]"
    End If
End If
If ActiveMovie1.FileName <> "" Then
    ProgressTmp = ActiveMovie1.CurrentPosition / ActiveMovie1.Duration
    Progress = ProgressTmp * 100
    ProgressBar1.Value = Progress
End If
End Sub

Private Sub Timer2_Timer()
If Slider1.Value < -3975 Then
    ActiveMovie1.Volume = -9640
Else
    ActiveMovie1.Volume = Slider1.Value
End If
    On Error Resume Next
    Prcent = Slider1.Value * -2.5 / 100
    Prcent2 = 100 - Prcent
    Prcent3 = Format$(Prcent2, "##")
    PercentVol = Prcent3 + "%"
    Label7.Caption = PercentVol
    If Slider1.Value = 0 Then Label7.Caption = "100%"
    If Slider1.Value < -3995 Then Label7.Caption = "0%"
End Sub

Private Sub Timer3_Timer()
If Check2.Value = 1 And WindowState <> 1 Then Form4.Visible = True
If Len(Label1.Caption) > 40 Then
    Label1.Caption = "..." + Right$(Label1.Caption, 37)
End If
Select Case ActiveMovie1.PlayState
    Case mpClosed
        Label4.Caption = "Idle"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &H800000
            Case &HFF00&   'Green
                Label3.ForeColor = &H8000&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &H8080&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &H808080
            Case &H80FF& 'Orange
                Label3.ForeColor = &H4080&
        End Select
    Case mpStopped
        Label4.Caption = "Stopped"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &HFF0000
            Case &HFF00&   'Green
                Label3.ForeColor = &HFF00&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &HC0C0&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &HC0C0C0
            Case &H80FF& 'Orange
                Label3.ForeColor = &H80FF&
        End Select
    Case mpPaused
        Label4.Caption = "Paused"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &HFF0000
            Case &HFF00&   'Green
                Label3.ForeColor = &HFF00&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &HC0C0&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &HC0C0C0
            Case &H80FF&    'Orange
                Label3.ForeColor = &H80FF&
        End Select
    Case mpPlaying
        Label4.Caption = "Playing"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &HFF0000
            Case &HFF00&   'Green
                Label3.ForeColor = &HFF00&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &HC0C0&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &HC0C0C0
            Case &H80FF&    'Orange
                Label3.ForeColor = &H80FF&
        End Select
    Case mpWaiting
        Label4.Caption = "Opening..."
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &H800000
            Case &HFF00&   'Green
                Label3.ForeColor = &H8000&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &H8080&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &H808080
            Case &H80FF&    'Orange
                Label3.ForeColor = &H4080&
        End Select
    Case mpScanForward
            Label4.Caption = "Scanning Forward"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &H800000
            Case &HFF00&   'Green
                Label3.ForeColor = &H8000&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &H8080&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &H808080
            Case &H80FF&    'Orange
                Label3.ForeColor = &H4080&
        End Select
    Case mpScanReverse
            Label4.Caption = "Scanning Back"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &H800000
            Case &HFF00&   'Green
                Label3.ForeColor = &H8000&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &H8080&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &H808080
            Case &H80FF&    'Orange
                Label3.ForeColor = &H4080&
        End Select
    Case Else
        Label4.Caption = "Status Unknown"
        Select Case Label1.ForeColor
            Case &HFF0000  'Blue
                Label3.ForeColor = &H800000
            Case &HFF00&   'Green
                Label3.ForeColor = &H8000&
            Case &HC0C0&   'Gold
                Label3.ForeColor = &H8080&
            Case &HC0C0C0  'Grey
                Label3.ForeColor = &H808080
            Case &H80FF&    'Orange
                Label3.ForeColor = &H4080&
        End Select
End Select
If Check1.Value = 1 Then
    ActiveMovie1.PlayCount = 0
Else
    ActiveMovie1.PlayCount = 1
End If
End Sub

Private Sub Timer4_Timer()
Label15.Visible = False
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
'Duration Code
DurationSec.Value = ActiveMovie1.Duration
If DurationSec.Value < 59 Then
    DurationMin.Value = 0
End If
If DurationSec.Value > 59 And DurationSec.Value < 120 Then
    DurationMin.Value = 1
    DurationSec.Value = DurationSec.Value - 60
End If
If DurationSec.Value > 119 And DurationSec.Value < 180 Then
    DurationMin.Value = 2
    DurationSec.Value = DurationSec.Value - 120
End If
If DurationSec.Value > 179 And DurationSec.Value < 240 Then
    DurationMin.Value = 3
    DurationSec.Value = DurationSec.Value - 180
End If
If DurationSec.Value > 239 And DurationSec.Value < 300 Then
    DurationMin.Value = 4
    DurationSec.Value = DurationSec.Value - 240
End If
If DurationSec.Value > 299 And DurationSec.Value < 360 Then
    DurationMin.Value = 5
    DurationSec.Value = DurationSec.Value - 300
End If
If DurationSec.Value > 359 And DurationSec.Value < 420 Then
    DurationMin.Value = 6
    DurationSec.Value = DurationSec.Value - 360
End If
If DurationSec.Value > 419 And DurationSec.Value < 480 Then
    DurationMin.Value = 7
    DurationSec.Value = DurationSec.Value - 420
End If
If DurationSec.Value > 479 And DurationSec.Value < 540 Then
    DurationMin.Value = 8
    DurationSec.Value = DurationSec.Value - 480
End If
If DurationSec.Value > 539 And DurationSec.Value < 600 Then
    DurationMin.Value = 9
    DurationSec.Value = DurationSec.Value - 540
End If
If DurationSec.Value > 599 And DurationSec.Value < 660 Then
    DurationMin.Value = 10
    DurationSec.Value = DurationSec.Value - 600
End If
If DurationSec.Value > 659 And DurationSec.Value < 720 Then
    DurationMin.Value = 11
    DurationSec.Value = DurationSec.Value - 660
End If
If DurationSec.Value > 719 Then
    DurationMin.Value = 12
    DurationSec.Value = DurationSec.Value - 720
End If

SecTmp$ = DurationSec.Value
SecDure$ = Left$(SecTmp$, 2)
MinTmp$ = DurationMin.Value
MinDure$ = Left$(MinTmp$, 3)
If MinDure$ = "0" And SecDure$ = "" Then
Else
    If MinDure$ = "" Then MinDure$ = "0"
    If MinDure$ <> "" And SecDure$ = "" Then SecDure$ = "00"
    If Len(SecDure$) = 1 Then SecDure$ = "0" + SecDure$
    DurationTime$ = "/ " + MinDure$ + ":" + SecDure$
    Label6.Caption = DurationTime$
End If



'Current Play Time Code
SecSlider.Value = ActiveMovie1.CurrentPosition
If SecSlider.Value < 59 Then
    MinSlider.Value = 0
End If
If SecSlider.Value > 59 And SecSlider.Value < 120 Then
    MinSlider.Value = 1
    SecSlider.Value = SecSlider.Value - 60
End If
If SecSlider.Value > 119 And SecSlider.Value < 180 Then
    MinSlider.Value = 2
    SecSlider.Value = SecSlider.Value - 120
End If
If SecSlider.Value > 179 And SecSlider.Value < 240 Then
    MinSlider.Value = 3
    SecSlider.Value = SecSlider.Value - 180
End If
If SecSlider.Value > 239 And SecSlider.Value < 300 Then
    MinSlider.Value = 4
    SecSlider.Value = SecSlider.Value - 240
End If
If SecSlider.Value > 299 And SecSlider.Value < 360 Then
    MinSlider.Value = 5
    SecSlider.Value = SecSlider.Value - 300
End If
If SecSlider.Value > 359 And SecSlider.Value < 420 Then
    MinSlider.Value = 6
    SecSlider.Value = SecSlider.Value = 360
End If
If SecSlider.Value > 419 And SecSlider.Value < 480 Then
    MinSlider.Value = 7
    SecSlider.Value = SecSlider.Value - 420
End If
If SecSlider.Value > 479 And SecSlider.Value < 540 Then
    MinSlider.Value = 8
    SecSlider.Value = SecSlider.Value - 480
End If
If SecSlider.Value > 539 And SecSlider.Value < 600 Then
    MinSlider.Value = 9
    SecSlider.Value = SecSlider.Value - 540
End If
If SecSlider.Value > 599 And SecSlider.Value < 660 Then
    MinSlider.Value = 10
    SecSlider.Value = SecSlider.Value - 600
End If
If SecSlider.Value > 659 And SecSlider.Value < 720 Then
    MinSlider.Value = 11
    SecSlider.Value = SecSlider.Value - 660
End If
' This one is the last.... to be continued
If SecSlider.Value > 719 Then
    MinSlider.Value = 12
    SecSlider.Value = SecSlider.Value - 720
End If

CurrentSec = Format(SecSlider.Value, "##")
CurrentMin = Format(MinSlider.Value, "##")
If CurrentMin = "" And CurrentSec = "" Then
Label13.Caption = "0:00"
Else
If CurrentMin = "" Then CurrentMin = "0"
If CurrentMin <> "" And CurrentSec = "" Then CurrentSec = "00"
If Len(CurrentSec) = 1 Then CurrentSec = "0" + CurrentSec
CurrentTime$ = CurrentMin + ":" + CurrentSec
Label13.Caption = CurrentTime$
Form7.Label2.Caption = CurrentTime$ + DurationTime$
End If
End Sub

Private Sub Timer6_Timer()
If Form4.Check1.Value = 1 Then ActiveMovie1.AutoStart = True
If Form4.Check2.Value = 1 Then ActiveMovie1.AutoRewind = True
ActiveMovie1.Balance = Form4.Slider1.Value
End Sub
