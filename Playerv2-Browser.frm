VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowser 
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Playerv2-Browser.frx":0000
   ScaleHeight     =   3630
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3600
      Top             =   960
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-Browser.frx":84DA
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   8
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-Browser.frx":BDA4
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   16
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
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-Browser.frx":C086
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   7
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-Browser.frx":FCB0
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   15
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
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-Browser.frx":FF92
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   6
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-Browser.frx":13E44
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   14
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
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   220
      Left            =   120
      Picture         =   "Playerv2-Browser.frx":14126
      ScaleHeight     =   225
      ScaleWidth      =   3855
      TabIndex        =   5
      Top             =   100
      Visible         =   0   'False
      Width           =   3855
      Begin VB.PictureBox Picture5 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3610
         Picture         =   "Playerv2-Browser.frx":17D08
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   13
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
         TabIndex        =   9
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
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
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3120
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3840
      ExtentX         =   6773
      ExtentY         =   5503
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5280
      Top             =   2400
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   4140
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Visible         =   0   'False
      Width           =   4140
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Text            =   "www3.sympatico.ca/michel.lamalice/dlmplay.htm"
         Top             =   600
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         Height          =   255
         Left            =   45
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4920
      Top             =   1680
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
            Picture         =   "Playerv2-Browser.frx":17FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-Browser.frx":182CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-Browser.frx":185AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-Browser.frx":18890
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-Browser.frx":18B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Playerv2-Browser.frx":18E54
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()
    
If Form4.Option1.Value = 1 Then
    Picture1.Visible = False
    Picture2.Visible = True
    Picture3.Visible = False
    Picture4.Visible = False
End If
If Form4.Option2.Value = 1 Then
    Picture1.Visible = True
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
End If
If Form4.Option3.Value = 1 Then
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = True
    Picture4.Visible = False
End If
If Form4.Option4.Value = 1 Then
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

    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'try to navigate to the starting address
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If

End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateComplete(ByVal URL As String)
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub Form_Resize()
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
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

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Ret&
    ReleaseCapture
    Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

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
frmBrowser.Visible = False
Form1.Check3.Value = 0
End Sub

Private Sub Picture6_Click()
frmBrowser.Visible = False
Form1.Check3.Value = 0
End Sub

Private Sub Picture7_Click()
frmBrowser.Visible = False
Form1.Check3.Value = 0
End Sub

Private Sub Picture8_Click()
frmBrowser.Visible = False
Form1.Check3.Value = 0
End Sub

Private Sub Timer1_Timer()
frmBrowser.Top = Form1.Top
frmBrowser.Left = Form1.Left + Form1.Width + 20
End Sub

Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Working..."
    End If
End Sub


