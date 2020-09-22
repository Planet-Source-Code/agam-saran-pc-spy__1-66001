VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B24801&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PC Spy"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PCSpy.xpcmdbutton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ListView lstHistory 
      Height          =   3015
      Left            =   2040
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483642
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "User"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Logged In"
         Object.Width           =   2911
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Logged Out"
         Object.Width           =   2910
      EndProperty
   End
   Begin VB.PictureBox fraSettings 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   2040
      ScaleHeight     =   3015
      ScaleWidth      =   5055
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
      Begin PCSpy.xpradiobutton optUnlimited 
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Unlimited Mode"
         BackColor       =   16777215
      End
      Begin PCSpy.xpradiobutton optLimited 
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
         Caption         =   "Limited Mode"
         BackColor       =   16777215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"Form1.frx":000C
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This will log upto 100 times."
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   2610
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"Form1.frx":00E4
         Height          =   1095
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   4815
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Written by: Agam Saran"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4800
      TabIndex        =   10
      Top             =   150
      Width           =   2265
   End
   Begin VB.Image imgHistory 
      Height          =   735
      Left            =   480
      Picture         =   "Form1.frx":01E4
      ToolTipText     =   "About"
      Top             =   480
      Width           =   765
   End
   Begin VB.Image imgWebsite 
      Height          =   735
      Left            =   600
      Picture         =   "Form1.frx":2002
      ToolTipText     =   "Website"
      Top             =   2120
      Width           =   705
   End
   Begin VB.Image imgSettings 
      Height          =   765
      Left            =   560
      Picture         =   "Form1.frx":3BD4
      ToolTipText     =   "Settings"
      Top             =   1240
      Width           =   765
   End
   Begin VB.Image imgHelp 
      Height          =   780
      Left            =   480
      Picture         =   "Form1.frx":5B2A
      ToolTipText     =   "Help"
      Top             =   3000
      Width           =   825
   End
   Begin VB.Image imgBullet 
      Height          =   120
      Left            =   360
      Picture         =   "Form1.frx":7D8C
      Top             =   4690
      Width           =   120
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Use Ctrl+Alt+F to show or hide this window"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   4650
      Width           =   4185
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00BFBFBF&
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   5535
   End
   Begin ComctlLib.ImageList imgList 
      Left            =   5040
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7E8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC History"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   2100
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   2280
      Picture         =   "Form1.frx":81E0
      Top             =   360
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00CC8859&
      Height          =   4335
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00CC8859&
      Height          =   4335
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CC8859&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00CC8859&
      FillColor       =   &H00CC8859&
      Height          =   3855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   4335
      Left            =   2880
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312

Private Enum states
    Normal = 0
    Disable = 1
    ReadOnly = 2
End Enum

Private bCancel As Boolean
Private Sub cmdOK_Click()
Me.Hide
End Sub
Private Sub ProcessMessages()
    Dim message As MSG
    Do While Not bCancel
       WaitMessage
    If PeekMessage(message, Me.hwnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
        If Me.Visible = False Then
            Me.Show
        Else
            Me.Hide
        End If
    End If
        DoEvents
    Loop
End Sub
Private Sub Form_Load()
Dim ret As Long, i As Integer
Dim strUserName As String
Dim TmpStr As String, Current As Integer
Dim Ctl As Control
gHW = Me.hwnd
Hook
bCancel = False
Open App.Path & "\spy.ini" For Input As #1
    Do Until EOF(1)
      Input #1, TmpStr
      Select Case Left(TmpStr, 1)
            Case "&"
                If Mid(TmpStr, 2, Len(TmpStr) - 1) = "True" Then
                    optLimited.Value = True
                    optUnlimited.Value = False
                Else
                    optLimited.Value = False
                    optUnlimited.Value = True
                End If
        End Select
    Loop
Close #1
TmpStr = ""
strUserName = String(100, Chr$(0))
GetUserName strUserName, 100
strUserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
Open App.Path & "\pc.log" For Append As #1
    Print #1, Code("~" & strUserName)
    Print #1, Code("+" & Date & " " & time)
Close #1
Open App.Path & "\pc.log" For Input As #1
    Do Until EOF(1)
      Input #1, TmpStr
      TmpStr = Code(TmpStr)
      Select Case Left(TmpStr, 1)
            Case "~"
                Current = Current + 1
                lstHistory.ListItems.Add , , Mid(TmpStr, 2, Len(TmpStr) - 1), , 1
            Case "+"
                lstHistory.ListItems(Current).SubItems(1) = Mid(TmpStr, 2, Len(TmpStr) - 1)
            Case "-"
                lstHistory.ListItems(Current).SubItems(2) = Mid(TmpStr, 2, Len(TmpStr) - 1)
        End Select
    Loop
    If optLimited.Value = True Then
        If lstHistory.ListItems.Count > 100 Then
            lstHistory.ListItems.Remove 1
            Close #1
            Open App.Path & "\pc.log" For Output As #1
            For i = 1 To lstHistory.ListItems.Count
                Print #1, Code("~" & lstHistory.ListItems(i).Text)
                Print #1, Code("+" & lstHistory.ListItems(i).SubItems(1))
                If lstHistory.ListItems(i).SubItems(2) <> "" Then
                    Print #1, Code("-" & lstHistory.ListItems(i).SubItems(2))
                End If
            Next
        End If
    End If
ret = RegisterHotKey(Me.hwnd, &HBFFF&, MOD_CONTROL + MOD_ALT, vbKeyF)

    For Each Ctl In Me.Controls
    Select Case Ctl.Name
      Case "imgHelp", "imgSettings", "imgWebsite", "imgHistory"
        Ctl.MousePointer = 99
        Ctl.MouseIcon = LoadPicture(App.Path & "\Hand.cur")
    End Select
    Next

    SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "PC Spy", App.Path & "\Spy.exe"
    App.TaskVisible = False
ProcessMessages
End Sub


Private Sub Form_Unload(Cancel As Integer)
bCancel = True
Call UnregisterHotKey(Me.hwnd, &HBFFF&)
Unhook

End Sub

Private Sub imgHelp_Click()
ShellExecute 0, "open", App.Path & "\Help.chm", "", "", 10
End Sub

Private Sub imgHistory_Click()
lblTitle.Caption = "PC History"
lstHistory.Visible = True
fraSettings.Visible = False

End Sub

Private Sub imgHistory_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgHistory.Top = 510
imgHistory.Left = 510
End Sub

Private Sub imgHistory_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgHistory.Top = 480
imgHistory.Left = 480
End Sub

Private Sub imgHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgHelp.Top = 3030
imgHelp.Left = 510
End Sub

Private Sub imgHelp_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgHelp.Top = 3000
imgHelp.Left = 480
End Sub

Private Sub imgSettings_Click()
lblTitle.Caption = "Settings"
lstHistory.Visible = False
fraSettings.Visible = True
End Sub

Private Sub imgSettings_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgSettings.Top = 1270
imgSettings.Left = 580
End Sub

Private Sub imgSettings_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgSettings.Top = 1240
imgSettings.Left = 560
End Sub

Private Sub imgWebsite_Click()
ShellExecute 0, "open", "http://www.agamsaran.bravehost.com/", "", "", 10
End Sub

Private Sub imgWebsite_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgWebsite.Top = 2140
imgWebsite.Left = 630
End Sub

Private Sub imgWebsite_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
imgWebsite.Top = 2120
imgWebsite.Left = 600
End Sub

