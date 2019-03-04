VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DX Joystick plug-in for WinAmp 2.X"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "&Enabled"
      Height          =   375
      Left            =   4920
      TabIndex        =   38
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buttons"
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   4455
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   15
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   14
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   13
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   12
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   11
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   10
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   9
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   8
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   7
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   6
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   5
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   4
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   3
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   2
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "F"
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   54
         ToolTipText     =   "Button F (15)"
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "E"
         Height          =   255
         Index           =   14
         Left            =   2280
         TabIndex        =   53
         ToolTipText     =   "Button E (14)"
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "D"
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   52
         ToolTipText     =   "Button D (13)"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "C"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   51
         ToolTipText     =   "Button C (12)"
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "B"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   50
         ToolTipText     =   "Button B (11)"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "A"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   49
         ToolTipText     =   "Button A (10)"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "9"
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   48
         ToolTipText     =   "Button 9"
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "8"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   47
         ToolTipText     =   "Button 8"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "7"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Button 7"
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "Button 6"
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "Button 5"
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Button 4"
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Button 3"
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   41
         ToolTipText     =   "Button 2"
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Button 1"
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Button 0"
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Axis"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   15
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   14
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   13
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   12
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   11
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   10
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   9
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   8
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   7
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   6
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   5
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "S2 -"
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   17
         ToolTipText     =   "Axis Slider 2 Negative"
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "S2 +"
         Height          =   255
         Index           =   14
         Left            =   2280
         TabIndex        =   16
         ToolTipText     =   "Axis Slider 2 Positive"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "S1 -"
         Height          =   255
         Index           =   13
         Left            =   2280
         TabIndex        =   15
         ToolTipText     =   "Axis Slider 1 Negative"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "S1 +"
         Height          =   255
         Index           =   12
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Axis Slider 1 Positive"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RZ -"
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   13
         ToolTipText     =   "Axis Z Rotation Negative"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RZ +"
         Height          =   255
         Index           =   10
         Left            =   2280
         TabIndex        =   12
         ToolTipText     =   "Axis Z Rotation Positive"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RY -"
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   11
         ToolTipText     =   "Axis Y Rotation Negative"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RY +"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   10
         ToolTipText     =   "Axis Y Rotation Positive"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RX -"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Axis X Rotation Negative"
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "RX +"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Axis X Rotation Positive"
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Z -"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Axis Z Negative"
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Z +"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Axis Z Positive"
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Y -"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Axis Y Negative"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Y +"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Axis Y Positive"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "X -"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Axis X Negative"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "X +"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Axis X Positive"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   4920
      Picture         =   "frmOptions.frx":0000
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Vince Milum Jr. AKA -=-Darkain Dragoon-=-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   22
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by:"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   6120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo ErrHan
  Dim i As Integer
  
  For i = 0 To 7
    Select Case Combo1(i * 2).ListIndex
      Case 1
        Amp_AxisPos(i) = Cmd_Play
      Case 2
        Amp_AxisPos(i) = Cmd_Pause
      Case 3
        Amp_AxisPos(i) = Cmd_Stop
      Case 4
        Amp_AxisPos(i) = Cmd_PrevTrack
      Case 5
        Amp_AxisPos(i) = Cmd_NextTrack
      Case 6
        Amp_AxisPos(i) = Cmd_FastForward
      Case 7
        Amp_AxisPos(i) = Cmd_Rewind
      Case 8
        Amp_AxisPos(i) = Cmd_Shuffle
      Case 9
        Amp_AxisPos(i) = Cmd_RepeatAll
      Case 10
        Amp_AxisPos(i) = Cmd_Up10
      Case 11
        Amp_AxisPos(i) = Cmd_Down10
      Case 12
        Amp_AxisPos(i) = Cmd_DblSize
      Case 13
        Amp_AxisPos(i) = Cmd_EQ
      Case 14
        Amp_AxisPos(i) = Cmd_Playlist
      Case 15
        Amp_AxisPos(i) = Cmd_MiniBrowser
      Case 16
        Amp_AxisPos(i) = Cmd_Vis
      Case Else
        Amp_AxisPos(i) = 0
    End Select
  
    
    Select Case Combo1(i * 2 + 1).ListIndex
      Case 1
        Amp_AxisNeg(i) = Cmd_Play
      Case 2
        Amp_AxisNeg(i) = Cmd_Pause
      Case 3
        Amp_AxisNeg(i) = Cmd_Stop
      Case 4
        Amp_AxisNeg(i) = Cmd_PrevTrack
      Case 5
        Amp_AxisNeg(i) = Cmd_NextTrack
      Case 6
        Amp_AxisNeg(i) = Cmd_FastForward
      Case 7
        Amp_AxisNeg(i) = Cmd_Rewind
      Case 8
        Amp_AxisNeg(i) = Cmd_Shuffle
      Case 9
        Amp_AxisNeg(i) = Cmd_RepeatAll
      Case 10
        Amp_AxisNeg(i) = Cmd_Up10
      Case 11
        Amp_AxisNeg(i) = Cmd_Down10
      Case 12
        Amp_AxisNeg(i) = Cmd_DblSize
      Case 13
        Amp_AxisNeg(i) = Cmd_EQ
      Case 14
        Amp_AxisNeg(i) = Cmd_Playlist
      Case 15
        Amp_AxisNeg(i) = Cmd_MiniBrowser
      Case 16
        Amp_AxisNeg(i) = Cmd_Vis
      Case Else
        Amp_AxisNeg(i) = 0
    End Select
  Next i
  
  
  For i = 0 To 15
    Select Case Combo2(i).ListIndex
      Case 1
        Amp_Buttons(i) = Cmd_Play
      Case 2
        Amp_Buttons(i) = Cmd_Pause
      Case 3
        Amp_Buttons(i) = Cmd_Stop
      Case 4
        Amp_Buttons(i) = Cmd_PrevTrack
      Case 5
        Amp_Buttons(i) = Cmd_NextTrack
      Case 6
        Amp_Buttons(i) = Cmd_FastForward
      Case 7
        Amp_Buttons(i) = Cmd_Rewind
      Case 8
        Amp_Buttons(i) = Cmd_Shuffle
      Case 9
        Amp_Buttons(i) = Cmd_RepeatAll
      Case 10
        Amp_Buttons(i) = Cmd_Up10
      Case 11
        Amp_Buttons(i) = Cmd_Down10
      Case 12
        Amp_Buttons(i) = Cmd_DblSize
      Case 13
        Amp_Buttons(i) = Cmd_EQ
      Case 14
        Amp_Buttons(i) = Cmd_Playlist
      Case 15
        Amp_Buttons(i) = Cmd_MiniBrowser
      Case 16
        Amp_Buttons(i) = Cmd_Vis
      Case Else
        Amp_Buttons(i) = 0
    End Select
  Next i
  
  If -(Running) <> Check1.Value Then
    If Running Then
      Running = False
    Else
      Running = True
      StartLoop
    End If
  End If
  
  Unload Me
Exit Sub

ErrHan:
  MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Public Sub LoadForm()
  Dim i As Integer
  
  For i = 0 To 15
    Combo1(i).AddItem "(Nothing)"
    Combo1(i).AddItem "Play"                '1
    Combo1(i).AddItem "Pause"               '2
    Combo1(i).AddItem "Stop"                '3
    Combo1(i).AddItem "Prev Track"          '4
    Combo1(i).AddItem "Next Track"          '5
    Combo1(i).AddItem "Fast Fwd"            '6
    Combo1(i).AddItem "Rewind"              '7
    Combo1(i).AddItem "Shuffle"             '8
    Combo1(i).AddItem "Repeate All"         '9
    Combo1(i).AddItem "Up by 10"           '10
    Combo1(i).AddItem "Down by 10"         '11
    Combo1(i).AddItem "Double Size"        '12
    Combo1(i).AddItem "Equilizer"          '13
    Combo1(i).AddItem "Playlist"           '14
    Combo1(i).AddItem "Mini Browser"       '15
    Combo1(i).AddItem "Vis Plug-in"        '16
    Combo1(i).ToolTipText = Label1(i).ToolTipText
  
    Combo2(i).AddItem "(Nothing)"
    Combo2(i).AddItem "Play"                '1
    Combo2(i).AddItem "Pause"               '2
    Combo2(i).AddItem "Stop"                '3
    Combo2(i).AddItem "Prev Track"          '4
    Combo2(i).AddItem "Next Track"          '5
    Combo2(i).AddItem "Fast Fwd"            '6
    Combo2(i).AddItem "Rewind"              '7
    Combo2(i).AddItem "Shuffle"             '8
    Combo2(i).AddItem "Repeate All"         '9
    Combo2(i).AddItem "Up by 10"           '10
    Combo2(i).AddItem "Down by 10"         '11
    Combo2(i).AddItem "Double Size"        '12
    Combo2(i).AddItem "Equilizer"          '13
    Combo2(i).AddItem "Playlist"           '14
    Combo2(i).AddItem "Mini Browser"       '15
    Combo2(i).AddItem "Vis Plug-in"        '16
    Combo2(i).ToolTipText = Label4(i).ToolTipText
  Next i
  
  For i = 0 To 7
    Select Case Amp_AxisPos(i)
      Case Cmd_Play
        Combo1(i * 2).ListIndex = 1
      Case Cmd_Pause
        Combo1(i * 2).ListIndex = 2
      Case Cmd_Stop
        Combo1(i * 2).ListIndex = 3
      Case Cmd_FastForward
        Combo1(i * 2).ListIndex = 6
      Case Cmd_Rewind
        Combo1(i * 2).ListIndex = 7
      Case Cmd_PrevTrack
        Combo1(i * 2).ListIndex = 4
      Case Cmd_NextTrack
        Combo1(i * 2).ListIndex = 5
      Case Cmd_Shuffle
        Combo1(i * 2).ListIndex = 8
      Case Cmd_RepeatAll
        Combo1(i * 2).ListIndex = 9
      Case Cmd_Vis
        Combo1(i * 2).ListIndex = 16
      Case Cmd_DblSize
        Combo1(i * 2).ListIndex = 12
      Case Cmd_EQ
        Combo1(i * 2).ListIndex = 13
      Case Cmd_Playlist
        Combo1(i * 2).ListIndex = 14
      Case Cmd_MiniBrowser
        Combo1(i * 2).ListIndex = 15
      Case Cmd_Up10
        Combo1(i * 2).ListIndex = 10
      Case Cmd_Down10
        Combo1(i * 2).ListIndex = 11
      Case Else
        Combo1(i * 2).ListIndex = 0
    End Select
  
  
    Select Case Amp_AxisNeg(i)
      Case Cmd_Play
        Combo1(i * 2 + 1).ListIndex = 1
      Case Cmd_Pause
        Combo1(i * 2 + 1).ListIndex = 2
      Case Cmd_Stop
        Combo1(i * 2 + 1).ListIndex = 3
      Case Cmd_FastForward
        Combo1(i * 2 + 1).ListIndex = 6
      Case Cmd_Rewind
        Combo1(i * 2 + 1).ListIndex = 7
      Case Cmd_PrevTrack
        Combo1(i * 2 + 1).ListIndex = 4
      Case Cmd_NextTrack
        Combo1(i * 2 + 1).ListIndex = 5
      Case Cmd_Shuffle
        Combo1(i * 2 + 1).ListIndex = 8
      Case Cmd_RepeatAll
        Combo1(i * 2 + 1).ListIndex = 9
      Case Cmd_Vis
        Combo1(i * 2 + 1).ListIndex = 16
      Case Cmd_DblSize
        Combo1(i * 2 + 1).ListIndex = 12
      Case Cmd_EQ
        Combo1(i * 2 + 1).ListIndex = 13
      Case Cmd_Playlist
        Combo1(i * 2 + 1).ListIndex = 14
      Case Cmd_MiniBrowser
        Combo1(i * 2 + 1).ListIndex = 15
      Case Cmd_Up10
        Combo1(i * 2 + 1).ListIndex = 10
      Case Cmd_Down10
        Combo1(i * 2 + 1).ListIndex = 11
      Case Else
        Combo1(i * 2 + 1).ListIndex = 0
    End Select
  Next i
  
  For i = 0 To 15
    Select Case Amp_Buttons(i)
      Case Cmd_Play
        Combo2(i).ListIndex = 1
      Case Cmd_Pause
        Combo2(i).ListIndex = 2
      Case Cmd_Stop
        Combo2(i).ListIndex = 3
      Case Cmd_FastForward
        Combo2(i).ListIndex = 6
      Case Cmd_Rewind
        Combo2(i).ListIndex = 7
      Case Cmd_PrevTrack
        Combo2(i).ListIndex = 4
      Case Cmd_NextTrack
        Combo2(i).ListIndex = 5
      Case Cmd_Shuffle
        Combo2(i).ListIndex = 8
      Case Cmd_RepeatAll
        Combo2(i).ListIndex = 9
      Case Cmd_Vis
        Combo2(i).ListIndex = 16
      Case Cmd_DblSize
        Combo2(i).ListIndex = 12
      Case Cmd_EQ
        Combo2(i).ListIndex = 13
      Case Cmd_Playlist
        Combo2(i).ListIndex = 14
      Case Cmd_MiniBrowser
        Combo2(i).ListIndex = 15
      Case Cmd_Up10
        Combo2(i).ListIndex = 10
      Case Cmd_Down10
        Combo2(i).ListIndex = 11
      Case Else
        Combo2(i).ListIndex = 0
    End Select
  Next i

  Check1.Value = -(Running)
End Sub

Private Sub Form_Activate()
  LoadForm
End Sub

