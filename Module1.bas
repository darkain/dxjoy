Attribute VB_Name = "Module1"
Option Explicit

Global dx As New DirectX7
Global di As DirectInput
Global diDev As DirectInputDevice
Global diDevEnum As DirectInputEnumDevices
Global EventHandle As Long
Global joyCaps As DIDEVCAPS
Global js As DIJOYSTATE
Global DiProp_Dead As DIPROPLONG
Global DiProp_Range As DIPROPRANGE
Global DiProp_Saturation As DIPROPLONG
Global AxisPresent(1 To 8) As Boolean
Global Running As Boolean

Global HasPads As Boolean


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, ByVal lparam As Long) As Long

Type POINTAPI
  x As Long
  y As Long
End Type

Type msg
  hwnd As Long
  message As Long
  wparam As Long
  lparam As Long
  time As Long
  pt As POINTAPI
End Type

Public Const WM_COMMAND = &H111
Public Const WM_USER = &H400
Global hwndWinamp As Long

Global Last_Button(15) As Integer
Global Last_Axis(7) As Integer

Global Amp_Buttons(15) As Long
Global Amp_AxisPos(7) As Long
Global Amp_AxisNeg(7) As Long
Global Const Cmd_Play = 40045
Global Const Cmd_Pause = 40046
Global Const Cmd_Stop = 40047
Global Const Cmd_FastForward = 40148
Global Const Cmd_Rewind = 40144
Global Const Cmd_PrevTrack = 40044
Global Const Cmd_NextTrack = 40048
Global Const Cmd_Shuffle = 40023
Global Const Cmd_RepeatAll = 40022
Global Const Cmd_Vis = 40192
Global Const Cmd_DblSize = 40165
Global Const Cmd_EQ = 40036
Global Const Cmd_Playlist = 40040
Global Const Cmd_MiniBrowser = 40298
Global Const Cmd_Up10 = 40197
Global Const Cmd_Down10 = 40195

Sub Main()

End Sub

Public Sub StartLoop()
  Unload frmOptions
  
  Do
    DoEvents
    DirectXEvent_DXCallback 0
    Sleep (10)
    
    hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
    If hwndWinamp = 0 Then
      Running = False
    End If
    
  Loop Until Running = False
  
  SaveSettings
End Sub

Public Sub SendCommand(ByVal CommandNum As Long)
  If CommandNum = 0 Then Exit Sub
  Dim a As Long
  a = SendMessage(hwndWinamp, WM_COMMAND, CommandNum, 0)
End Sub

Public Sub SaveSettings(Optional DefaultSettings As Boolean = False)
On Error GoTo ErrHan
  Dim i As Integer
    
  If DefaultSettings Then
    Amp_AxisPos(0) = Cmd_PrevTrack
    Amp_AxisNeg(0) = Cmd_NextTrack
    Amp_Buttons(0) = Cmd_Play
    Amp_Buttons(1) = Cmd_Stop
    Amp_Buttons(2) = Cmd_Pause
  End If
  
  Open App.Path & "\ControlAMP.INI" For Output As #1
    Print #1, "[Joypad1]"
    
    For i = 0 To 7
      Print #1, "AxisPos" & i & "=" & Amp_AxisPos(i)
      Print #1, "AxisNeg" & i & "=" & Amp_AxisNeg(i)
    Next i
      
    For i = 0 To 15
      Print #1, "Buttons" & Hex(i) & "=" & Amp_Buttons(i)
    Next i
  Close #1
Exit Sub

ErrHan:
  MsgBox Err.Number & "(SaveSettings) - " & Err.Description
End Sub

Public Sub LoadSettings()
On Error GoTo ErrHan
  Dim i As Integer
  
  If Not FileExist(App.Path & "\ControlAMP.INI") Then
    SaveSettings (True)
  End If
  
  Dim a As String
  Dim Section As Integer
  Dim StrLoc As Integer
  
  Open App.Path & "\ControlAMP.INI" For Input As #1
    Do
      Input #1, a
      If InStr(a, ";") Then
        a = Left(a, InStr(a, ";") - 1)
      End If
            
      a = Trim$(a)
      If Left(a, 1) = "[" And Right(a, 1) = "]" Then
        a = UCase(Mid(a, 2, Len(a) - 2))
        Select Case a
          Case "JOYPAD1"
            Section = 1
        End Select
      
      Else
        Select Case UCase(Left(a, 7))
          Case "BUTTONS"
            Amp_Buttons(CInt("&H" & Mid(a, 8, 1))) = Val(Mid(a, InStr(a, "=") + 1))
          Case "AXISPOS"
            Amp_AxisPos(CInt("&H" & Mid(a, 8, 1))) = Val(Mid(a, InStr(a, "=") + 1))
          Case "AXISNEG"
            Amp_AxisNeg(CInt("&H" & Mid(a, 8, 1))) = Val(Mid(a, InStr(a, "=") + 1))
        End Select
      End If
    Loop Until EOF(1)
  
  Close #1
Exit Sub

ErrHan:
  MsgBox Err.Number & "(LoadSettings) - " & Err.Description
End Sub


Public Function FileExist(FileName As String)
  On Error Resume Next
  Open FileName For Input As #1
  Close #1
  
  If Err.Number Then
    Err.Clear
    FileExist = False
  Else
    FileExist = True
  End If
End Function





Public Sub InitX()
  On Local Error Resume Next
    
  'Create the joystick device
  Set diDev = Nothing
  Set diDev = di.CreateDevice(diDevEnum.GetItem(1).GetGuidInstance)
  diDev.SetCommonDataFormat DIFORMAT_JOYSTICK
  'diDev.SetCooperativeLevel Me.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  diDev.SetCooperativeLevel frmOptions.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    
  ' Find out what device objects it has
  diDev.GetCapabilities joyCaps
  Call IdentifyAxes(diDev)
    
  ' Ask for notification of events
  Call diDev.SetEventNotification(EventHandle)

  ' Set deadzone for X and Y axis to 10 percent of the range of travel
  With DiProp_Dead
    .lData = 1000
    .lObj = DIJOFS_X
    .lSize = Len(DiProp_Dead)
    .lHow = DIPH_BYOFFSET
    .lObj = DIJOFS_X
    diDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    .lObj = DIJOFS_Y
    diDev.SetProperty "DIPROP_DEADZONE", DiProp_Dead
  End With
    
  ' Set saturation zones for X and Y axis to 5 percent of the range
  With DiProp_Saturation
    .lData = 9500
    .lHow = DIPH_BYOFFSET
    .lSize = Len(DiProp_Saturation)
    .lObj = DIJOFS_X
     diDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
    .lObj = DIJOFS_Y
     diDev.SetProperty "DIPROP_SATURATION", DiProp_Saturation
  End With
    
  SetProp
    
    
  diDev.Acquire
    
  ' Get the list of current properties
  ' USB joysticks wont call this callback until you play with the joystick
  ' so we call the callback ourselves the first time
  DirectXEvent_DXCallback 0
    
  ' Poll the device so that events are sure to be signaled.
  ' Usually this would be done in Sub Main or in the game rendering loop.
    
'    While running = True
'        DoEvents
'        diDev.Poll
'    Wend
End Sub

Sub SetProp()
  ' Set range for all axes
  With DiProp_Range
    .lHow = DIPH_DEVICE
    .lSize = Len(DiProp_Range)
    .lMin = 0
    .lMax = 10000
  End With
  diDev.SetProperty "DIPROP_RANGE", DiProp_Range
End Sub

Sub IdentifyAxes(diDev As DirectInputDevice)

 ' It's not enough to count axes; we need to know which in particular
 ' are present.
   
  Dim didoEnum As DirectInputEnumDeviceObjects
  Dim dido As DirectInputDeviceObjectInstance
  Dim i As Integer
   
  For i = 1 To 8
    AxisPresent(i) = False
  Next
   
  ' Enumerate the axes
  Set didoEnum = diDev.GetDeviceObjectsEnum(DIDFT_AXIS)
   
  ' Check data offset of each axis to learn what it is
   
  For i = 1 To didoEnum.GetCount
    Set dido = didoEnum.GetItem(i)
      
    Select Case dido.GetOfs
      Case DIJOFS_X
        AxisPresent(1) = True
      Case DIJOFS_Y
        AxisPresent(2) = True
      Case DIJOFS_Z
        AxisPresent(3) = True
      Case DIJOFS_RX
        AxisPresent(4) = True
      Case DIJOFS_RY
        AxisPresent(5) = True
      Case DIJOFS_RZ
        AxisPresent(6) = True
      Case DIJOFS_SLIDER0
        AxisPresent(7) = True
      Case DIJOFS_SLIDER1
        AxisPresent(8) = True
    End Select
   
  Next
End Sub



Public Sub DirectXEvent_DXCallback(ByVal eventid As Long)
  Dim i As Integer
  Dim ListPos As Integer
  Dim S As String
    
  If diDev Is Nothing Then Exit Sub
        
  '' Get the device info
  On Local Error Resume Next
  diDev.GetDeviceStateJoystick js
  If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
    diDev.Acquire
    Exit Sub
  End If
'  diDev.GetDeviceStateJoystick js
  diDev.Poll
    
    
  On Error GoTo err_out
    
  ' Display axis coordinates
  ListPos = 0
  For i = 1 To 8
    If AxisPresent(i) Then
      
      Select Case i
        Case 1
          If js.x > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.x < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If

        
        Case 2
          If js.y > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.y < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If

        
        Case 3
          If js.z > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.z < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If

        
        Case 4
          If js.rx > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.rx < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If


        Case 5
          If js.ry > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.ry < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If


        Case 6
          If js.rz > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.rz < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If

        
        Case 7
          If js.slider(0) > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.slider(0) < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If


        Case 8
          If js.slider(1) > 7500 Then
            If Last_Axis(i - 1) <> 1 Then
              Call SendCommand(Amp_AxisPos(i - 1))
            End If
            Last_Axis(i - 1) = 1
          
          ElseIf js.slider(1) < 2500 Then
            If Last_Axis(i - 1) <> -1 Then
              Call SendCommand(Amp_AxisNeg(i - 1))
            End If
            Last_Axis(i - 1) = -1
          
          Else
            Last_Axis(i - 1) = 0
          End If

      End Select
    End If
  Next
  
  ' Buttons
  For i = 0 To joyCaps.lButtons - 1
    If Last_Button(i) <> js.buttons(i) Then
      If js.buttons(i) <> 0 Then
        Call SendCommand(Amp_Buttons(i))
      End If
      
      Last_Button(i) = js.buttons(i)
    End If
  Next
        
  ' Hats
  For i = 0 To joyCaps.lPOVs - 1
'    lstHat.List(i) = "POV " + CStr(i + 1) + ": " + CStr(js.POV(i))
  Next
    
Exit Sub
    
err_out:
    MsgBox Err.Description & " : " & Err.Number, vbApplicationModal
'    End
End Sub



Sub InitDirectInput()
  Set di = dx.DirectInputCreate()
  Set diDevEnum = di.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
  If diDevEnum.GetCount = 0 Then
    MsgBox "No joystick attached."
    HasPads = False
    Running = False
  End If
    
  'Add attached joysticks to the listbox
  Dim i As Integer
  For i = 1 To diDevEnum.GetCount
    'diDevEnum.GetItem(i).GetInstanceName
    'a list of ALL gamepads
  Next
    
  HasPads = True
Exit Sub
    
Error_Out:
  MsgBox "Error initializing DirectInput."
  HasPads = False
  Running = False
End Sub
