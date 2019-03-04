VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IRjlWinAmpGenPlugin
Private m_Wrapper As IRjlWinAmpGenWrapper

Private Sub IRjlWinAmpGenPlugin_Initialize(ByVal Wrapper As RjlWinAmpGenLib.IRjlWinAmpGenWrapper)
  Set m_Wrapper = Wrapper
  m_Wrapper.Description = "WinAmp DX Joystick Control (Gen_DX_Joy.Plugin.dll)"

  LoadSettings
  hwndWinamp = FindWindow("Winamp v1.x", vbNullString)
  
  Call InitDirectInput
  Call InitX
  
  Running = True
  
  StartLoop
End Sub

Private Sub IRjlWinAmpGenPlugin_Configure()
'  Load frmOptions
'  frmOptions.LoadForm
  frmOptions.Show 1
End Sub

Private Sub IRjlWinAmpGenPlugin_Quit()
  Running = False
End Sub
