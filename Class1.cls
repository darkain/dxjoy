VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IRjlWinAmpGenPlugin
Private m_Wrapper As IRjlWinAmpGenWrapper

Private Sub IRjlWinAmpGenPlugin_Initialize(ByVal Wrapper As RjlWinAmpGenLib.IRjlWinAmpGenWrapper)
  Set m_Wrapper = Wrapper
  m_Wrapper.Description = "WinAmp DX Joystick Control (Gen_DX_Joy.Plugin.dll)"

  LoadSettings
  frmMain.Show
  frmMain.Hide
End Sub

Private Sub IRjlWinAmpGenPlugin_Configure()
  'FrmConfig.Show 1
  Call MsgBox("Configuration is not needed with this, so why bother?", vbOKOnly, "WinAmp RM-900 Plug-In V1.4")
End Sub

Private Sub IRjlWinAmpGenPlugin_Quit()
  Unload frmMain
End Sub


