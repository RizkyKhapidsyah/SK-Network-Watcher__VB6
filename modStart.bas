Attribute VB_Name = "modStart"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public NetWatchVersion As String  'use for name program + version

Sub Main()
'frmMenu.Show
frmStartup.Show
'Wait 1 seconds..
DoEvents
Sleep 1000
Unload frmStartup
'ModSysTrayMenu.Skin = False
NetWatchVersion = " Net Watch V2.1"
AdaptorPopup = 0
Load frmMenu
frmMenu.Caption = NetWatchVersion

End Sub
