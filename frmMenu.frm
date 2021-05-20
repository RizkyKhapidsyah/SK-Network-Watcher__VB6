VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   Caption         =   "Net Watch V2.1"
   ClientHeight    =   900
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3735
   FillStyle       =   0  'Solid
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrPopup 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   1920
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2880
      Picture         =   "frmMenu.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      ToolTipText     =   "Net Watch"
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lpICON As Long  'holds address ICON part

Private Sub Start()
Dim cnt As Integer   'counter for loop
On Error GoTo Err1
'Dim X() As frmSpeedMeter
Dim TotalAdaptors As New clsAdaptors
ReDim Xmenu(1 To (TotalAdaptors.Count_adaptors + 1)) As frmMain
Max_Adapters = TotalAdaptors.Count_adaptors
'Check if this application runs the first time...
modINI.SearchINIfile
For cnt = 1 To TotalAdaptors.Count_adaptors
    Set Xmenu(cnt) = New frmMain
    Load Xmenu(cnt)
Err1:
    'If ip = 0.0.0.0 => unload it => Err1...
Next

'Object was unloaded because the ip is not valid..
 '  was "O.O.O.O"
 Me.Hide
 Me.Visible = False
End Sub


Private Sub ExitProgram_Click()
Unload Me
End Sub


Private Sub Form_Load()
modInfoAdaptor.m_AdaptorCnt = 0     'set the max counted adapters to ZERO
modInfoAdaptor.m_MaxNewForm = 1
 CreateIcon
 Start
 Me.Hide
 frmMenu.Hide
 Y = 0
 Me.ScaleMode = vbPixels
 Hook Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tel As Byte
For tel = 1 To Max_Adapters
    
    Unload Xmenu(tel)
    
Next
DeleteIcon
End Sub

Public Sub CreateIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = "NET watch V2.1" & Chr$(0)
    lpICON = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Public Sub DeleteIcon()
    Dim Tic As NOTIFYICONDATA
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    lpICON = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   X = X / Screen.TwipsPerPixelX


    Select Case X
        Case WM_LBUTTONDOWN
        MenuTrack Me
        Case WM_RBUTTONDOWN
   '
     '  PopupMenu Menu1
      MenuTrack Me
        Case WM_MOUSEMOVE
    '
        Case WM_LBUTTONDBLCLK
    '
    End Select

End Sub

Sub ReloadMe()
Dim tel As Byte
For tel = 1 To Max_Adapters
Unload Xmenu(tel)
Next
DeleteIcon
Me.Caption = " Net Watch V2.1"
modInfoAdaptor.m_AdaptorCnt = 0
modInfoAdaptor.m_MaxNewForm = 1
 CreateIcon
 Start
 Me.Hide
 'frmMenu.Hide
 Y = 0
 Me.ScaleMode = vbPixels
 Hook Me
End Sub

Private Sub tmrPopup_Timer()
' I'm using this timer function for detecting if the mouse it out the popup window..
' then the popup must be hide !
Dim MPA As PointAPI
Dim test&
On Error GoTo Err1
ModSysTrayMenu.GetCursorPos MPA

If MPA.X < (MP.X - RectPopup.Right + RectPopup.Bottom) Or MPA.X > (MP.X + RectPopup.Bottom) Then
   'Hide the popupmenu...
    
    frmPop.Show
    DoEvents
    Unload frmPop
    tmrPopup.Enabled = False

End If
If MPA.Y < MP.Y - ModSysTrayMenu.mnuHeight Then
    'frmPop.Show
    frmPop.Show
    DoEvents
    Unload frmPop
    tmrPopup.Enabled = False
    
End If
Err1:
End Sub
