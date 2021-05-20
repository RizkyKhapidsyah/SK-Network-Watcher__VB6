VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IP address"
   ClientHeight    =   5760
   ClientLeft      =   8100
   ClientTop       =   465
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   Begin NetWatch.ctlDigital ctlDigUPL 
      Height          =   315
      Left            =   3720
      TabIndex        =   23
      Top             =   360
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
   End
   Begin NetWatch.ctlDigital ctlDigDNL 
      Height          =   315
      Left            =   3720
      TabIndex        =   21
      Top             =   120
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H00000000&
      Height          =   2475
      Left            =   30
      TabIndex        =   10
      Top             =   1860
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label lblUnkProtocols 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   1800
         TabIndex        =   36
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblErrSND 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   35
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblErrRCV 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   1500
         TabIndex        =   34
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Unknow protocols Rcv:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   33
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Packets Snd:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Error Packets Rcv :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   31
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kb/s"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   30
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Kb/s"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   2100
         TabIndex        =   29
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblMaxDownLoadScale 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1560
         TabIndex        =   28
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lblMaxUP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1380
         TabIndex        =   27
         Top             =   2040
         Width           =   675
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Download scale:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   2280
         Width           =   1635
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Upload scale: "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection State"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblBytesSends 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1320
         TabIndex        =   18
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblBytesSend 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Send: "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblBytesReceived 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bytes Received:"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lblSpeedConnection 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1500
         TabIndex        =   14
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Connection :"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblAdaptorType 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Type adaptor : "
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Frame FraScale 
      BackColor       =   &H80000008&
      Caption         =   "Scale settings"
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   30
      TabIndex        =   2
      Top             =   4500
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton cmdScaleAccept 
         Caption         =   "Accept Settings"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   840
         Width           =   2355
      End
      Begin VB.TextBox txtDownl_KBps 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1260
         TabIndex        =   8
         Text            =   "56"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtUpload_KBps 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1260
         TabIndex        =   4
         Text            =   "0"
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "kB/s"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "kB/s"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Download:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Upload:"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   540
         Width           =   975
      End
   End
   Begin NetWatch.ctlSpeedometer ctlUploader 
      Height          =   1770
      Left            =   1860
      TabIndex        =   1
      Top             =   60
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3122
   End
   Begin NetWatch.ctlSpeedometer ctlDownloader 
      Height          =   1770
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3122
   End
   Begin VB.Timer tmrCount 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6120
      Top             =   2820
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      DrawMode        =   5  'Not Copy Pen
      X1              =   3660
      X2              =   6300
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Kb/s Upl"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5580
      TabIndex        =   24
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Kb/s Dnl"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5580
      TabIndex        =   22
      Top             =   300
      Width           =   1275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As _
Long, ByVal cy As Long, ByVal wFlags As Long)
'-----------------------------------------------------------------------------------------------------------------------


Private WithEvents Test1 As clsIPStatistics
Attribute Test1.VB_VarHelpID = -1
Private cnt1, cnt2 As Long
Private m_Adaptor As Long
Private NewAdapter As frmMain            'handle for new form in form
Dim NewAdap As Boolean
Dim OldUp, OldDn As Double          'to keep the old value upload/download speed
Dim second As Byte                         'show the digital meter 1 per sec / average
Dim blnChange  As Boolean          ' use for lower overhead.. put meters on zero..


Const WindowH = 2460 - 300
Const WindowW = 3825

'Settings from & to the ini file..
Dim posX As Integer   ' X position for the window
Dim posY As Integer ' Y position for the window
Dim WinState As Byte ' State window 1 = show , 0 = hide
Dim SkinX As String * 10    ' Digital - analoge skin ..
Dim MaxUpload As Long   'Max Upload scale value
Dim MaxDownload As Long 'Max download scale value
Dim oldCntAdaptors As Integer 'hold the previous adapters  (dial up users)

Dim NewApplication As frmMenu
Dim blnValueMeter As Boolean  'Check if the meter selected is analog or digital..

Private Sub GetSettings()
    posX = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "posX", "?")
    posY = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "posY", "?")
    SkinX = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "Skin", "?")
    WinState = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "ShowWindow", "?")
    MaxDownload = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "MaxDownloadSpeed", "?")
    MaxUpload = modINI.sGetINI(CurDir & "\" & "SettingsNet.INI", "Adaptor" & m_Adaptor, "MaxUploadSpeed", "?")
End Sub



Private Sub Start()
'Me.Cls
    Test1.Choose_Adaptor = m_Adaptor
    Test1.Update_Adaptors_Stat

    cnt1 = Test1.BytesRecieved
    cnt2 = Test1.BytesSends
    tmrCount.Enabled = True
    ctlDownloader.MaxValue = MaxDownload
    ctlDownloader.SpeedometerCaption = "Download"
    ctlUploader.MaxValue = MaxUpload
    ctlUploader.SpeedometerCaption = "Upload"
    ctlDigUPL.MaxVal = ctlUploader.MaxValue
    ctlDigDNL.MaxVal = ctlDownloader.MaxValue

End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or _
    SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub cmdHideME_Click()
    fraInfo.Visible = False
    FraScale.Visible = False
    frmMain.Height = WindowsH
End Sub

Private Sub cmdScaleAccept_Click()
        'Accept the settings .. and resize the scales..
    ctlDownloader.MaxValue = val(txtDownl_KBps)
    ctlUploader.MaxValue = val(txtUpload_KBps)
        'Set the digital meter.
    ctlDigUPL.MaxVal = ctlUploader.MaxValue
    ctlDigDNL.MaxVal = ctlDownloader.MaxValue
        'Set the dims.. Maxupload & download..
    MaxUpload = val(txtUpload_KBps)
    MaxDownload = val(txtDownl_KBps)
        'Save the settings to the ini file
    modINI.WriteSettings m_Adaptor, "MaxDownloadSpeed", (CStr(MaxDownload))
    modINI.WriteSettings m_Adaptor, "MaxUploadSpeed", CStr(MaxUpload)
        'Close the window part
    modInfoAdaptor.Xmenu(m_Adaptor).ScaleSettings_Click
End Sub

Private Sub Form_Load()
    Set Test1 = New clsIPStatistics
        NewAdap = False

    AdaptorPopup = AdaptorPopup + 1  'use for the scaling part of popupwindow
    second = 0
    blnChange = False   ' no change captured in overhead part timer..

    modInfoAdaptor.m_AdaptorCnt = modInfoAdaptor.m_AdaptorCnt + 1
    m_Adaptor = m_AdaptorCnt

'Get the old settings...
    GetSettings

    Test1.Choose_Adaptor = modInfoAdaptor.m_AdaptorCnt
    Test1.Update_Adaptors_Stat
'
    oldCntAdaptors = Test1.Found_Adaptors
    Me.Caption = Test1.Local_IP
'Settings Window Default
    Me.Width = WindowW
    Me.Height = WindowH
    Me.Top = posY
    Me.Left = posX

    lblAdaptorType = Test1.Interface_Type
    lblSpeedConnection = Test1.Connection_Speed
    lblBytesReceived = Test1.BytesRecieved
    lblBytesSends = Test1.BytesSends
    lblStatus = Test1.OperationStatus
    lblErrRCV = Test1.ErrorPacketsRcv
    lblErrSND = Test1.ErrorPacketsSnd
    lblUnkProtocols = Test1.UnknowProtocolsRvc

    ctlDownloader.ActiveSpeed_line = 0
    ctlUploader.ActiveSpeed_line = 0

    ctlDigUPL.MaxVal = ctlUploader.MaxValue
    ctlDigDNL.MaxVal = ctlDownloader.MaxValue


'check if previous skin was anloge or digital
    If InStr(1, SkinX, "ANALOGE ", vbString) > 0 Then
        ModSysTrayMenu.Skin(modInfoAdaptor.m_AdaptorCnt) = False
        Me.ShowAnalog_Digital False
    Else
        ModSysTrayMenu.Skin(modInfoAdaptor.m_AdaptorCnt) = True
        Me.ShowAnalog_Digital True
    End If
'Check if the old state , window was visible or not
    If WinState = 1 Then
        Me.Show
    Else
        Me.Hide
    End If


    If Mid(Trim(Test1.Local_IP), 1, 1) = "0" Then
    'MsgBox "Funcky Zero not alowed .."
        Me.Hide
        Me.Visible = False
        AdaptorPopup = AdaptorPopup - 1
    End If
    Start
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save the positions on the screen ...
    modINI.WriteSettings m_Adaptor, "posY", Me.Top
    modINI.WriteSettings m_Adaptor, "posX", Me.Left
'save the window state visible or not ...
    If Me.Visible = True Then
        modINI.WriteSettings m_Adaptor, "ShowWindow", 1
    Else
        modINI.WriteSettings m_Adaptor, "ShowWindow", 0
    End If
' save the skin  analoge or digital
    If ModSysTrayMenu.Skin(m_Adaptor) = False Then
        modINI.WriteSettings m_Adaptor, "Skin", "ANALOGE"
    Else
        modINI.WriteSettings m_Adaptor, "Skin", "DIGITAL"
    End If

    Set Test1 = Nothing
        If NewAdap = True Then
            Unload NewAdapter
        Else
            ' no unload needed , doesn't exist
        End If

End Sub



Public Sub ScaleSettings_Click()
Dim Y As Integer

If blnValueMeter = False Then
    Y = 0
Else
    Y = 1250
End If
If FraScale.Visible = False Then
    'Rescale the form window..
            fraInfo.Visible = False
            FraScale.Visible = True
            FraScale.Top = 1860 - Y
            Me.Height = 3450 - Y
            txtDownl_KBps = CStr(ctlDownloader.MaxValue)
            txtUpload_KBps = CStr(ctlUploader.MaxValue)
            'Set the digital meter.
            ctlDigUPL.MaxVal = ctlUploader.MaxValue
            ctlDigDNL.MaxVal = ctlDownloader.MaxValue
Else
             fraInfo.Visible = False
            FraScale.Visible = False
            FraScale.Top = 1860 - Y
            Me.Height = WindowH - Y
        
End If
End Sub

Private Sub Test1_AdaptorError(ByVal ErrorDiscription As Integer)
Select Case ErrorDiscription
    Case 1
        ' Adaptor to high .. disconnected dial up ??
        tmrCount.Enabled = False
        'm_AdaptorCnt = m_AdaptorCnt - 1
        modInfoAdaptor.m_AdaptorCnt = modInfoAdaptor.m_AdaptorCnt - 1
        DoEvents
        Unload Me
        DoEvents
        'MsgBox "Change captured , connection lost !! " & vbCrLf & "Program will automatic update ", vbInformation
        frmMenu.ReloadMe
        Exit Sub
        'lblStatus = "Disconnected"
End Select

End Sub

Private Sub tmrCount_Timer()
Dim OldRCV, OldSND As Long         ' use it for make te CPU overhead lower
Dim newAdaptors As Integer

'second = second + 1
On Error GoTo Error1
    newAdaptors = Test1.Found_Adaptors 'check it there are new adapters...
    OldRCV = Test1.BytesRecieved ' this part will bes used to get an average of the real download for the click/dig watch
    OldSND = Test1.BytesSends ' this part will bes used to get an average of the real upload for the click/dig watch
    Test1.Choose_Adaptor = m_Adaptor  'choose the adapter for updating the data...
    Test1.Update_Adaptors_Stat  'update the adapter now

'Set the data in the more info...
    lblMaxDownLoadScale = ctlDownloader.MaxValue  'put the setting from download into the label
    lblMaxUP = ctlUploader.MaxValue
'Check if some data is changed.. send - recieved.. if not...
' then exit the subroutine => so on that case we got lower overhead !
    If OldRCV = Test1.BytesRecieved And OldSND = Test1.BytesSends Then

'================== if a new adapter is added to the list. new dial up ? =========
    If newAdaptors > oldCntAdaptors Then
    'New Adaptor , dial up found..
        DoEvents
        Set NewAdapter = New frmMain
        NewAdap = True
        'put some code to unload by Menu Applic
        DoEvents
        frmMenu.ReloadMe
    
    End If
   'check if the meters - aren't on zero  because there is no data send/recieved...
   If blnChange = False Then
        blnChange = True
        ctlDownloader.NumericSpeed = "0"
        ctlDownloader.ActiveSpeed_line = 0
        ctlUploader.ActiveSpeed_line = 0
        ctlUploader.NumericSpeed = "0"
        ctlDigUPL.CurrentValue = 0
        ctlDigDNL.CurrentValue = 0
        ctlDigUPL.txtSpeed = "0"
        ctlDigDNL.txtSpeed = "0"
   End If
   Exit Sub             ' No change in total send & recieved..
Else
    blnChange = False
End If

second = second + 1
'========Check if new adaptors is less than previous ===========================
If newAdaptors < oldCntAdaptors Then
    'New Adaptor , dial up found..
    DoEvents
    Set NewAdapter = New frmMain
    NewAdap = True
    'put some code to unload by Menu Applic
    NewAdapter.Show
    DoEvents
    
     Set NewApplication = New frmMenu
        DoEvents
        Unload frmMenu
End If
'Download section
cnt1 = Test1.BytesRecieved - cnt1
ctlDownloader.NumericSpeed = CStr(Round(((cnt1 / 1024) * 2), 1))
ctlDigDNL.CurrentValue = CLng(ctlDownloader.ActiveSpeed_line)
ctlDigDNL.txtSpeed = ctlDownloader.NumericSpeed
If second = 1 Then
    OldDn = (Round(((cnt1 / 1000) * 2), 1))
End If
'Capture per second..
If second = 2 Then
    ctlDownloader.ActiveSpeed_line = (OldDn + (cnt1 / 1024) * 2) / 2
End If
cnt1 = Test1.BytesRecieved

'Upload Section
cnt2 = Test1.BytesSends - cnt2
ctlUploader.NumericSpeed = CStr(Round((cnt2 / 1024) * 2, 1))
ctlDigUPL.CurrentValue = CLng(ctlUploader.ActiveSpeed_line)
ctlDigUPL.txtSpeed = ctlUploader.NumericSpeed
'Capture per second
If second = 1 Then
    OldUp = (Round(((cnt2 / 1000) * 2), 1))
End If
If second = 2 Then
    ctlUploader.ActiveSpeed_line = (OldUp + (cnt2 / 1024) * 2) / 2
    second = 0
End If
cnt2 = Test1.BytesSends

'Aditional information
lblAdaptorType = Test1.Interface_Type
lblSpeedConnection = Test1.Connection_Speed
lblBytesReceived = Test1.BytesRecieved
lblBytesSends = Test1.BytesSends
lblStatus = Test1.OperationStatus
lblErrRCV = Test1.ErrorPacketsRcv
lblErrSND = Test1.ErrorPacketsSnd
lblUnkProtocols = Test1.UnknowProtocolsRvc


Exit Sub
Error1:
    Unload Me
End Sub

Public Sub View_more_Click()
Dim Y As Integer

If blnValueMeter = False Then
    Y = 0
Else
    Y = 1250
End If


If fraInfo.Visible = False Then
    fraInfo.Visible = True
    FraScale.Visible = False
    fraInfo.Top = 1860 - Y
    Me.Height = 4700 - Y
    Line1.Y1 = 1800 - Y
    Line1.Y2 = Line1.Y1
    Line1.X1 = 1
    Line1.X2 = 2600
    Line1.Visible = True
Else
    fraInfo.Visible = False
    FraScale.Visible = False
    fraInfo.Top = 1860 - Y
    Me.Height = 2100 - Y
    Line1.Visible = False
End If
End Sub

Sub ShowThisWindow()
 Me.Visible = Not Me.Visible
End Sub

Sub ShowAnalog_Digital(value As Boolean)
blnValueMeter = value
'--------------------------------- digital View ------------------------------------------------
If value = True Then
    'Digital view
    ctlDownloader.Visible = False
    ctlUploader.Visible = False
    ctlDigUPL.Top = 1
    ctlDigUPL.Left = 5
    ctlDigDNL.Top = 250
    ctlDigDNL.Left = 5
    ctlDigUPL.Visible = True
    ctlDigDNL.Visible = True
    Me.Height = 850
    Me.Width = 2600
    Label8.Left = 1850
    Label9.Left = 1850
    Label8.Visible = True
    Label9.Visible = True
    ModSysTrayMenu.SkinText(modInfoAdaptor.m_AdaptorCnt) = "Show Analog"
    'ModSysTrayMenu.Skin = True
    ' ----------------------------------------Analoge view -------------------------------
Else
    Me.Height = WindowH
    Me.Width = WindowW
    ctlDownloader.Visible = True
    ctlUploader.Visible = True
     ctlDigUPL.Visible = False
    ctlDigDNL.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    ModSysTrayMenu.SkinText(modInfoAdaptor.m_AdaptorCnt) = "Show Digital"
    'ModSysTrayMenu.Skin = False
End If
FraScale.Visible = False
fraInfo.Visible = False
End Sub

Private Sub txtDownl_KBps_LostFocus()
    txtUpload_KBps.SetFocus
End Sub
