VERSION 5.00
Begin VB.UserControl ctlSpeedometer 
   Alignable       =   -1  'True
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   MaskColor       =   &H80000007&
   MaskPicture     =   "ctlSpeedometer.ctx":0000
   ScaleHeight     =   1800
   ScaleWidth      =   1860
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   660
      TabIndex        =   5
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label lblBits 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Kb/s"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   435
   End
   Begin VB.Line lineSpeed 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   300
      X2              =   900
      Y1              =   910
      Y2              =   910
   End
   Begin VB.Label lbSpeedoVal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1140
      TabIndex        =   3
      Top             =   900
      Width           =   315
   End
   Begin VB.Label lbSpeedoVal 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   420
      Width           =   375
   End
   Begin VB.Label lbSpeedoVal 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   660
      Width           =   315
   End
   Begin VB.Label lblSpeedoCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speedometer"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Image ImgMeter 
      Appearance      =   0  'Flat
      Height          =   1800
      Left            =   0
      Picture         =   "ctlSpeedometer.ctx":AEA2
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "ctlSpeedometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_MaxValue As Double     'Max Scale value
Private m_LineMeter As Double   ' line meter on the speed-display
Private m_CurrentSpeed As String   ' datatransfer..speed
Private m_LenSpeedMeter As Integer   'defines the length of the meter..
Private m_X, m_Y As Integer   'Hold the startcoord for the meter..
Private ConvPI As Double  'conversion to Radius..

Private Const PI = 3.14159265358979  'Define PI .. needed for the sin & cos
Private Const ScaleLinearityFactor = 0.75  'scalefactor bmp is no

Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


Private Sub ImgMeter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub



Private Sub lblBits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbSpeedoVal_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtSpeed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
'Repaint the picture..for a ambient underground..
Dim TEMP_COLOR As Long
    m_MaxValue = 100
    lblSpeedoCaption = "."
    m_LenSpeedMeter = lineSpeed.X2 - lineSpeed.X1
    m_X = lineSpeed.X2
    m_Y = lineSpeed.Y2
    lblSpeed = "0"
    'Set the angle - 45 degrees = -pi/4

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1770
    UserControl.Width = 1815
End Sub

Public Property Let MaxValue(val As Long)
     m_MaxValue = val
     'change the counter data...
      ChangeDisplay CLng(m_MaxValue)
End Property

Public Property Get MaxValue() As Long
                MaxValue = CLng(m_MaxValue)
End Property


Public Property Get SpeedometerCaption() As String
        SpeedometerCaption = lblSpeedoCaption
End Property

Public Property Let SpeedometerCaption(ByVal vNewValue As String)
    lblSpeedoCaption = vNewValue
End Property

Public Property Get ActiveSpeed_line() As Double
'Displays the line correct in the scale
    ActiveSpeed_line = m_LineMeter
End Property

Public Property Let ActiveSpeed_line(ByVal vNewValue As Double)
'Displays the line correct in the scale
If vNewValue <= m_MaxValue Then
     m_LineMeter = vNewValue
     ConvPI = ConvertMaxToPi(m_LineMeter, m_MaxValue)
     SetMeter (ConvPI)
Else
    m_LineMeter = m_MaxValue
    ConvPI = ConvertMaxToPi(m_LineMeter, m_MaxValue)
    SetMeter (ConvPI)
End If
     'Check if the linespeed not > max value display..
End Property


Public Property Get NumericSpeed() As String
        NumericSpeed = m_CurrentSpeed
End Property

Public Property Let NumericSpeed(ByVal vNewValue As String)
    m_CurrentSpeed = vNewValue
    lblSpeed = m_CurrentSpeed
End Property

Private Sub ChangeDisplay(max As Long)
'Change the labels that indicate a part of the speed on the display.
'lbSpeedoVal(0) = "0"
lbSpeedoVal(1) = CStr(Round((max / 4), 1)) '1/4 of the speed
lbSpeedoVal(2) = CStr(Round((max / 2), 1)) '1/2 of the speed
lbSpeedoVal(3) = CStr(Round(((max / 4) * 3), 1)) ' 3/4 of the speed
'lbSpeedoVal(4) = CStr(max)  'Max speed
End Sub

Private Function ConvertMaxToPi(val As Double, max As Double) As Double
'convert the val -> max to Radius...
Dim div1 As Double
If val > max Then val = max
div1 = val / max
' div = 2 => PI/2
If val > max / 2 Then
    ConvertMaxToPi = (PI * div1 + (PI / 4) * ScaleLinearityFactor)  '-45 degrees..
    If val > (max * (3 / 4)) Then
        'autoscale correction
       ConvertMaxToPi = (PI * div1 + (PI / 4) * (ScaleLinearityFactor + 0.2))
    End If
Else
If val < (max / 9) Then
    ConvertMaxToPi = (PI * div1 - (PI / 4)) '* ScaleLinearityFactor)  '-45 degrees..
    Else
    'autoscale correction
    ConvertMaxToPi = (PI * div1 - (PI / 4) * ScaleLinearityFactor) '-45 degrees..
    End If
End If
End Function

Private Sub SetMeter(ByVal val1 As Double)
Dim vall As Double     'use it to correct the scale..
vall = vall * ScaleLinearityFactor
'End If
'Set the analoge meter on the display..
lineSpeed.X1 = lineSpeed.X2 - (CDbl(m_LenSpeedMeter * (Cos(val1)))) 'ok

lineSpeed.Y1 = lineSpeed.Y2 - (CDbl(m_LenSpeedMeter * (Sin(val1))))
End Sub
