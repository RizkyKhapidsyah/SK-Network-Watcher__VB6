VERSION 5.00
Begin VB.UserControl ctlDigital 
   BackStyle       =   0  'Transparent
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   300
   ScaleWidth      =   2325
   Begin VB.TextBox txtSpeedval 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   12
      Text            =   "0"
      Top             =   0
      Width           =   555
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   600
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      BackColor       =   &H80000012&
      Height          =   255
      Index           =   1
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
   Begin VB.PictureBox picLED 
      Height          =   375
      Index           =   0
      Left            =   1140
      ScaleHeight     =   315
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   420
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kb/s"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1860
      TabIndex        =   11
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "ctlDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_MaxVal As Long
Private m_CurVal As Long
Private OldVal As Byte

Public Property Get MaxVal() As Long
    MaxVal = m_MaxVal
End Property

Public Property Let MaxVal(ByVal vNewValue As Long)
     m_MaxVal = vNewValue
End Property

Public Property Get CurrentValue() As Double
    CurrentValue = m_CurVal
End Property

Public Property Let CurrentValue(ByVal vNewValue As Double)
Dim ValLED As Double
Dim cnt As Byte
        m_CurVal = vNewValue
        If m_CurVal > m_MaxVal Then m_CurVal = m_MaxVal
        'Set the Value..
        ValLED = CDbl(m_MaxVal / 10)
        'Clear the old values... digits..
        If CByte(m_CurVal / ValLED) < OldVal Then
            'Clear the upper leds..
            For cnt = CByte(m_CurVal / ValLED) To OldVal
                picLED(cnt).BackColor = vbBlack
            Next
        End If
        'Displays the leds..
        For cnt = 1 To (CByte(m_CurVal / ValLED))
            If cnt < 6 Then
                picLED(cnt).BackColor = vbGreen
            End If
            If cnt >= 6 And cnt < 9 Then
                picLED(cnt).BackColor = vbYellow
            End If
            If cnt >= 9 Then
                picLED(cnt).BackColor = vbRed
            End If
        Next
        'Set the speed in the textbox.
         OldVal = CByte(m_CurVal / ValLED)
End Property


Private Sub UserControl_Initialize()
    OldVal = 0  'No display.. active LEDS
End Sub

Public Property Get txtSpeed() As String

End Property

Public Property Let txtSpeed(ByVal vNewValue As String)
    txtSpeedval = vNewValue
End Property
