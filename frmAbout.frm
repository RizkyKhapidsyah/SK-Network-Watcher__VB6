VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  About"
   ClientHeight    =   3195
   ClientLeft      =   3450
   ClientTop       =   3030
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   120
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   1755
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   720
      Width           =   1875
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   3
         X1              =   900
         X2              =   360
         Y1              =   960
         Y2              =   660
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "verburgh.peter@skynet.be"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2100
      TabIndex        =   5
      Top             =   2040
      Width           =   2475
   End
   Begin VB.Label Label2 
      Caption         =   "Questions ?"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2100
      TabIndex        =   4
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Measure the speed of your network connection by analoge/digital meters.      Look at the information about your connection.."
      Height          =   795
      Left            =   2100
      TabIndex        =   3
      Top             =   780
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "netwatch"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About Net Watch ...."
    lblVersion = modStart.NetWatchVersion
End Sub

