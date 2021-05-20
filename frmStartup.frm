VERSION 5.00
Begin VB.Form frmStartup 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   1875
      Left            =   60
      Picture         =   "frmStartup.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   1815
      TabIndex        =   0
      Top             =   60
      Width           =   1875
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         ForeColor       =   &H80000018&
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Net Watch V2.1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000018&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   1995
      Left            =   0
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
