VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2724
   ClientLeft      =   3276
   ClientTop       =   2460
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2724
   ScaleWidth      =   5760
   Begin Project1.ColorWheel ColorWheel1 
      Height          =   1680
      Left            =   2760
      TabIndex        =   2
      Top             =   336
      Width           =   2196
      _ExtentX        =   3874
      _ExtentY        =   2963
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E41712&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   588
      Left            =   984
      ScaleHeight     =   564
      ScaleWidth      =   1452
      TabIndex        =   0
      Top             =   840
      Width           =   1476
   End
   Begin VB.Label Label1 
      Caption         =   "Color:"
      Height          =   348
      Left            =   384
      TabIndex        =   1
      Top             =   984
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ColorWheel1_ColorChange()
    Picture1.BackColor = ColorWheel1.Color
End Sub

Private Sub Form_Load()
    ColorWheel1.Color = Picture1.BackColor
End Sub
