VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3336
   ClientLeft      =   2652
   ClientTop       =   2160
   ClientWidth     =   4908
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
   ScaleHeight     =   3336
   ScaleWidth      =   4908
   Begin VB.CommandButton Command1 
      Caption         =   "Change color"
      Height          =   444
      Left            =   2736
      TabIndex        =   1
      Top             =   1248
      Width           =   1308
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1356
      Left            =   984
      ScaleHeight     =   1308
      ScaleWidth      =   1356
      TabIndex        =   0
      Top             =   888
      Width           =   1404
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim iDlg As New ColorDialog
    
    SetColorDialogCaptions iDlg

 '   iDlg.SetCompact
    iDlg.SelectionParameter = cdParameterHue
    iDlg.SetComplete True
    iDlg.Color = Picture1.BackColor
    If iDlg.Show Then
        Picture1.BackColor = iDlg.Color
    End If
    
End Sub


