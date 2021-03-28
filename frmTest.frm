VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wheel Color Dialog test"
   ClientHeight    =   8208
   ClientLeft      =   2652
   ClientTop       =   2160
   ClientWidth     =   5892
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8208
   ScaleWidth      =   5892
   Begin VB.Frame Frame10 
      Caption         =   "Complete && Big"
      Height          =   1332
      Left            =   3024
      TabIndex        =   27
      Top             =   6700
      Width           =   2508
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   29
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   28
         Top             =   456
         Width           =   1164
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Use HSL color system"
      Height          =   1332
      Left            =   312
      TabIndex        =   24
      Top             =   6700
      Width           =   2508
      Begin VB.CommandButton Command9 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   26
         Top             =   456
         Width           =   1164
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   25
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Complete"
      Height          =   1332
      Left            =   3024
      TabIndex        =   21
      Top             =   5110
      Width           =   2508
      Begin VB.CommandButton Command8 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   23
         Top             =   456
         Width           =   1164
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   22
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Compact"
      Height          =   1332
      Left            =   312
      TabIndex        =   18
      Top             =   5110
      Width           =   2508
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   20
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   19
         Top             =   456
         Width           =   1164
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Hue selection"
      Height          =   1332
      Left            =   312
      TabIndex        =   15
      Top             =   3548
      Width           =   2508
      Begin VB.CommandButton Command6 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   17
         Top             =   456
         Width           =   1164
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H0059C5F9&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   16
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Hide parameters section"
      Height          =   1332
      Left            =   312
      TabIndex        =   12
      Top             =   1992
      Width           =   2508
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4CE33&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   14
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   13
         Top             =   456
         Width           =   1164
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Hide recent colors"
      Height          =   1332
      Left            =   3024
      TabIndex        =   9
      Top             =   1992
      Width           =   2508
      Begin VB.CommandButton Command4 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   11
         Top             =   456
         Width           =   1164
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0081FAF7&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   10
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Saturation selection"
      Height          =   1332
      Left            =   3024
      TabIndex        =   6
      Top             =   3548
      Width           =   2508
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H002404FF&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   8
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   7
         Top             =   456
         Width           =   1164
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Big"
      Height          =   1332
      Left            =   3024
      TabIndex        =   3
      Top             =   408
      Width           =   2508
      Begin VB.CommandButton Command2 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   5
         Top             =   456
         Width           =   1164
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0061DC7C&
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
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   4
         Top             =   336
         Width           =   756
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Standard"
      Height          =   1332
      Left            =   312
      TabIndex        =   0
      Top             =   408
      Width           =   2508
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   168
         ScaleHeight     =   684
         ScaleWidth      =   732
         TabIndex        =   2
         Top             =   336
         Width           =   756
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change color"
         Height          =   396
         Left            =   1080
         TabIndex        =   1
         Top             =   456
         Width           =   1164
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.Color = Picture1.BackColor
    If oDlg.Show Then
        Picture1.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command2_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.DialogSizeBig = True
    
    oDlg.Color = Picture2.BackColor
    If oDlg.Show Then
        Picture2.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command3_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SelectionParametersAvailable = cdSelectionParametersNone
    oDlg.SelectionParameter = cdParameterSaturation

    oDlg.Color = Picture3.BackColor
    If oDlg.Show Then
        Picture3.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command4_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.RecentColorsVisible = False

    oDlg.Color = Picture4.BackColor
    If oDlg.Show Then
        Picture4.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command5_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.ColorParametersVisible = False

    oDlg.Color = Picture5.BackColor
    If oDlg.Show Then
        Picture5.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command6_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SelectionParametersAvailable = cdSelectionParametersNone
    oDlg.SelectionParameter = cdParameterHue

    oDlg.Color = Picture6.BackColor
    If oDlg.Show Then
        Picture6.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command7_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetCompact

    oDlg.Color = Picture7.BackColor
    If oDlg.Show Then
        Picture7.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command8_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetComplete

    oDlg.Color = Picture8.BackColor
    If oDlg.Show Then
        Picture8.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command9_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.ColorSystem = cdColorSystemHSL

    oDlg.Color = Picture9.BackColor
    If oDlg.Show Then
        Picture9.BackColor = oDlg.Color
    End If
End Sub

Private Sub Command10_Click()
    Dim oDlg As New ColorDialog
    
    oDlg.SetComplete True
    
    oDlg.Color = Picture10.BackColor
    oDlg.DrawFixed = False
    If oDlg.Show Then
        Picture10.BackColor = oDlg.Color
    End If
End Sub


