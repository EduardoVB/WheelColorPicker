VERSION 5.00
Begin VB.Form frmWheelColorDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Color Picker"
   ClientHeight    =   6264
   ClientLeft      =   2676
   ClientTop       =   2160
   ClientWidth     =   5580
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
   ScaleHeight     =   6264
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrDoNotShowTT 
      Enabled         =   0   'False
      Interval        =   59000
      Left            =   960
      Top             =   5616
   End
   Begin VB.Timer tmrHideTT 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   576
      Top             =   5616
   End
   Begin VB.PictureBox picParameterLabel 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   170
      Left            =   3864
      ScaleHeight     =   168
      ScaleWidth      =   732
      TabIndex        =   24
      Top             =   3900
      Visible         =   0   'False
      Width           =   732
      Begin VB.Label lblParameter 
         AutoSize        =   -1  'True
         Caption         =   "Lum."
         Height          =   192
         Left            =   180
         TabIndex        =   25
         Top             =   0
         Width           =   348
      End
   End
   Begin VB.PictureBox picParametersContainer 
      BorderStyle     =   0  'None
      Height          =   1040
      Left            =   120
      ScaleHeight     =   1044
      ScaleWidth      =   3276
      TabIndex        =   11
      Top             =   4272
      Width           =   3276
      Begin VB.TextBox txtLum 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   17
         Top             =   720
         Width           =   450
      End
      Begin VB.TextBox txtSat 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   16
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtHue 
         Height          =   300
         Left            =   2184
         MaxLength       =   3
         TabIndex        =   15
         Top             =   0
         Width           =   450
      End
      Begin VB.TextBox txtBlue 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   14
         Top             =   720
         Width           =   450
      End
      Begin VB.TextBox txtGreen 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   13
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtRed 
         Height          =   300
         Left            =   624
         MaxLength       =   3
         TabIndex        =   12
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lblLum 
         Alignment       =   1  'Right Justify
         Caption         =   "Lum.:"
         Height          =   300
         Left            =   1536
         TabIndex        =   23
         Top             =   768
         Width           =   588
      End
      Begin VB.Label lblSat 
         Alignment       =   1  'Right Justify
         Caption         =   "Sat.:"
         Height          =   300
         Left            =   1536
         TabIndex        =   22
         Top             =   408
         Width           =   588
      End
      Begin VB.Label lblHue 
         Alignment       =   1  'Right Justify
         Caption         =   "Hue:"
         Height          =   300
         Left            =   1536
         TabIndex        =   21
         Top             =   48
         Width           =   588
      End
      Begin VB.Label lblBlue 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         Height          =   300
         Left            =   0
         TabIndex        =   20
         Top             =   768
         Width           =   588
      End
      Begin VB.Label lblGreen 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         Height          =   300
         Left            =   0
         TabIndex        =   19
         Top             =   408
         Width           =   588
      End
      Begin VB.Label lblRed 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         Height          =   300
         Left            =   0
         TabIndex        =   18
         Top             =   48
         Width           =   588
      End
   End
   Begin VB.PictureBox picRecentContainer 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   4728
      ScaleHeight     =   3900
      ScaleWidth      =   780
      TabIndex        =   8
      Top             =   60
      Width           =   780
      Begin VB.PictureBox picRecent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   324
         Index           =   0
         Left            =   168
         ScaleHeight     =   324
         ScaleWidth      =   444
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   444
      End
      Begin VB.Label lblRecent 
         Alignment       =   2  'Center
         Caption         =   "Recent:"
         Height          =   228
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   804
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3888
      TabIndex        =   1
      Top             =   5700
      Width           =   1284
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   372
      Left            =   2496
      TabIndex        =   0
      Top             =   5700
      Width           =   1284
   End
   Begin VB.PictureBox picSelection 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   3888
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4320
      Width           =   800
   End
   Begin VB.Timer tmrHexChange 
      Interval        =   3000
      Left            =   144
      Top             =   5616
   End
   Begin VB.TextBox txtHex 
      Height          =   300
      Left            =   744
      MaxLength       =   11
      TabIndex        =   4
      Top             =   5352
      Width           =   768
   End
   Begin ColorDialogTest.ColorWheel ColorWheel1 
      Height          =   3888
      Left            =   120
      TabIndex        =   2
      Top             =   144
      Width           =   4596
      _ExtentX        =   8107
      _ExtentY        =   6858
   End
   Begin VB.Label lblTT 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Hold the Control key down to navigate Saturation with the mouse wheel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   168
      TabIndex        =   26
      Top             =   5736
      Visible         =   0   'False
      Width           =   2172
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      Caption         =   "current"
      Height          =   228
      Left            =   3888
      TabIndex        =   7
      Top             =   5110
      Width           =   804
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      Caption         =   "new"
      Height          =   228
      Left            =   3888
      TabIndex        =   6
      Top             =   4104
      Width           =   804
   End
   Begin VB.Label lblHex 
      Alignment       =   1  'Right Justify
      Caption         =   "Hex:"
      Height          =   228
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   588
   End
   Begin VB.Menu mnuPopupRecent 
      Caption         =   "mnuPopupRecent"
      Visible         =   0   'False
      Begin VB.Menu mnuForgetRecent 
         Caption         =   "Forget"
      End
      Begin VB.Menu mnuClearAllRecent 
         Caption         =   "Clear recent colors"
      End
   End
End
Attribute VB_Name = "frmWheelColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private mCurrentColor As Long
Private mSelectedColor As Long
Private mSettingCurrent As Boolean
Private mOKPressed As Boolean
Private mIndexRecentToRemove As Long
Private mCurrentColorSet As Boolean
Private mContext As String
Private mHexValueVisible As Boolean
Private mColorParametersVisible As Boolean
Private mRecentColorsVisible As Boolean
Private mDialogSizeBig As Boolean
Private mSelectionParametersAvailable As CDSelectionParametersAvailable
Private mDrawFixedControlVisible As Boolean
Private mColorSystemControlVisible As Boolean
Private mDrawFixed As Boolean
Private mColorSystem As CDColorSystemConstants
Private mSelectionParameter As CDColorWheelParameterConstants
Private mSelectionDrawHorizontal As Boolean
Private mInvalidColorMessage As String
Private mCaptionColor As String
Private mCaptionColorSet As Boolean
Private mSettingParameters As Boolean
Private mNavigatedRadially As Boolean
Private mToolTipMouseWheelFirstPart As String
Private mToolTipMouseWheelLastPart As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveRecent
    mOKPressed = True
    Unload Me
End Sub

Private Sub ColorWheel1_MouseWheelNavigation(Axis As CDMouseWheelNavigation)
    If Axis = cdMouseWheelNavigatingAxial Then
        If Not mNavigatedRadially Then
            If Not tmrDoNotShowTT.Enabled Then
                If Not lblTT.Visible Then
                    SetMouseWheelTTText
                    lblTT.Visible = True
                    tmrHideTT.Enabled = True
                End If
            End If
        End If
    ElseIf Axis = cdMouseWheelNavigatingRadial Then
        mNavigatedRadially = True
    End If
End Sub

Private Sub SetMouseWheelTTText()
    Dim iCaptionID As CDColorWheelCaptionsIDConstants
    Dim iMove As Boolean
    
    If mToolTipMouseWheelFirstPart = "" Then
        mToolTipMouseWheelFirstPart = "Hold the Control key down to navigate"
    End If
    If mToolTipMouseWheelLastPart = "" Then
        mToolTipMouseWheelLastPart = "with the mouse wheel, Shift to go slowly"
    End If
    
    If ColorWheel1.RadialParameter = cdParameterLuminance Then
        If ColorWheel1.ColorSystem = cdColorSystemHSV Then
            iCaptionID = cdCWCaptionVal
        Else
            iCaptionID = cdCWCaptionLum
        End If
    Else
        iCaptionID = ColorWheel1.RadialParameter
    End If
    
    lblTT.Caption = Trim$(mToolTipMouseWheelFirstPart) & " " & ColorWheel1.GetCaption(iCaptionID) & " " & mToolTipMouseWheelLastPart
    
    lblTT.AutoSize = False
    lblTT.AutoSize = True
    If txtHex.Visible And (Not mRecentColorsVisible) Then
        lblTT.Width = cmdOK.Left - 150
    Else
        lblTT.Width = cmdOK.Left - 300
    End If
    lblTT.AutoSize = False
    lblTT.AutoSize = True
    If txtHex.Visible Then
        If (lblTT.Height + 120) > (Me.ScaleHeight - (txtHex.Top + txtHex.Height + 30)) Then
            lblTT.Width = cmdOK.Left - 150
            lblTT.FontSize = 6.5
            lblTT.AutoSize = False
            lblTT.AutoSize = True
            iMove = True
        End If
    ElseIf mSelectionDrawHorizontal Then
        If (lblTT.Height + 120) > (Me.ScaleHeight - (picSelection.Top + picSelection.Height + 30)) Then
            lblTT.Width = cmdOK.Left - 150
            lblTT.FontSize = 6.5
            lblTT.AutoSize = False
            lblTT.AutoSize = True
            iMove = True
        End If
    End If
    If iMove Then
        lblTT.Move 90, Me.ScaleHeight - lblTT.Height - 90
    Else
        lblTT.Move 120, Me.ScaleHeight - lblTT.Height - 120
    End If
End Sub

Private Sub ColorWheel1_ParameterValueChange()
    Dim iStr As String
    Dim iSS As Long
    Dim iSL As Long
        
    mSettingParameters = True
    mSelectedColor = ColorWheel1.Color
    
    iSS = txtHex.SelStart
    iSL = txtHex.SelLength
    
    txtHex.Text = LCase$(Hex(ColorWheel1.Color))
    
    On Error Resume Next
    txtHex.SelStart = iSS
    txtHex.SelLength = iSL
    On Error GoTo 0
    
    txtRed.Text = ColorWheel1.R
    txtGreen.Text = ColorWheel1.G
    txtBlue.Text = ColorWheel1.B
    txtHue.Text = Round(ColorWheel1.H)
    txtSat.Text = Round(ColorWheel1.S)
    txtLum.Text = Round(ColorWheel1.L)
    
    mSettingParameters = False
    
    If Not mSettingCurrent Then
        ShowSelection
    End If
End Sub

Private Sub ColorWheel1_ColorSystemChange()
    lblLum.Caption = EnsureEnding(IIf(ColorWheel1.ColorSystem = cdColorSystemHSV, ColorWheel1.GetCaption(cdCWCaptionVal), ColorWheel1.GetCaption(cdCWCaptionLum)), ":")
    
    If picParameterLabel.Visible Then
        If ColorWheel1.SelectionParameter = cdParameterLuminance Then
            lblParameter.Caption = lblLum.Caption
            PositionlblParameter
        End If
    End If
    lblTT.Visible = False
    tmrHideTT.Enabled = False
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub ColorWheel1_SelectionParameterChange()
    lblTT.Visible = False
    tmrHideTT.Enabled = False
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub Form_Load()
    Set Me.Icon = Nothing
    ColorWheel1.Redraw = False
    ColorWheel1.SelectionParametersAvailable = mSelectionParametersAvailable
    ColorWheel1.DrawFixedControlVisible = mDrawFixedControlVisible
    ColorWheel1.ColorSystemControlVisible = mColorSystemControlVisible
    ColorWheel1.DrawFixed = mDrawFixed
    ColorWheel1.ColorSystem = mColorSystem
    ColorWheel1.SelectionParameter = mSelectionParameter
    ColorWheel1_ColorSystemChange
    PositionControls
    ColorWheel1_ParameterValueChange
    mOKPressed = False
    LoadRecent
    PositionForm
    ColorWheel1.Redraw = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (mSelectionParametersAvailable <> cdSelectionParametersNone) Or mDrawFixedControlVisible Or mColorSystemControlVisible Then
        If mOKPressed Then
            SaveSetting RegKey, "Initialize", "LastColor" & IIf(mContext <> "", "_" & mContext, ""), mSelectedColor
        End If
        If (mSelectionParametersAvailable <> cdSelectionParametersNone) Then
            SaveSetting RegKey, "Initialize", "SelectionParameter" & IIf(mContext <> "", "_" & mContext, ""), CStr(ColorWheel1.SelectionParameter)
        End If
        If mDrawFixedControlVisible Then
            SaveSetting RegKey, "Initialize", "DrawFixed" & IIf(mContext <> "", "_" & mContext, ""), CStr(CLng(ColorWheel1.DrawFixed))
        End If
        If mColorSystemControlVisible Then
            SaveSetting RegKey, "Initialize", "ColorSystem" & IIf(mContext <> "", "_" & mContext, ""), CStr(CLng(ColorWheel1.ColorSystem))
        End If
    End If
End Sub

Private Sub mnuClearAllRecent_Click()
    ClearRecent
End Sub

Private Sub mnuForgetRecent_Click()
    picRecent(mIndexRecentToRemove).BackColor = vbWindowBackground
    picRecent(mIndexRecentToRemove).Tag = ""

    If GetSetting(RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(mIndexRecentToRemove + 1), "-") <> "-" Then
        DeleteSetting RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(mIndexRecentToRemove + 1)
    End If
End Sub

Private Sub picRecent_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picRecent(Index).Tag <> "" Then
            ColorWheel1.Color = Val(picRecent(Index).Tag)
        End If
    ElseIf Button = vbRightButton Then
        mnuForgetRecent.Visible = picRecent(Index).Tag <> ""
        mIndexRecentToRemove = Index
        PopupMenu mnuPopupRecent
    End If
End Sub

Private Sub picSelection_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        ColorWheel1.Color = picSelection.Point(X, Y)
    End If
End Sub

Private Sub tmrDoNotShowTT_Timer()
    tmrDoNotShowTT.Enabled = False
End Sub

Private Sub tmrHexChange_Timer()
    Dim iStr As String
    
    tmrHexChange.Enabled = False
    iStr = UCase(txtHex.Text)
    If Right$(iStr, 1) <> "&" Then
        iStr = iStr & "&"
    End If
    If Left$(iStr, 2) <> "&H" Then
        iStr = "&H" & iStr
    End If
    If IsValidOLE_COLOR(Val(iStr)) Then
        ColorWheel1.Color = Val(iStr)
    End If
End Sub

Private Sub tmrHideTT_Timer()
    tmrDoNotShowTT.Enabled = True
    tmrHideTT.Enabled = False
    lblTT.Visible = False
End Sub

Private Sub txtBlue_GotFocus()
    SelectTxtOnGotFocus txtBlue
End Sub

Private Sub txtHex_Change()
    tmrHexChange.Enabled = False
    tmrHexChange.Enabled = True
End Sub

Private Sub txtHex_GotFocus()
    If txtHex.SelStart = 0 Then
        txtHex.SelStart = Len(txtHex.Text)
    End If
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
    Dim iStr As String
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        iStr = UCase(txtHex.Text)
        If Right$(iStr, 1) <> "&" Then
            iStr = iStr & "&"
        End If
        If Left$(iStr, 2) <> "&H" Then
            iStr = "&H" & iStr
        End If
        If IsValidOLE_COLOR(Val(iStr)) Then
            ColorWheel1.Color = Val(iStr)
            If Hex(ColorWheel1.Color) <> Hex(Val(iStr)) Then
                If mInvalidColorMessage = "" Then
                    mInvalidColorMessage = "The Hex color is not valid."
                End If
                MsgBox mInvalidColorMessage, vbExclamation
                txtHex.Text = Hex(ColorWheel1.Color)
            End If
        Else
            If mInvalidColorMessage = "" Then
                mInvalidColorMessage = "The Hex color is not valid."
            End If
            MsgBox mInvalidColorMessage, vbExclamation
            txtHex.Text = Hex(ColorWheel1.Color)
        End If
        
        tmrHexChange_Timer
    End If
End Sub

Private Sub txtGreen_GotFocus()
    SelectTxtOnGotFocus txtGreen
End Sub

Private Sub txtHue_GotFocus()
    SelectTxtOnGotFocus txtHue
End Sub

Private Sub txtLum_GotFocus()
    SelectTxtOnGotFocus txtLum
End Sub

Private Sub txtRed_Change()
    Dim iVal As Long
    
    iVal = Val(txtRed.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorWheel1.R = iVal
End Sub

Private Sub txtGreen_Change()
    Dim iVal As Long
    
    iVal = Val(txtGreen.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorWheel1.G = iVal
End Sub

Private Sub txtBlue_Change()
    Dim iVal As Long
    
    iVal = Val(txtBlue.Text)
    If iVal < 0 Then iVal = 0
    If iVal > 255 Then iVal = 255
    ColorWheel1.B = iVal
End Sub

Private Sub txtHue_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtHue.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorWheel1.HMax Then iVal = ColorWheel1.HMax
    ColorWheel1.H = iVal
End Sub

Private Sub txtLum_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtLum.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorWheel1.LMax Then iVal = ColorWheel1.LMax
    ColorWheel1.L = iVal
End Sub

Private Sub txtRed_GotFocus()
    SelectTxtOnGotFocus txtRed
End Sub

Private Sub txtSat_Change()
    Dim iVal As Long
    
    If mSettingParameters Then Exit Sub
    
    iVal = Val(txtSat.Text)
    If iVal < 0 Then iVal = 0
    If iVal > ColorWheel1.SMax Then iVal = ColorWheel1.SMax
    ColorWheel1.S = iVal
End Sub

Private Function IsValidOLE_COLOR(nColor As Long) As Boolean
    Dim iLng As Long
    
    IsValidOLE_COLOR = True
    If nColor > &H100FFFF Then
        IsValidOLE_COLOR = False
    ElseIf nColor < 0 Then
        If (nColor And &HFF000000) = &H80000000 Then
            iLng = nColor And &HFFFF
            If iLng > 18 Then
                IsValidOLE_COLOR = False
            End If
        Else
            IsValidOLE_COLOR = False
        End If
    End If
End Function

Private Sub SelectTxtOnGotFocus(nTextBox As Control)
    If nTextBox.SelStart = 0 Then
        If nTextBox.SelLength = 0 Then
            nTextBox.SelLength = Len(nTextBox.Text)
        End If
    End If
End Sub

Private Sub txtSat_GotFocus()
    SelectTxtOnGotFocus txtSat
End Sub

Public Property Let CurrentColor(nColor As Long)
    Dim iStr As String
    Dim iLng As Long
    
    If Not IsValidOLE_COLOR(nColor) Then
        Err.Raise 1234, "ColorWheelDialog", "Invalid OLE color."
        Exit Property
    End If
    mSettingCurrent = True
    ColorWheel1.Redraw = False
    mCurrentColor = nColor
    ColorWheel1.Color = mCurrentColor
    mSettingCurrent = False
    mCurrentColorSet = True
    ShowSelection
    
    If (mSelectionParametersAvailable <> cdSelectionParametersNone) Or mDrawFixedControlVisible Or mColorSystemControlVisible Then
        iStr = GetSetting(RegKey, "Initialize", "LastColor" & IIf(mContext <> "", "_" & mContext, ""), "-")
        If iStr <> "-" Then
            If Val(iStr) = mCurrentColor Then
                If mDrawFixedControlVisible Then
                    ColorWheel1.DrawFixed = CBool(CLng(GetSetting(RegKey, "Initialize", "DrawFixed" & IIf(mContext <> "", "_" & mContext, ""), CStr(CLng(mDrawFixed)))))
                    mDrawFixed = ColorWheel1.DrawFixed
                End If
                If (mSelectionParametersAvailable <> cdSelectionParametersNone) Then
                    iLng = GetSetting(RegKey, "Initialize", "SelectionParameter" & IIf(mContext <> "", "_" & mContext, ""), CStr(mSelectionParameter))
                    If (iLng >= cdParameterHue) And (iLng <= cdParameterBlue) Then
                        ColorWheel1.SelectionParameter = iLng
                        mSelectionParameter = iLng
                    End If
                End If
                If mColorSystemControlVisible Then
                    iLng = GetSetting(RegKey, "Initialize", "ColorSystem" & IIf(mContext <> "", "_" & mContext, ""), CStr(mColorSystem))
                    If (iLng = cdColorSystemHSV) Or (iLng = cdColorSystemHSL) Then
                        ColorWheel1.ColorSystem = iLng
                        mColorSystem = iLng
                    End If
                End If
            End If
        End If
        SaveSetting RegKey, "Initialize", "LastColor" & IIf(mContext <> "", "_" & mContext, ""), mCurrentColor
    End If
    ColorWheel1.Redraw = True
End Property

Public Property Let Context(nContext As String)
    mContext = nContext
End Property

Private Sub ShowSelection()
    picSelection.Cls

    If mCurrentColorSet Then
        If mSelectionDrawHorizontal Then
            picSelection.Line (0, 0)-(picSelection.ScaleWidth / 2, picSelection.ScaleHeight), mCurrentColor, BF
            picSelection.Line (picSelection.ScaleWidth / 2, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mSelectedColor, BF
        Else
            picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight / 2), mSelectedColor, BF
            picSelection.Line (0, picSelection.ScaleHeight / 2)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mCurrentColor, BF
        End If
    Else
        If mCaptionColor = "" Then
            mCaptionColor = "Color:"
        End If
        lblNew.Caption = mCaptionColor
        mCaptionColorSet = True
        lblCurrent.Visible = False
        picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), mSelectedColor, BF
    End If
    picSelection.Line (0, 0)-(picSelection.ScaleWidth, picSelection.ScaleHeight), vbActiveBorder, B
End Sub

Private Sub LoadRecent()
    Dim c As Long
    Dim iStr As String
    Dim c2 As Long
    Dim iStep As Long
    
    iStep = IIf(mDialogSizeBig, 398, 405)
    For c = 1 To IIf(mDialogSizeBig, 11, 8)
        Load picRecent(c)
        picRecent(c).Top = picRecent(c - 1).Top + iStep
        picRecent(c).BackColor = vbWindowBackground
        picRecent(c).Visible = True
    Next c
    picRecentContainer.Height = picRecent(picRecent.UBound).Top + picRecent(picRecent.UBound).Height + 30
    
    For c = 1 To picRecent.Count
        iStr = GetSetting(RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(c), "-")
        If iStr <> "-" Then
            On Error Resume Next
            picRecent(c2).BackColor = Val(iStr)
            picRecent(c2).Tag = picRecent(c2).BackColor
            On Error GoTo 0
            c2 = c2 + 1
        End If
    Next c
End Sub

Private Sub SaveRecent()
    Dim c As Long
    Dim iList() As Long
    Dim c2 As Long
    
    ReDim iList(picRecent.Count)
    
    iList(0) = mSelectedColor
    For c = 1 To picRecent.Count
        iList(c) = -1
        If picRecent(c - 1).Tag <> "" Then
            iList(c) = Val(picRecent(c - 1).Tag)
        End If
    Next c
    For c = 1 To picRecent.Count
        If iList(c) <> -1 Then
            For c2 = 0 To c - 1
                If iList(c2) = iList(c) Then
                    iList(c) = -1
                    Exit For
                End If
            Next c2
        End If
    Next c
    c2 = 0
    For c = 0 To picRecent.Count
        If iList(c) <> -1 Then
            c2 = c2 + 1
            SaveSetting RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(c2), CStr(iList(c))
        End If
    Next c
End Sub

Private Sub ClearRecent()
    Dim c As Long
    
    For c = 1 To picRecent.UBound + 1
        If GetSetting(RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(c), "-") <> "-" Then
            DeleteSetting RegKey, "RecentColors" & IIf(mContext <> "", "_" & mContext, ""), CStr(c)
        End If
        picRecent(c - 1).BackColor = vbWindowBackground
        picRecent(c - 1).Tag = ""
    Next c
End Sub

Public Property Get OKPressed() As Boolean
    OKPressed = mOKPressed
End Property

Public Property Get SelectedColor() As Long
    SelectedColor = mSelectedColor
End Property


Public Property Let HexValueVisible(nValue As Boolean)
    mHexValueVisible = nValue
End Property

Public Property Let ColorParametersVisible(nValue As Boolean)
    mColorParametersVisible = nValue
End Property

Public Property Let RecentColorsVisible(nValue As Boolean)
    mRecentColorsVisible = nValue
End Property

Public Property Let DialogSizeBig(nValue As Boolean)
    mDialogSizeBig = nValue
End Property

Public Property Let SelectionParametersAvailable(nValue As CDSelectionParametersAvailable)
    mSelectionParametersAvailable = nValue
End Property

Public Property Let DrawFixedControlVisible(nValue As Boolean)
    mDrawFixedControlVisible = nValue
End Property

Public Property Let ColorSystemControlVisible(nValue As Boolean)
    mColorSystemControlVisible = nValue
End Property

Public Property Let DrawFixed(nValue As Boolean)
    mDrawFixed = nValue
End Property

Public Property Let ColorSystem(nValue As CDColorSystemConstants)
    mColorSystem = nValue
End Property

Public Property Let SelectionParameter(nValue As CDColorWheelParameterConstants)
    mSelectionParameter = nValue
End Property


Private Sub PositionControls()
    If mDialogSizeBig Then
        ColorWheel1.Height = 5000
        picParametersContainer.Top = ColorWheel1.Height + 420
        txtHex.Top = picParametersContainer.Top + picParametersContainer.Height + 40
        lblHex.Top = txtHex.Top + 50
    End If
    
    Me.Height = ColorWheel1.Height + 2500 + (Me.Height - Me.ScaleHeight)
    
    If Not mRecentColorsVisible Then
        picRecentContainer.Visible = False
        If mDialogSizeBig Then
            picSelection.Left = 3900
        Else
            picSelection.Left = ColorWheel1.Left + ColorWheel1.ParameterSelectorLeft + ColorWheel1.ParameterSelectorWidth - picSelection.Width
        End If
        Me.Width = ColorWheel1.Width + 240
    Else
        picRecentContainer.Left = ColorWheel1.Width + 170
        Me.Width = picRecentContainer.Left + picRecentContainer.Width + 240
        If (mSelectionParametersAvailable <> cdSelectionParametersNone) Then
            picSelection.Left = ColorWheel1.Left + ColorWheel1.SelectionParameterControlLeft + ColorWheel1.SelectionParameterControlWidth / 2 - ColorWheel1.SelectionParameterControlWidth / 2
            picSelection.Width = ColorWheel1.SelectionParameterControlWidth
        Else
            picSelection.Left = ColorWheel1.Left + ColorWheel1.ParameterSelectorLeft + ColorWheel1.ParameterSelectorWidth / 2 - picSelection.Width / 2
        End If
    End If
    lblNew.Left = picSelection.Left
    lblCurrent.Left = picSelection.Left
    If (Not mDialogSizeBig) And (Not mRecentColorsVisible) And mColorParametersVisible And mHexValueVisible Then
        cmdOK.Left = Me.ScaleWidth - 3030 + 120
    Else
        cmdOK.Left = Me.ScaleWidth - 3030
    End If
    If Not mColorParametersVisible Then
        picParametersContainer.Visible = False
        mHexValueVisible = False
        picSelection.Move ColorWheel1.Left + ColorWheel1.WheelCenterLeft - 1700 / 2, ColorWheel1.Height + 450, 1700, 400
        lblNew.Alignment = vbLeftJustify
        lblCurrent.Alignment = vbRightJustify
        lblCurrent.Move picSelection.Left - lblNew.Width - 60, picSelection.Top + picSelection.Height / 2 - lblNew.Height / 2
        lblNew.Move picSelection.Left + picSelection.Width + 60, lblCurrent.Top
        Me.Height = ColorWheel1.Height + 1750 + (Me.Height - Me.ScaleHeight)
        mSelectionDrawHorizontal = True
    Else
        picSelection.Top = Me.ScaleHeight - 1932
        lblNew.Top = picSelection.Top - 220
        lblCurrent.Top = picSelection.Top + picSelection.Height + 10
    End If
    If Not mHexValueVisible Then
        lblHex.Visible = False
        txtHex.Visible = False
    End If
    If (mSelectionParametersAvailable = cdSelectionParametersNone) Then
        lblParameter.Move 0, 0
        PositionlblParameter
        picParameterLabel.Visible = True
    End If
    
    cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 200
    cmdCancel.Move cmdOK.Left + 1400, cmdOK.Top

End Sub

Private Sub PositionlblParameter()
    If ColorWheel1.SelectionParameter = cdParameterLuminance Then
        If ColorWheel1.ColorSystem = cdColorSystemHSV Then
            lblParameter.Caption = WithoutEnding(ColorWheel1.GetCaption(cdCWCaptionVal), ":")
        Else
            lblParameter.Caption = WithoutEnding(ColorWheel1.GetCaption(cdCWCaptionLum), ":")
        End If
    Else
        lblParameter.Caption = WithoutEnding(ColorWheel1.GetCaption(ColorWheel1.SelectionParameter), ":")
    End If
    lblParameter.AutoSize = False
    lblParameter.AutoSize = True
    picParameterLabel.Move ColorWheel1.Left + ColorWheel1.ParameterSelectorLeft + ColorWheel1.ParameterSelectorWidth / 2 - lblParameter.Width / 2, ColorWheel1.Height + 60, lblParameter.Width
End Sub

Public Sub SetCaptions(nCaptionsArray)
    Dim c As CDCaptionsIDConstants
    
    For c = 0 To UBound(nCaptionsArray)
        If nCaptionsArray(c) <> "" Then
            If c = cdCaptionHue Then
                lblHue.Caption = EnsureEnding(nCaptionsArray(c), ":")
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionLum Then
                If mColorSystem = cdColorSystemHSL Then
                    lblLum.Caption = EnsureEnding(nCaptionsArray(c), ":")
                End If
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionSat Then
                lblSat.Caption = EnsureEnding(nCaptionsArray(c), ":")
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionVal Then
                If mColorSystem = cdColorSystemHSV Then
                    lblLum.Caption = EnsureEnding(nCaptionsArray(c), ":")
                End If
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionRed Then
                lblRed.Caption = EnsureEnding(nCaptionsArray(c), ":")
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionGreen Then
                lblGreen.Caption = EnsureEnding(nCaptionsArray(c), ":")
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionBlue Then
                lblBlue.Caption = EnsureEnding(nCaptionsArray(c), ":")
                ColorWheel1.SetCaption c, WithoutEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionFixed Then
                ColorWheel1.SetCaption c, CStr(nCaptionsArray(c))
            ElseIf c = cdCaptionFixedToolTipText Then
                ColorWheel1.SetCaption c, CStr(nCaptionsArray(c))
            ElseIf c = cdCaptionSelectionParameterToolTipText Then
                ColorWheel1.SetCaption c, CStr(nCaptionsArray(c))
            ElseIf c = cdCaptionCancel Then
                cmdCancel.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionCurrent Then
                lblCurrent.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionHex Then
                lblHex.Caption = EnsureEnding(nCaptionsArray(c), ":")
            ElseIf c = cdCaptionInvalidColorMessage Then
                mInvalidColorMessage = EnsureEnding(nCaptionsArray(c), ".")
            ElseIf c = cdCaptionNew Then
                lblNew.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionOK Then
                cmdOK.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionRecent Then
                lblRecent.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionColor Then
                mCaptionColor = EnsureEnding(nCaptionsArray(c), ":")
                If mCaptionColorSet Then
                    lblNew.Caption = mCaptionColor
                End If
            ElseIf c = cdCaptionMenuForgetRecent Then
                mnuForgetRecent.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionMenuClearAllRecent Then
                mnuClearAllRecent.Caption = nCaptionsArray(c)
            ElseIf c = cdCaptionMode Then
                ColorWheel1.SetCaption cdCWCaptionMode, CStr(nCaptionsArray(c))
            ElseIf c = cdCaptionToolTipMouseWheelBeginning Then
                mToolTipMouseWheelFirstPart = nCaptionsArray(c)
            ElseIf c = cdCaptionToolTipMouseWheelEnding Then
                mToolTipMouseWheelLastPart = nCaptionsArray(c)
            End If
        End If
    Next c
    PositionlblParameter
    
    AssignAccelerators
    
End Sub

Private Function EnsureEnding(nText As Variant, nEnding As String)
    EnsureEnding = nText
    If Right$(EnsureEnding, Len(nEnding)) <> nEnding Then
        EnsureEnding = EnsureEnding & nEnding
    End If
End Function

Private Function WithoutEnding(nText As Variant, nEnding As String)
    WithoutEnding = nText
    If Right$(WithoutEnding, Len(nEnding)) = nEnding Then
        WithoutEnding = Left$(WithoutEnding, Len(WithoutEnding) - 1)
    End If
End Function

Private Sub PositionForm()
    Dim iAFHwnd As Long
    Dim iRc As RECT
    Dim iPt As POINTAPI
    Dim iShift As Long
    
    iAFHwnd = GetActiveWindow
    If iAFHwnd <> 0 Then
        GetWindowRect iAFHwnd, iRc
        If iRc.Top < (Screen.Height / Screen.TwipsPerPixelY) And iRc.Left < (Screen.Width / Screen.TwipsPerPixelX) Then
            If (iRc.Top + 100 + Me.Height / Screen.TwipsPerPixelY) > (Screen.Height / Screen.TwipsPerPixelY - 100) Then
                 iRc.Top = (Screen.Height / Screen.TwipsPerPixelY - 100) - Me.Height / Screen.TwipsPerPixelY - 100
            End If
            If (iRc.Left + 150 + Me.Width / Screen.TwipsPerPixelX) > (Screen.Width / Screen.TwipsPerPixelX) Then
                iRc.Left = Screen.Width / Screen.TwipsPerPixelX - Me.Width / Screen.TwipsPerPixelX - 150
            End If
        End If
        Me.Move ScaleX(iRc.Left + 100, vbPixels, vbTwips), ScaleY(iRc.Top + 100, vbPixels, vbTwips)
    Else
        GetCursorPos iPt
        iPt.X = iPt.X - 15
        If iPt.X < 10 Then iPt.X = 10
        iPt.Y = iPt.Y + 20
        
        If iPt.Y < (Screen.Height / Screen.TwipsPerPixelY) And iPt.X < (Screen.Width / Screen.TwipsPerPixelX) Then
            If (iPt.Y + Me.Height / Screen.TwipsPerPixelY) > (Screen.Height / Screen.TwipsPerPixelY - 100) Then
                 iPt.Y = (Screen.Height / Screen.TwipsPerPixelY - 100) - Me.Height / Screen.TwipsPerPixelY
            End If
            If (iPt.X + 50 + Me.Width / Screen.TwipsPerPixelX) > (Screen.Width / Screen.TwipsPerPixelX) Then
                iPt.X = Screen.Width / Screen.TwipsPerPixelX - Me.Width / Screen.TwipsPerPixelX - 50
            End If
        End If
        Me.Move ScaleX(iPt.X, vbPixels, vbTwips), ScaleY(iPt.Y, vbPixels, vbTwips)
    End If
End Sub

Private Function AssignAcceleratorToCaption(nCaption As String, nUsedLetters As String, Optional nAssignRepeatedIfNeccesary As Boolean = True, Optional nLetterAssigned As String) As String
    Dim c As Long
    Dim iLU As String
    Dim iCap As String
    Dim iChar As String
    
    If HasLetters(nCaption) Then
        AssignAcceleratorToCaption = nCaption
        iCap = LCase$(nCaption)
        iLU = LCase$(nUsedLetters)
        For c = 1 To Len(iCap)
            iChar = Mid$(iCap, c, 1)
            If IsLetter(iChar) Then
                If InStr(iLU, iChar) = 0 Then
                    AssignAcceleratorToCaption = ""
                    If c > 1 Then
                        AssignAcceleratorToCaption = Left$(nCaption, c - 1)
                    End If
                    AssignAcceleratorToCaption = AssignAcceleratorToCaption & "&"
                    nLetterAssigned = UCase(Left$(Right$(nCaption, Len(nCaption) - c + 1), 1))
                    AssignAcceleratorToCaption = AssignAcceleratorToCaption & Right$(nCaption, Len(nCaption) - c + 1)
                    Exit Function
                End If
            End If
        Next c
        If nAssignRepeatedIfNeccesary Then
            AssignAcceleratorToCaption = AssignAcceleratorToCaption(nCaption, "")
        End If
    End If
End Function

Private Function HasLetters(nTexto As String) As Boolean
    Dim c As Long
    Dim iChar As String
    
    For c = 1 To Len(nTexto)
        iChar = Mid$(nTexto, c, 1)
        If IsLetter(iChar) Then
            HasLetters = True
            Exit Function
        End If
    Next c
End Function

Private Function IsLetter(nCharacter As String) As Boolean
    Dim iAsc As Long
    
    If nCharacter = "" Then Exit Function
    
    iAsc = Asc(UCase$(nCharacter))
    
    If (iAsc >= 65) And (iAsc <= 90) Then
        IsLetter = True
    End If
End Function

Private Sub AssignAccelerators()
    Dim iUsed As String
    Dim iStr As String
    
    cmdOK.Caption = AssignAcceleratorToCaption(cmdOK.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    cmdCancel.Caption = AssignAcceleratorToCaption(cmdCancel.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblRed.Caption = AssignAcceleratorToCaption(lblRed.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblGreen.Caption = AssignAcceleratorToCaption(lblGreen.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblBlue.Caption = AssignAcceleratorToCaption(lblBlue.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblLum.Caption = AssignAcceleratorToCaption(lblLum.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblHue.Caption = AssignAcceleratorToCaption(lblHue.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblSat.Caption = AssignAcceleratorToCaption(lblSat.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    lblHex.Caption = AssignAcceleratorToCaption(lblHex.Caption, iUsed, , iStr): iUsed = iUsed & iStr
    ColorWheel1.SetCaption cdCWCaptionFixed, AssignAcceleratorToCaption(ColorWheel1.GetCaption(cdCWCaptionFixed), iStr): iUsed = iUsed & iStr
    ColorWheel1.SetCaption cdCWCaptionMode, AssignAcceleratorToCaption(ColorWheel1.GetCaption(cdCWCaptionMode), iStr): iUsed = iUsed & iStr
End Sub

Private Function RegKey() As String
    RegKey = App.Title & "\ColorDialog"
End Function
