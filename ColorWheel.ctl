VERSION 5.00
Begin VB.UserControl ColorWheel 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   ToolboxBitmap   =   "ColorWheel.ctx":0000
   Begin VB.PictureBox picSlider 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2124
      Left            =   2904
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   0
      Top             =   120
      Width           =   132
   End
   Begin VB.Timer tmrFixSize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   984
      Top             =   2400
   End
   Begin VB.PictureBox picColorSystem 
      BorderStyle     =   0  'None
      Height          =   228
      Left            =   72
      ScaleHeight     =   228
      ScaleWidth      =   780
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2256
      Width           =   780
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode:"
         Height          =   192
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   432
      End
   End
   Begin VB.ComboBox cboColorSystem 
      Height          =   288
      ItemData        =   "ColorWheel.ctx":0312
      Left            =   72
      List            =   "ColorWheel.ctx":031C
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2496
      Width           =   700
   End
   Begin VB.Timer tmrDraw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1392
      Top             =   2400
   End
   Begin VB.ComboBox cboSelectionParameter 
      Height          =   288
      ItemData        =   "ColorWheel.ctx":032A
      Left            =   2328
      List            =   "ColorWheel.ctx":032C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select parameter"
      Top             =   2328
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CheckBox chkDrawFixed 
      Caption         =   "Fixed"
      Height          =   348
      Left            =   96
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Reflects color changes visually or not"
      Top             =   72
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picAux 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   720
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   -48
      Visible         =   0   'False
      Width           =   324
   End
   Begin VB.PictureBox picShades 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
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
      Height          =   2124
      Left            =   2496
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   144
      Width           =   324
   End
   Begin VB.PictureBox picWheel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
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
      Height          =   2124
      Left            =   144
      ScaleHeight     =   177
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   168
      Width           =   2244
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   0
         X1              =   10
         X2              =   24
         Y1              =   74
         Y2              =   74
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   1
         X1              =   50
         X2              =   64
         Y1              =   74
         Y2              =   74
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   3
         X1              =   36
         X2              =   36
         Y1              =   82
         Y2              =   96
      End
      Begin VB.Line linPointer 
         BorderWidth     =   3
         DrawMode        =   5  'Not Copy Pen
         Index           =   2
         X1              =   36
         X2              =   36
         Y1              =   54
         Y2              =   68
      End
   End
End
Attribute VB_Name = "ColorWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const cSUBCLASS_IN_IDE As Boolean = True

Public Event ColorChange()
Attribute ColorChange.VB_MemberFlags = "200"
Public Event ParameterValueChange()
Public Event SelectionParameterChange()
Public Event DrawFixedChange()
Public Event ColorSystemChange()
Public Event MouseWheelNavigation(Axis As CDMouseWheelNavigation)

Public Enum CDColorSystemConstants
    cdColorSystemHSV
    cdColorSystemHSL
End Enum

Public Enum CDColorWheelParameterConstants
    cdParameterHue = 0
    cdParameterLuminance = 1
    cdParameterValue = 1
    cdParameterSaturation = 2
    cdParameterRed = 3
    cdParameterGreen = 4
    cdParameterBlue = 5
End Enum

Public Enum CDColorWheelCaptionsIDConstants
    cdCWCaptionHue ' Hue
    cdCWCaptionLum ' Lum
    cdCWCaptionSat ' Sat
    cdCWCaptionRed ' Red
    cdCWCaptionGreen ' Green
    cdCWCaptionBlue ' Blue
    cdCWCaptionVal ' Val.
    cdCWCaptionFixed ' Fixed
    cdCWCaptionFixedToolTipText ' Reflects color changes visually or not
    cdCWCaptionSelectionParameterToolTipText ' Select parameter
    cdCWCaptionMode ' Mode
End Enum

Public Enum CDMouseWheelNavigation
    cdMouseWheelNavigatingSlider
    cdMouseWheelNavigatingAxial
    cdMouseWheelNavigatingRadial
End Enum

Public Enum CDSelectionParametersAvailable
    cdSelectionParametersNone
    cdSelectionParametersLumAndSat
    cdSelectionParametersHueLumAndSat
    cdSelectionParametersAll
End Enum

#Const ADD_ANTIALIASED_BORDER = False

'--- for MST subclassing (1)
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)

Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Private m_pSubclass         As IUnknown
'--- End for MST subclassing (1)
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Type RGBQuad
    R As Byte
    G As Byte
    B As Byte
    a As Byte
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Sub ColorRGBToHLS Lib "shlwapi" (ByVal clrRGB As Long, ByRef pwHue As Integer, ByRef pwLuminance As Integer, ByRef pwSaturation As Integer)
Private Declare Function ColorHLSToRGB Lib "shlwapi" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Private Const DIB_RGB_COLORS    As Long = 0

Private Const Pi = 3.14159265358979

Private Const cpicShadesWidth = 280

Private mCx As Long
Private mCy As Long
Private mDiameter As Long
Private mRadius As Long
Private mBMPiH As BITMAPINFOHEADER
Private mPixelsBytes() As Byte
Private mPixelsBytes2() As Byte
Private mBorderPixels() As Long
Private mBorderPixels_Alpha() As Byte
Private mPixelsAreInCircle() As Boolean
Private mPixelsAngle() As Double
Private mPixelsRadius() As Double
Private mBytesStride As Long
Private mBytesCount As Long
Private mBMPHeight As Long
Private mBMPWidth As Long
Private mPureColor As Long
Private mSettingColor As Boolean
Private mChangingShade As Boolean
Private mChangingHue As Boolean
Private mChangingLuminance As Boolean
Private mChangingSaturation As Boolean
Private mSelectingColor As Boolean
Private mClickingWheel As Boolean
Private mPointerX As Single
Private mPointerY As Single
Private mUserControlShown As Boolean
Private mChangingColorSystemOrInitializing As Boolean
Private mCaptionLum As String
Private mCaptionVal As String
Private mAmbientUserMode As Boolean
Private mDrawEnabled As Boolean
Private mSettingSlider As Boolean
Private mRedraw As Boolean
Private mDrawPending As Boolean
Private mWheelColorsStored As Boolean
Private mNewHeight As Long
Private mNewWidth As Long
Private mPropertiesAreSet As Boolean
Private mChangingParameter As Boolean
Private mRaiseEvents As Boolean
Private mInitialized As Boolean
Private mUserControlHwnd As Long
Private mRadialParameter As CDColorWheelParameterConstants ' this parameter increases its value from the center of the wheel to the periphery
Private mAxialParameter As CDColorWheelParameterConstants ' this parameter changes its value with the different angles to the center of the wheel
Private mParametersCaptions(5) As String

Private Const cDefaultColor As Long = &H808080
Private Const cDefaultSelectionParametersAvailable As Long = cdSelectionParametersLumAndSat
Private Const cDefaultDrawFixedControlVisible As Boolean = False
Private Const cDefaultColorSystemControlVisible As Boolean = False
Private Const cDefaultDrawFixed As Boolean = True
Private Const cDefaultSelectionParameter As Long = cdParameterLuminance
Private Const cDefaultColorSystem As Long = cdColorSystemHSV
Private Const cDefaultBackColor As Long = vbButtonFace

Private mColor As Long
Private mSelectionParametersAvailable As CDSelectionParametersAvailable
Private mDrawFixedControlVisible As Boolean
Private mColorSystemControlVisible As Boolean
Private mDrawFixed As Boolean
Private mSelectionParameter As CDColorWheelParameterConstants
Private mColorSystem As CDColorSystemConstants
Private mBackColor As Long

Private mH As Double
Private mL As Double
Private mS As Double
Private mR As Long
Private mG As Long
Private mB As Long

Private mH_Max As Long
Private mL_Max As Long
Private mS_Max As Long
Private mH_Fixed As Double
Private mL_Fixed As Double
Private mS_Fixed As Double

' Slider control
Private mSliderMin As Long
Private mSliderMax As Long
Private mSliderValue As Long
Private mGripLenght As Long
Private mGripWidth As Long

Private Sub cboColorSystem_Click()
    ColorSystem = cboColorSystem.ListIndex
End Sub

Private Sub cboSelectionParameter_Click()
    SelectionParameter = cboSelectionParameter.ItemData(cboSelectionParameter.ListIndex)
End Sub

Private Sub chkDrawFixed_Click()
    DrawFixed = (chkDrawFixed.Value = 1)
End Sub

Private Sub mForm_Load()
    mDrawEnabled = True
    StoreWheelColors
    DrawWheel
    DrawShades
    ShowSelectedColor
End Sub

Private Sub picShades_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRect As RECT
    Dim iPt As POINTAPI
    
    If Button = 1 Then
        GetClientRect picShades.hWnd, iRect
        iPt.X = iRect.Left
        iPt.Y = iRect.Top
        ClientToScreen picShades.hWnd, iPt
        OffsetRect iRect, iPt.X, iPt.Y
        ClipCursor iRect
        picShades_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub picShades_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SliderValue = mSliderMax / (picShades.ScaleHeight - 1) * (Y - 1)
    End If
End Sub

Private Sub picShades_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClipCursor ByVal 0&
End Sub

Private Sub picWheel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iRect As RECT
    Dim iPt As POINTAPI
    
    If Button = 1 Then
        If PixelIsInCircle(X, Y) Then
            GetClientRect picWheel.hWnd, iRect
            iPt.X = iRect.Left
            iPt.Y = iRect.Top
            ClientToScreen picWheel.hWnd, iPt
            OffsetRect iRect, iPt.X, iPt.Y
            ClipCursor iRect
            PointerVisible = False
            mSelectingColor = True
            picWheel_MouseMove Button, Shift, X, Y
        End If
    End If
End Sub

Private Sub SliderChange()
    If mSettingColor Then Exit Sub
    If mSettingSlider Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    
    mChangingShade = True
    
    If mSelectionParameter = cdParameterLuminance Then
        mChangingLuminance = True
    ElseIf mSelectionParameter = cdParameterHue Then
        mChangingHue = True
    ElseIf mSelectionParameter = cdParameterSaturation Then
        mChangingSaturation = True
    End If
    
    If mSelectionParameter = cdParameterLuminance Then
        mL = mSliderMax - SliderValue
    ElseIf mSelectionParameter = cdParameterHue Then
        mH = mSliderMax - SliderValue
        If mH = mH_Max Then mH = 0
    ElseIf mSelectionParameter = cdParameterSaturation Then
        mS = mSliderMax - SliderValue
    ElseIf mSelectionParameter = cdParameterRed Then
        mR = mSliderMax - SliderValue
    ElseIf mSelectionParameter = cdParameterGreen Then
        mG = mSliderMax - SliderValue
    ElseIf mSelectionParameter = cdParameterBlue Then
        mB = mSliderMax - SliderValue
    End If
    
    If (mSelectionParameter = cdParameterLuminance) Or (mSelectionParameter = cdParameterSaturation) Then
        If (Not mDrawFixed) Then
            tmrDraw.Enabled = True
        End If
    Else
        tmrDraw.Enabled = True
    End If
        
    If Not SetColor(GetShadedColor) Then
        RaiseEvent ParameterValueChange
    End If
    
    If mSelectionParameter = cdParameterLuminance Then
        mChangingLuminance = False
    ElseIf mSelectionParameter = cdParameterHue Then
        mChangingHue = False
    ElseIf mSelectionParameter = cdParameterSaturation Then
        mChangingSaturation = False
    End If
    
    mChangingShade = False
    
End Sub

Private Sub tmrDraw_Timer()
    tmrDraw.Enabled = False
    DrawWheel
End Sub

Private Sub tmrFixSize_Timer()
    tmrFixSize.Enabled = False
    UserControl.Size mNewWidth, mNewHeight
End Sub

Private Sub UserControl_Initialize()
    mRedraw = True
    mParametersCaptions(0) = "Hue"
    mCaptionLum = "Lum."
    mCaptionVal = "Value"
    mParametersCaptions(2) = "Sat."
    mParametersCaptions(3) = "Red"
    mParametersCaptions(4) = "Green"
    mParametersCaptions(5) = "Blue"
    mDrawEnabled = False
    mRadialParameter = cdParameterSaturation
    mAxialParameter = cdParameterHue
End Sub

Private Sub UserControl_InitProperties()
    mColor = cDefaultColor
    mSelectionParametersAvailable = cDefaultSelectionParametersAvailable
    mDrawFixedControlVisible = cDefaultDrawFixedControlVisible
    mColorSystemControlVisible = cDefaultColorSystemControlVisible
    mDrawFixed = cDefaultDrawFixed
    mSelectionParameter = cDefaultSelectionParameter
    LoadcboSelectionParameter
    If Not SelectInListByItemData(cboSelectionParameter, mSelectionParameter) Then cboSelectionParameter.ListIndex = 0
    mColorSystem = cDefaultColorSystem
    cboColorSystem.ListIndex = mColorSystem
    mBackColor = cDefaultBackColor
    SetBackColor
    
    On Error Resume Next
    mAmbientUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    If mAmbientUserMode Then
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    If mForm Is Nothing Then mDrawEnabled = True
    mPropertiesAreSet = True
    Init
    mRaiseEvents = True
    pvSubclass
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If picSlider.Visible Then
        If KeyCode = vbKeyUp Then
            If SliderValue > mSliderMin Then
                SliderValue = SliderValue - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            If SliderValue < mSliderMax Then
                SliderValue = SliderValue + 1
            End If
        End If
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mColor = PropBag.ReadProperty("Color", cDefaultColor)
    mSelectionParametersAvailable = PropBag.ReadProperty("SelectionParametersAvailable", cDefaultSelectionParametersAvailable)
    mDrawFixedControlVisible = PropBag.ReadProperty("DrawFixedControlVisible", cDefaultDrawFixedControlVisible)
    mColorSystemControlVisible = PropBag.ReadProperty("ColorSystemControlVisible", cDefaultColorSystemControlVisible)
    mDrawFixed = PropBag.ReadProperty("DrawFixed", cDefaultDrawFixed)
    mSelectionParameter = PropBag.ReadProperty("SelectionParameter", cDefaultSelectionParameter)
    LoadcboSelectionParameter
    If Not SelectInListByItemData(cboSelectionParameter, mSelectionParameter) Then cboSelectionParameter.ListIndex = 0
    mColorSystem = PropBag.ReadProperty("ColorSystem", cDefaultColorSystem)
    cboColorSystem.ListIndex = mColorSystem
    mBackColor = PropBag.ReadProperty("BackColor", cDefaultBackColor)
    SetBackColor
    On Error Resume Next
    mAmbientUserMode = Ambient.UserMode
    mUserControlHwnd = UserControl.hWnd
    If mAmbientUserMode Then
        If TypeOf Parent Is Form Then
            Set mForm = Parent
        End If
    End If
    If mForm Is Nothing Then mDrawEnabled = True
    mPropertiesAreSet = True
    Init
    mRaiseEvents = True
    pvSubclass
End Sub

Private Sub UserControl_Resize()
    Dim iAdditionalWidth As Long
    Dim ipicShadesWidth As Long
    Static sInside As Long
    Dim iNewHeight As Long
    Dim iNewWidth As Long
    
    sInside = sInside + 1
    
    If IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) >= 2400 Then
        ipicShadesWidth = cpicShadesWidth
        iAdditionalWidth = ipicShadesWidth + 150 + 75 + 75 + picSlider.Width
    Else
        ipicShadesWidth = cpicShadesWidth - 80
        iAdditionalWidth = ipicShadesWidth + 30 + 75 + 75 + picSlider.Width
    End If
    
    If (IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) + Screen.TwipsPerPixelX) < (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - iAdditionalWidth) Then
        iNewHeight = (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - iAdditionalWidth)
    End If
    
    If (IIf(iNewHeight <> 0, iNewHeight, UserControl.ScaleHeight) / Screen.TwipsPerPixelY) Mod 2 <> 0 Then
        iNewHeight = Round(IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) / Screen.TwipsPerPixelY / 2) * Screen.TwipsPerPixelY * 2
    End If
    
    If (mSelectionParametersAvailable <> cdSelectionParametersNone) Or mDrawFixedControlVisible Or mColorSystemControlVisible Then
        If (IIf(iNewHeight <> 0, iNewHeight, UserControl.ScaleHeight) + Screen.TwipsPerPixelY) < 3700 Then
            iNewHeight = 3700
        End If
    End If
    
    If IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) < iAdditionalWidth + 300 Then
        iNewHeight = iAdditionalWidth + 300
    End If
    If (IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) + Screen.TwipsPerPixelX) < (IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) + iAdditionalWidth) Then
        iNewWidth = (IIf(iNewHeight <> 0, iNewHeight, UserControl.Height) + iAdditionalWidth)
    End If
    If Abs(IIf(iNewWidth <> 0, iNewWidth, UserControl.Width) - (UserControl.Height + iAdditionalWidth)) > (Screen.TwipsPerPixelX * 1.3) Then
        iNewWidth = UserControl.Height + iAdditionalWidth
    End If
    If ((iNewHeight <> 0) Or (iNewWidth <> 0)) Then
        If iNewHeight = 0 Then iNewHeight = iNewWidth - iAdditionalWidth
        If iNewWidth = 0 Then iNewWidth = iNewHeight + iAdditionalWidth
        mNewWidth = iNewWidth
        mNewHeight = iNewHeight
        UserControl.Size iNewWidth, iNewHeight
        tmrFixSize.Enabled = True
    End If
    
    picWheel.Move 0, 0, UserControl.ScaleHeight, UserControl.ScaleHeight
    
    SetPicShades
    cboColorSystem.Move chkDrawFixed.Left, cboSelectionParameter.Top
    lblMode.Move 0, 0
    lblMode.AutoSize = False
    lblMode.AutoSize = True
    picColorSystem.Move chkDrawFixed.Left, cboSelectionParameter.Top - lblMode.Height - 45, lblMode.Width, lblMode.Height
    chkDrawFixed.Visible = mDrawFixedControlVisible
    picColorSystem.Visible = mColorSystemControlVisible
    cboColorSystem.Visible = mColorSystemControlVisible
    
    sInside = sInside - 1

    If (sInside = 0) And mPropertiesAreSet Then
        InitWheel
        DrawWheel
        DrawShades
        ShowSelectedColor
        picSlider.Width = (mGripWidth + 2) * Screen.TwipsPerPixelX
        DrawSliderGrip
    End If
End Sub

Private Sub SetPicShades()
    Dim ipicShadesHeight As Long
    Dim ipicShadesWidth As Long
    
    If UserControl.Height >= 2400 Then
        ipicShadesWidth = cpicShadesWidth
    Else
        ipicShadesWidth = cpicShadesWidth - 80
    End If
    If mSelectionParametersAvailable <> cdSelectionParametersNone Then
        ipicShadesHeight = UserControl.ScaleHeight - mGripLenght * Screen.TwipsPerPixelY - cboSelectionParameter.Height - 45
        chkDrawFixed.Move 90, 90
    Else
        ipicShadesHeight = UserControl.ScaleHeight - mGripLenght * Screen.TwipsPerPixelY
    End If
    picShades.Move picWheel.Width + IIf(UserControl.Height >= 2400, 150, 30), mGripLenght / 2 * Screen.TwipsPerPixelY, ipicShadesWidth, ipicShadesHeight
    picSlider.Move picShades.Left + picShades.Width + 25, 0, picSlider.Width, picShades.Height + (mGripLenght - 1) * Screen.TwipsPerPixelY
    cboSelectionParameter.Move picShades.Left + picShades.Width - cboSelectionParameter.Width, UserControl.ScaleHeight - cboSelectionParameter.Height
    cboSelectionParameter.Visible = mSelectionParametersAvailable <> cdSelectionParametersNone
End Sub

Private Sub Init()
    Dim iColor As Long
    Dim iRedrawPrev As Boolean
    
    TranslateColor mColor, 0, iColor
    
    iRedrawPrev = Redraw
    Redraw = False
    InitSlider
    InitWheel
    SetMaxAndFixedvalues
    SetSelectionParameter
    
    mR = iColor And 255
    mG = (iColor \ 256) And 255
    mB = (iColor \ 65536) And 255
    ColorRGBToCurrentColorSystem iColor, mH, mL, mS
    SetColor iColor
    
    ShowSelectedColor
    
    SetPicShades
    picColorSystem.Visible = mColorSystemControlVisible
    cboColorSystem.Visible = mColorSystemControlVisible
    chkDrawFixed.Visible = mDrawFixedControlVisible
    chkDrawFixed.Value = Abs(CLng(mDrawFixed))
    
    iColor = mColor
    mColor = -1
    mChangingColorSystemOrInitializing = True
    SetColor iColor
    mChangingColorSystemOrInitializing = False
    Redraw = iRedrawPrev
    mInitialized = True
End Sub

Private Sub InitWheel()
    Dim iBMP As BITMAP
    Dim iPic As StdPicture
    Dim c As Long
    Dim iUb As Long
    Dim iX As Long
    Dim iY As Long
    Dim iNear As Boolean
    Dim iX2 As Long
    Dim iY2 As Long
    Dim i As Long
    Dim iB1 As Long
    Dim iB2 As Long
    Dim iBackColor As Long
    Dim iBMPiH As BITMAPINFOHEADER
    Dim iPixelsBytes() As Byte
    Dim iPix_XY()  As POINTAPI
    Dim iPixToCheck() As Long
    Dim iUbp As Long
    Dim iIndexPixBorder As Long
    Dim iIndexPixToCheck As Long
    Dim iPixToCheckCount As Long
    Dim iCenterX As Long
    Dim iCenterY As Long
    Dim iRadius As Long
    Dim iDistanceToCircumference As Long
    
    If mDiameter = picWheel.ScaleWidth - 8 Then Exit Sub
    
    mCx = picWheel.ScaleWidth / 2
    mCy = picWheel.ScaleHeight / 2
    mDiameter = picWheel.ScaleWidth - 16
    mRadius = mDiameter / 2
    
    picAux.Move picAux.Left, picAux.Top, picWheel.Width, picWheel.Height
    picAux.FillStyle = vbFSSolid
    picAux.DrawWidth = 1
    picAux.BackColor = UserControl.BackColor
    iBackColor = picAux.BackColor
    TranslateColor iBackColor, 0&, iBackColor
    iB1 = (iBackColor \ 65536) And 255
    If iB1 = 255 Then
        iB2 = 200
    Else
        iB2 = 255
    End If
    picAux.FillColor = RGB(255, 255, iB2)
    picAux.Cls
    picAux.Circle (picAux.ScaleWidth / 2 - 1, picAux.ScaleHeight / 2 - 1), mRadius + 2, RGB(255, 255, iB2)
    Set iPic = picAux.Image
    picAux.Cls
    
    GetObject iPic.Handle, Len(iBMP), iBMP
    With mBMPiH
        .biSize = Len(mBMPiH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = iBMP.bmWidth
        .biHeight = iBMP.bmHeight
        .biSizeImage = ((.biWidth * 4 + 3) And &HFFFFFFFC) * .biHeight
    End With
    ReDim mPixelsBytes(Len(mBMPiH) + mBMPiH.biSizeImage)
    GetDIBits picAux.hDC, iPic.Handle, 0, iBMP.bmHeight, mPixelsBytes(0), mBMPiH, DIB_RGB_COLORS
    
    mBMPHeight = iBMP.bmHeight
    mBMPWidth = iBMP.bmWidth
    ReDim mPixelsAreInCircle(UBound(mPixelsBytes) / 4)
    ReDim mPixelsAngle(UBound(mPixelsAreInCircle))
    ReDim mPixelsRadius(UBound(mPixelsAreInCircle))
    ReDim mPixelsBytes2(UBound(mPixelsBytes))
    
    mBytesStride = mBMPWidth * 4
    mBytesCount = mBMPiH.biSizeImage - 1
    
    For c = 0 To mBytesCount - 4 Step 4
        If mPixelsBytes(c) = iB2 Then
            mPixelsAreInCircle(c / 4) = True
        End If
    Next c
    mWheelColorsStored = False
    
    ' Prepare the info to add an anti-alias border
    picAux.Cls
    picAux.FillStyle = vbFSTransparent
    picAux.DrawWidth = 5
    picAux.Circle (picAux.ScaleWidth / 2 - 1, picAux.ScaleHeight / 2 - 1), mRadius + 2, RGB(255, 255, iB2)
    Set iPic = picAux.Image
    picAux.Cls
    
    GetObject iPic.Handle, Len(iBMP), iBMP
    With iBMPiH
        .biSize = Len(mBMPiH)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = iBMP.bmWidth
        .biHeight = iBMP.bmHeight
        .biSizeImage = ((.biWidth * 4 + 3) And &HFFFFFFFC) * .biHeight
    End With
    ReDim iPixelsBytes(Len(iBMPiH) + iBMPiH.biSizeImage)
    GetDIBits picAux.hDC, iPic.Handle, 0, iBMP.bmHeight, iPixelsBytes(0), iBMPiH, DIB_RGB_COLORS
    
    iCenterX = mCx * 100 - 50
    iCenterY = mCy * 100 - 50
    iRadius = (mRadius + 2) * 100 + 55
    ' find one of the pixels that belong to the border
    iX = 0
    iY = 0
    i = (iY + 1) * mBMPHeight
    i = i + iX
    i = i * 4
    Do Until iPixelsBytes(i) = iB2
        iX = iX + 1
        iY = iY + 1
        i = (iY + 1) * mBMPHeight
        i = i + iX
        i = i * 4
        If i > mBytesCount Then Exit Sub
    Loop
    iUb = 1000
    ReDim mBorderPixels(iUb)
    ReDim mBorderPixels_Alpha(iUb)
    iIndexPixBorder = 0
    If (iX > 0) And (iY > 0) Then
        ReDim iPix_XY(mBytesCount)
        iUbp = 1000
        ReDim iPixToCheck(iUbp)
        iIndexPixToCheck = 1
        iPixToCheckCount = 1
        iPix_XY(i).X = iX
        iPix_XY(i).Y = iY
        iPixToCheck(iIndexPixToCheck) = i
        iDistanceToCircumference = Sqr((iCenterX - iX * 100 - 50) ^ 2 + (iCenterY - iY * 100 - 50) ^ 2) - iRadius
        iDistanceToCircumference = iDistanceToCircumference
        If iDistanceToCircumference < 0 Then iDistanceToCircumference = 0
        If iDistanceToCircumference > 255 Then iDistanceToCircumference = 255
        mBorderPixels_Alpha(iIndexPixBorder) = 255 - iDistanceToCircumference
        ' find the pixels of the border and save then in a vector (without scanning the whole picture)
        ' for each pixel, check all nearby whether they belong to the circunsference border or not
        Do While iIndexPixToCheck <= iPixToCheckCount
            iX = iPix_XY(iPixToCheck(iIndexPixToCheck)).X
            iY = iPix_XY(iPixToCheck(iIndexPixToCheck)).Y
            For iX2 = iX - 1 To iX + 1
                For iY2 = iY - 1 To iY + 1
                    i = (iY2 + 1) * mBMPHeight
                    i = i + iX2
                    i = i * 4
                    If (i >= 0) And (i <= mBytesCount) Then
                        If iPixelsBytes(i) = iB2 Then ' it is part of the border
                            If iPix_XY(i).X = 0 Then ' if it is not already added
                                If Not mPixelsAreInCircle(i / 4) Then ' if it is not inside the wheel
                                    iPixToCheckCount = iPixToCheckCount + 1
                                    If iPixToCheckCount > iUbp Then
                                        iUbp = iUbp + 1000
                                        ReDim Preserve iPixToCheck(iUbp)
                                    End If
                                    iPix_XY(i).X = iX2
                                    iPix_XY(i).Y = iY2
                                    iPixToCheck(iPixToCheckCount) = i
                                    iIndexPixBorder = iIndexPixBorder + 1
                                    If iIndexPixBorder > iUb Then
                                        iUb = iUb + 1000
                                        ReDim Preserve mBorderPixels(iUb)
                                        ReDim Preserve mBorderPixels_Alpha(iUb)
                                    End If
                                    mBorderPixels(iIndexPixBorder) = i
                                    iDistanceToCircumference = Sqr((iCenterX - iX2 * 100 - 50) ^ 2 + (iCenterY - iY2 * 100 - 50) ^ 2) - iRadius
                                    iDistanceToCircumference = iDistanceToCircumference * 1.5
                                    If iDistanceToCircumference < 0 Then iDistanceToCircumference = 0
                                    If iDistanceToCircumference > 255 Then iDistanceToCircumference = 255
                                    mBorderPixels_Alpha(iIndexPixBorder) = 255 - iDistanceToCircumference
                                End If
                            End If
                        End If
                    End If
                Next iY2
            Next iX2
            iIndexPixToCheck = iIndexPixToCheck + 1
        Loop
    End If
    ReDim Preserve mBorderPixels(iIndexPixBorder)
End Sub

Private Sub StoreWheelColors()
    Dim c As Long
    Dim iX As Long
    Dim iY As Long
    Dim iHorz As Long
    Dim iVert As Long
    Dim iAngle As Single
    Dim iRadius As Single
    Dim iColor As Long
    Dim iL1 As Long
    Dim i As Long
    Dim iP1 As Long
    Dim iP2 As Long
    Dim iRGB As RGBQuad
    
    If Not mDrawEnabled Then Exit Sub
    
    If mColorSystem = cdColorSystemHSV Then
        For c = 0 To mBytesCount - 4 Step 4
            iX = (c Mod mBytesStride) / 4
            iY = mBMPHeight - (c / mBMPHeight / 4 - 0.4999) - 1
            i = c / 4
            If mPixelsAreInCircle(i) Then
                iHorz = iX - mCx
                iVert = mCy - iY
                If iHorz = 0 Then
                    iAngle = 90 * Pi / 180 ' angle is hue
                Else
                    iAngle = Atn(iVert / iHorz)
                End If
                iAngle = 180 * iAngle / Pi
                
                If (iHorz >= 0) And (iVert >= 0) Then
                    ' ok
                ElseIf (iHorz < 0) And (iVert >= 0) Then
                    iAngle = 180 - iAngle * -1
                ElseIf (iHorz <= 0) And (iVert < 0) Then
                    iAngle = iAngle + 180
                Else
                    iAngle = iAngle + 360
                End If
                
                iRadius = Sqr(iHorz ^ 2 + iVert ^ 2) ' iRadius is saturation
                
                If mSelectionParameter = cdParameterLuminance Then
                    iAngle = iAngle / 360 * mH_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                    If iRadius > mS_Max Then iRadius = mS_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHSVToRGB(mPixelsAngle(i), mL_Fixed, mPixelsRadius(i))
                ElseIf mSelectionParameter = cdParameterHue Then
                    iAngle = iAngle / 360 * mL_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                    If iRadius > mS_Max Then iRadius = mS_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHSVToRGB(mH_Fixed, mPixelsAngle(i), mPixelsRadius(i))
                ElseIf mSelectionParameter = cdParameterSaturation Then
                    iAngle = iAngle / 360 * mH_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mL_Max * 2)
                    If iRadius > mL_Max Then iRadius = mL_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHSVToRGB(mPixelsAngle(i), mPixelsRadius(i), mS_Fixed)
                ElseIf mSelectionParameter = cdParameterRed Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(0, iP1, iP2)
                ElseIf mSelectionParameter = cdParameterGreen Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(iP1, 0, iP2)
                ElseIf mSelectionParameter = cdParameterBlue Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(iP1, iP2, 0)
                End If
                
                CopyMemory iRGB, iColor, 4
                
                mPixelsBytes(c + 2) = iRGB.R
                mPixelsBytes(c + 1) = iRGB.G
                mPixelsBytes(c) = iRGB.B
                
            Else
                mPixelsBytes2(c) = mPixelsBytes(c)
                mPixelsBytes2(c + 1) = mPixelsBytes(c + 1)
                mPixelsBytes2(c + 2) = mPixelsBytes(c + 2)
            End If
        Next c
    Else
        For c = 0 To mBytesCount - 4 Step 4
            iX = (c Mod mBytesStride) / 4
            iY = mBMPHeight - (c / mBMPHeight / 4 - 0.4999) - 1
            i = c / 4
            If mPixelsAreInCircle(i) Then
                iHorz = iX - mCx
                iVert = mCy - iY
                If iHorz = 0 Then
                    iAngle = 90 * Pi / 180 ' angle is hue
                Else
                    iAngle = Atn(iVert / iHorz)
                End If
                iAngle = 180 * iAngle / Pi
                
                If (iHorz >= 0) And (iVert >= 0) Then
                    ' ok
                ElseIf (iHorz < 0) And (iVert >= 0) Then
                    iAngle = 180 - iAngle * -1
                ElseIf (iHorz <= 0) And (iVert < 0) Then
                    iAngle = iAngle + 180
                Else
                    iAngle = iAngle + 360
                End If
                
                iRadius = Sqr(iHorz ^ 2 + iVert ^ 2) ' iRadius is saturation
                
                If mSelectionParameter = cdParameterLuminance Then
                    iAngle = iAngle / 360 * mH_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                    If iRadius > mS_Max Then iRadius = mS_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHLSToRGB(mPixelsAngle(i), mL_Fixed, mPixelsRadius(i))
                ElseIf mSelectionParameter = cdParameterHue Then
                    iAngle = iAngle / 360 * mL_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mS_Max * 2)
                    If iRadius > mS_Max Then iRadius = mS_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHLSToRGB(mH_Fixed, mPixelsAngle(i), mPixelsRadius(i))
                ElseIf mSelectionParameter = cdParameterSaturation Then
                    iAngle = iAngle / 360 * mH_Max
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * mL_Max * 2)
                    If iRadius > mL_Max Then iRadius = mL_Max
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iColor = ColorHLSToRGB(mPixelsAngle(i), mPixelsRadius(i), mS_Fixed)
                ElseIf mSelectionParameter = cdParameterRed Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(0, iP1, iP2)
                ElseIf mSelectionParameter = cdParameterGreen Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(iP1, 0, iP2)
                ElseIf mSelectionParameter = cdParameterBlue Then
                    iAngle = iAngle / 360 * 255
                    mPixelsAngle(i) = iAngle
                    iRadius = Int(iRadius / mDiameter * 255 * 2)
                    If iRadius > 255 Then iRadius = 255
                    If iRadius < 1 Then iRadius = 1
                    mPixelsRadius(i) = iRadius
                    iP1 = mPixelsAngle(i): If iP1 > 255 Then iP1 = 255
                    iP2 = mPixelsRadius(i): If iP2 > 255 Then iP2 = 255
                    iColor = RGB(iP1, iP2, 0)
                End If
                
                CopyMemory iRGB, iColor, 4
                
                mPixelsBytes(c + 2) = iRGB.R
                mPixelsBytes(c + 1) = iRGB.G
                mPixelsBytes(c) = iRGB.B
                
            Else
                mPixelsBytes2(c) = mPixelsBytes(c)
                mPixelsBytes2(c + 1) = mPixelsBytes(c + 1)
                mPixelsBytes2(c + 2) = mPixelsBytes(c + 2)
            End If
        Next c
    End If
    mWheelColorsStored = True
    
End Sub

Private Sub DrawWheel()
    Dim c As Long
    Dim iColor As Long
    Dim iP1 As Long
    Dim i As Long
    Dim iRGB As RGBQuad
    Dim iP1b As Double
    Dim iR1 As Long
    Dim iG1 As Long
    Dim iB1 As Long
    Dim iX As Long
    Dim iY As Long
    Dim iX2 As Long
    Dim iY2 As Long
    Dim i2 As Long
    Dim c2 As Single
    Dim iDo As Boolean
    Dim iBackColor As RGBQuad
    Dim iLng As Long
    Dim t As Long
    
    If Not mRedraw Then
        mDrawPending = True
        Exit Sub
    End If
    If Not mDrawEnabled Then Exit Sub
    
    tmrDraw.Enabled = False
    mDrawPending = False
    If mChangingColorSystemOrInitializing Then Exit Sub
    If Not mWheelColorsStored Then
        StoreWheelColors
    End If
    
    If mSelectionParameter = cdParameterLuminance Then
        If mColorSystem = cdColorSystemHSV Then
            If mDrawFixed Then
                iP1b = mL_Fixed
            Else
                iP1b = mL
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    iColor = ColorHSVToRGB(mPixelsAngle(i), iP1b, mPixelsRadius(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            If mDrawFixed Then
                iP1 = mL_Fixed
            Else
                iP1 = mL
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    iColor = ColorHLSToRGB(mPixelsAngle(i), iP1, mPixelsRadius(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSelectionParameter = cdParameterHue Then
        If mColorSystem = cdColorSystemHSV Then
            iP1b = mH
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    iColor = ColorHSVToRGB(iP1b, mPixelsAngle(i), mPixelsRadius(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            iP1 = mH
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    iColor = ColorHLSToRGB(iP1, mPixelsAngle(i), mPixelsRadius(i))
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSelectionParameter = cdParameterSaturation Then
        If mColorSystem = cdColorSystemHSV Then
            If mDrawFixed Then
                iP1b = mS_Fixed
            Else
                iP1b = mS
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    iColor = ColorHSVToRGB(mPixelsAngle(i), mPixelsRadius(i), iP1b)
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        Else
            If mDrawFixed Then
                iP1 = mS_Fixed
            Else
                iP1 = mS
            End If
            For c = 0 To mBytesCount - 4 Step 4
                i = c / 4
                If mPixelsAreInCircle(i) Then
                    If (Round(mPixelsAngle(i)) = 160) And (iP1 = 0) Then
                        iColor = 0
                    Else
                        iColor = ColorHLSToRGB(mPixelsAngle(i), mPixelsRadius(i), iP1)
                    End If
                    CopyMemory iRGB, iColor, 4
                    mPixelsBytes2(c + 2) = iRGB.R
                    mPixelsBytes2(c + 1) = iRGB.G
                    mPixelsBytes2(c) = iRGB.B
                End If
            Next c
        End If
    ElseIf mSelectionParameter = cdParameterRed Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInCircle(i) Then
                If mDrawFixed Then
                    iP1 = 0
                Else
                    iP1 = mR
                End If
                mPixelsBytes2(c + 2) = iP1 'R
                mPixelsBytes2(c + 1) = mPixelsAngle(i) 'G
                mPixelsBytes2(c) = mPixelsRadius(i) 'B
            End If
        Next c
    ElseIf mSelectionParameter = cdParameterGreen Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInCircle(i) Then
                If mDrawFixed Then
                    iP1 = 0
                Else
                    iP1 = mG
                End If
                mPixelsBytes2(c + 2) = mPixelsAngle(i) 'R
                mPixelsBytes2(c + 1) = iP1 'G
                mPixelsBytes2(c) = mPixelsRadius(i) 'B
            End If
        Next c
    ElseIf mSelectionParameter = cdParameterBlue Then
        For c = 0 To mBytesCount - 4 Step 4
            i = c / 4
            If mPixelsAreInCircle(i) Then
                If mDrawFixed Then
                    iP1 = 0
                Else
                    iP1 = mB
                End If
                mPixelsBytes2(c + 2) = mPixelsAngle(i) 'R
                mPixelsBytes2(c + 1) = mPixelsRadius(i) 'G
                mPixelsBytes2(c) = iP1 'B
            End If
        Next c
    End If
    
    ' Add a anti-aliased border
    Call OleTranslateColor(UserControl.BackColor, 0, iBackColor)
    
    For t = 1 To 3
        For c = 1 To UBound(mBorderPixels)
            i = mBorderPixels(c)
            iX = (i Mod mBytesStride) / 4
            iY = i / mBMPHeight / 4 - 0.4999 - 1
            iR1 = 0
            iG1 = 0
            iB1 = 0
            c2 = 0
            For iX2 = iX - 1 To iX + 1
                For iY2 = iY - 1 To iY + 1
                    i2 = (iY2 + 1) * mBMPHeight
                    i2 = i2 + iX2
                    i2 = i2 * 4
                    If (i2 > 0) And (i2 <= (mBytesCount - 4)) Then
                        If t = 1 Then
                            iDo = mPixelsAreInCircle(i2 / 4)
                        Else
                            iDo = True
                        End If
                        If iDo Then
                            If (iX2 = iX) Or (iY2 = iY) Then
                                c2 = c2 + 0.7
                                iR1 = iR1 + mPixelsBytes2(i2 + 2) * 0.7
                                iG1 = iG1 + mPixelsBytes2(i2 + 1) * 0.7
                                iB1 = iB1 + mPixelsBytes2(i2) * 0.7
                            Else
                                iR1 = iR1 + mPixelsBytes2(i2 + 2)
                                iG1 = iG1 + mPixelsBytes2(i2 + 1)
                                iB1 = iB1 + mPixelsBytes2(i2)
                                c2 = c2 + 1
                            End If
                        End If
                    End If
                Next
            Next
            
            If c2 > 0 Then
                iLng = iR1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.R) * (255 - mBorderPixels_Alpha(c)) / 255
                If iLng > 255 Then iLng = 255
                mPixelsBytes2(i + 2) = iLng
                iLng = iG1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.G) * (255 - mBorderPixels_Alpha(c)) / 255
                If iLng > 255 Then iLng = 255
                mPixelsBytes2(i + 1) = iLng
                iLng = iB1 / c2 * mBorderPixels_Alpha(c) / 255 + CLng(iBackColor.B) * (255 - mBorderPixels_Alpha(c)) / 255
                If iLng > 255 Then iLng = 255
                mPixelsBytes2(i) = iLng
            End If
        Next c
    Next t

    SetDIBitsToDevice picWheel.hDC, 0, 0, mBMPWidth, mBMPHeight, 0, 0, 0, mBMPHeight, mPixelsBytes2(0), mBMPiH, DIB_RGB_COLORS
    picWheel.Refresh
    SetPointer
End Sub

Private Function PixelIsInCircle(ByVal X As Single, ByVal Y As Single) As Boolean
    Dim i As Long
    
  '  Y = Y + 1
    
    X = Int(X)
    Y = Int(Y)
    If (X >= 0) And (X <= mBMPWidth) And (Y >= 0) And (Y <= mBMPHeight) Then
        i = (Y + 1) * mBMPHeight
        i = i + X
        If (i >= 0) And (i <= UBound(mPixelsAreInCircle)) Then
            PixelIsInCircle = mPixelsAreInCircle(i)
        End If
    End If
End Function

Private Sub UserControl_Terminate()
    pvUnsubclass
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Color", mColor, cDefaultColor
    PropBag.WriteProperty "SelectionParametersAvailable", mSelectionParametersAvailable, cDefaultSelectionParametersAvailable
    PropBag.WriteProperty "DrawFixedControlVisible", mDrawFixedControlVisible, cDefaultDrawFixedControlVisible
    PropBag.WriteProperty "ColorSystemControlVisible", mColorSystemControlVisible, cDefaultColorSystemControlVisible
    PropBag.WriteProperty "DrawFixed", mDrawFixed, cDefaultDrawFixed
    PropBag.WriteProperty "SelectionParameter", mSelectionParameter, cDefaultSelectionParameter
    PropBag.WriteProperty "ColorSystem", mColorSystem, cDefaultColorSystem
    PropBag.WriteProperty "BackColor", mBackColor, cDefaultBackColor
End Sub

Private Sub picWheel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSelectingColor Then
        If PixelIsInCircle(X, Y) Then
            mPureColor = GetWheelColor(X, Y)
            If mSelectionParameter = cdParameterLuminance Then
                ColorRGBToCurrentColorSystem mPureColor, mH, 0&, mS
            ElseIf mSelectionParameter = cdParameterHue Then
                ColorRGBToCurrentColorSystem mPureColor, 0&, mL, mS
            ElseIf mSelectionParameter = cdParameterSaturation Then
                ColorRGBToCurrentColorSystem mPureColor, mH, mL, 0&
            ElseIf mSelectionParameter = cdParameterRed Then
                mG = (mPureColor \ 256) And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSelectionParameter = cdParameterGreen Then
                mR = mPureColor And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSelectionParameter = cdParameterBlue Then
                mR = mPureColor And 255
                mG = (mPureColor \ 256) And 255
            End If
            mClickingWheel = True
            SetColor GetShadedColor
            mClickingWheel = False
            PointerVisible = False
        Else
            GetXYSameAngleInsideCircle X, Y
            mPureColor = GetWheelColor(X, Y)
            If mSelectionParameter = cdParameterLuminance Then
                ColorRGBToCurrentColorSystem mPureColor, mH, 0&, mS
            ElseIf mSelectionParameter = cdParameterHue Then
                ColorRGBToCurrentColorSystem mPureColor, 0&, mL, mS
            ElseIf mSelectionParameter = cdParameterSaturation Then
                ColorRGBToCurrentColorSystem mPureColor, mH, mL, 0&
            ElseIf mSelectionParameter = cdParameterRed Then
                mG = (mPureColor \ 256) And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSelectionParameter = cdParameterGreen Then
                mR = mPureColor And 255
                mB = (mPureColor \ 65536) And 255
            ElseIf mSelectionParameter = cdParameterBlue Then
                mR = mPureColor And 255
                mG = (mPureColor \ 256) And 255
            End If
            mClickingWheel = True
            SetColor GetShadedColor
            mClickingWheel = False
            SetPointer X, Y
            PointerVisible = True
        End If
    End If
End Sub

Private Sub picWheel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mSelectingColor Then
        picWheel_MouseMove Button, Shift, X, Y
        ClipCursor ByVal 0&
        mSelectingColor = False
        If PixelIsInCircle(X, Y) Then SetPointer X, Y
        PointerVisible = True
    End If
End Sub

Public Property Let Color(Value As OLE_COLOR)
    mChangingParameter = True
    SetColor Value
    mChangingParameter = False
End Property

Public Property Get Color() As OLE_COLOR
Attribute Color.VB_MemberFlags = "200"
    Color = mColor
End Property

Private Function SetColor(Value As Long) As Boolean
    Dim iPrev As Long
    Dim iColor As Long
    Dim iH1 As Double
    Dim iL1 As Double
    Dim iS1 As Double
    Dim iRGB As RGBQuad
    
    If Value = -1 Then Exit Function
    
    iPrev = mColor
    mColor = Value
    If (mColor <> iPrev) Or mChangingColorSystemOrInitializing Then
        mSettingColor = True
        TranslateColor mColor, 0, iColor
        CopyMemory iRGB, iColor, 4
        If (Not mSelectionParameter = cdParameterRed) Or mChangingParameter Then
            mR = iRGB.R
        End If
        If (Not mSelectionParameter = cdParameterGreen) Or mChangingParameter Then
            mG = iRGB.G
        End If
        If (Not mSelectionParameter = cdParameterBlue) Or mChangingParameter Then
            mB = iRGB.B
        End If
        ColorRGBToCurrentColorSystem iColor, iH1, iL1, iS1
        If (ColorCurrentColorSystemToRGB(iH1, iL1, iS1) <> iPrev) Or mChangingColorSystemOrInitializing Then
            If Not (mChangingHue Or mChangingLuminance Or mChangingSaturation) Then
                If (Not mSelectionParameter = cdParameterSaturation) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    mS = iS1
                End If
                'If Not ((iH1 = 160) And (mS = 0)) Then
                If Not (mSelectionParameter = cdParameterHue) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    mH = iH1
                End If
                If mH = mH_Max Then mH = 0
                If Not (mSelectionParameter = cdParameterLuminance) Or mChangingColorSystemOrInitializing Or mChangingParameter Then
                    mL = iL1
                End If
                'End If
            End If
            If (Not mChangingShade) And (Not mClickingWheel) Then
                If mSelectionParameter = cdParameterLuminance Then
                    SliderValue = mL_Max - mL
                    mPureColor = ColorCurrentColorSystemToRGB(mH, mL_Fixed, mS)
                ElseIf mSelectionParameter = cdParameterHue Then
                    SliderValue = mH_Max - mH
                    mPureColor = ColorCurrentColorSystemToRGB(mH_Fixed, mL, mS)
                ElseIf mSelectionParameter = cdParameterSaturation Then
                    SliderValue = mS_Max - mS
                    mPureColor = ColorCurrentColorSystemToRGB(mH, mL, mS_Fixed)
                Else
                    If mSelectionParameter = cdParameterRed Then
                        SliderValue = 255 - mR
                        mPureColor = RGB(0, mG, mB)
                    ElseIf mSelectionParameter = cdParameterGreen Then
                        SliderValue = 255 - mG
                        mPureColor = RGB(mR, 0, mB)
                    ElseIf mSelectionParameter = cdParameterBlue Then
                        SliderValue = 255 - mB
                        mPureColor = RGB(mR, mG, 0)
                    End If
                End If
            End If
            DrawShades
            ShowSelectedColor
        End If
        If Not tmrDraw.Enabled Then
            If Not mClickingWheel Then
                If (mSelectionParameter = cdParameterLuminance) Or (mSelectionParameter = cdParameterSaturation) Then
                    If (Not mDrawFixed) Then
                        DrawWheel
                    End If
                Else
                    DrawWheel
            End If
            End If
        End If
        mSettingColor = False
        If mRaiseEvents Then RaiseEvent ColorChange
        If mRaiseEvents Then RaiseEvent ParameterValueChange
        If mInitialized Then PropertyChanged "Color"
    End If
End Function

Private Property Get SliderValue() As Long
    SliderValue = mSliderValue
End Property

Private Property Let SliderValue(ByVal nValue As Long)
    If mSliderValue <> nValue Then
        If nValue > mSliderMax Then nValue = mSliderMax
        If nValue < mSliderMin Then nValue = mSliderMin
        If mSliderValue <> nValue Then
            mSliderValue = nValue
            DrawSliderGrip
            SliderChange
        End If
    End If
End Property

Public Property Get H() As Integer
    H = mH
End Property

Public Property Let H(Value As Integer)
    If Value <> mH Then
        If (Value < 0) Or (Value > mH_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mH = Value
        If mH = mH_Max Then mH = 0
        mChangingHue = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSelectionParameter = cdParameterHue Then
                SliderValue = mH_Max - mH
            End If
            DrawWheel
            DrawShades
            ShowSelectedColor
            RaiseEvent ParameterValueChange
        End If
        mChangingParameter = False
        mChangingHue = False
    End If
End Property


Public Property Get L() As Integer
    L = mL
End Property

Public Property Let L(Value As Integer)
    If Value <> mL Then
        If (Value < 0) Or (Value > mL_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mL = Value
        mChangingLuminance = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSelectionParameter = cdParameterLuminance Then
                SliderValue = mL_Max - mL
            End If
            DrawWheel
            DrawShades
            ShowSelectedColor
            RaiseEvent ParameterValueChange
        End If
        mChangingParameter = False
        mChangingLuminance = False
    End If
End Property


Public Property Get V() As Integer
    V = mL
End Property

Public Property Let V(Value As Integer)
    L = Value
End Property


Public Property Get S() As Integer
    S = mS
End Property

Public Property Let S(Value As Integer)
    If Value <> mS Then
        If (Value < 0) Or (Value > mS_Max) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mS = Value
        mChangingSaturation = True
        mChangingParameter = True
        If Not SetColor(ColorCurrentColorSystemToRGB(mH, mL, mS)) Then
            If mSelectionParameter = cdParameterSaturation Then
                SliderValue = mS_Max - mS
            End If
            DrawWheel
            DrawShades
            ShowSelectedColor
            RaiseEvent ParameterValueChange
        End If
        mChangingParameter = False
        mChangingSaturation = False
    End If
End Property


Public Property Get R() As Integer
    R = mR
End Property

Public Property Let R(Value As Integer)
    If Value <> mR Then
        If (Value < 0) Or (Value > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mR = Value
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property


Public Property Get G() As Integer
    G = mG
End Property

Public Property Let G(Value As Integer)
    If Value <> mG Then
        If (Value < 0) Or (Value > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mG = Value
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property


Public Property Get B() As Integer
    B = mB
End Property

Public Property Let B(Value As Integer)
    If Value <> mB Then
        If (Value < 0) Or (Value > 255) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        mB = Value
        mChangingParameter = True
        SetColor RGB(mR, mG, mB)
        mChangingParameter = False
        DrawShades
    End If
End Property

Private Sub DrawShades()
    Dim iY As Long
    Dim iColor As Long
    Dim iH1 As Double
    Dim iS1 As Double
    Dim iL1 As Double
    Dim iHeight As Long
    Dim iWidth As Long
    
    If mDrawPending Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    If mChangingColorSystemOrInitializing Then Exit Sub
    
    picShades.Cls
    iHeight = picShades.ScaleHeight - 2
    iWidth = picShades.ScaleWidth - 1
    ColorRGBToCurrentColorSystem mPureColor, iH1, iL1, iS1
    
    If mSelectionParameter = cdParameterLuminance Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - 1
                iColor = ColorHSVToRGB(iH1, iY / iHeight * mL_Max, iS1)
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - 1
                iColor = ColorHLSToRGB(iH1, iY / iHeight * mL_Max, iS1)
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSelectionParameter = cdParameterHue Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - 1
                If mDrawFixed Then
                    iColor = ColorHSVToRGB(iY / iHeight * mH_Max, CDbl(mL_Fixed), CDbl(mS_Fixed))
                Else
                    iColor = ColorHSVToRGB(iY / iHeight * mH_Max, iL1, iS1)
                End If
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - 1
                If mDrawFixed Then
                    iColor = ColorHLSToRGB(iY / iHeight * mH_Max, mL_Fixed, mS_Fixed)
                Else
                    If Round(iY / iHeight * mH_Max) = 160 And (iS1 = 0) Then
                        iColor = 0
                    Else
                        iColor = ColorHLSToRGB(iY / iHeight * mH_Max, iL1, iS1) 'aca
                    End If
                End If
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSelectionParameter = cdParameterSaturation Then
        If mColorSystem = cdColorSystemHSV Then
            For iY = 0 To iHeight - 1
                iColor = ColorHSVToRGB(iH1, iL1, iY / iHeight * mS_Max)
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        Else
            For iY = 0 To iHeight - 1
                iColor = ColorHLSToRGB(iH1, iL1, iY / iHeight * mS_Max)
                picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
            Next iY
        End If
    ElseIf mSelectionParameter = cdParameterRed Then
        For iY = 0 To iHeight - 1
            If mDrawFixed Then
                iColor = RGB(iY / iHeight * 255, mG, mB)
            Else
                iColor = RGB(iY / iHeight * 255, 0, 0)
            End If
            picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    ElseIf mSelectionParameter = cdParameterGreen Then
        For iY = 0 To iHeight - 1
            If mDrawFixed Then
                iColor = RGB(mR, iY / iHeight * 255, mB)
            Else
                iColor = RGB(0, iY / iHeight * 255, 0)
            End If
            picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    ElseIf mSelectionParameter = cdParameterBlue Then
        For iY = 0 To iHeight - 1
            If mDrawFixed Then
                iColor = RGB(mR, mG, iY / iHeight * 255)
            Else
                iColor = RGB(0, 0, iY / iHeight * 255)
            End If
            picShades.Line (1, iHeight - iY)-(iWidth, iHeight - iY), iColor
        Next iY
    End If
    picShades.Refresh
End Sub

Private Function GetShadedColor()
    If (mSelectionParameter = cdParameterLuminance) Or (mSelectionParameter = cdParameterHue) Or (mSelectionParameter = cdParameterSaturation) Then
        GetShadedColor = ColorCurrentColorSystemToRGB(mH, mL, mS)
    Else
        GetShadedColor = RGB(mR, mG, mB)
    End If
End Function

Private Function GetWheelColor(ByVal X As Single, ByVal Y As Single) As Long
    Dim i As Long
    
    'Y = Y + 1
    
    X = Int(X)
    Y = Int(Y)
    i = (mBMPHeight - Y - 1) * mBMPHeight
    i = i + X
    i = i * 4
    
    GetWheelColor = RGB(mPixelsBytes(i + 2), mPixelsBytes(i + 1), mPixelsBytes(i))
End Function

Private Sub ShowSelectedColor()
    Dim iX As Single
    Dim iY As Single
    Dim iX2 As Long
    Dim iY2 As Long
    Dim iFound As Boolean
    Dim iR1 As Long
    Dim iG1 As Long
    Dim iB1 As Long
    Dim iColor As Long
    Dim iTolerance  As Long
    Dim c As Long
    Dim iP1 As Long
    Dim iP2 As Long
    Dim iP1Max As Long
    Dim iP2Max As Long
    Dim iRGB As RGBQuad
    Dim iListOfPossible_X() As Long
    Dim iListOfPossible_Y() As Long
    Dim iUb As Long
    Dim iCount As Long
    Dim iNearest_Index As Long
    Dim iNearest_Distance As Single
    Dim iDistance As Single
    
    If mDrawPending Then Exit Sub
    If Not mDrawEnabled Then Exit Sub
    If mClickingWheel Or mChangingShade Then Exit Sub
    
    If mH_Max = 0 Then SetMaxAndFixedvalues
    If mSelectionParameter = cdParameterLuminance Then
        iP1 = mS ' iP1 is radius
        iP2 = mH ' iP2 is angle
        iP1Max = mS_Max
        iP2Max = mH_Max
    ElseIf mSelectionParameter = cdParameterHue Then
        iP1 = mS
        iP2 = mL
        iP1Max = mS_Max
        iP2Max = mL_Max
    ElseIf mSelectionParameter = cdParameterSaturation Then
        iP1 = mL
        iP2 = mH
        iP1Max = mL_Max
        iP2Max = mH_Max
    ElseIf mSelectionParameter = cdParameterRed Then
        iP1 = mB
        iP2 = mG
        iP1Max = 255
        iP2Max = 255
    ElseIf mSelectionParameter = cdParameterGreen Then
        iP1 = mB
        iP2 = mR
        iP1Max = 255
        iP2Max = 255
    ElseIf mSelectionParameter = cdParameterBlue Then
        iP1 = mG
        iP2 = mR
        iP1Max = 255
        iP2Max = 255
    End If
    
    iX = (iP1 * Cos(Pi / 180 * (1 - iP2 / iP2Max) * 360)) / iP1Max * mRadius + mCx
    iY = (iP1 * Sin(Pi / 180 * (1 - iP2 / iP2Max) * 360)) / iP1Max * mRadius + mCy
    If mColor = cDefaultColor Then
        iX = mCx
        iY = mCy
        SetPointer iX, iY
        Exit Sub
    End If
    iX = Int(iX)
    iY = Int(iY)
    If Not PixelIsInCircle(iX, iY) Then
        iX = iX + 1
        iY = iY + 1
        If Not PixelIsInCircle(iX, iY) Then
            iX = iX + 1
            iY = iY + 1
        End If
    End If
    If Not PixelIsInCircle(iX, iY) Then
        iX = (iP1 * Cos(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.99 / 2 + mCx
        iY = (iP1 * Sin(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.99 / 2 + mCy
        If Not PixelIsInCircle(iX, iY) Then
            iX = (iP1 * Cos(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.98 / 2 + mCx
            iY = (iP1 * Sin(Pi * 2 / 360 * (360 - iP2 / iP2Max * 360))) / iP1Max * mDiameter * 0.98 / 2 + mCy
        End If
    End If
    If PixelIsInCircle(iX, iY) Then
'        For iX2 = -10 To 10
'            For iY2 = -10 To 10
'                If PixelIsInCircle(iX + iX2, iY + iY2) Then
'                    If GetWheelColor(iX, iY) = mPureColor Then
'                        iFound = True
'                        Exit For
'                    End If
'                End If
'            Next iY2
'            If iFound Then Exit For
'        Next iX2
'        If Not iFound Then
        For iTolerance = 0 To 10
            iUb = 100
            ReDim iListOfPossible_X(iUb)
            ReDim iListOfPossible_Y(iUb)
            iCount = -1
            For iX2 = -10 To 10
                For iY2 = -10 To 10
                    If PixelIsInCircle(iX + iX2, iY + iY2) Then
                        iColor = GetWheelColor(iX + iX2, iY + iY2)
                        CopyMemory iRGB, iColor, 4
                        If (Abs(mR - iRGB.R) + Abs(mG - iRGB.G) + Abs(mB - iRGB.B)) <= iTolerance Then
                            iFound = True
                            iCount = iCount + 1
                            If iCount > iUb Then
                                iUb = iUb + 100
                                ReDim Preserve iListOfPossible_X(iUb)
                                ReDim Preserve iListOfPossible_Y(iUb)
                            End If
                            iListOfPossible_X(iCount) = iX2
                            iListOfPossible_Y(iCount) = iY2
                            'Exit For
                        End If
                    End If
                Next iY2
                'If iFound Then Exit For
            Next iX2
            If iFound Then
                iNearest_Distance = mRadius
                For c = 0 To iCount
                    iDistance = Sqr(iListOfPossible_X(c) ^ 2 + iListOfPossible_Y(c) ^ 2)
                    If iDistance < iNearest_Distance Then
                        iNearest_Distance = iDistance
                        iNearest_Index = c
                    End If
                Next c
                iX2 = iListOfPossible_X(iNearest_Index)
                iY2 = iListOfPossible_Y(iNearest_Index)
                Exit For
            End If
        Next iTolerance
'        End If
        
        If iFound Then
            iX = iX + iX2
            iY = iY + iY2
        End If
        
        SetPointer iX, iY
    End If
End Sub

Private Sub GetXYSameAngleInsideCircle(X As Single, Y As Single)
    Dim H As Single
    Dim a As Single
    Dim B As Single
    Dim iAngle As Single
    Dim iSin As Single
    Dim iHorz As Single
    Dim iVert As Single
    Dim iX2 As Single
    Dim iY2 As Single
    Dim R As Single
    Dim iRadius As Single
    
    iHorz = X - mCx
    iVert = Y - mCy
    If iHorz = 0 Then
        
        If Y < mCy Then
            iAngle = 270 * Pi / 180 ' angle is hue
        Else
            iAngle = 90 * Pi / 180
        End If
    Else
        iAngle = Atn(iVert / iHorz)
    End If
    
    iRadius = mRadius * 1.02
    iX2 = Cos(iAngle) * iRadius + mCx
    iY2 = Sin(iAngle) * iRadius + mCy
    If iHorz < 0 Then
        iX2 = mCx - (iX2 - mCx)
        iY2 = mCy - (iY2 - mCy)
    End If
    
    Do Until PixelIsInCircle(iX2, iY2)
        iRadius = iRadius * 0.9999
        iX2 = Cos(iAngle) * iRadius + CSng(mCx)
        iY2 = Sin(iAngle) * iRadius + CSng(mCy)
        If iHorz < 0 Then
            iX2 = mCx - (iX2 - mCx)
            iY2 = mCy - (iY2 - mCy)
        End If
    Loop
    
    X = iX2
    Y = iY2
End Sub

Public Property Get ColorHex() As String
    ColorHex = Hex(mColor)
    If Len(ColorHex) < 6 Then ColorHex = String$(6 - Len(ColorHex), "0") & ColorHex
    ColorHex = "&H" & ColorHex & "&"
End Property

Private Function DistanceFromCenter(X As Single, Y As Single) As Single
    DistanceFromCenter = Sqr((X - mCx) ^ 2 + (Y - mCy) ^ 2)
End Function

Private Property Let PointerVisible(ByVal nValue As Boolean)
    linPointer(0).Visible = nValue
    linPointer(1).Visible = nValue
    linPointer(2).Visible = nValue
    linPointer(3).Visible = nValue
End Property

Private Sub SetPointer(Optional X As Single = -1, Optional Y As Single)
    Dim c As Long
    Dim iPointerColor As Long
    Dim iX2 As Single
    Dim iY2 As Single
    Dim iDrawMode As Long
    Dim iColorBrightness As Long
    
    If X <> -1 Then
        mPointerX = X
        mPointerY = Y
    End If
    
    iX2 = mPointerX
    iY2 = mPointerY
'    If iX2 < mCx Then
'        iX2 = iX2 + 1
'    Else
'        iX2 = iX2 - 1
'    End If
'    If iY2 < mCy Then
'        iY2 = iY2 + 1
'    Else
'        iY2 = iY2 - 1
'    End If
    
    iColorBrightness = GetColorBrightness(mColor)
    If iColorBrightness > 110 Then
        If (iColorBrightness > 200) Then
            iPointerColor = vbWhite
            iDrawMode = vbMaskPenNot
        Else
            iPointerColor = vbBlack
            iDrawMode = vbCopyPen
        End If
    Else
        iPointerColor = vbWhite
        If mRadius - DistanceFromCenter(iX2, iY2) < ((linPointer(0).X2 - linPointer(0).X1) * 1.5) Then
            iDrawMode = vbMaskPenNot
        Else
            If (iColorBrightness < 50) Then
                iDrawMode = vbMaskPenNot
            Else
                iDrawMode = vbCopyPen
            End If
        End If
    End If
    
    For c = 0 To 3
        linPointer(c).BorderColor = iPointerColor
        linPointer(c).DrawMode = iDrawMode
    Next c
    
    linPointer(0).X1 = mPointerX - 14 * 15 / Screen.TwipsPerPixelX - 0.5
    linPointer(0).X2 = linPointer(0).X1 + 8 * 15 / Screen.TwipsPerPixelX
    linPointer(0).Y1 = mPointerY - 0.5
    linPointer(0).Y2 = mPointerY - 0.5

    linPointer(1).X1 = mPointerX + 14 * 15 / Screen.TwipsPerPixelX - 0.5
    linPointer(1).X2 = linPointer(1).X1 - 8 * 15 / Screen.TwipsPerPixelX
    linPointer(1).Y1 = mPointerY - 0.5
    linPointer(1).Y2 = mPointerY - 0.5

    linPointer(2).Y1 = mPointerY - 14 * 15 / Screen.TwipsPerPixelY - 0.5
    linPointer(2).Y2 = linPointer(2).Y1 + 8 * 15 / Screen.TwipsPerPixelY
    linPointer(2).X1 = mPointerX - 0.5
    linPointer(2).X2 = mPointerX - 0.5

    linPointer(3).Y1 = mPointerY + 14 * 15 / Screen.TwipsPerPixelY - 0.5
    linPointer(3).Y2 = linPointer(3).Y1 - 8 * 15 / Screen.TwipsPerPixelY
    linPointer(3).X1 = mPointerX - 0.5
    linPointer(3).X2 = mPointerX - 0.5

End Sub


Public Property Let SelectionParametersAvailable(nValue As CDSelectionParametersAvailable)
    Dim iPrev As CDSelectionParametersAvailable
    
    If nValue <> mSelectionParametersAvailable Then
        If (nValue < cdSelectionParametersNone) Or (nValue > cdSelectionParametersAll) Then
            Err.Raise 380, TypeName(Me)
            Exit Property
        End If
        iPrev = mSelectionParametersAvailable
        mSelectionParametersAvailable = nValue
        If mInitialized Then PropertyChanged "SelectionParametersAvailable"
        If (mSelectionParametersAvailable <> cdSelectionParametersNone) And (iPrev = cdSelectionParametersNone) Or (mSelectionParametersAvailable = cdSelectionParametersNone) And (iPrev <> cdSelectionParametersNone) Then
            LoadcboSelectionParameter
            UserControl_Resize
        Else
            LoadcboSelectionParameter
        End If
        SetPicShades
    End If
End Property

Public Property Get SelectionParametersAvailable() As CDSelectionParametersAvailable
    SelectionParametersAvailable = mSelectionParametersAvailable
End Property


Public Property Let DrawFixedControlVisible(nValue As Boolean)
    If nValue <> mDrawFixedControlVisible Then
        mDrawFixedControlVisible = nValue
        chkDrawFixed.Visible = mDrawFixedControlVisible
        If mInitialized Then PropertyChanged "DrawFixedControlVisible"
        UserControl_Resize
    End If
End Property

Public Property Get DrawFixedControlVisible() As Boolean
    DrawFixedControlVisible = mDrawFixedControlVisible
End Property


Public Property Let ColorSystemControlVisible(nValue As Boolean)
    If nValue <> mColorSystemControlVisible Then
        mColorSystemControlVisible = nValue
        UserControl_Resize
        picColorSystem.Visible = mColorSystemControlVisible
        cboColorSystem.Visible = mColorSystemControlVisible
        If mInitialized Then PropertyChanged "ColorSystemControlVisible"
    End If
End Property

Public Property Get ColorSystemControlVisible() As Boolean
    ColorSystemControlVisible = mColorSystemControlVisible
End Property


Public Property Let DrawFixed(nValue As Boolean)
    If nValue <> mDrawFixed Then
        mDrawFixed = nValue
        If mInitialized Then PropertyChanged "DrawFixed"
        chkDrawFixed.Value = Abs(CLng(mDrawFixed))
        DrawWheel
        DrawShades
        If mRaiseEvents Then RaiseEvent DrawFixedChange
    End If
End Property

Public Property Get DrawFixed() As Boolean
    DrawFixed = mDrawFixed
End Property


Public Property Let SelectionParameter(nValue As CDColorWheelParameterConstants)
    If nValue <> mSelectionParameter Then
        mSelectionParameter = nValue
        If mInitialized Then PropertyChanged "SelectionParameter"
        If SelectionParametersAvailable <> cdSelectionParametersNone Then
            If Not IsItemDataInList(cboSelectionParameter, mSelectionParameter) Then
                If mSelectionParameter = cdParameterHue Then
                    SelectionParametersAvailable = cdSelectionParametersHueLumAndSat
                Else
                    SelectionParametersAvailable = cdSelectionParametersAll
                End If
            End If
        End If
        SetSelectionParameter
        If mRaiseEvents Then RaiseEvent SelectionParameterChange
    End If
End Property

Private Sub SetSelectionParameter()
    mSettingSlider = True
    SelectInListByItemData cboSelectionParameter, mSelectionParameter
    StoreWheelColors
    If mSelectionParameter = cdParameterLuminance Then
        mSliderMax = mL_Max
        SliderValue = mL_Max - mL
        mPureColor = ColorCurrentColorSystemToRGB(mH, mL_Fixed, mS)
        mRadialParameter = cdParameterSaturation
        mAxialParameter = cdParameterHue
    ElseIf mSelectionParameter = cdParameterHue Then
        mSliderMax = mH_Max
        SliderValue = mH_Max - mH
        mPureColor = ColorCurrentColorSystemToRGB(mH_Fixed, mL, mS)
        mRadialParameter = cdParameterSaturation
        mAxialParameter = cdParameterLuminance
    ElseIf mSelectionParameter = cdParameterSaturation Then
        mSliderMax = mS_Max
        SliderValue = mS_Max - mS
        mPureColor = ColorCurrentColorSystemToRGB(mH, mL, mS_Fixed)
        mRadialParameter = cdParameterLuminance
        mAxialParameter = cdParameterHue
    ElseIf mSelectionParameter = cdParameterRed Then
        mSliderMax = 255
        SliderValue = 255 - mR
        mPureColor = RGB(0, mG, mB)
        mRadialParameter = cdParameterBlue
        mAxialParameter = cdParameterGreen
    ElseIf mSelectionParameter = cdParameterGreen Then
        mSliderMax = 255
        SliderValue = 255 - mG
        mPureColor = RGB(mR, 0, mB)
        mRadialParameter = cdParameterBlue
        mAxialParameter = cdParameterRed
    Else ' cdParameterBlue
        mSliderMax = 255
        SliderValue = 255 - mB
        mPureColor = RGB(mR, mG, 0)
        mRadialParameter = cdParameterGreen
        mAxialParameter = cdParameterRed
    End If
    mSettingSlider = False
    DrawWheel
    DrawShades
    ShowSelectedColor
End Sub

Public Property Get SelectionParameter() As CDColorWheelParameterConstants
    SelectionParameter = mSelectionParameter
End Property


Public Property Let ColorSystem(nValue As CDColorSystemConstants)
    Dim iColor As Long
    Dim c As Long
    
    If nValue <> mColorSystem Then
        iColor = mColor
        mColor = -1
        mColorSystem = nValue
        ColorRGBToCurrentColorSystem iColor, mH, mL, mS
        SetMaxAndFixedvalues
        StoreWheelColors
        mChangingColorSystemOrInitializing = True
        For c = 0 To cboSelectionParameter.ListCount - 1
            If cboSelectionParameter.ItemData(c) = cdParameterLuminance Then
                If mColorSystem = cdColorSystemHSV Then
                    cboSelectionParameter.List(c) = mCaptionVal
                Else
                    cboSelectionParameter.List(c) = mCaptionLum
                End If
                Exit For
            End If
        Next c
        SetSelectionParameter
        SetColor iColor
        cboColorSystem.ListIndex = mColorSystem
        mChangingColorSystemOrInitializing = False
        DrawWheel
        DrawShades
        If mInitialized Then PropertyChanged "ColorSystem"
        If mRaiseEvents Then RaiseEvent ColorSystemChange
    End If
End Property

Public Property Get ColorSystem() As CDColorSystemConstants
    ColorSystem = mColorSystem
End Property

Public Property Let BackColor(Value As OLE_COLOR)
    If Value <> mBackColor Then
        mBackColor = Value
        SetBackColor
        mDiameter = 0
        Init
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Private Sub SetBackColor()
    UserControl.BackColor = mBackColor
    picSlider.BackColor = mBackColor
    picColorSystem.BackColor = mBackColor
    If GetColorBrightness(UserControl.BackColor) > 170 Then
        lblMode.ForeColor = vbWindowText
    Else
        lblMode.ForeColor = vbWindowBackground
    End If
End Sub


Private Sub SetMaxAndFixedvalues()
    If mColorSystem = cdColorSystemHSV Then
        mH_Max = 360
        mL_Max = 100
        mS_Max = 100
        mH_Fixed = 240
        mL_Fixed = mL_Max
        mS_Fixed = mS_Max
    Else
        mH_Max = 240
        mL_Max = 240
        mS_Max = 240
        mH_Fixed = 160
        mL_Fixed = 120
        mS_Fixed = mS_Max
    End If
End Sub

Private Function GetColorBrightness(ByVal nColor As Long) As Long
    Dim iRGB As RGBQuad
    
    TranslateColor nColor, 0&, nColor
    CopyMemory iRGB, nColor, 4
    GetColorBrightness = (0.2125 * iRGB.R + 0.7154 * iRGB.G + 0.0721 * iRGB.B)
End Function

Public Sub SetCaption(ByVal CaptionID As CDColorWheelCaptionsIDConstants, nCaption As String)
    Dim iIsCurrent As Boolean
    
    If CaptionID = cdCWCaptionFixed Then
        chkDrawFixed.Caption = nCaption
    ElseIf CaptionID = cdCWCaptionFixedToolTipText Then
        chkDrawFixed.ToolTipText = nCaption
    ElseIf CaptionID = cdCWCaptionSelectionParameterToolTipText Then
        cboSelectionParameter.ToolTipText = nCaption
    ElseIf CaptionID = cdCWCaptionMode Then
        lblMode.Caption = nCaption
    Else
        If CaptionID = cdCWCaptionLum Then
            mCaptionLum = nCaption
        ElseIf CaptionID = cdCWCaptionVal Then
            mCaptionVal = nCaption
            CaptionID = cdCWCaptionLum
        End If
        mParametersCaptions(CaptionID) = nCaption
        LoadcboSelectionParameter
    End If
End Sub

Private Sub LoadcboSelectionParameter()
    Dim c As Long
    Dim iCurrent As Long
    Dim iFrom As Long
    Dim iTo As Long
    
    If mSelectionParametersAvailable = cdSelectionParametersNone Then Exit Sub
    
    If cboSelectionParameter.ListIndex > -1 Then
        iCurrent = cboSelectionParameter.ItemData(cboSelectionParameter.ListIndex)
    Else
        iCurrent = -1
    End If
    cboSelectionParameter.Clear
    If mSelectionParametersAvailable = cdSelectionParametersAll Then
        iFrom = 0
        iTo = 5
    ElseIf mSelectionParametersAvailable = cdSelectionParametersHueLumAndSat Then
        iFrom = 0
        iTo = 2
    ElseIf mSelectionParametersAvailable = cdSelectionParametersLumAndSat Then
        iFrom = 1
        iTo = 2
    End If
    For c = iFrom To iTo
        If c = cdParameterLuminance Then
            If mColorSystem = cdColorSystemHSV Then
                cboSelectionParameter.AddItem mCaptionVal
            Else
                cboSelectionParameter.AddItem mCaptionLum
            End If
        Else
            cboSelectionParameter.AddItem mParametersCaptions(c)
        End If
        cboSelectionParameter.ItemData(cboSelectionParameter.NewIndex) = c
    Next c
    If iCurrent > -1 Then
        If Not SelectInListByItemData(cboSelectionParameter, iCurrent) Then
            cboSelectionParameter.ListIndex = 0
        End If
    End If
End Sub

Private Function SelectInListByItemData(nListControl As Object, nItemData As Long) As Boolean
    Dim c As Long
    
    For c = 0 To nListControl.ListCount - 1
        If nListControl.ItemData(c) = nItemData Then
            nListControl.ListIndex = c
            SelectInListByItemData = True
            Exit For
        End If
    Next c
End Function

Private Function IsItemDataInList(nListControl As Object, nItemData As Long) As Boolean
    Dim c As Long
    
    For c = 0 To nListControl.ListCount - 1
        If nListControl.ItemData(c) = nItemData Then
            IsItemDataInList = True
            Exit For
        End If
    Next c
End Function

Public Function GetCaption(CaptionID As CDColorWheelCaptionsIDConstants) As String
    If CaptionID = cdCWCaptionFixed Then
        GetCaption = chkDrawFixed.Caption
    ElseIf CaptionID = cdCWCaptionFixedToolTipText Then
        GetCaption = chkDrawFixed.ToolTipText
    ElseIf CaptionID = cdCWCaptionSelectionParameterToolTipText Then
        GetCaption = cboSelectionParameter.ToolTipText
    ElseIf CaptionID = cdCWCaptionLum Then
        GetCaption = mCaptionLum
    ElseIf CaptionID = cdCWCaptionVal Then
        GetCaption = mCaptionVal
    ElseIf CaptionID = cdCWCaptionMode Then
        GetCaption = lblMode.Caption
    Else
        GetCaption = mParametersCaptions(CaptionID)
    End If
End Function
    
Private Sub ColorRGBToCurrentColorSystem(nColorRGB As Long, nHue As Double, nLuminance As Double, nSaturation As Double)
    If mColorSystem = cdColorSystemHSL Then
        Dim iH1 As Integer
        Dim iL1 As Integer
        Dim iS1 As Integer
        
        ColorRGBToHLS nColorRGB, iH1, iL1, iS1
        nHue = CDbl(iH1)
        nLuminance = CDbl(iL1)
        nSaturation = CDbl(iS1)
    Else
        ColorRGBToHSV nColorRGB, nHue, nLuminance, nSaturation
    End If
End Sub
    
Private Sub ColorRGBToHSV(ByVal nColorRGB As Long, nHue As Double, nValue As Double, nSaturation As Double)
'--- based on wqweto (Vlad Vissoultchev)'s code  from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=36529
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an RGB value to the HSB color model. Adapted from Java.awt.pvColor.java
    Dim nTemp           As Double
    Dim lMin            As Long
    Dim LMax            As Long
    Dim lDelta          As Long
    Dim rgbValue        As RGBQuad
    
    Call OleTranslateColor(nColorRGB, 0, rgbValue)
    LMax = IIf(rgbValue.R > rgbValue.G, IIf(rgbValue.R > rgbValue.B, rgbValue.R, rgbValue.B), IIf(rgbValue.G > rgbValue.B, rgbValue.G, rgbValue.B))
    lMin = IIf(rgbValue.R < rgbValue.G, IIf(rgbValue.R < rgbValue.B, rgbValue.R, rgbValue.B), IIf(rgbValue.G < rgbValue.B, rgbValue.G, rgbValue.B))
    lDelta = LMax - lMin
    nValue = (LMax * 100) / 255
    If LMax > 0 Then
        nSaturation = (lDelta / LMax) * 100
        If lDelta > 0 Then
            If LMax = rgbValue.R Then
                nTemp = (CLng(rgbValue.G) - rgbValue.B) / lDelta
            ElseIf LMax = rgbValue.G Then
                nTemp = 2 + (CLng(rgbValue.B) - rgbValue.R) / lDelta
            Else
                nTemp = 4 + (CLng(rgbValue.R) - rgbValue.G) / lDelta
            End If
            nHue = nTemp * 60
            If nHue < 0 Then
                nHue = nHue + 360
            End If
        End If
    End If
End Sub
    
Private Function ColorCurrentColorSystemToRGB(nHue As Double, nLuminance As Double, nSaturation As Double) As Long
    If mColorSystem = cdColorSystemHSL Then
        Dim iH1 As Long
        Dim iL1 As Long
        Dim iS1 As Long
        
        iH1 = CLng(nHue)
        iL1 = CLng(nLuminance)
        iS1 = CLng(nSaturation)
        
        ColorCurrentColorSystemToRGB = ColorHLSToRGB(iH1, iL1, iS1)
    Else
        ColorCurrentColorSystemToRGB = ColorHSVToRGB(nHue, nLuminance, nSaturation)
    End If
End Function

Private Function ColorHSVToRGB(nHue As Double, nValue As Double, nSaturation As Double) As Long
'--- based on wqweto (Vlad Vissoultchev)'s code  from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=36529
'--- based on *cool* code by Branco Medeiros (http://www.myrealbox.com/branco_medeiros)
'--- Converts an HSB value to the RGB color model. Adapted from Java.awt.pvColor.java
    Dim nH              As Double
    Dim nS              As Double
    Dim nL              As Double
    Dim nF              As Double
    Dim nP              As Double
    Dim nQ              As Double
    Dim nT              As Double
    Dim lH              As Long
    Dim clrConv         As RGBQuad
    
    With clrConv
        If nSaturation > 0 Then
            nH = nHue / 60
            nL = nValue / 100
            nS = nSaturation / 100
            lH = Int(nH)
            nF = nH - lH
            nP = nL * (1 - nS)
            nQ = nL * (1 - nS * nF)
            nT = nL * (1 - nS * (1 - nF))
            Select Case lH
            Case 0
                .R = nL * 255
                .G = nT * 255
                .B = nP * 255
            Case 1
                .R = nQ * 255
                .G = nL * 255
                .B = nP * 255
            Case 2
                .R = nP * 255
                .G = nL * 255
                .B = nT * 255
            Case 3
                .R = nP * 255
                .G = nQ * 255
                .B = nL * 255
            Case 4
                .R = nT * 255
                .G = nP * 255
                .B = nL * 255
            Case 5
                .R = nL * 255
                .G = nP * 255
                .B = nQ * 255
            End Select
        Else
            .R = (nValue * 255) / 100
            .G = .R
            .B = .R
        End If
    End With
    '--- return long
    CopyMemory lH, clrConv, 4
    ColorHSVToRGB = lH
End Function

Public Property Get HMax() As Long
    HMax = mH_Max
End Property

Public Property Get LMax() As Long
    LMax = mL_Max
End Property

Public Property Get SMax() As Long
    SMax = mS_Max
End Property


Public Property Let Redraw(nValue As Boolean)
    If nValue <> mRedraw Then
        mRedraw = nValue
        If mRedraw Then
            If mDrawPending Then
                mDrawPending = False
                DrawWheel
                DrawShades
                ShowSelectedColor
            End If
        End If
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Public Property Get ParameterSelectorLeft() As Single
    ParameterSelectorLeft = Round(UserControl.ScaleX(picShades.Left, UserControl.ScaleMode, vbContainerPosition))
End Property

Public Property Get ParameterSelectorWidth() As Single
    ParameterSelectorWidth = Round(UserControl.ScaleX(picShades.Width, UserControl.ScaleMode, vbContainerSize))
End Property

Public Property Get SelectionParameterControlLeft() As Single
    SelectionParameterControlLeft = Round(UserControl.ScaleX(cboSelectionParameter.Left, UserControl.ScaleMode, vbContainerPosition))
End Property

Public Property Get SelectionParameterControlWidth() As Single
    SelectionParameterControlWidth = Round(UserControl.ScaleX(cboSelectionParameter.Width, UserControl.ScaleMode, vbContainerSize))
End Property

Public Property Get WheelCenterLeft() As Single
    WheelCenterLeft = Round(UserControl.ScaleX(mCx, vbPixels, vbContainerPosition))
End Property
    
Public Property Get WheelCenterTop() As Single
    WheelCenterTop = Round(UserControl.ScaleY(mCy, vbPixels, vbContainerPosition))
End Property
    
' Slider control
Private Sub picSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSlider_MouseMove Button, Shift, X, Y
End Sub

Private Sub picSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SliderValue = Y / picSlider.ScaleHeight * (mSliderMax - mSliderMin) + mSliderMin
        If SliderValue > mSliderMax Then SliderValue = mSliderMax
        If SliderValue < mSliderMin Then SliderValue = mSliderMin
    End If
End Sub

Private Sub DrawSliderGrip()
    Dim iPoints() As POINTAPI
    
    If mSliderMax <= mSliderMin Then mSliderMax = mSliderMin + 1
    ReDim iPoints(2)
    iPoints(0).Y = (picSlider.ScaleHeight - mGripLenght) / (mSliderMax - mSliderMin) * (mSliderValue - mSliderMin) + mGripLenght / 2
    iPoints(0).X = 0
    iPoints(1).Y = iPoints(0).Y - mGripLenght / 2
    iPoints(1).X = picSlider.ScaleWidth - 2
    iPoints(2).Y = iPoints(0).Y + mGripLenght / 2
    iPoints(2).X = picSlider.ScaleWidth - 2
    
    picSlider.ForeColor = &HFF0000
    picSlider.FillColor = &H6E6E6E
    picSlider.FillStyle = vbFSSolid
    picSlider.DrawStyle = vbSolid
    picSlider.DrawWidth = 1
    picSlider.Cls
    Polygon picSlider.hDC, iPoints(0), UBound(iPoints) + 1
    picSlider.Refresh
End Sub

' Slider
Private Sub InitSlider()
    mSliderMin = 0
    mSliderMax = 100
    SliderValue = 50
    picSlider.BackColor = mBackColor
    mGripLenght = 14 * 15 / Screen.TwipsPerPixelY
    mGripWidth = 7 * 15 / Screen.TwipsPerPixelY
    picSlider.AutoRedraw = True
    picSlider.Width = (mGripWidth + 2) * Screen.TwipsPerPixelX
    DrawSliderGrip
End Sub

Private Property Let RadialValue(ByVal nValue As Double)
    If mRadialParameter = cdParameterHue Then
        If nValue < 0 Then nValue = 0
        If nValue > mH_Max Then nValue = mH_Max
        H = nValue
    ElseIf mRadialParameter = cdParameterLuminance Then
        If nValue < 0 Then nValue = 0
        If nValue > mL_Max Then nValue = mL_Max
        L = nValue
    ElseIf mRadialParameter = cdParameterSaturation Then
        If nValue < 0 Then nValue = 0
        If nValue > mS_Max Then nValue = mS_Max
        S = nValue
    ElseIf mRadialParameter = cdParameterRed Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        R = nValue
    ElseIf mRadialParameter = cdParameterGreen Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        G = nValue
    ElseIf mRadialParameter = cdParameterBlue Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        B = nValue
    End If
End Property

Private Property Get RadialValue() As Double
    If mRadialParameter = cdParameterHue Then
        RadialValue = mH
    ElseIf mRadialParameter = cdParameterLuminance Then
        RadialValue = mL
    ElseIf mRadialParameter = cdParameterSaturation Then
        RadialValue = mS
    ElseIf mRadialParameter = cdParameterRed Then
        RadialValue = mR
    ElseIf mRadialParameter = cdParameterGreen Then
        RadialValue = mG
    ElseIf mRadialParameter = cdParameterBlue Then
        RadialValue = mB
    End If
End Property

Private Property Get RadialMax() As Long
    If mRadialParameter = cdParameterHue Then
        RadialMax = mH_Max
    ElseIf mRadialParameter = cdParameterLuminance Then
        RadialMax = mL_Max
    ElseIf mRadialParameter = cdParameterSaturation Then
        RadialMax = mS_Max
    ElseIf mRadialParameter = cdParameterRed Then
        RadialMax = 255
    ElseIf mRadialParameter = cdParameterGreen Then
        RadialMax = 255
    ElseIf mRadialParameter = cdParameterBlue Then
        RadialMax = 255
    End If
End Property


Private Property Let AxialValue(ByVal nValue As Double)
    If mAxialParameter = cdParameterHue Then
        If nValue < 0 Then nValue = 0
        If nValue > mH_Max Then nValue = mH_Max
        H = nValue
    ElseIf mAxialParameter = cdParameterLuminance Then
        If nValue < 0 Then nValue = 0
        If nValue > mL_Max Then nValue = mL_Max
        L = nValue
    ElseIf mAxialParameter = cdParameterSaturation Then
        If nValue < 0 Then nValue = 0
        If nValue > mS_Max Then nValue = mS_Max
        S = nValue
    ElseIf mAxialParameter = cdParameterRed Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        R = nValue
    ElseIf mAxialParameter = cdParameterGreen Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        G = nValue
    ElseIf mAxialParameter = cdParameterBlue Then
        If nValue < 0 Then nValue = 0
        If nValue > 255 Then nValue = 255
        B = nValue
    End If
End Property

Private Property Get AxialValue() As Double
    If mAxialParameter = cdParameterHue Then
        AxialValue = mH
    ElseIf mAxialParameter = cdParameterLuminance Then
        AxialValue = mL
    ElseIf mAxialParameter = cdParameterSaturation Then
        AxialValue = mS
    ElseIf mAxialParameter = cdParameterRed Then
        AxialValue = mR
    ElseIf mAxialParameter = cdParameterGreen Then
        AxialValue = mG
    ElseIf mAxialParameter = cdParameterBlue Then
        AxialValue = mB
    End If
End Property

Private Property Get AxialMax() As Long
    If mAxialParameter = cdParameterHue Then
        AxialMax = mH_Max
    ElseIf mAxialParameter = cdParameterLuminance Then
        AxialMax = mL_Max
    ElseIf mAxialParameter = cdParameterSaturation Then
        AxialMax = mS_Max
    ElseIf mAxialParameter = cdParameterRed Then
        AxialMax = 255
    ElseIf mAxialParameter = cdParameterGreen Then
        AxialMax = 255
    ElseIf mAxialParameter = cdParameterBlue Then
        AxialMax = 255
    End If
End Property

Public Property Get RadialParameter() As CDColorWheelParameterConstants
    RadialParameter = mRadialParameter
End Property

Public Property Get AxialParameter() As CDColorWheelParameterConstants
    AxialParameter = mAxialParameter
End Property


'--- for MST subclassing (2)
'Autor: wqweto http://www.vbforums.com/showthread.php?872819
'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================
Private Sub pvSubclass()
    If mUserControlHwnd <> 0 Then
        If (Not InIDE) Or cSUBCLASS_IN_IDE Then
            Set m_pSubclass = InitSubclassingThunk(mUserControlHwnd, InitAddressOfMethod().SubclassProc(0, 0, 0, 0, 0))
        End If
    End If
End Sub

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
End Sub

Private Function InitAddressOfMethod() As ColorWheel
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(Me), 5, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, ByVal pfnCallback As Long) As IUnknown
Attribute InitSubclassingThunk.VB_MemberFlags = "40"
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEDMAV1aLdCQUg8YIgz4AdC+L+oHH/BEzAIvCBQQRMwCri8IFQBEzAKuLwgVQETMAq4vCBXgRMwCruQkAAADzpYHC/BEzAFJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4MBIzAC1wEDMAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBACQVYvsi1UYiwqLQSyFwHQ1Uv/QWoP4AXdUg/gAdQmBfQwDAgAAdEaLClL/UTBahcB1O4sKUmrw/3Ek/1EoWqkAAAAIdShSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/dQj/cgz/UhBZWFqFyXURiwr/dRT/dRD/dQz/dQj/USBdwhgADx8A" ' 29.3.2019 13:04:54
    Const THUNK_SIZE    As Long = 448
    Dim hThunk          As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(Me)
    aParams(1) = pfnCallback
    hThunk = GetProp(pvGetGlobalHwnd(), "InitSubclassingThunk")
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        Call SetProp(pvGetGlobalHwnd(), "InitSubclassingThunk", hThunk)
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function

Private Function pvGetGlobalHwnd() As Long
    pvGetGlobalHwnd = FindWindowEx(0, 0, "STATIC", App.hInstance & ":" & App.ThreadID & ":MST Global Data")
    If pvGetGlobalHwnd = 0 Then
        pvGetGlobalHwnd = CreateWindowEx(0, "STATIC", App.hInstance & ":" & App.ThreadID & ":MST Global Data", _
            0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0)
    End If
End Function

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    Dim iDelta As Long
    Dim iHandle As Boolean
    Dim iPt As POINTAPI
    Dim iLng As Long
    Dim iShiftPressed As Boolean
    
    #If hWnd And wParam And lParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
        Case WM_MOUSEWHEEL
            If (wParam And 128) = 0 Then ' if not already handled
                GetCursorPos iPt
                ScreenToClient mUserControlHwnd, iPt
                iShiftPressed = (GetAsyncKeyState(vbKeyShift) < 0)
                If Sqr((iPt.X - mCx) ^ 2 + (iPt.Y - mCy) ^ 2) <= mRadius Then ' if inside the wheel
                    iDelta = WordHi(wParam)
                    If (GetAsyncKeyState(vbKeyControl) < 0) Then
                        If iDelta > 1 Then
                            iLng = RadialValue + IIf(iShiftPressed, 1, RadialMax / 15)
                            If iLng > RadialMax Then iLng = RadialMax
                            RadialValue = iLng
                        Else
                            iLng = RadialValue - IIf(iShiftPressed, 1, RadialMax / 15)
                            If iLng < 0 Then iLng = 0
                            RadialValue = iLng
                        End If
                        RaiseEvent MouseWheelNavigation(cdMouseWheelNavigatingRadial)
                    Else
                        If iDelta > 1 Then
                            iLng = AxialValue + IIf(iShiftPressed, 1, AxialMax / 30)
                            If iLng > AxialMax Then iLng = iLng - AxialMax
                            AxialValue = iLng
                        Else
                            iLng = AxialValue - IIf(iShiftPressed, 1, AxialMax / 30)
                            If iLng < 0 Then iLng = iLng + AxialMax
                            AxialValue = iLng
                        End If
                        RaiseEvent MouseWheelNavigation(cdMouseWheelNavigatingAxial)
                    End If
                    wParam = wParam Or 128
                Else
                    iHandle = False
                    If (iPt.X >= (picShades.Left - picShades.Width) / Screen.TwipsPerPixelX) Then 'And (iPt.X <= (picSlider.Left + picSlider.Width + picShades.Width) / Screen.TwipsPerPixelX) Then
                        If (iPt.Y >= picShades.Top / Screen.TwipsPerPixelY) And (iPt.Y <= (picShades.Top + picShades.Height) / Screen.TwipsPerPixelY) Then
                            iHandle = True
                        End If
                    End If
                    If iHandle Then ' if inside or near the slider
                        iDelta = WordHi(wParam)
                        If iDelta > 1 Then
                            SliderValue = SliderValue - IIf(iShiftPressed, 1, mSliderMax / 30)
                        Else
                            SliderValue = SliderValue + IIf(iShiftPressed, 1, mSliderMax / 30)
                        End If
                        RaiseEvent MouseWheelNavigation(cdMouseWheelNavigatingSlider)
                        wParam = wParam Or 128
                    End If
                End If
            End If
    End Select
    If Not mAmbientUserMode Then
        Handled = True
        SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
    End If
End Function

Private Function WordHi(ByVal LongIn As Long) As Integer
    ' Mask off low word then do integer divide to
    ' shift right by 16.
    
    WordHi = (LongIn And &HFFFF0000) \ &H10000
End Function

'--- End for MST subclassing (2)

Private Function InIDE() As Boolean
    Static sValue As Long
    Dim iErrNumber As Integer
    Dim iErrDesc As String
    Dim iErrSource As String
    
    If sValue = 0 Then
        iErrNumber = Err.Number
        iErrDesc = Err.Description
        iErrSource = Err.Source
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
        Err.Number = iErrNumber
        Err.Description = iErrDesc
        Err.Source = iErrSource
    End If
    InIDE = (sValue = 1)
End Function

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
