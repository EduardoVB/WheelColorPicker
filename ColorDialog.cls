VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum CDCaptionsIDConstants
    cdCaptionHue ' Hue
    cdCaptionLum ' Lum
    cdCaptionSat ' Sat
    cdCaptionRed ' Red
    cdCaptionGreen ' Green
    cdCaptionBlue ' Blue
    cdCaptionVal ' Val
    cdCaptionFixed ' Fixed
    cdCaptionFixedToolTipText ' Reflects color changes visually or not
    cdCaptionSelectionParameterToolTipText ' Select parameter
    cdCaptionRecent ' recent
    cdCaptionNew     ' new
    cdCaptionCurrent  ' current
    cdCaptionColor ' Color
    cdCaptionHex ' Hex
    cdCaptionInvalidColorMessage ' The Hex color is not valid.
    cdCaptionOK ' OK
    cdCaptionCancel ' Cancel
    cdCaptionMenuForgetRecent ' Forget
    cdCaptionMenuClearAllRecent ' Clear all recent
    cdCaptionMode ' Mode
    cdCaptionToolTipMouseWheelBeginning ' Hold the Control key down to navigate
    cdCaptionToolTipMouseWheelEnding ' with the mouse wheel
End Enum

Private mCanceled As Boolean
Private mChanged As Boolean
Private mColor As Long
Private mColorSet As Boolean
Private mDialogTitle As String
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
Private mCaptions(22) As String
Private mContext As String

Public Function Show() As Boolean
    Dim iFrm As New frmWheelColorDialog
    
    mCanceled = False
    mChanged = False
    iFrm.HexValueVisible = mHexValueVisible
    iFrm.ColorParametersVisible = mColorParametersVisible
    iFrm.RecentColorsVisible = mRecentColorsVisible
    iFrm.DialogSizeBig = mDialogSizeBig
    iFrm.SelectionParametersAvailable = mSelectionParametersAvailable
    iFrm.DrawFixedControlVisible = mDrawFixedControlVisible
    iFrm.ColorSystemControlVisible = mColorSystemControlVisible
    iFrm.DrawFixed = mDrawFixed
    iFrm.ColorSystem = mColorSystem
    iFrm.SelectionParameter = mSelectionParameter
    iFrm.Context = mContext
    If mColorSet Then
        iFrm.CurrentColor = mColor
    End If
    If mDialogTitle <> "" Then
        iFrm.Caption = " " & mDialogTitle
    End If
    iFrm.SetCaptions mCaptions
    iFrm.Show 1
    mCanceled = Not iFrm.OKPressed
    If Not mCanceled Then
        If mColorSet Then
            mChanged = iFrm.SelectedColor <> mColor
        Else
            mChanged = True
        End If
        If mChanged Then
            mColor = iFrm.SelectedColor
        End If
    End If
    Show = mChanged
End Function


Public Property Let Color(NewColor As OLE_COLOR)
    If Not IsValidOLE_COLOR(NewColor) Then
        Err.Raise 380, TypeName(Me)
        Exit Property
    End If
    mColor = NewColor
    mColorSet = True
End Property

Public Property Get Color() As OLE_COLOR
    Color = mColor
End Property


Public Property Let DialogTitle(NewTitle As String)
    mDialogTitle = NewTitle
End Property

Public Property Get DialogTitle() As String
    DialogTitle = mDialogTitle
End Property

Public Property Get Canceled() As Boolean
    Canceled = mCanceled
End Property


Public Property Get Changed() As Boolean
    Changed = mChanged
End Property


Public Property Let HexValueVisible(nValue As Boolean)
    mHexValueVisible = nValue
End Property

Public Property Get HexValueVisible() As Boolean
    HexValueVisible = mHexValueVisible
End Property


Public Property Let ColorParametersVisible(nValue As Boolean)
    mColorParametersVisible = nValue
End Property

Public Property Get ColorParametersVisible() As Boolean
    ColorParametersVisible = mColorParametersVisible
End Property


Public Property Let RecentColorsVisible(nValue As Boolean)
    mRecentColorsVisible = nValue
End Property

Public Property Get RecentColorsVisible() As Boolean
    RecentColorsVisible = mRecentColorsVisible
End Property


Public Property Let DialogSizeBig(nValue As Boolean)
    mDialogSizeBig = nValue
End Property

Public Property Get DialogSizeBig() As Boolean
    DialogSizeBig = mDialogSizeBig
End Property


Public Property Let SelectionParametersAvailable(nValue As CDSelectionParametersAvailable)
    mSelectionParametersAvailable = nValue
End Property

Public Property Get SelectionParametersAvailable() As CDSelectionParametersAvailable
    SelectionParametersAvailable = mSelectionParametersAvailable
End Property


Public Property Let DrawFixedControlVisible(nValue As Boolean)
    mDrawFixedControlVisible = nValue
End Property

Public Property Get DrawFixedControlVisible() As Boolean
    DrawFixedControlVisible = mDrawFixedControlVisible
End Property


Public Property Let ColorSystemControlVisible(nValue As Boolean)
    mColorSystemControlVisible = nValue
End Property

Public Property Get ColorSystemControlVisible() As Boolean
    ColorSystemControlVisible = mColorSystemControlVisible
End Property


Public Property Let DrawFixed(nValue As Boolean)
    mDrawFixed = nValue
End Property

Public Property Get DrawFixed() As Boolean
    DrawFixed = mDrawFixed
End Property


Public Property Let ColorSystem(nValue As CDColorSystemConstants)
    mColorSystem = nValue
End Property

Public Property Get ColorSystem() As CDColorSystemConstants
    ColorSystem = mColorSystem
End Property


Public Property Let SelectionParameter(nValue As CDColorWheelParameterConstants)
    mSelectionParameter = nValue
End Property

Public Property Get SelectionParameter() As CDColorWheelParameterConstants
    SelectionParameter = mSelectionParameter
End Property


Public Property Let Context(nValue As String)
    mContext = nValue
End Property

Public Property Get Context() As String
    Context = mContext
End Property


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

 Sub Class_Initialize()
    mHexValueVisible = True
    mColorParametersVisible = True
    mRecentColorsVisible = True
    mDialogSizeBig = False
    mSelectionParametersAvailable = cdSelectionParametersLumAndSat
    mDrawFixedControlVisible = False
    mColorSystemControlVisible = False
    mDrawFixed = True
    mColorSystem = cdColorSystemHSV
    mSelectionParameter = cdParameterLuminance
End Sub

Public Sub SetCaption(CaptionID As CDCaptionsIDConstants, nCaption As String)
    mCaptions(CaptionID) = nCaption
End Sub

Public Sub SetComplete(Optional DialogSizeBig As Boolean = False, Optional SelectionParametersAvailable As CDSelectionParametersAvailable = cdSelectionParametersAll, Optional DrawFixedControlVisible As Boolean = True, Optional ColorSystemControlVisible As Boolean = True)
    mDialogSizeBig = DialogSizeBig
    mSelectionParametersAvailable = SelectionParametersAvailable
    mDrawFixedControlVisible = DrawFixedControlVisible
    mColorSystemControlVisible = ColorSystemControlVisible
End Sub

Public Sub SetCompact(Optional RecentColorsVisible As Boolean = False, Optional ColorParametersVisible As Boolean = False, Optional SelectionParametersAvailable As CDSelectionParametersAvailable = cdSelectionParametersNone)
    mRecentColorsVisible = RecentColorsVisible
    mColorParametersVisible = ColorParametersVisible
    mSelectionParametersAvailable = SelectionParametersAvailable
End Sub
