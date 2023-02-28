Attribute VB_Name = "Module1"
Option Explicit

Public Sub SetColorDialogCaptions(nDialog As ColorDialog)
    With nDialog
        If .DialogTitle = "" Then .DialogTitle = "Seleccionar color"
        .SetCaption cdCaptionBlue, "Azul"
        .SetCaption cdCaptionCancel, "Cancelar"
        .SetCaption cdCaptionColor, "Color"
        .SetCaption cdCaptionCurrent, "actual"
        .SetCaption cdCaptionFixed, "Fijos"
        .SetCaption cdCaptionFixedToolTipText, "Refleja o no visualmente los cambios de colores"
        .SetCaption cdCaptionGreen, "Verde"
        .SetCaption cdCaptionHex, "Hex"
        .SetCaption cdCaptionHue, "Matiz"
        .SetCaption cdCaptionInvalidColorMessage, "El color no es válido"
        .SetCaption cdCaptionLum, "Lum."
        .SetCaption cdCaptionMenuForgetRecent, "Olvidar"
        .SetCaption cdCaptionMenuClearAllRecent, "Borrar colores recientes"
        .SetCaption cdCaptionMode, "Modo"
        .SetCaption cdCaptionNew, "nuevo"
        .SetCaption cdCaptionOK, "Aceptar"
        .SetCaption cdCaptionRecent, "Recientes"
        .SetCaption cdCaptionRed, "Rojo"
        .SetCaption cdCaptionSat, "Sat."
        .SetCaption cdCaptionSelectionParameterToolTipText, "Seleccionar parámetro"
        .SetCaption cdCaptionVal, "Valor"
        .SetCaption cdCaptionToolTipMouseWheelBeginning, "Mantenga presionada la tecla Control para navegar"
        .SetCaption cdCaptionToolTipMouseWheelEnding, "con la rueda del mouse, Mayúsculas para ir lento"
    End With
End Sub
