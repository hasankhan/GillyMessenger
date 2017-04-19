Attribute VB_Name = "modGradient"
Option Explicit

Private Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer 'Ushort value
    Green As Integer 'Ushort value
    Blue As Integer 'ushort value
    Alpha As Integer 'ushort
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long  'In reality this is a UNSIGNED Long
    LowerRight As Long 'In reality this is a UNSIGNED Long
End Type

Const GRADIENT_FILL_RECT_H As Long = &H0 'In this mode, two endpoints describe a rectangle. The rectangle is
'defined to have a constant color (specified by the TRIVERTEX structure) for the left and right edges. GDI interpolates
'the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_RECT_V  As Long = &H1 'In this mode, two endpoints describe a rectangle. The rectangle
' is defined to have a constant color (specified by the TRIVERTEX structure) for the top and bottom edges. GDI interpolates
' the color from the top to bottom edge and fills the interior.
Const GRADIENT_FILL_TRIANGLE As Long = &H2 'In this mode, an array of TRIVERTEX structures is passed to GDI
'along with a list of array indexes that describe separate triangles. GDI performs linear interpolation between triangle vertices
'and fills the interior. Drawing is done directly in 24- and 32-bpp modes. Dithering is performed in 16-, 8.4-, and 1-bpp mode.
Const GRADIENT_FILL_OP_FLAG As Long = &HFF

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

Public Sub GradientFill(hDC As Long, X_Start As Long, Y_Start As Long, X_Stop As Long, Y_Stop As Long, StartColor As String, StopColor As String, Vertical As Boolean)
Dim vert(1) As TRIVERTEX
    Dim gRect As GRADIENT_RECT

    'from black
    With vert(0)
        .X = X_Start
        .Y = Y_Start
        .Red = Val("&H" & Left(StartColor, 2) & "00")
        .Green = Val("&H" & Mid(StartColor, 3, 2) & "00")
        .Blue = Val("&H" & Right(StartColor, 2) & "00")
        .Alpha = 0&
    End With

    'to blue
    With vert(1)
        .X = X_Stop
        .Y = Y_Stop
        .Red = Val("&H" & Left(StopColor, 2) & "00")
        .Green = Val("&H" & Mid(StopColor, 3, 2) & "00")
        .Blue = Val("&H" & Right(StopColor, 2) & "00")
        .Alpha = 0&
    End With

    gRect.UpperLeft = 0
    gRect.LowerRight = 1

    If Vertical Then
        GradientFillRect hDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V
    Else
        GradientFillRect hDC, vert(0), 2, gRect, 1, GRADIENT_FILL_RECT_H
    End If
End Sub

