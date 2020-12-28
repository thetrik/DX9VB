Attribute VB_Name = "D3DX_COLOR"
Option Explicit

' // Adds two color values together to create a new color value.
Public Sub D3DXColorAdd(pOut As D3DCOLORVALUE, _
                        pC1 As D3DCOLORVALUE, _
                        pC2 As D3DCOLORVALUE)
    pOut.r = pC1.r + pC2.r
    pOut.g = pC1.g + pC2.g
    pOut.B = pC1.B + pC2.B
    pOut.a = pC1.a + pC2.a
End Sub

' // Adjusts the contrast value of a color.
Public Sub D3DXColorAdjustContrast(pOut As D3DCOLORVALUE, _
                                   pC As D3DCOLORVALUE, _
                                   ByVal c As Single)
    pOut.r = 0.5! + c * (pC.r - 0.5!)
    pOut.g = 0.5! + c * (pC.g - 0.5!)
    pOut.B = 0.5! + c * (pC.B - 0.5!)
    pOut.a = pC.a
End Sub

' // Adjusts the saturation value of a color.
Public Sub D3DXColorAdjustSaturation(pOut As D3DCOLORVALUE, _
                                     pC As D3DCOLORVALUE, _
                                     ByVal s As Single)
    Dim grey    As Single
    
    grey = pC.r * 0.2125! + pC.g * 0.7154! + pC.B * 0.0721!
    pOut.r = grey + s * (pC.r - grey)
    pOut.g = grey + s * (pC.g - grey)
    pOut.B = grey + s * (pC.B - grey)
    pOut.a = grey + s * (pC.a - grey)
End Sub

' // Uses linear interpolation to create a color value.
Public Sub D3DXColorLerp(pOut As D3DCOLORVALUE, _
                         pC1 As D3DCOLORVALUE, _
                         pC2 As D3DCOLORVALUE, _
                         ByVal s As Single)
    pOut.r = pC1.r + s * (pC2.r - pC1.r)
    pOut.g = pC1.g + s * (pC2.g - pC1.g)
    pOut.B = pC1.B + s * (pC2.B - pC1.B)
    pOut.a = pC1.a + s * (pC2.a - pC1.a)
End Sub

' // Blends two colors.
Public Sub D3DXColorModulate(pOut As D3DCOLORVALUE, _
                             pC1 As D3DCOLORVALUE, _
                             pC2 As D3DCOLORVALUE)
    pOut.r = pC1.r * pC2.r
    pOut.g = pC1.g * pC2.g
    pOut.B = pC1.B * pC2.B
    pOut.a = pC1.a * pC2.a
End Sub

' // Creates the negative color value of a color value.
Public Sub D3DXColorNegative(pOut As D3DCOLORVALUE, _
                             pC As D3DCOLORVALUE)
    pOut.r = 1! - pC.r
    pOut.g = 1! - pC.g
    pOut.B = 1! - pC.B
    pOut.a = pC.a
End Sub

' // Scales a color value.
Public Sub D3DXColorScale(pOut As D3DCOLORVALUE, _
                          pC As D3DCOLORVALUE, ByVal s As Single)
    pOut.r = pC.r * s
    pOut.g = pC.g * s
    pOut.B = pC.B * s
    pOut.a = pC.a * s
End Sub

' // Subtracts two color values to create a new color value.
Public Sub D3DXColorSubtract(pOut As D3DCOLORVALUE, _
                             pC1 As D3DCOLORVALUE, _
                             pC2 As D3DCOLORVALUE)
    pOut.r = pC1.r - pC2.r
    pOut.g = pC1.g - pC2.g
    pOut.B = pC1.B - pC2.B
    pOut.a = pC1.a - pC2.a
End Sub


