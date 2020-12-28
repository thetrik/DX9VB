Attribute VB_Name = "D3DX_VECTOR4"
Option Explicit

' // Adds two 4D vectors.
Public Sub D3DXVec4Add(pOut As D3DVECTOR4, _
                       pV1 As D3DVECTOR4, _
                       pV2 As D3DVECTOR4)
          
    pOut.X = pV1.X + pV2.X
    pOut.Y = pV1.Y + pV2.Y
    pOut.z = pV1.z + pV2.z
    pOut.w = pV1.w + pV2.w
    
End Sub

' // Returns a point in Barycentric coordinates, using the specified 4D vectors.
Public Sub D3DXVec4BaryCentric(pOut As D3DVECTOR4, _
                               pV1 As D3DVECTOR4, _
                               pV2 As D3DVECTOR4, _
                               pV3 As D3DVECTOR4, _
                               ByVal f As Single, _
                               ByVal g As Single)
    Dim tmp As Single
    
    tmp = (1! - f - g)
    pOut.X = tmp * (pV1.X) + f * (pV2.X) + g * (pV3.X)
    pOut.Y = tmp * (pV1.Y) + f * (pV2.Y) + g * (pV3.Y)
    pOut.z = tmp * (pV1.z) + f * (pV2.z) + g * (pV3.z)
    pOut.w = tmp * (pV1.w) + f * (pV2.w) + g * (pV3.w)
    
End Sub

' // Performs a Catmull-Rom interpolation, using the specified 4D vectors.
Public Sub D3DXVec4CatmullRom(pOut As D3DVECTOR4, _
                              pV0 As D3DVECTOR4, _
                              pV1 As D3DVECTOR4, _
                              pV2 As D3DVECTOR4, _
                              pV3 As D3DVECTOR4, _
                              ByVal s As Single)
          
    pOut.X = 0.5! * (2! * pV1.X + (pV2.X - pV0.X) * s + _
                    (2! * pV0.X - 5! * pV1.X + 4! * pV2.X - pV3.X) * s * s + _
                    (pV3.X - 3! * pV2.X + 3! * pV1.X - pV0.X) * s * s * s)
    pOut.Y = 0.5! * (2! * pV1.Y + (pV2.Y - pV0.Y) * s + _
                    (2! * pV0.Y - 5! * pV1.Y + 4! * pV2.Y - pV3.Y) * s * s + _
                    (pV3.Y - 3! * pV2.Y + 3! * pV1.Y - pV0.Y) * s * s * s)
    pOut.z = 0.5! * (2! * pV1.z + (pV2.z - pV0.z) * s + _
                    (2! * pV0.z - 5! * pV1.z + 4! * pV2.z - pV3.z) * s * s + _
                    (pV3.z - 3! * pV2.z + 3! * pV1.z - pV0.z) * s * s * s)
    pOut.w = 0.5! * (2! * pV1.w + (pV2.w - pV0.w) * s + _
                    (2! * pV0.w - 5! * pV1.w + 4! * pV2.w - pV3.w) * s * s + _
                    (pV3.w - 3! * pV2.w + 3! * pV1.w - pV0.w) * s * s * s)
    
End Sub

' // Determines the cross-product in four dimensions.
Public Sub D3DXVec4Cross(pOut As D3DVECTOR4, _
                         pV1 As D3DVECTOR4, _
                         pV2 As D3DVECTOR4, _
                         pV3 As D3DVECTOR4)
                            
    Dim out As D3DVECTOR4
    
    
    out.X = pV1.Y * (pV2.z * pV3.w - pV3.z * pV2.w) - pV1.z * (pV2.Y * pV3.w - pV3.Y * pV2.w) + pV1.w * (pV2.Y * pV3.z - pV2.z * pV3.Y)
    out.Y = -(pV1.X * (pV2.z * pV3.w - pV3.z * pV2.w) - pV1.z * (pV2.X * pV3.w - pV3.X * pV2.w) + pV1.w * (pV2.X * pV3.z - pV3.X * pV2.z))
    out.z = pV1.X * (pV2.Y * pV3.w - pV3.Y * pV2.w) - pV1.Y * (pV2.X * pV3.w - pV3.X * pV2.w) + pV1.w * (pV2.X * pV3.Y - pV3.X * pV2.Y)
    out.w = -(pV1.X * (pV2.Y * pV3.z - pV3.Y * pV2.z) - pV1.Y * (pV2.X * pV3.z - pV3.X * pV2.z) + pV1.z * (pV2.X * pV3.Y - pV3.X * pV2.Y))

    pOut = out
End Sub

' // Determines the dot product of two 4D vectors.
Public Function D3DXVec4Dot(pV1 As D3DVECTOR4, _
                            pV2 As D3DVECTOR4) As Single
                            
    D3DXVec4Dot = pV1.X * pV2.X + pV1.Y * pV2.Y + pV1.z * pV2.z + pV1.w * pV2.w
    
End Function

' // Performs a Hermite spline interpolation, using the specified 4D vectors.
Public Sub D3DXVec4Hermite(pOut As D3DVECTOR4, _
                           pV1 As D3DVECTOR4, _
                           pT1 As D3DVECTOR4, _
                           pV2 As D3DVECTOR4, _
                           pT2 As D3DVECTOR4, _
                           ByVal s As Single)
          
    Dim h1  As Single
    Dim h2  As Single
    Dim h3  As Single
    Dim h4  As Single

    h1 = 2! * s * s * s - 3! * s * s + 1!
    h2 = s * s * s - 2! * s * s + s
    h3 = -2! * s * s * s + 3! * s * s
    h4 = s * s * s - s * s

    pOut.X = h1 * (pV1.X) + h2 * (pT1.X) + h3 * (pV2.X) + h4 * (pT2.X)
    pOut.Y = h1 * (pV1.Y) + h2 * (pT1.Y) + h3 * (pV2.Y) + h4 * (pT2.Y)
    pOut.z = h1 * (pV1.z) + h2 * (pT1.z) + h3 * (pV2.z) + h4 * (pT2.z)
    pOut.w = h1 * (pV1.w) + h2 * (pT1.w) + h3 * (pV2.w) + h4 * (pT2.w)

End Sub


' // Returns the length of a 4D vector.
Public Function D3DXVec4Length(pV As D3DVECTOR4) As Single

    D3DXVec4Length = Sqr(pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z + pV.w * pV.w)
    
End Function

' // Returns the square of the length of a 4D vector.
Public Function D3DXVec4LengthSq(pV As D3DVECTOR4) As Single

    D3DXVec4LengthSq = pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z + pV.w * pV.w
    
End Function

' // Performs a linear interpolation between two 4D vectors.
Public Sub D3DXVec4Lerp(pOut As D3DVECTOR4, _
                        pV1 As D3DVECTOR4, _
                        pV2 As D3DVECTOR4, _
                        ByVal s As Single)
    Dim s1  As Single
    
    s1 = 1 - s
    pOut.X = s1 * pV1.X + s * pV2.X
    pOut.Y = s1 * pV1.Y + s * pV2.Y
    pOut.z = s1 * pV1.z + s * pV2.z
    pOut.w = s1 * pV1.w + s * pV2.w
    
End Sub

' // Returns a 4D vector that is made up of the largest components of two 4D vectors.
Public Sub D3DXVec4Maximize(pOut As D3DVECTOR4, _
                            pV1 As D3DVECTOR4, _
                            pV2 As D3DVECTOR4)
    
    If pV1.X > pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y > pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    If pV1.z > pV2.z Then pOut.z = pV1.z Else pOut.z = pV2.z
    If pV1.w > pV2.w Then pOut.w = pV1.w Else pOut.w = pV2.w
    
End Sub

' // Returns a 4D vector that is made up of the smallest components of two 4D vectors.
Public Sub D3DXVec4Minimize(pOut As D3DVECTOR4, _
                            pV1 As D3DVECTOR4, _
                            pV2 As D3DVECTOR4)
    
    If pV1.X < pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y < pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    If pV1.z < pV2.z Then pOut.z = pV1.z Else pOut.z = pV2.z
    If pV1.w < pV2.w Then pOut.w = pV1.w Else pOut.w = pV2.w
    
End Sub

' // Returns the normalized version of a 4D vector.
Public Sub D3DXVec4Normalize(pOut As D3DVECTOR4, _
                             pV As D3DVECTOR4)
    Dim norm    As Single
    
    norm = Sqr(pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z + pV.w * pV.w)
    
    If norm = 0! Then
        pOut.X = 0!:    pOut.Y = 0!:    pOut.z = 0!:    pOut.w = 0
    Else
        pOut.X = pV.X / norm
        pOut.Y = pV.Y / norm
        pOut.z = pV.z / norm
        pOut.w = pV.w / norm
    End If
    
End Sub

' // Scales a 4D vector.
Public Sub D3DXVec4Scale(pOut As D3DVECTOR4, _
                         pV As D3DVECTOR4, _
                         ByVal s As Single)

    pOut.X = pV.X * s
    pOut.Y = pV.Y * s
    pOut.z = pV.z * s
    pOut.w = pV.w * s
    
End Sub

' // Subtracts two 4D vectors.
Public Sub D3DXVec4Subtract(pOut As D3DVECTOR4, _
                            pV1 As D3DVECTOR4, _
                            pV2 As D3DVECTOR4)
                            
    pOut.X = pV1.X - pV2.X
    pOut.Y = pV1.Y - pV2.Y
    pOut.z = pV1.z - pV2.z
    pOut.w = pV1.w - pV2.w
    
End Sub

' // Transforms a 4D vector by a given matrix.
Public Sub D3DXVec4Transform(pOut As D3DVECTOR4, _
                             pV As D3DVECTOR4, _
                             pM As D3DMATRIX)
          
    pOut.X = pM.m11 * pV.X + pM.m21 * pV.Y + pM.m31 * pV.z + pM.m41 * pV.w
    pOut.Y = pM.m12 * pV.X + pM.m22 * pV.Y + pM.m32 * pV.z + pM.m42 * pV.w
    pOut.z = pM.m13 * pV.X + pM.m23 * pV.Y + pM.m33 * pV.z + pM.m43 * pV.w
    pOut.w = pM.m14 * pV.X + pM.m24 * pV.Y + pM.m34 * pV.z + pM.m44 * pV.w
    
End Sub
