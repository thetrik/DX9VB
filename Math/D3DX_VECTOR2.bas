Attribute VB_Name = "D3DX_VECTOR2"
Option Explicit

' // Adds two 2D vectors.
Public Sub D3DXVec2Add(pOut As D3DVECTOR2, _
                       pV1 As D3DVECTOR2, _
                       pV2 As D3DVECTOR2)
          
    pOut.X = pV1.X + pV2.X
    pOut.Y = pV1.Y + pV2.Y
          
End Sub

' // Returns a point in Barycentric coordinates, using the specified 2D vectors.
Public Sub D3DXVec2BaryCentric(pOut As D3DVECTOR2, _
                               pV1 As D3DVECTOR2, _
                               pV2 As D3DVECTOR2, _
                               pV3 As D3DVECTOR2, _
                               ByVal f As Single, _
                               ByVal g As Single)
    Dim tmp As Single
    
    tmp = (1! - f - g)
    
    pOut.X = tmp * (pV1.X) + f * (pV2.X) + g * (pV3.X)
    pOut.Y = tmp * (pV1.Y) + f * (pV2.Y) + g * (pV3.Y)
    
End Sub

' // Returns the z-component by taking the cross product of two 2D vectors.
Public Function D3DXVec2CCW(pV1 As D3DVECTOR2, _
                            pV2 As D3DVECTOR2) As Single
         
    D3DXVec2CCW = pV1.X * pV2.Y - pV1.Y * pV2.X
    
End Function

' // Determines the dot product of two 2D vectors.
Public Function D3DXVecDot(pV1 As D3DVECTOR2, _
                           pV2 As D3DVECTOR2) As Single
         
    D3DXVecDot = pV1.X * pV2.X + pV1.Y * pV2.Y
    
End Function

' // Performs a Catmull-Rom interpolation, using the specified 2D vectors.
Public Sub D3DXVec2CatmullRom(pOut As D3DVECTOR2, _
                              pV0 As D3DVECTOR2, _
                              pV1 As D3DVECTOR2, _
                              pV2 As D3DVECTOR2, _
                              pV3 As D3DVECTOR2, _
                              ByVal s As Single)
          
    pOut.X = 0.5! * (2! * pV1.X + (pV2.X - pV0.X) * s + _
                    (2! * pV0.X - 5! * pV1.X + 4! * pV2.X - pV3.X) * s * s + _
                    (pV3.X - 3! * pV2.X + 3! * pV1.X - pV0.X) * s * s * s)
    pOut.Y = 0.5! * (2! * pV1.Y + (pV2.Y - pV0.Y) * s + _
                    (2! * pV0.Y - 5! * pV1.Y + 4! * pV2.Y - pV3.Y) * s * s + _
                    (pV3.Y - 3! * pV2.Y + 3! * pV1.Y - pV0.Y) * s * s * s)
          
End Sub

' // Performs a Hermite spline interpolation, using the specified 2D vectors.
Public Sub D3DXVec2Hermite(pOut As D3DVECTOR2, _
                           pV1 As D3DVECTOR2, _
                           pT1 As D3DVECTOR2, _
                           pV2 As D3DVECTOR2, _
                           pT2 As D3DVECTOR2, _
                           ByVal s As Single)
          
    Dim h1  As Single
    Dim h2  As Single
    Dim h3  As Single
    Dim h4  As Single

    h1 = 2! * s * s * s - 3! * s * s + 1!
    h2 = s * s * s - 2! * s * s + s
    h3 = -2! * s * s * s + 3! * s * s
    h4 = s * s * s - s * s

    pOut.X = h1 * pV1.X + h2 * pT1.X + h3 * pV2.X + h4 * pT2.X
    pOut.Y = h1 * pV1.Y + h2 * pT1.Y + h3 * pV2.Y + h4 * pT2.Y

End Sub

' // Returns the length of a 2D vector.
Public Function D3DXVec2Length(pV As D3DVECTOR2) As Single
    
    D3DXVec2Length = Sqr(pV.X * pV.X + pV.Y * pV.Y)
    
End Function

' // Returns the square of the length of a 2D vector.
Public Function D3DXVec2LengthSq(pV As D3DVECTOR2) As Single
    
    D3DXVec2LengthSq = pV.X * pV.X + pV.Y + pV.Y
    
End Function

' // Performs a linear interpolation between two 2D vectors.
Public Sub D3DXVec2Lerp(pOut As D3DVECTOR2, _
                        pV1 As D3DVECTOR2, _
                        pV2 As D3DVECTOR2, _
                        ByVal s As Single)
    Dim s1  As Single
    
    s1 = 1 - s
    pOut.X = s1 * pV1.X + s * pV2.X
    pOut.Y = s1 * pV1.Y + s * pV2.Y

End Sub

' // Returns a 2D vector that is made up of the largest components of two 2D vectors.
Public Sub D3DXVec2Maximize(pOut As D3DVECTOR2, _
                            pV1 As D3DVECTOR2, _
                            pV2 As D3DVECTOR2)
    
    If pV1.X > pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y > pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    
End Sub

' // Returns a 2D vector that is made up of the smallest components of two 2D vectors.
Public Sub D3DXVec2Minimize(pOut As D3DVECTOR2, _
                            pV1 As D3DVECTOR2, _
                            pV2 As D3DVECTOR2)
    
    If pV1.X < pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y < pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    
End Sub


' // Returns the normalized version of a 2D vector.
Public Sub D3DXVec2Normalize(pOut As D3DVECTOR2, _
                             pV As D3DVECTOR2)
                  
    Dim norm    As Single

    norm = D3DXVec2Length(pV)
    
    If norm = 0! Then
        pOut.X = 0!
        pOut.Y = 0!
    Else
        pOut.X = pV.X / norm
        pOut.Y = pV.Y / norm
    End If

End Sub

' // Scales a 2D vector.
Public Sub D3DXVec2Scale(pOut As D3DVECTOR2, _
                         pV As D3DVECTOR2, _
                         ByVal s As Single)

    pOut.X = pV.X * s
    pOut.Y = pV.Y * s

End Sub

' // Subtracts two 2D vectors.
Public Sub D3DXVec2Subtract(pOut As D3DVECTOR2, _
                            pV1 As D3DVECTOR2, _
                            pV2 As D3DVECTOR2)
          
    pOut.X = pV1.X - pV2.X
    pOut.Y = pV1.Y - pV2.Y
          
End Sub

' // Transforms a 2D vector by a given matrix.
Public Sub D3DXVec2Transform(pOut As D3DVECTOR4, _
                             pV As D3DVECTOR2, _
                             pM As D3DMATRIX)
          
    pOut.X = pM.m11 * pV.X + pM.m21 * pV.Y + pM.m41
    pOut.Y = pM.m12 * pV.X + pM.m22 * pV.Y + pM.m42
    pOut.z = pM.m13 * pV.X + pM.m23 * pV.Y + pM.m43
    pOut.w = pM.m14 * pV.X + pM.m24 * pV.Y + pM.m44
    
End Sub

' // Transforms a 2D vector by a given matrix, projecting the result back into w = 1.
Public Sub D3DXVec2TransformCoord(pOut As D3DVECTOR2, _
                                  pV As D3DVECTOR2, _
                                  pM As D3DMATRIX)
          
    Dim v       As D3DVECTOR2
    Dim norm    As Single

    v = pV
    norm = pM.m14 * pV.X + pM.m24 * pV.Y + pM.m44

    pOut.X = (pM.m11 * v.X + pM.m21 * v.Y + pM.m41) / norm
    pOut.Y = (pM.m12 * v.X + pM.m22 * v.Y + pM.m42) / norm

End Sub

' // Transforms the 2D vector normal by the given matrix.
Public Sub D3DXVec2TransformNormal(pOut As D3DVECTOR2, _
                                   pV As D3DVECTOR2, _
                                   pM As D3DMATRIX)
          
    Dim v       As D3DVECTOR2

    v = pV
    pOut.X = pM.m11 * v.X + pM.m21 * v.Y
    pOut.Y = pM.m12 * v.X + pM.m22 * v.Y

End Sub


