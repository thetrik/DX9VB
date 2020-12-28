Attribute VB_Name = "D3DX_VECTOR3"
Option Explicit

' // Adds two 3D vectors.
Public Sub D3DXVec3Add(pOut As D3DVECTOR, _
                       pV1 As D3DVECTOR, _
                       pV2 As D3DVECTOR)
          
    pOut.X = pV1.X + pV2.X
    pOut.Y = pV1.Y + pV2.Y
    pOut.z = pV1.z + pV2.z
    
End Sub

' // Returns a point in Barycentric coordinates, using the specified 3D vectors.
Public Sub D3DXVec3BaryCentric(pOut As D3DVECTOR, _
                               pV1 As D3DVECTOR, _
                               pV2 As D3DVECTOR, _
                               pV3 As D3DVECTOR, _
                               ByVal f As Single, _
                               ByVal g As Single)
    Dim tmp As Single
    
    tmp = (1! - f - g)
    
    pOut.X = tmp * (pV1.X) + f * (pV2.X) + g * (pV3.X)
    pOut.Y = tmp * (pV1.Y) + f * (pV2.Y) + g * (pV3.Y)
    pOut.z = tmp * (pV1.z) + f * (pV2.z) + g * (pV3.z)
    
End Sub

' // Performs a Catmull-Rom interpolation, using the specified 3D vectors.
Public Sub D3DXVec3CatmullRom(pOut As D3DVECTOR, _
                              pV0 As D3DVECTOR, _
                              pV1 As D3DVECTOR, _
                              pV2 As D3DVECTOR, _
                              pV3 As D3DVECTOR, _
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
    
End Sub

' // Determines the cross-product of two 3D vectors.
Public Sub D3DXVec3Cross(pOut As D3DVECTOR, _
                         pV1 As D3DVECTOR, _
                         pV2 As D3DVECTOR)
    Dim v   As D3DVECTOR
    
    v.X = pV1.Y * pV2.z - pV1.z * pV2.Y
    v.Y = pV1.z * pV2.X - pV1.X * pV2.z
    v.z = pV1.X * pV2.Y - pV1.Y * pV2.X
    
    pOut = v
    
End Sub

' // Determines the dot product of two 3D vectors.
Public Function D3DXVec3Dot(pV1 As D3DVECTOR, _
                            pV2 As D3DVECTOR) As Single
                            
    D3DXVec3Dot = pV1.X * pV2.X + pV1.Y * pV2.Y + pV1.z * pV2.z
    
End Function

' // Performs a Hermite spline interpolation, using the specified 3D vectors.
Public Sub D3DXVec3Hermite(pOut As D3DVECTOR, _
                           pV1 As D3DVECTOR, _
                           pT1 As D3DVECTOR, _
                           pV2 As D3DVECTOR, _
                           pT2 As D3DVECTOR, _
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
    pOut.z = h1 * pV1.z + h2 * pT1.z + h3 * pV2.z + h4 * pT2.z

End Sub

' // Returns the length of a 3D vector.
Public Function D3DXVec3Length(pV As D3DVECTOR) As Single

    D3DXVec3Length = Sqr(pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z)
    
End Function

' // Returns the square of the length of a 3D vector.
Public Function D3DXVec3LengthSq(pV As D3DVECTOR) As Single

    D3DXVec3LengthSq = pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z
    
End Function

' // Performs a linear interpolation between two 3D vectors.
Public Sub D3DXVec3Lerp(pOut As D3DVECTOR, _
                        pV1 As D3DVECTOR, _
                        pV2 As D3DVECTOR, _
                        ByVal s As Single)
    Dim s1  As Single
    
    s1 = 1 - s
    pOut.X = s1 * pV1.X + s * pV2.X
    pOut.Y = s1 * pV1.Y + s * pV2.Y
    pOut.z = s1 * pV1.z + s * pV2.z
    
End Sub

' // Returns a 3D vector that is made up of the largest components of two 3D vectors.
Public Sub D3DXVec3Maximize(pOut As D3DVECTOR, _
                            pV1 As D3DVECTOR, _
                            pV2 As D3DVECTOR)
    
    If pV1.X > pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y > pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    If pV1.z > pV2.z Then pOut.z = pV1.z Else pOut.z = pV2.z
    
End Sub

' // Returns a 3D vector that is made up of the smallest components of two 3D vectors.
Public Sub D3DXVec3Minimize(pOut As D3DVECTOR, _
                            pV1 As D3DVECTOR, _
                            pV2 As D3DVECTOR)
    
    If pV1.X < pV2.X Then pOut.X = pV1.X Else pOut.X = pV2.X
    If pV1.Y < pV2.Y Then pOut.Y = pV1.Y Else pOut.Y = pV2.Y
    If pV1.z < pV2.z Then pOut.z = pV1.z Else pOut.z = pV2.z
    
End Sub

' // Returns the normalized version of a 3D vector.
Public Sub D3DXVec3Normalize(pOut As D3DVECTOR, _
                             pV As D3DVECTOR)
    Dim norm    As Single
    
    norm = Sqr(pV.X * pV.X + pV.Y * pV.Y + pV.z * pV.z)
    
    If norm = 0! Then
        pOut.X = 0!:    pOut.Y = 0!: pOut.z = 0!
    Else
        pOut.X = pV.X / norm
        pOut.Y = pV.Y / norm
        pOut.z = pV.z / norm
    End If
    
End Sub

' // Projects a 3D vector from object space into screen space.
Public Sub D3DXVec3Project(pOut As D3DVECTOR, _
                           pV As D3DVECTOR, _
                           pViewport As D3DVIEWPORT9, _
                           pProjection As D3DMATRIX, _
                           pView As D3DMATRIX, _
                           pWorld As D3DMATRIX)

    Dim m   As D3DMATRIX
    Dim out As D3DVECTOR

    D3DXMatrixMultiply m, pWorld, pView
    D3DXMatrixMultiply m, m, pProjection
    D3DXVec3TransformCoord out, pV, m
    out.X = pViewport.X + (1! + out.X) * pViewport.Width / 2!
    out.Y = pViewport.Y + (1! - out.Y) * pViewport.Height / 2!
    out.z = pViewport.MinZ + out.z * (pViewport.MaxZ - pViewport.MinZ)
    
    pOut = out
    
End Sub

' // Scales a 3D vector.
Public Sub D3DXVec3Scale(pOut As D3DVECTOR, _
                         pV As D3DVECTOR, _
                         ByVal s As Single)

    pOut.X = pV.X * s
    pOut.Y = pV.Y * s
    pOut.z = pV.z * s
    
End Sub

' // Subtracts two 3D vectors.
Public Sub D3DXVec3Subtract(pOut As D3DVECTOR, _
                            pV1 As D3DVECTOR, _
                            pV2 As D3DVECTOR)
                            
    pOut.X = pV1.X - pV2.X
    pOut.Y = pV1.Y - pV2.Y
    pOut.z = pV1.z - pV2.z
    
End Sub

' // Transforms vector (x, y, z, 1) by a given matrix.
Public Sub D3DXVec3Transform(pOut As D3DVECTOR4, _
                             pV As D3DVECTOR, _
                             pM As D3DMATRIX)
          
    pOut.X = pM.m11 * pV.X + pM.m21 * pV.Y + pM.m31 * pV.z + pM.m41
    pOut.Y = pM.m12 * pV.X + pM.m22 * pV.Y + pM.m32 * pV.z + pM.m42
    pOut.z = pM.m13 * pV.X + pM.m23 * pV.Y + pM.m33 * pV.z + pM.m43
    pOut.w = pM.m14 * pV.X + pM.m24 * pV.Y + pM.m34 * pV.z + pM.m44
    
End Sub

' // Transforms a 3D vector by a given matrix, projecting the result back into w = 1.
Public Sub D3DXVec3TransformCoord(pOut As D3DVECTOR, _
                                  pV As D3DVECTOR, _
                                  pM As D3DMATRIX)
          
    Dim out     As D3DVECTOR
    Dim norm    As Single

    norm = pM.m14 * pV.X + pM.m24 * pV.Y + pM.m34 * pV.z + pM.m44

    out.X = (pM.m11 * pV.X + pM.m21 * pV.Y + pM.m31 * pV.z + pM.m41) / norm
    out.Y = (pM.m12 * pV.X + pM.m22 * pV.Y + pM.m32 * pV.z + pM.m42) / norm
    out.z = (pM.m13 * pV.X + pM.m23 * pV.Y + pM.m33 * pV.z + pM.m43) / norm
    
    pOut = out
    
End Sub

' // Transforms the 3D vector normal by the given matrix.
Public Sub D3DXVec3TransformNormal(pOut As D3DVECTOR, _
                                   pV As D3DVECTOR, _
                                   pM As D3DMATRIX)
          
    Dim v       As D3DVECTOR

    v = pV
    
    pOut.X = pM.m11 * v.X + pM.m21 * v.Y + pM.m31 * v.z
    pOut.Y = pM.m12 * v.X + pM.m22 * v.Y + pM.m32 * v.z
    pOut.z = pM.m13 * v.X + pM.m23 * v.Y + pM.m33 * v.z

End Sub

' // Projects a vector from screen space into object space.
Public Sub D3DXVec3Unproject(pOut As D3DVECTOR, _
                             pV As D3DVECTOR, _
                             pViewport As D3DVIEWPORT9, _
                             pProjection As D3DMATRIX, _
                             pView As D3DMATRIX, _
                             pWorld As D3DMATRIX)

    Dim m   As D3DMATRIX
    Dim out As D3DVECTOR

    D3DXMatrixMultiply m, pWorld, pView
    D3DXMatrixMultiply m, m, pProjection
    D3DXMatrixInverse m, 0!, m
    
    out.X = 2! * (pV.X - pViewport.X) / pViewport.Width - 1!
    out.Y = 1! - 2! * (pV.Y - pViewport.Y) / pViewport.Height
    out.z = (pV.z - pViewport.MinZ) / (pViewport.MaxZ - pViewport.MinZ)
    
    D3DXVec3TransformCoord out, out, m
    pOut = out
    
End Sub



