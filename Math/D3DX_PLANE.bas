Attribute VB_Name = "D3DX_PLANE"
Option Explicit

' // Normalizes the plane coefficients so that the plane normal has unit length.
Public Sub D3DXPlaneNormalize(pOut As D3DPLANE, _
                              pP As D3DPLANE)
    Dim norm    As Single
    
    norm = Sqr(pP.a * pP.a + pP.b * pP.b + pP.c * pP.c)
    
    If norm = 0! Then
    
        pOut.a = 0!
        pOut.b = 0!
        pOut.c = 0!
        pOut.d = 0!
        
    Else
    
        pOut.a = pP.a / norm
        pOut.b = pP.b / norm
        pOut.c = pP.c / norm
        pOut.d = pP.d / norm
        
    End If
    
End Sub

' // Computes the dot product of a plane and a 4D vector.
Public Function D3DXPlaneDot(pP As D3DPLANE, _
                             pV As D3DVECTOR4) As Single
                             
    D3DXPlaneDot = pP.a * pV.X + pP.b * pV.Y + pP.c * pV.z + pP.d * pV.w
    
End Function

' // Computes the dot product of a plane and a 3D vector. The w parameter of the vector is assumed to be 1.
Public Function D3DXPlaneDotCoord(pP As D3DPLANE, _
                                  pV As D3DVECTOR) As Single

    D3DXPlaneDotCoord = pP.a * pV.X + pP.b * pV.Y + pP.c * pV.z + pP.d
    
End Function

' // Computes the dot product of a plane and a 3D vector. The w parameter of the vector is assumed to be 0.
Public Function D3DXPlaneDotNormal(pP As D3DPLANE, _
                                   pV As D3DVECTOR) As Single

    D3DXPlaneDotNormal = pP.a * pV.X + pP.b * pV.Y + pP.c * pV.z
    
End Function

' // Constructs a plane from a point and a normal.
Public Sub D3DXPlaneFromPointNormal(pOut As D3DPLANE, _
                                    pPoint As D3DVECTOR, _
                                    pNormal As D3DVECTOR)
                                    
    pOut.a = pNormal.X
    pOut.b = pNormal.Y
    pOut.c = pNormal.z
    pOut.d = -D3DXVec3Dot(pPoint, pNormal)
    
End Sub

' // Constructs a plane from three points.
Public Sub D3DXPlaneFromPoints(pOut As D3DPLANE, _
                               pV1 As D3DVECTOR, _
                               pV2 As D3DVECTOR, _
                               pV3 As D3DVECTOR)
    Dim edge1   As D3DVECTOR
    Dim edge2   As D3DVECTOR
    Dim normal  As D3DVECTOR
    Dim Nnormal As D3DVECTOR

    D3DXVec3Subtract edge1, pV2, pV1
    D3DXVec3Subtract edge2, pV3, pV1
    D3DXVec3Cross normal, edge1, edge2
    D3DXVec3Normalize Nnormal, normal
    D3DXPlaneFromPointNormal pOut, pV1, Nnormal
    
End Sub

' // Finds the intersection between a plane and a line.
Public Function D3DXPlaneIntersectLine(pOut As D3DVECTOR, _
                                       pP As D3DPLANE, _
                                       pV1 As D3DVECTOR, _
                                       pV2 As D3DVECTOR) As Boolean
    Dim direction   As D3DVECTOR
    Dim normal      As D3DVECTOR
    Dim dot         As Single
    Dim temp        As Single
    
    normal.X = pP.a
    normal.Y = pP.b
    normal.z = pP.c
    
    direction.X = pV2.X - pV1.X
    direction.Y = pV2.Y - pV1.Y
    direction.z = pV2.z - pV1.z
    
    dot = D3DXVec3Dot(normal, direction)
    
    If dot = 0 Then Exit Function
    
    temp = (pP.d + D3DXVec3Dot(normal, pV1)) / dot
    
    pOut.X = pV1.X - temp * direction.X
    pOut.Y = pV1.Y - temp * direction.Y
    pOut.z = pV1.z - temp * direction.z
    
    D3DXPlaneIntersectLine = True
    
End Function
 
 ' // Scale the plane with the given scaling factor.
Public Sub D3DXPlaneScale(pOut As D3DPLANE, _
                          pP As D3DPLANE, _
                          ByVal s As Single)
                          
    pOut.a = pP.a * s
    pOut.b = pP.b * s
    pOut.c = pP.c * s
    pOut.d = pP.d * s
    
End Sub

' // Transforms a plane by a matrix. The input matrix is the inverse transpose of the actual transformation.
Public Sub D3DXPlaneTransform(pOut As D3DPLANE, _
                              pP As D3DPLANE, _
                              pM As D3DMATRIX)
    Dim plane   As D3DPLANE
    
    plane = pP
    
    pOut.a = pM.m11 * plane.a + pM.m21 * plane.b + pM.m31 * plane.c + pM.m41 * plane.d
    pOut.b = pM.m12 * plane.a + pM.m22 * plane.b + pM.m32 * plane.c + pM.m42 * plane.d
    pOut.c = pM.m13 * plane.a + pM.m23 * plane.b + pM.m33 * plane.c + pM.m43 * plane.d
    pOut.d = pM.m14 * plane.a + pM.m24 * plane.b + pM.m34 * plane.c + pM.m44 * plane.d

End Sub

