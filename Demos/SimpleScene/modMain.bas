Attribute VB_Name = "modMain"
' //
' // Common functions
' //

Option Explicit

' // Fast 3D vector creation
Public Function vec3( _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal Z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.Z = Z
End Function

' // Fast 2D vector creation
Public Function vec2( _
                ByVal X As Single, _
                ByVal Y As Single) As D3DVECTOR2
    vec2.X = X: vec2.Y = Y
End Function

' // Fast color creation
Public Function color( _
                ByVal r As Single, _
                ByVal g As Single, _
                ByVal b As Single, _
                Optional ByVal a As Single = 1) As D3DCOLORVALUE
    
    color.r = r
    color.g = g
    color.b = b
    color.a = a
    
End Function

' // Check if a ray intersects a triangle
Public Function IsIntersected( _
                ByRef tp1 As D3DVECTOR, _
                ByRef tp2 As D3DVECTOR, _
                ByRef tp3 As D3DVECTOR, _
                ByRef trfrom As D3DVECTOR, _
                ByRef trdir As D3DVECTOR) As Boolean
    Dim tEdge(1)    As D3DVECTOR
    Dim tVec(2)     As D3DVECTOR
    Dim fDet        As Single
    Dim fU          As Single
    Dim fV          As Single
    
    D3DXVec3Subtract tEdge(0), tp2, tp1
    D3DXVec3Subtract tEdge(1), tp3, tp1
    D3DXVec3Cross tVec(0), trdir, tEdge(1)
    
    fDet = D3DXVec3Dot(tEdge(0), tVec(0))
    
    If fDet < 0.00001! Then Exit Function
    
    D3DXVec3Subtract tVec(1), trfrom, tp1
    
    fU = D3DXVec3Dot(tVec(0), tVec(1))
    
    If fU < 0 Or fU > fDet Then Exit Function
    
    D3DXVec3Cross tVec(2), tVec(1), tEdge(0)
    
    fV = D3DXVec3Dot(trdir, tVec(2))
    
    If fV < 0 Or fV + fU > fDet Then Exit Function
    
    IsIntersected = True
    
End Function

' // Euler angles from matrix
Public Function MatrixToEuler( _
                ByRef tMtx As D3DMATRIX) As D3DVECTOR
    Dim fCy As Single
    
    fCy = Sqr(tMtx.m33 * tMtx.m33 + tMtx.m31 * tMtx.m31)
    
    If fCy > 1.175494351E-38 * 10 Then
    
        MatrixToEuler.X = Atan2(-tMtx.m32, fCy)
        MatrixToEuler.Y = Atan2(tMtx.m31, tMtx.m33)
        MatrixToEuler.Z = Atan2(tMtx.m12, tMtx.m22)
    
    Else
    
        MatrixToEuler.X = Atan2(-tMtx.m32, fCy)
        MatrixToEuler.Y = 0
        MatrixToEuler.Z = Atan2(-tMtx.m21, tMtx.m11)
        
    End If
    
End Function

