Attribute VB_Name = "D3DX_MATRICES"
Option Explicit

' // Builds a 3D affine transformation matrix.
Public Sub D3DXMatrixAffineTransformation(pOut As D3DMATRIX, _
                                          ByVal Scaling As Single, _
                                          pRotationCenter As D3DVECTOR, _
                                          pRotation As D3DQUATERNION, _
                                          pTranslation As D3DVECTOR)
    Dim m1  As D3DMATRIX
    Dim m2  As D3DMATRIX
    Dim m3  As D3DMATRIX
    Dim m4  As D3DMATRIX
    Dim m5  As D3DMATRIX
    
    D3DXMatrixScaling m1, Scaling, Scaling, Scaling
    D3DXMatrixTranslation m2, -pRotationCenter.X, -pRotationCenter.Y, -pRotationCenter.z
    D3DXMatrixTranslation m4, pRotationCenter.X, pRotationCenter.Y, pRotationCenter.z
    D3DXMatrixRotationQuaternion m3, pRotation
    D3DXMatrixTranslation m5, pTranslation.X, pTranslation.Y, pTranslation.z
    
    D3DXMatrixMultiply m1, m1, m2
    D3DXMatrixMultiply m1, m1, m3
    D3DXMatrixMultiply m1, m1, m4
    D3DXMatrixMultiply pOut, m1, m5
    
End Sub

' // Builds a 2D affine transformation matrix in the xy plane.
Public Sub D3DXMatrixAffineTransformation2D(pOut As D3DMATRIX, _
                                            ByVal Scaling As Single, _
                                            pRotationCenter As D3DVECTOR2, _
                                            ByVal Rotation As Single, _
                                            pTranslation As D3DVECTOR2)
    Dim m1          As D3DMATRIX
    Dim m2          As D3DMATRIX
    Dim m3          As D3DMATRIX
    Dim m4          As D3DMATRIX
    Dim m5          As D3DMATRIX
    Dim rot         As D3DQUATERNION
    Dim rot_center  As D3DVECTOR
    Dim trans       As D3DVECTOR
    
    rot.w = Cos(Rotation / 2!)
    rot.z = Sin(Rotation / 2!)
    
    rot_center.X = pRotationCenter.X
    rot_center.Y = pRotationCenter.Y
    trans.X = pTranslation.X
    trans.Y = pTranslation.Y

    
    D3DXMatrixScaling m1, Scaling, Scaling, 1!
    D3DXMatrixTranslation m2, -rot_center.X, -rot_center.Y, -rot_center.z
    D3DXMatrixTranslation m4, rot_center.X, rot_center.Y, rot_center.z
    D3DXMatrixRotationQuaternion m3, rot
    D3DXMatrixTranslation m5, trans.X, trans.Y, trans.z

    D3DXMatrixMultiply m1, m1, m2
    D3DXMatrixMultiply m1, m1, m3
    D3DXMatrixMultiply m1, m1, m4
    D3DXMatrixMultiply pOut, m1, m5
    
End Sub

' // Breaks down a general 3D transformation matrix into its scalar, rotational, and translational components.
Public Sub D3DXMatrixDecompose(pOutScale As D3DVECTOR, _
                               pOutRotation As D3DQUATERNION, _
                               pOutTranslation As D3DVECTOR, _
                               pM As D3DMATRIX)
    Dim normalized  As D3DMATRIX
    Dim vec         As D3DVECTOR
    Dim det         As Single
    
    vec.X = pM.m11: vec.Y = pM.m12: vec.z = pM.m13
    pOutScale.X = D3DXVec3Length(vec)
    vec.X = pM.m21: vec.Y = pM.m22: vec.z = pM.m23
    pOutScale.Y = D3DXVec3Length(vec)
    vec.X = pM.m31: vec.Y = pM.m32: vec.z = pM.m33
    pOutScale.z = D3DXVec3Length(vec)
    
    pOutTranslation.X = pM.m41: pOutTranslation.Y = pM.m42: pOutTranslation.z = pM.m43
    
    If ((pOutScale.X = 0!) Or (pOutScale.Y = 0!) Or (pOutScale.z = 0!)) Then
        Err.Raise D3DERR.D3DERR_INVALIDCALL
        Exit Sub
    End If

    normalized.m11 = pM.m11 / pOutScale.X
    normalized.m12 = pM.m12 / pOutScale.X
    normalized.m13 = pM.m13 / pOutScale.X
    normalized.m21 = pM.m21 / pOutScale.Y
    normalized.m22 = pM.m22 / pOutScale.Y
    normalized.m23 = pM.m23 / pOutScale.Y
    normalized.m31 = pM.m31 / pOutScale.z
    normalized.m32 = pM.m32 / pOutScale.z
    normalized.m33 = pM.m33 / pOutScale.z
    normalized.m44 = 1
    
    D3DXQuaternionRotationMatrix pOutRotation, normalized
 
End Sub

' // Returns the determinant of a matrix.
Public Function D3DXMatrixDeterminant(pM As D3DMATRIX) As Single
    Dim minor   As D3DVECTOR4
    Dim v1      As D3DVECTOR4
    Dim v2      As D3DVECTOR4
    Dim v3      As D3DVECTOR4

    v1.X = pM.m11:  v1.Y = pM.m21:  v1.z = pM.m31:  v1.w = pM.m41
    v2.X = pM.m12:  v2.Y = pM.m22:  v2.z = pM.m32:  v2.w = pM.m42
    v3.X = pM.m13:  v3.Y = pM.m23:  v3.z = pM.m33:  v3.w = pM.m43

    D3DXVec4Cross minor, v1, v2, v3
    
    D3DXMatrixDeterminant = -(pM.m14 * minor.X + pM.m24 * minor.Y + pM.m34 * minor.z + pM.m44 * minor.w)

End Function

' // Calculates the inverse of a matrix.
Public Function D3DXMatrixInverse(pOut As D3DMATRIX, _
                                  pDeterminant As Single, _
                                  pM As D3DMATRIX) As Boolean
    Dim a       As Long
    Dim i       As Long
    Dim j       As Long
    Dim out     As D3DMATRIX
    Dim v       As D3DVECTOR4
    Dim vec(2)  As D3DVECTOR4
    Dim det     As Single
    Dim m(3, 3) As Single
    Dim o(3, 3) As Single
    Dim sign    As Single
    
    det = D3DXMatrixDeterminant(pM)
    If det = 0 Then Exit Function
    
    pDeterminant = det
    memcpy m(0, 0), pM, 64
    
    For i = 0 To 3
    
        For j = 0 To 3
            If j <> i Then
                a = j
                If j > i Then a = a - 1
                vec(a).X = m(j, 0)
                vec(a).Y = m(j, 1)
                vec(a).z = m(j, 2)
                vec(a).w = m(j, 3)
            End If
        Next
        
        D3DXVec4Cross v, vec(0), vec(1), vec(2)
        
        sign = IIf(i And 1, -1, 1)
        
        o(0, i) = sign * v.X / det
        o(1, i) = sign * v.Y / det
        o(2, i) = sign * v.z / det
        o(3, i) = sign * v.w / det

    Next
    
    memcpy pOut, o(0, 0), 64
    
    D3DXMatrixInverse = True
    
End Function

' // Creates an identity matrix.
Public Sub D3DXMatrixIdentity(pOut As D3DMATRIX)

    pOut.m11 = 1!:   pOut.m12 = 0!:   pOut.m13 = 0!:   pOut.m14 = 0!
    pOut.m21 = 0!:   pOut.m22 = 1!:   pOut.m23 = 0!:   pOut.m24 = 0!
    pOut.m31 = 0!:   pOut.m32 = 0!:   pOut.m33 = 1!:   pOut.m34 = 0!
    pOut.m41 = 0!:   pOut.m42 = 0!:   pOut.m43 = 0!:   pOut.m44 = 1!
    
End Sub

' // Determines if a matrix is an identity matrix.
Public Function D3DXMatrixIsIdentity(pM As D3DMATRIX) As Boolean

    If Abs(1! - pM.m11) > 0.0001! Then Exit Function
    If Abs(pM.m12) > 0.0001! Then Exit Function
    If Abs(pM.m13) > 0.0001! Then Exit Function
    If Abs(pM.m14) > 0.0001! Then Exit Function
    If Abs(pM.m21) > 0.0001! Then Exit Function
    If Abs(1! - pM.m22) > 0.0001! Then Exit Function
    If Abs(pM.m23) > 0.0001! Then Exit Function
    If Abs(pM.m24) > 0.0001! Then Exit Function
    If Abs(pM.m31) > 0.0001! Then Exit Function
    If Abs(pM.m32) > 0.0001! Then Exit Function
    If Abs(1! - pM.m33) > 0.0001! Then Exit Function
    If Abs(pM.m34) > 0.0001! Then Exit Function
    If Abs(pM.m41) > 0.0001! Then Exit Function
    If Abs(pM.m42) > 0.0001! Then Exit Function
    If Abs(pM.m43) > 0.0001! Then Exit Function
    If Abs(1! - pM.m44) > 0.0001! Then Exit Function
    
    D3DXMatrixIsIdentity = True
    
End Function

' // Builds a left-handed, look-at matrix.
Public Sub D3DXMatrixLookAtLH(pOut As D3DMATRIX, _
                              pEye As D3DVECTOR, _
                              pAt As D3DVECTOR, _
                              pUp As D3DVECTOR)
    Dim zaxis   As D3DVECTOR
    Dim xaxis   As D3DVECTOR
    Dim yaxis   As D3DVECTOR
    Dim vec     As D3DVECTOR
    
    D3DXVec3Subtract vec, pAt, pEye
    D3DXVec3Normalize zaxis, vec
    D3DXVec3Cross vec, pUp, zaxis
    D3DXVec3Normalize xaxis, vec
    D3DXVec3Cross yaxis, zaxis, xaxis
    
    pOut.m11 = xaxis.X
    pOut.m12 = yaxis.X
    pOut.m13 = zaxis.X
    pOut.m14 = 0!
    
    pOut.m21 = xaxis.Y
    pOut.m22 = yaxis.Y
    pOut.m23 = zaxis.Y
    pOut.m24 = 0!
    
    pOut.m31 = xaxis.z
    pOut.m32 = yaxis.z
    pOut.m33 = zaxis.z
    pOut.m34 = 0!
    
    pOut.m41 = -D3DXVec3Dot(xaxis, pEye)
    pOut.m42 = -D3DXVec3Dot(yaxis, pEye)
    pOut.m43 = -D3DXVec3Dot(zaxis, pEye)
    pOut.m44 = 1!
    
End Sub

' // Builds a right-handed, look-at matrix.
Public Sub D3DXMatrixLookAtRH(pOut As D3DMATRIX, _
                              pEye As D3DVECTOR, _
                              pAt As D3DVECTOR, _
                              pUp As D3DVECTOR)
    Dim zaxis   As D3DVECTOR
    Dim xaxis   As D3DVECTOR
    Dim yaxis   As D3DVECTOR
    Dim vec     As D3DVECTOR
    
    D3DXVec3Subtract vec, pEye, pAt
    D3DXVec3Normalize zaxis, vec
    D3DXVec3Cross vec, pUp, zaxis
    D3DXVec3Normalize xaxis, vec
    D3DXVec3Cross yaxis, zaxis, xaxis
    
    pOut.m11 = xaxis.X
    pOut.m12 = yaxis.X
    pOut.m13 = zaxis.X
    pOut.m14 = 0!
    
    pOut.m21 = xaxis.Y
    pOut.m22 = yaxis.Y
    pOut.m23 = zaxis.Y
    pOut.m24 = 0!
    
    pOut.m31 = xaxis.z
    pOut.m32 = yaxis.z
    pOut.m33 = zaxis.z
    pOut.m34 = 0!
    
    pOut.m41 = -D3DXVec3Dot(xaxis, pEye)
    pOut.m42 = -D3DXVec3Dot(yaxis, pEye)
    pOut.m43 = -D3DXVec3Dot(zaxis, pEye)
    pOut.m44 = 1!

End Sub

' // Calculates the transposed product of two matrices.
Public Sub D3DXMatrixMultiplyTranspose(pOut As D3DMATRIX, _
                                       pM1 As D3DMATRIX, _
                                       pM2 As D3DMATRIX)
    
    D3DXMatrixMultiply pOut, pM1, pM2
    D3DXMatrixTranspose pOut, pOut
    
End Sub

' // Builds a left-handed orthographic projection matrix.
Public Sub D3DXMatrixOrthoLH(pOut As D3DMATRIX, _
                             ByVal w As Single, _
                             ByVal h As Single, _
                             ByVal zn As Single, _
                             ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! / w
    pOut.m22 = 2! / h
    pOut.m33 = 1! / (zf - zn)
    pOut.m43 = zn / (zn - zf)

End Sub

' // Builds a customized, left-handed orthographic projection matrix.
Public Sub D3DXMatrixOrthoOffCenterLH(pOut As D3DMATRIX, _
                                      ByVal l As Single, _
                                      ByVal r As Single, _
                                      ByVal b As Single, _
                                      ByVal t As Single, _
                                      ByVal zn As Single, _
                                      ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! / (r - l)
    pOut.m22 = 2! / (t - b)
    pOut.m33 = 1! / (zf - zn)
    pOut.m41 = -1! - 2! * l / (r - l)
    pOut.m42 = 1! + 2! * t / (b - t)
    pOut.m43 = zn / (zn - zf)
    
End Sub

' // Builds a customized, right-handed orthographic projection matrix.
Public Sub D3DXMatrixOrthoOffCenterRH(pOut As D3DMATRIX, _
                                      ByVal l As Single, _
                                      ByVal r As Single, _
                                      ByVal b As Single, _
                                      ByVal t As Single, _
                                      ByVal zn As Single, _
                                      ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! / (r - l)
    pOut.m22 = 2! / (t - b)
    pOut.m33 = 1! / (zn - zf)
    pOut.m41 = -1! - 2! * l / (r - l)
    pOut.m42 = 1! + 2! * t / (b - t)
    pOut.m43 = zn / (zn - zf)
    
End Sub

' // Builds a right-handed orthographic projection matrix.
Public Sub D3DXMatrixOrthoRH(pOut As D3DMATRIX, _
                             ByVal w As Single, _
                             ByVal h As Single, _
                             ByVal zn As Single, _
                             ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! / w
    pOut.m22 = 2! / h
    pOut.m33 = 1! / (zn - zf)
    pOut.m43 = zn / (zn - zf)

End Sub

' // Builds a left-handed perspective projection matrix based on a field of view.
Public Sub D3DXMatrixPerspectiveFovLH(pOut As D3DMATRIX, _
                                      ByVal fovy As Single, _
                                      ByVal aspect As Single, _
                                      ByVal zn As Single, _
                                      ByVal zf As Single)
    
    D3DXMatrixIdentity pOut
    
    pOut.m11 = 1! / (aspect * Tan(fovy / 2!))
    pOut.m22 = 1! / Tan(fovy / 2!)
    pOut.m33 = zf / (zf - zn)
    pOut.m34 = 1!
    pOut.m43 = (zf * zn) / (zn - zf)
    pOut.m44 = 0!
    
End Sub

' // Builds a right-handed perspective projection matrix based on a field of view.
Public Sub D3DXMatrixPerspectiveFovRH(pOut As D3DMATRIX, _
                                      ByVal fovy As Single, _
                                      ByVal aspect As Single, _
                                      ByVal zn As Single, _
                                      ByVal zf As Single)
    
    D3DXMatrixIdentity pOut
    
    pOut.m11 = 1! / (aspect * Tan(fovy / 2!))
    pOut.m22 = 1! / Tan(fovy / 2!)
    pOut.m33 = zf / (zn - zf)
    pOut.m34 = -1!
    pOut.m43 = (zf * zn) / (zn - zf)
    pOut.m44 = 0!
    
End Sub

' // Builds a left-handed perspective projection matrix
Public Sub D3DXMatrixPerspectiveLH(pOut As D3DMATRIX, _
                                   ByVal w As Single, _
                                   ByVal h As Single, _
                                   ByVal zn As Single, _
                                   ByVal zf As Single)
            
    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! * zn / w
    pOut.m22 = 2! * zn / h
    pOut.m33 = zf / (zf - zn)
    pOut.m43 = (zn * zf) / (zn - zf)
    pOut.m34 = 1!
    pOut.m44 = 0!
    
End Sub

' // Builds a customized, left-handed perspective projection matrix.
Public Sub D3DXMatrixPerspectiveOffCenterLH(pOut As D3DMATRIX, _
                                            ByVal l As Single, _
                                            ByVal r As Single, _
                                            ByVal b As Single, _
                                            ByVal t As Single, _
                                            ByVal zn As Single, _
                                            ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! * zn / (r - l)
    pOut.m22 = -2! * zn / (b - t)
    pOut.m31 = -1! - 2! * l / (r - l)
    pOut.m32 = 1! + 2! * t / (b - t)
    pOut.m33 = -zf / (zn - zf)
    pOut.m43 = (zn * zf) / (zn - zf)
    pOut.m34 = 1!
    pOut.m44 = 0!

End Sub

' // Builds a customized, right-handed perspective projection matrix.
Public Sub D3DXMatrixPerspectiveOffCenterRH(pOut As D3DMATRIX, _
                                            ByVal l As Single, _
                                            ByVal r As Single, _
                                            ByVal b As Single, _
                                            ByVal t As Single, _
                                            ByVal zn As Single, _
                                            ByVal zf As Single)

    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! * zn / (r - l)
    pOut.m22 = -2! * zn / (b - t)
    pOut.m31 = 1! + 2! * l / (r - l)
    pOut.m32 = -1! - 2! * t / (b - t)
    pOut.m33 = zf / (zn - zf)
    pOut.m43 = (zn * zf) / (zn - zf)
    pOut.m34 = -1!
    pOut.m44 = 0!

End Sub

' // Builds a right-handed perspective projection matrix
Public Sub D3DXMatrixPerspectiveRH(pOut As D3DMATRIX, _
                                   ByVal w As Single, _
                                   ByVal h As Single, _
                                   ByVal zn As Single, _
                                   ByVal zf As Single)
            
    D3DXMatrixIdentity pOut
    
    pOut.m11 = 2! * zn / w
    pOut.m22 = 2! * zn / h
    pOut.m33 = zf / (zn - zf)
    pOut.m43 = (zn * zf) / (zn - zf)
    pOut.m34 = -1!
    pOut.m44 = 0!
    
End Sub

' // Builds a matrix that reflects the coordinate system about a plane.
Public Sub D3DXMatrixReflect(pOut As D3DMATRIX, _
                             pPlane As D3DPLANE)
    Dim nPlane  As D3DPLANE
    
    D3DXPlaneNormalize nPlane, pPlane
    D3DXMatrixIdentity pOut
    
    pOut.m11 = 1! - 2! * nPlane.a * nPlane.a
    pOut.m12 = -2! * nPlane.a * nPlane.b
    pOut.m13 = -2! * nPlane.a * nPlane.c
    pOut.m21 = -2! * nPlane.a * nPlane.b
    pOut.m22 = 1! - 2! * nPlane.b * nPlane.b
    pOut.m23 = -2! * nPlane.b * nPlane.c
    pOut.m31 = -2! * nPlane.c * nPlane.a
    pOut.m32 = -2! * nPlane.c * nPlane.b
    pOut.m33 = 1! - 2! * nPlane.c * nPlane.c
    pOut.m41 = -2! * nPlane.d * nPlane.a
    pOut.m42 = -2! * nPlane.d * nPlane.b
    pOut.m43 = -2! * nPlane.d * nPlane.c
    
End Sub

' // Builds a matrix that rotates around an arbitrary axis.
Public Sub D3DXMatrixRotationAxis(pOut As D3DMATRIX, _
                                  pV As D3DVECTOR, _
                                  ByVal angle As Single)
                           
    Dim v   As D3DVECTOR
    Dim c   As Single
    Dim s   As Single
    
    D3DXMatrixIdentity pOut
    D3DXVec3Normalize v, pV
    c = Cos(angle)
    s = Sin(angle)
    
    pOut.m11 = (1! - c) * v.X * v.X + c
    pOut.m21 = (1! - c) * v.X * v.Y - s * v.z
    pOut.m31 = (1! - c) * v.X * v.z + s * v.Y
    pOut.m12 = (1! - c) * v.Y * v.X + s * v.z
    pOut.m22 = (1! - c) * v.Y * v.Y + c
    pOut.m32 = (1! - c) * v.Y * v.z - s * v.X
    pOut.m13 = (1! - c) * v.z * v.X - s * v.Y
    pOut.m23 = (1! - c) * v.z * v.Y + s * v.X
    pOut.m33 = (1! - c) * v.z * v.z + c
    
End Sub

' // Returns the matrix transpose of a matrix.
Public Sub D3DXMatrixTranspose(pOut As D3DMATRIX, _
                               pM As D3DMATRIX)
    Dim out As D3DMATRIX
    
    out.m11 = pM.m11:  out.m12 = pM.m21:  out.m13 = pM.m31:  out.m14 = pM.m41
    out.m21 = pM.m12:  out.m22 = pM.m22:  out.m23 = pM.m32:  out.m24 = pM.m42
    out.m31 = pM.m13:  out.m32 = pM.m23:  out.m33 = pM.m33:  out.m34 = pM.m43
    out.m41 = pM.m14:  out.m42 = pM.m24:  out.m43 = pM.m34:  out.m44 = pM.m44
    
    pOut = out
    
End Sub

' // Builds a matrix that scales along the x-axis, the y-axis, and the z-axis.
Public Sub D3DXMatrixScaling(pOut As D3DMATRIX, _
                             ByVal sx As Single, _
                             ByVal sy As Single, _
                             ByVal sz As Single)
                             
    D3DXMatrixIdentity pOut
    pOut.m11 = sx:  pOut.m22 = sy:  pOut.m33 = sz
    
End Sub

' // Builds a matrix using the specified offsets.
Public Sub D3DXMatrixTranslation(pOut As D3DMATRIX, _
                                 ByVal X As Single, _
                                 ByVal Y As Single, _
                                 ByVal z As Single)
                                 
    D3DXMatrixIdentity pOut
    pOut.m41 = X:   pOut.m42 = Y:   pOut.m43 = z
    
End Sub


' // Builds a rotation matrix from a quaternion.
Public Sub D3DXMatrixRotationQuaternion(pOut As D3DMATRIX, _
                                        pQ As D3DQUATERNION)
                                        
    pOut.m11 = 1! - 2! * (pQ.Y * pQ.Y + pQ.z * pQ.z)
    pOut.m12 = 2! * (pQ.X * pQ.Y + pQ.z * pQ.w)
    pOut.m13 = 2! * (pQ.X * pQ.z - pQ.Y * pQ.w)
    pOut.m14 = 0!
    pOut.m21 = 2! * (pQ.X * pQ.Y - pQ.z * pQ.w)
    pOut.m22 = 1! - 2! * (pQ.X * pQ.X + pQ.z * pQ.z)
    pOut.m23 = 2! * (pQ.Y * pQ.z + pQ.X * pQ.w)
    pOut.m24 = 0!
    pOut.m31 = 2! * (pQ.X * pQ.z + pQ.Y * pQ.w)
    pOut.m32 = 2! * (pQ.Y * pQ.z - pQ.X * pQ.w)
    pOut.m33 = 1! - 2! * (pQ.X * pQ.X + pQ.Y * pQ.Y)
    pOut.m34 = 0!
    pOut.m41 = 0!
    pOut.m42 = 0!
    pOut.m43 = 0!
    pOut.m44 = 1!
    
End Sub

' // Builds a matrix that rotates around the x-axis.
Public Sub D3DXMatrixRotationX(pOut As D3DMATRIX, _
                               angle As Single)
    Dim s   As Single
    Dim c   As Single

    D3DXMatrixIdentity pOut
    s = Sin(angle)
    c = Cos(angle)
    
    pOut.m22 = c
    pOut.m33 = c
    pOut.m23 = s
    pOut.m32 = -s

End Sub

' // Builds a matrix that rotates around the y-axis.
Public Sub D3DXMatrixRotationY(pOut As D3DMATRIX, _
                               angle As Single)
    Dim s   As Single
    Dim c   As Single

    D3DXMatrixIdentity pOut
    s = Sin(angle)
    c = Cos(angle)

    pOut.m11 = c
    pOut.m33 = c
    pOut.m13 = -s
    pOut.m31 = s

End Sub

' // Builds a matrix that rotates around the z-axis.
Public Sub D3DXMatrixRotationZ(pOut As D3DMATRIX, _
                               angle As Single)
    Dim s   As Single
    Dim c   As Single

    D3DXMatrixIdentity pOut
    s = Sin(angle)
    c = Cos(angle)
    
    pOut.m11 = c
    pOut.m22 = c
    pOut.m12 = s
    pOut.m21 = -s

End Sub

' // Builds a matrix with a specified yaw, pitch, and roll.
Public Sub D3DXMatrixRotationYawPitchRoll(pOut As D3DMATRIX, _
                                          ByVal yaw As Single, _
                                          ByVal pitch As Single, _
                                          ByVal roll As Single)
    Dim m   As D3DMATRIX
    
    D3DXMatrixIdentity pOut
    D3DXMatrixRotationZ m, roll
    D3DXMatrixMultiply pOut, pOut, m
    D3DXMatrixRotationX m, pitch
    D3DXMatrixMultiply pOut, pOut, m
    D3DXMatrixRotationY m, yaw
    D3DXMatrixMultiply pOut, pOut, m
    
End Sub

' // Builds a matrix that flattens geometry into a plane.
Public Sub D3DXMatrixShadow(pOut As D3DMATRIX, _
                            pLight As D3DVECTOR4, _
                            pPlane As D3DPLANE)
    Dim nPlane  As D3DPLANE
    Dim dot     As Single

    D3DXPlaneNormalize nPlane, pPlane
    dot = D3DXPlaneDot(nPlane, pLight)
    pOut.m11 = dot - nPlane.a * pLight.X
    pOut.m12 = -nPlane.a * pLight.Y
    pOut.m13 = -nPlane.a * pLight.z
    pOut.m14 = -nPlane.a * pLight.w
    pOut.m21 = -nPlane.b * pLight.X
    pOut.m22 = dot - nPlane.b * pLight.Y
    pOut.m23 = -nPlane.b * pLight.z
    pOut.m24 = -nPlane.b * pLight.w
    pOut.m31 = -nPlane.c * pLight.X
    pOut.m32 = -nPlane.c * pLight.Y
    pOut.m33 = dot - nPlane.c * pLight.z
    pOut.m34 = -nPlane.c * pLight.w
    pOut.m41 = -nPlane.d * pLight.X
    pOut.m42 = -nPlane.d * pLight.Y
    pOut.m43 = -nPlane.d * pLight.z
    pOut.m44 = dot - nPlane.d * pLight.w

End Sub

' // Builds a transformation matrix.
Public Sub D3DXMatrixTransformation(pOut As D3DMATRIX, _
                                    pScalingCenter As D3DVECTOR, _
                                    pScalingRotation As D3DQUATERNION, _
                                    pScaling As D3DVECTOR, _
                                    pRotationCenter As D3DVECTOR, _
                                    pRotation As D3DQUATERNION, _
                                    pTranslation As D3DVECTOR)
    Dim m1  As D3DMATRIX
    Dim m2  As D3DMATRIX
    Dim m3  As D3DMATRIX

    m3.m11 = pScaling.X
    m3.m22 = pScaling.Y
    m3.m33 = pScaling.z
    m3.m44 = 1!
    
    D3DXMatrixRotationQuaternion m2, pScalingRotation
    D3DXMatrixTranspose m1, m2
    D3DXMatrixIdentity pOut
    
    pOut.m41 = -pScalingCenter.X
    pOut.m42 = -pScalingCenter.Y
    pOut.m43 = -pScalingCenter.z
    
    D3DXMatrixMultiply pOut, pOut, m1
    D3DXMatrixMultiply pOut, pOut, m3
    D3DXMatrixMultiply pOut, pOut, m2

    pOut.m41 = pOut.m41 + pScalingCenter.X - pRotationCenter.X
    pOut.m42 = pOut.m42 + pScalingCenter.Y - pRotationCenter.Y
    pOut.m43 = pOut.m43 + pScalingCenter.z - pRotationCenter.z
    
    D3DXMatrixRotationQuaternion m2, pRotation
    D3DXMatrixMultiply pOut, pOut, m2
    
    pOut.m41 = pOut.m41 + pRotationCenter.X + pTranslation.X
    pOut.m42 = pOut.m42 + pRotationCenter.Y + pTranslation.Y
    pOut.m43 = pOut.m43 + pRotationCenter.z + pTranslation.z

End Sub

' // Builds a 2D transformation matrix that represents transformations in the xy plane.
Public Sub D3DXMatrixTransformation2D(pOut As D3DMATRIX, _
                                      pScalingCenter As D3DVECTOR2, _
                                      ByVal ScalingRotation As Single, _
                                      pScaling As D3DVECTOR2, _
                                      pRotationCenter As D3DVECTOR2, _
                                      ByVal Rotation As Single, _
                                      pTranslation As D3DVECTOR2)
    Dim m1  As D3DMATRIX
    Dim m2  As D3DMATRIX
    Dim m3  As D3DMATRIX

    m3.m11 = pScaling.X
    m3.m22 = pScaling.Y
    m3.m33 = 1!
    m3.m44 = 1!
    
    D3DXMatrixRotationZ m2, ScalingRotation
    D3DXMatrixTranspose m1, m2
    D3DXMatrixIdentity pOut
    
    pOut.m41 = -pScalingCenter.X
    pOut.m42 = -pScalingCenter.Y
        
    D3DXMatrixMultiply pOut, pOut, m1
    D3DXMatrixMultiply pOut, pOut, m3
    D3DXMatrixMultiply pOut, pOut, m2
    
    pOut.m41 = pScalingCenter.X + pOut.m41
    pOut.m42 = pScalingCenter.Y + pOut.m42
        
    If Rotation <> 0 Then
    
        D3DXMatrixRotationZ m2, Rotation
        
        pOut.m41 = pOut.m41 - pRotationCenter.X
        pOut.m42 = pOut.m42 - pRotationCenter.Y
        
        D3DXMatrixMultiply pOut, pOut, m2
        
        pOut.m41 = pRotationCenter.X + pOut.m41
        pOut.m42 = pRotationCenter.Y + pOut.m42
        
    End If
    
    pOut.m41 = pTranslation.X + pOut.m41
    pOut.m42 = pTranslation.Y + pOut.m42
    
End Sub

' // Determines the product of two matrices.
Public Sub D3DXMatrixMultiply(pOut As D3DMATRIX, _
                              pM1 As D3DMATRIX, _
                              pM2 As D3DMATRIX)
    Dim out As D3DMATRIX
    
    out.m11 = pM1.m11 * pM2.m11 + pM1.m12 * pM2.m21 + pM1.m13 * pM2.m31 + pM1.m14 * pM2.m41
    out.m12 = pM1.m11 * pM2.m12 + pM1.m12 * pM2.m22 + pM1.m13 * pM2.m32 + pM1.m14 * pM2.m42
    out.m13 = pM1.m11 * pM2.m13 + pM1.m12 * pM2.m23 + pM1.m13 * pM2.m33 + pM1.m14 * pM2.m43
    out.m14 = pM1.m11 * pM2.m14 + pM1.m12 * pM2.m24 + pM1.m13 * pM2.m34 + pM1.m14 * pM2.m44
    
    out.m21 = pM1.m21 * pM2.m11 + pM1.m22 * pM2.m21 + pM1.m23 * pM2.m31 + pM1.m24 * pM2.m41
    out.m22 = pM1.m21 * pM2.m12 + pM1.m22 * pM2.m22 + pM1.m23 * pM2.m32 + pM1.m24 * pM2.m42
    out.m23 = pM1.m21 * pM2.m13 + pM1.m22 * pM2.m23 + pM1.m23 * pM2.m33 + pM1.m24 * pM2.m43
    out.m24 = pM1.m21 * pM2.m14 + pM1.m22 * pM2.m24 + pM1.m23 * pM2.m34 + pM1.m24 * pM2.m44
  
    out.m31 = pM1.m31 * pM2.m11 + pM1.m32 * pM2.m21 + pM1.m33 * pM2.m31 + pM1.m34 * pM2.m41
    out.m32 = pM1.m31 * pM2.m12 + pM1.m32 * pM2.m22 + pM1.m33 * pM2.m32 + pM1.m34 * pM2.m42
    out.m33 = pM1.m31 * pM2.m13 + pM1.m32 * pM2.m23 + pM1.m33 * pM2.m33 + pM1.m34 * pM2.m43
    out.m34 = pM1.m31 * pM2.m14 + pM1.m32 * pM2.m24 + pM1.m33 * pM2.m34 + pM1.m34 * pM2.m44
    
    out.m41 = pM1.m41 * pM2.m11 + pM1.m42 * pM2.m21 + pM1.m43 * pM2.m31 + pM1.m44 * pM2.m41
    out.m42 = pM1.m41 * pM2.m12 + pM1.m42 * pM2.m22 + pM1.m43 * pM2.m32 + pM1.m44 * pM2.m42
    out.m43 = pM1.m41 * pM2.m13 + pM1.m42 * pM2.m23 + pM1.m43 * pM2.m33 + pM1.m44 * pM2.m43
    out.m44 = pM1.m41 * pM2.m14 + pM1.m42 * pM2.m24 + pM1.m43 * pM2.m34 + pM1.m44 * pM2.m44
    
    pOut = out
    
End Sub


