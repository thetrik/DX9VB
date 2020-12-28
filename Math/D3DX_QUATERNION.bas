Attribute VB_Name = "D3DX_QUATERNION"
Option Explicit

' // Returns a quaternion in barycentric coordinates.
Public Sub D3DXQuaternionBaryCentric(pOut As D3DQUATERNION, _
                                     pQ1 As D3DQUATERNION, _
                                     pQ2 As D3DQUATERNION, _
                                     pQ3 As D3DQUATERNION, _
                                     ByVal f As Single, _
                                     ByVal g As Single)
    Dim temp1   As D3DQUATERNION
    Dim temp2   As D3DQUATERNION
    
    D3DXQuaternionSlerp temp1, pQ1, pQ2, f + g
    D3DXQuaternionSlerp temp2, pQ1, pQ3, f + g
    
    D3DXQuaternionSlerp pOut, temp1, temp2, g / (f + g)

End Sub

' // Calculates the exponential.
Public Sub D3DXQuaternionExp(pOut As D3DQUATERNION, _
                             pQ As D3DQUATERNION)
    Dim norm    As Single

    norm = Sqr(pQ.X * pQ.X + pQ.Y * pQ.Y + pQ.z * pQ.z)
    
    If (norm <> 0) Then
    
        pOut.X = Sin(norm) * pQ.X / norm
        pOut.Y = Sin(norm) * pQ.Y / norm
        pOut.z = Sin(norm) * pQ.z / norm
        pOut.w = Cos(norm)
        
    Else
    
        pOut.X = 0!
        pOut.Y = 0!
        pOut.z = 0!
        pOut.w = 1!
        
    End If

End Sub

' // Conjugates and renormalizes a quaternion.
Public Sub D3DXQuaternionInverse(pOut As D3DQUATERNION, _
                                 pQ As D3DQUATERNION)
    Dim norm    As Single

    norm = D3DXQuaternionLengthSq(pQ)

    pOut.X = -pQ.X / norm
    pOut.Y = -pQ.Y / norm
    pOut.z = -pQ.z / norm
    pOut.w = pQ.w / norm

End Sub

' // Returns the dot product of two quaternions.
Public Function D3DXQuaternionDot(pQ1 As D3DQUATERNION, _
                                  pQ2 As D3DQUATERNION) As Single
                     
    D3DXQuaternionDot = pQ1.X * pQ2.X + pQ1.Y * pQ2.Y + pQ1.z * pQ2.z + pQ1.w * pQ2.w
    
End Function

' // Determines if a quaternion is an identity quaternion.
Public Function D3DXQuaternionIsIdentity(pQ As D3DQUATERNION) As Single
    
    D3DXQuaternionIsIdentity = pQ.X = 0 And pQ.Y = 0 And pQ.z = 0 And pQ.w = 1!
    
End Function

' // Returns the length of a quaternion.
Public Function D3DXQuaternionLength(pQ As D3DQUATERNION) As Single
    
    D3DXQuaternionLength = Sqr(pQ.X * pQ.X + pQ.Y * pQ.Y + pQ.z * pQ.z + pQ.w * pQ.w)
    
End Function

' // Returns the square of the length of a quaternion.
Public Function D3DXQuaternionLengthSq(pQ As D3DQUATERNION) As Single
    
    D3DXQuaternionLengthSq = pQ.X * pQ.X + pQ.Y * pQ.Y + pQ.z * pQ.z + pQ.w * pQ.w
    
End Function

' // Calculates the natural logarithm.
Public Sub D3DXQuaternionLn(pOut As D3DQUATERNION, _
                            pQ As D3DQUATERNION)
    Dim v3  As Single
    Dim v8  As Single
    Dim v5  As Boolean
    Dim v7  As Single
    
    If pQ.w < 1! Then
        v3 = acos(pQ.w)
        v8 = Sin(v3)
        
        v5 = Not (v8 >= -0.00000011920929 And v8 <= 0.00000011920929)
    End If

    If Not v5 Then
        pOut.X = pQ.X
        pOut.Y = pQ.Y
        pOut.z = pQ.z
    Else
        v7 = v3 / v8
        pOut.X = v7 * pQ.X
        pOut.Y = v7 * pQ.Y
        pOut.z = v7 * pQ.z
    End If
    
    pOut.w = 0
    
End Sub

' // Multiplies two quaternions.
Public Sub D3DXQuaternionMultiply(pOut As D3DQUATERNION, _
                                  pQ1 As D3DQUATERNION, _
                                  pQ2 As D3DQUATERNION)

    Dim out As D3DQUATERNION
    
    out.X = pQ2.w * pQ1.X + pQ2.X * pQ1.w + pQ2.Y * pQ1.z - pQ2.z * pQ1.Y
    out.Y = pQ2.w * pQ1.Y - pQ2.X * pQ1.z + pQ2.Y * pQ1.w + pQ2.z * pQ1.X
    out.z = pQ2.w * pQ1.z + pQ2.X * pQ1.Y - pQ2.Y * pQ1.X + pQ2.z * pQ1.w
    out.w = pQ2.w * pQ1.w - pQ2.X * pQ1.X - pQ2.Y * pQ1.Y - pQ2.z * pQ1.z
    
    pOut = out
    
End Sub

' // Computes a unit length quaternion.
Public Sub D3DXQuaternionNormalize(pOut As D3DQUATERNION, _
                                   pQ As D3DQUATERNION)
    
    Dim norm    As Single
    
    norm = D3DXQuaternionLength(pQ)
    
    pOut.X = pQ.X / norm
    pOut.Y = pQ.Y / norm
    pOut.z = pQ.z / norm
    pOut.w = pQ.w / norm
    
End Sub

' // Rotates a quaternion about an arbitrary axis.
Public Sub D3DXQuaternionRotationAxis(pOut As D3DQUATERNION, _
                                      pV As D3DVECTOR, _
                                      ByVal angle As Single)
    Dim temp    As D3DVECTOR

    D3DXVec3Normalize temp, pV
    
    pOut.X = Sin(angle / 2!) * temp.X
    pOut.Y = Sin(angle / 2!) * temp.Y
    pOut.z = Sin(angle / 2!) * temp.z
    pOut.w = Cos(angle / 2!)

End Sub

' // Builds a quaternion from a rotation matrix.
Public Sub D3DXQuaternionRotationMatrix(pOut As D3DQUATERNION, _
                                        pM As D3DMATRIX)
    Dim i       As Long
    Dim maxi    As Long
    Dim maxdiag As Single
    Dim s       As Single
    Dim trace   As Single
    Dim sqrt    As Single
    
    trace = pM.m11 + pM.m22 + pM.m33 + 1!
    
    If trace > 1! Then
        sqrt = Sqr(trace)
        pOut.X = (pM.m23 - pM.m32) / (2! * sqrt)
        pOut.Y = (pM.m31 - pM.m13) / (2! * sqrt)
        pOut.z = (pM.m12 - pM.m21) / (2! * sqrt)
        pOut.w = sqrt / 2
        Exit Sub
    End If
    
    maxi = 0:   maxdiag = pM.m11
    
    If pM.m22 > maxdiag Then
        maxi = 1
        maxdiag = pM.m22
    End If
    
    If pM.m33 > maxdiag Then
        maxi = 2
        maxdiag = pM.m33
    End If
    
    Select Case maxi
    Case 0
        s = 2! * Sqr(1! + pM.m11 - pM.m22 - pM.m33)
        pOut.X = 0.25! * s
        pOut.Y = (pM.m12 + pM.m21) / s
        pOut.z = (pM.m13 + pM.m31) / s
        pOut.w = (pM.m23 + pM.m32) / s
    Case 1
        s = 2! * Sqr(1! + pM.m22 - pM.m11 - pM.m33)
        pOut.X = (pM.m12 + pM.m21) / s
        pOut.Y = 0.25! * s
        pOut.z = (pM.m23 + pM.m32) / s
        pOut.w = (pM.m31 + pM.m13) / s
    Case 2
        s = 2! * Sqr(1! + pM.m33 - pM.m11 - pM.m22)
        pOut.X = (pM.m13 + pM.m31) / s
        pOut.Y = (pM.m23 + pM.m32) / s
        pOut.z = 0.25! * s
        pOut.w = (pM.m12 + pM.m21) / s
    End Select
    
End Sub

' // Builds a quaternion with the given yaw, pitch, and roll.
Public Sub D3DXQuaternionRotationYawPitchRoll(pOut As D3DQUATERNION, _
                                              ByVal yaw As Single, _
                                              ByVal pitch As Single, _
                                              ByVal roll As Single)

    pOut.X = Sin(yaw / 2!) * Cos(pitch / 2!) * Sin(roll / 2!) + Cos(yaw / 2!) * Sin(pitch / 2!) * Cos(roll / 2!)
    pOut.Y = Sin(yaw / 2!) * Cos(pitch / 2!) * Cos(roll / 2!) - Cos(yaw / 2!) * Sin(pitch / 2!) * Sin(roll / 2!)
    pOut.z = Cos(yaw / 2!) * Cos(pitch / 2!) * Sin(roll / 2!) - Sin(yaw / 2!) * Sin(pitch / 2!) * Cos(roll / 2!)
    pOut.w = Cos(yaw / 2!) * Cos(pitch / 2!) * Cos(roll / 2!) + Sin(yaw / 2!) * Sin(pitch / 2!) * Sin(roll / 2!)
    
End Sub

' // Interpolates between two quaternions, using spherical linear interpolation.
Public Sub D3DXQuaternionSlerp(pOut As D3DQUATERNION, _
                               pQ1 As D3DQUATERNION, _
                               pQ2 As D3DQUATERNION, _
                               ByVal t As Single)
    Dim dot     As Single
    Dim epsilon As Single
    Dim temp    As Single
    Dim theta   As Single
    Dim u       As Single

    epsilon = 1!
    temp = 1! - t
    u = t
    
    dot = D3DXQuaternionDot(pQ1, pQ2)
    
    If (dot < 0!) Then
    
        epsilon = -1!
        dot = -dot
        
    End If
    
    If 1! - dot > 0.001! Then

        theta = acos(dot)
        temp = Sin(theta * temp) / Sin(theta)
        u = Sin(theta * u) / Sin(theta)
        
    End If
    
    pOut.X = temp * pQ1.X + epsilon * u * pQ2.X
    pOut.Y = temp * pQ1.Y + epsilon * u * pQ2.Y
    pOut.z = temp * pQ1.z + epsilon * u * pQ2.z
    pOut.w = temp * pQ1.w + epsilon * u * pQ2.w

End Sub

' // Interpolates between quaternions, using spherical quadrangle interpolation.
Public Sub D3DXQuaternionSquad(pOut As D3DQUATERNION, _
                               pQ1 As D3DQUATERNION, _
                               pA As D3DQUATERNION, _
                               pB As D3DQUATERNION, _
                               pC As D3DQUATERNION, _
                               ByVal t As Single)
                               
    Dim temp1   As D3DQUATERNION
    Dim temp2   As D3DQUATERNION

    D3DXQuaternionSlerp temp1, pQ1, pC, t
    D3DXQuaternionSlerp temp2, pA, pB, t
    
    D3DXQuaternionSlerp pOut, temp1, temp2, 2! * t * (1! - t)

End Sub

' // Sets up control points for spherical quadrangle interpolation.
Public Sub D3DXQuaternionSquadSetup(pAOut As D3DQUATERNION, _
                                    pBOut As D3DQUATERNION, _
                                    pCOut As D3DQUATERNION, _
                                    pQ0 As D3DQUATERNION, _
                                    pQ1 As D3DQUATERNION, _
                                    pQ2 As D3DQUATERNION, _
                                    pQ3 As D3DQUATERNION)
    Dim q     As D3DQUATERNION
    Dim temp1 As D3DQUATERNION
    Dim temp2 As D3DQUATERNION
    Dim temp3 As D3DQUATERNION
    Dim zero  As D3DQUATERNION
  
    If (D3DXQuaternionDot(pQ0, pQ1) < 0!) Then
        
        temp2.X = -pQ0.X
        temp2.Y = -pQ0.Y
        temp2.z = -pQ0.z
        temp2.w = -pQ0.w

    Else: temp2 = pQ0
    End If
    
    If (D3DXQuaternionDot(pQ1, pQ2) < 0!) Then
        
        pCOut.X = -pQ2.X
        pCOut.Y = -pQ2.Y
        pCOut.z = -pQ2.z
        pCOut.w = -pQ2.w

    Else: pCOut = pQ2
    End If
    
    If (D3DXQuaternionDot(pCOut, pQ3) < 0!) Then
        
        temp3.X = -pQ3.X
        temp3.Y = -pQ3.Y
        temp3.z = -pQ3.z
        temp3.w = -pQ3.w

    Else: temp3 = pQ3
    End If
    
    D3DXQuaternionInverse temp1, pQ1
    D3DXQuaternionMultiply temp2, temp1, temp2
    D3DXQuaternionLn temp2, temp2
    D3DXQuaternionMultiply q, temp1, pCOut
    D3DXQuaternionLn q, q
    
    temp1.X = temp2.X + q.X
    temp1.Y = temp2.Y + q.Y
    temp1.z = temp2.z + q.z
    temp1.w = temp2.w + q.w
    
    temp1.X = temp1.X * -0.25!
    temp1.Y = temp1.Y * -0.25!
    temp1.z = temp1.z * -0.25!
    temp1.w = temp1.w * -0.25!
    
    D3DXQuaternionExp temp1, temp1
    D3DXQuaternionMultiply pAOut, pQ1, temp1
    D3DXQuaternionInverse temp1, pCOut
    D3DXQuaternionMultiply temp2, temp1, pQ1
    D3DXQuaternionLn temp2, temp2
    D3DXQuaternionMultiply q, temp1, temp3
    D3DXQuaternionLn q, q
    
    temp1.X = temp2.X + q.X
    temp1.Y = temp2.Y + q.Y
    temp1.z = temp2.z + q.z
    temp1.w = temp2.w + q.w
    
    temp1.X = temp1.X * -0.25!
    temp1.Y = temp1.Y * -0.25!
    temp1.z = temp1.z * -0.25!
    temp1.w = temp1.w * -0.25!
    
    D3DXQuaternionExp temp1, temp1
    D3DXQuaternionMultiply pBOut, pCOut, temp1

End Sub

' // Computes a quaternion's axis and angle of rotation.
Public Sub D3DXQuaternionToAxisAngle(pQ As D3DQUATERNION, _
                                     pAxis As D3DVECTOR, _
                                     pAngle As Single)
    Dim fNorm   As Single
    
    fNorm = D3DXQuaternionLength(pQ)
    
    pAngle = 0
    
    If fNorm <> 0! Then
    
        pAxis.X = pQ.X / fNorm
        pAxis.Y = pQ.Y / fNorm
        pAxis.z = pQ.z / fNorm
        
        If Abs(pQ.w <= 1!) Then
            pAngle = 2! * acos(pQ.w)
        End If
        
    Else
    
        pAxis.X = 1!
        pAxis.Y = 0!
        pAxis.z = 0!
        
    End If
        
End Sub

Private Function acos(ByVal Value As Single) As Single

    If Value = -1! Then acos = PI:  Exit Function
    If Value = 1! Then acos = 0:    Exit Function
    acos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
    
End Function
