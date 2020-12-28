VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Camera class
' //

Option Explicit

' // Raised when a property has been changed
Public Event Changed()

Private mfDistance      As Single           ' // Distance from target
Private mtTarget        As D3DVECTOR        ' // Target position
Private mtOrientation   As D3DQUATERNION    ' // Orientation
Private mfFOV           As Single           ' // Field of view
Private mtViewMatrix    As D3DMATRIX        ' // View matrix

Dim bUpdate As Boolean  ' // If any property has been changed update view matrix

' // Field of view proerty
Public Property Let FOV( _
                    ByVal fValue As Single)
                    
    mfFOV = fValue
    RaiseEvent Changed
    
End Property

Public Property Get FOV() As Single
    FOV = mfFOV
End Property

' // Camera position
Public Property Let Position( _
                    ByRef tValue As D3DVECTOR)
    LookAt tValue, mtTarget, vec3(0, 1, 0)
End Property

Public Property Get Position() As D3DVECTOR
    Dim tMtx    As D3DMATRIX
    
    ' // Get camera position from orientation, target and distance
    D3DXMatrixRotationQuaternion tMtx, mtOrientation
    
    Position = vec3(tMtx.m13, tMtx.m23, tMtx.m33)
    
    D3DXVec3Scale Position, Position, -mfDistance
    D3DXVec3Add Position, Position, mtTarget
    
End Property

' // Camera target
Public Property Let Target( _
                    ByRef tValue As D3DVECTOR)
    LookAt Position, tValue, vec3(0, 1, 0)
End Property

Public Property Get Target() As D3DVECTOR
    Target = mtTarget
End Property

' // Get view matrix
Public Property Get ViewMatrix() As D3DMATRIX
    
    If bUpdate Then
        Update
    End If
    
    ViewMatrix = mtViewMatrix
    
End Property

' // Place camera by specified point
Public Sub LookAt( _
           ByRef tEye As D3DVECTOR, _
           ByRef tTarget As D3DVECTOR, _
           ByRef tUp As D3DVECTOR)
    Dim tMtx    As D3DMATRIX
    Dim tDist   As D3DVECTOR
    
    D3DXMatrixLookAtLH tMtx, tEye, tTarget, tUp
    D3DXQuaternionRotationMatrix mtOrientation, tMtx
    D3DXQuaternionNormalize mtOrientation, mtOrientation
    
    D3DXVec3Subtract tDist, tEye, tTarget
    
    mfDistance = D3DXVec3Length(tDist)
    mtTarget = tTarget

    bUpdate = True
    
    RaiseEvent Changed
    
End Sub

' // Relative rotation
Public Sub RotateRel( _
           ByRef tVec As D3DVECTOR)
    Dim tQ  As D3DQUATERNION

    If (tVec.Y <> 0!) Then

        D3DXQuaternionRotationAxis tQ, vec3(0, 1, 0), tVec.Y
        D3DXQuaternionMultiply mtOrientation, tQ, mtOrientation
    
    End If

    If (tVec.X <> 0!) Then
    
        D3DXQuaternionRotationAxis tQ, vec3(1, 0, 0), tVec.X
        D3DXQuaternionMultiply mtOrientation, mtOrientation, tQ
        
    End If
        
    D3DXQuaternionNormalize mtOrientation, mtOrientation
    
    bUpdate = True
    RaiseEvent Changed
    
End Sub

' // Zoom
Public Sub Zoom( _
           ByVal fValue As Single)

    mfDistance = mfDistance + fValue
    
    bUpdate = True
    RaiseEvent Changed
              
End Sub

' // Panning
Public Sub Pan( _
           ByVal fX As Single, _
           ByVal fY As Single)
    Dim tDirX   As D3DVECTOR
    Dim tDirY   As D3DVECTOR
    Dim tMtx    As D3DMATRIX
    
    D3DXMatrixRotationQuaternion tMtx, mtOrientation
    
    tDirX = vec3(tMtx.m11, tMtx.m21, tMtx.m31)
    tDirY = vec3(tMtx.m12, tMtx.m22, tMtx.m32)
    
    D3DXVec3Scale tDirX, tDirX, fX
    D3DXVec3Scale tDirY, tDirY, fY
    
    D3DXVec3Add mtTarget, mtTarget, tDirX
    D3DXVec3Add mtTarget, mtTarget, tDirY
    
    bUpdate = True
    RaiseEvent Changed
              
End Sub

' // Update view matrix
Private Sub Update()
    Dim tMtx        As D3DMATRIX
    Dim tMtx2       As D3DMATRIX
    Dim tAxis(2)    As D3DVECTOR
    Dim tDir        As D3DVECTOR
    Dim tEye        As D3DVECTOR

    D3DXMatrixRotationQuaternion mtViewMatrix, mtOrientation

    tAxis(0) = vec3(mtViewMatrix.m11, mtViewMatrix.m21, mtViewMatrix.m31)
    tAxis(1) = vec3(mtViewMatrix.m12, mtViewMatrix.m22, mtViewMatrix.m32)
    tAxis(2) = vec3(mtViewMatrix.m13, mtViewMatrix.m23, mtViewMatrix.m33)
    
    tDir = tAxis(2)
    
    D3DXVec3Scale tDir, tDir, -mfDistance
    D3DXVec3Add tEye, tDir, mtTarget

    mtViewMatrix.m41 = -D3DXVec3Dot(tAxis(0), tEye)
    mtViewMatrix.m42 = -D3DXVec3Dot(tAxis(1), tEye)
    mtViewMatrix.m43 = -D3DXVec3Dot(tAxis(2), tEye)

    bUpdate = False
    
End Sub

Private Sub Class_Initialize()

    LookAt vec3(0, 1, -1), vec3(0, 0, 0), vec3(0, 1, 0)
    mfFOV = PI / 180 * 60
    
End Sub
