VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Mesh class
' //

Option Explicit

' // Vertex format of mesh
Private Type tVertexFormat
    tPosition   As D3DVECTOR
    tNormal     As D3DVECTOR
End Type

Private Const SIZE_OF_VERTEX    As Long = 24        ' // Size of vertex in bytes

Private mcVertexBuffer  As IDirect3DVertexBuffer9   ' // Vertex buffer
Private mtPosition      As D3DVECTOR                ' // Position of mesh
Private mtPivot         As D3DVECTOR                ' // Pivot point
Private mtOrientation   As D3DQUATERNION            ' // Orientation
Private mlColor         As Long                     ' // Color
Private mcMaterial      As CMaterial                ' // Material
Private mtWorld         As D3DMATRIX                ' // Transform matrix

Dim mbChanged   As Boolean  ' // If any property has been changed update matrix

' // Position of mesh
Public Property Let Position( _
                    ByRef tValue As D3DVECTOR)
                    
    mtPosition = tValue
    mbChanged = True
    
End Property
Public Property Get Position() As D3DVECTOR
    Position = mtPosition
End Property

' // Pivot point (mesh turns around that point)
Public Property Let PivotPoint( _
                    ByRef tValue As D3DVECTOR)
                    
    mtPivot = tValue
    mbChanged = True
    
End Property
Public Property Get PivotPoint() As D3DVECTOR
    PivotPoint = mtPivot
End Property

' // Orientation quaternion
Public Property Let Orientation( _
                    ByRef tValue As D3DQUATERNION)
                    
    mtOrientation = tValue
    mbChanged = True
    
End Property
Public Property Get Orientation() As D3DQUATERNION
    Orientation = mtOrientation
End Property

' // Vertext buffer
Public Property Get VertexBuffer() As IDirect3DVertexBuffer9
    Set VertexBuffer = mcVertexBuffer
End Property

' // Material
Public Property Get Material() As CMaterial
    Set Material = mcMaterial
End Property

' // Get mesh bounds
Public Sub GetBounds( _
           ByRef tV1 As D3DVECTOR, _
           ByRef tV2 As D3DVECTOR)
    Dim tVertex()   As tVertexFormat
    Dim lIndex      As Long
    
    GetVertexData tVertex()

    For lIndex = 0 To UBound(tVertex)
        
        With tVertex(lIndex).tPosition
        
            ' // Search for min/max vertices positions
            If lIndex Then
                
                If .X < tV1.X Then
                    tV1.X = .X
                ElseIf .X > tV2.X Then
                    tV2.X = .X
                End If
                
                If .Y < tV1.Y Then
                    tV1.Y = .Y
                ElseIf .Y > tV2.Y Then
                    tV2.Y = .Y
                End If
            
                If .Z < tV1.Z Then
                    tV1.Z = .Z
                ElseIf .Z > tV2.Z Then
                    tV2.Z = .Z
                End If
                
            Else
                
                tV1 = tVertex(lIndex).tPosition
                tV2 = tVertex(lIndex).tPosition
                
            End If
            
        End With
        
    Next

End Sub

' // Get transform matrix
Public Property Get WorldTransform() As D3DMATRIX
    Dim tMtx    As D3DMATRIX

    If mbChanged Then
        
        D3DXMatrixRotationQuaternion tMtx, mtOrientation
        D3DXMatrixTranslation mtWorld, mtPosition.X, mtPosition.Y, mtPosition.Z
        D3DXMatrixMultiply mtWorld, tMtx, mtWorld
        D3DXMatrixTranslation tMtx, mtPivot.X, mtPivot.Y, mtPivot.Z
        D3DXMatrixMultiply mtWorld, tMtx, mtWorld
        
        mbChanged = False
        
    End If
    
    WorldTransform = mtWorld
    
End Property

' // Check if a ray intersects mesh
Public Function HitTest( _
                ByRef tRayFrom As D3DVECTOR, _
                ByRef tRayTo As D3DVECTOR) As Boolean
    Dim tOrigin     As D3DVECTOR
    Dim tDir        As D3DVECTOR
    Dim tMtx1       As D3DMATRIX
    Dim tMtx2       As D3DMATRIX
    Dim tDesc       As D3DVERTEXBUFFER_DESC
    Dim pData       As Long
    Dim tVertex()   As tVertexFormat
    Dim lIndex      As Long
    Dim tQinv       As D3DQUATERNION
    
'
'    D3DXMatrixTranslation tMtx1, mtPivot.X, mtPivot.Y, mtPivot.z
'    D3DXMatrixRotationQuaternion tMtx2, mtOrientation
'    D3DXMatrixMultiply tMtx2, tMtx1, tMtx2
'    D3DXMatrixTranslation tMtx1, mtPosition.X, mtPosition.Y, mtPosition.z
'    D3DXMatrixMultiply tMtx2, tMtx2, tMtx1
'    D3DXMatrixInverse tMtx2, 0, tMtx2
'
    ' // Transform ray to object space
    D3DXMatrixTranslation tMtx1, -mtPosition.X, -mtPosition.Y, -mtPosition.Z
    D3DXQuaternionInverse tQinv, mtOrientation
    D3DXMatrixRotationQuaternion tMtx2, tQinv
    D3DXMatrixMultiply tMtx2, tMtx1, tMtx2
    D3DXMatrixTranslation tMtx1, -mtPivot.X, -mtPivot.Y, -mtPivot.Z
    D3DXMatrixMultiply tMtx2, tMtx2, tMtx1
    
    D3DXVec3TransformCoord tOrigin, tRayFrom, tMtx2
    D3DXVec3TransformCoord tDir, tRayTo, tMtx2
    
    ' // Get normalized direction
    D3DXVec3Subtract tDir, tDir, tOrigin
    D3DXVec3Normalize tDir, tDir
    
    GetVertexData tVertex()
    
    ' // Go thru triangles
    For lIndex = 0 To UBound(tVertex) Step 3
        
        If IsIntersected(tVertex(lIndex).tPosition, _
                         tVertex(lIndex + 1).tPosition, _
                         tVertex(lIndex + 2).tPosition, _
                         tOrigin, tDir) Then

            HitTest = True
            Exit Function
            
        End If
        
    Next
    
End Function

' // Render mesh
Public Sub Render()
    Dim tDesc   As D3DVERTEXBUFFER_DESC
    Dim tMtx    As D3DMATRIX
    Dim cDevice As IDirect3DDevice9
    
    Set cDevice = mcVertexBuffer.GetDevice
    
    tMtx = WorldTransform()
    
    cDevice.SetTransform D3DTS_WORLD, tMtx
    
    mcVertexBuffer.GetDesc tDesc

    cDevice.SetStreamSource 0, mcVertexBuffer, 0, SIZE_OF_VERTEX
    cDevice.SetFVF tDesc.FVF
        
    cDevice.SetMaterial mcMaterial.Material
    
    ' // Use specular
    cDevice.SetRenderState D3DRS_SPECULARENABLE, 1
    cDevice.SetRenderState D3DRS_SPECULARMATERIALSOURCE, D3DMCS_MATERIAL
    
    cDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, tDesc.Size \ (SIZE_OF_VERTEX * 3)

End Sub

' // Render normals (unoptimized, for debugging)
Public Sub RenderNormals()
    Dim cdev        As IDirect3DDevice9
    Dim cBuff       As IDirect3DVertexBuffer9
    Dim tNormals()  As tVertexFormat
    Dim tVertex()   As tVertexFormat
    Dim lIndex      As Long
    Dim lBufLen     As Long
    Dim pData       As Long
    
    GetVertexData tVertex()
    
    Set cdev = mcVertexBuffer.GetDevice

    ReDim tNormals(UBound(tVertex))

    For lIndex = 0 To UBound(tNormals)

        tNormals(lIndex).tPosition = tVertex(lIndex).tPosition
        D3DXVec3Scale tVertex(lIndex).tNormal, tVertex(lIndex).tNormal, 0.5
        D3DXVec3Add tNormals(lIndex).tNormal, tVertex(lIndex).tPosition, tVertex(lIndex).tNormal

    Next

    lBufLen = (UBound(tNormals) + 1) * SIZE_OF_VERTEX

    cdev.CreateVertexBuffer lBufLen, D3DUSAGE_NONE, D3DFVF_XYZ, D3DPOOL_DEFAULT, cBuff

    cBuff.Lock 0, lBufLen, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tNormals(0), lBufLen
    cBuff.Unlock

    cdev.SetStreamSource 0, cBuff, 0, 12
    cdev.SetFVF D3DFVF_XYZ

    cdev.DrawPrimitive D3DPT_LINELIST, 0, lBufLen / 24

End Sub

' // Create sphere to vertex buffer
Public Sub CreateSphere( _
           ByVal cDevice As IDirect3DDevice9, _
           ByVal lNumOfSegments As Long, _
           ByVal fSize As Single, _
           ByVal lColor As Long)
    Dim tVertex()   As tVertexFormat
    Dim pData       As Long
    Dim lBufferSize As Long
    
    GenerateSphere lNumOfSegments, fSize, tVertex
    
    lBufferSize = (UBound(tVertex) + 1) * Len(tVertex(0))
    
    ' // Write data to vertex buffer
    cDevice.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_NORMAL, _
                                D3DPOOL_DEFAULT, mcVertexBuffer

    mcVertexBuffer.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tVertex(0), lBufferSize
    mcVertexBuffer.Unlock
    
    mcMaterial.DiffuseColor = lColor
    
End Sub

' // Create cone to vertex buffer
Public Sub CreateCone( _
           ByVal cDevice As IDirect3DDevice9, _
           ByVal lNumOfSegments As Long, _
           ByVal fHeight As Single, _
           ByVal fRadius1 As Single, _
           ByVal fRadius2 As Single, _
           ByVal lColor As Long, _
           Optional ByVal bCap As Boolean = True)
    Dim tVertex()   As tVertexFormat
    Dim pData       As Long
    Dim lBufferSize As Long

    GenerateCone lNumOfSegments, fHeight, fRadius1, fRadius2, tVertex, bCap
    
    lBufferSize = (UBound(tVertex) + 1) * Len(tVertex(0))
    
    ' // Write data to vertex buffer
    cDevice.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_NORMAL, _
                                D3DPOOL_DEFAULT, mcVertexBuffer
    
    mcVertexBuffer.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tVertex(0), lBufferSize
    mcVertexBuffer.Unlock
    
    mcMaterial.DiffuseColor = lColor
     
End Sub

' // Create box to vertex buffer
Public Sub CreateBox( _
           ByVal cDevice As IDirect3DDevice9, _
           ByVal fWidth As Single, _
           ByVal fHeight As Single, _
           ByVal fDepth As Single, _
           ByVal lColor As Long)
    Dim tVertex()   As tVertexFormat
    Dim pData       As Long
    Dim lBufferSize As Long

    GenerateBox fWidth, fHeight, fDepth, tVertex
    
    lBufferSize = (UBound(tVertex) + 1) * Len(tVertex(0))
    
    ' // Write data to vertex buffer
    cDevice.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_NORMAL, _
                                D3DPOOL_DEFAULT, mcVertexBuffer
    
    mcVertexBuffer.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tVertex(0), lBufferSize
    mcVertexBuffer.Unlock
    
    mcMaterial.DiffuseColor = lColor
     
End Sub

' // Extract vertext data from buffer
Private Sub GetVertexData( _
            ByRef tVertex() As tVertexFormat)
    Dim pData   As Long
    Dim tDesc   As D3DVERTEXBUFFER_DESC
    
    mcVertexBuffer.GetDesc tDesc
    
    If tDesc.Size < SIZE_OF_VERTEX Then Exit Sub
    
    ReDim tVertex(tDesc.Size \ SIZE_OF_VERTEX - 1)

    mcVertexBuffer.Lock 0, tDesc.Size, pData, D3DLOCK_READONLY
    memcpy tVertex(0), ByVal pData, tDesc.Size
    mcVertexBuffer.Unlock

End Sub

' // Generate box vertices
Private Sub GenerateBox( _
            ByVal fWidth As Single, _
            ByVal fHeight As Single, _
            ByVal fDepth As Single, _
            ByRef tVertex() As tVertexFormat)
    Dim lSideIndex  As Long
    Dim lVtxIndex   As Long
    Dim tVec(3)     As D3DVECTOR
    Dim tNormal     As D3DVECTOR
    Dim tMtx        As D3DMATRIX
    
    ' // 1. Create unit box (1 x 1 x 1)
    ' // Each face contains 2 triangles with 3 points. Cube contains 6 faces.
    ReDim tVertex(6 * 3 * 2 - 1)
    
    ' // Create initial vectors
    tVec(0) = vec3(0.5, 0, 0.5)
    tVec(1) = vec3(0.5, 1, 0.5)
    tNormal = vec3(1, 0, -1)
    
    D3DXVec3Normalize tNormal, tNormal
    D3DXMatrixRotationY tMtx, PI / 2
    
    For lSideIndex = 0 To 3
        
        tVec(2) = tVec(0)
        tVec(3) = tVec(1)
        
        ' // Rotate vectors by 90 degrees
        D3DXVec3TransformCoord tVec(0), tVec(0), tMtx
        D3DXVec3TransformCoord tVec(1), tVec(1), tMtx

        tVertex(lVtxIndex).tPosition = tVec(1):     tVertex(lVtxIndex).tNormal = tNormal
        tVertex(lVtxIndex + 1).tPosition = tVec(2): tVertex(lVtxIndex + 1).tNormal = tNormal
        tVertex(lVtxIndex + 2).tPosition = tVec(0): tVertex(lVtxIndex + 2).tNormal = tNormal

        tVertex(lVtxIndex + 3).tPosition = tVec(1): tVertex(lVtxIndex + 3).tNormal = tNormal
        tVertex(lVtxIndex + 4).tPosition = tVec(3): tVertex(lVtxIndex + 4).tNormal = tNormal
        tVertex(lVtxIndex + 5).tPosition = tVec(2): tVertex(lVtxIndex + 5).tNormal = tNormal
        
        D3DXVec3TransformCoord tNormal, tNormal, tMtx
        
        lVtxIndex = lVtxIndex + 6
        
    Next
    
    ' // Up cap
    tNormal = vec3(0, 1, 0)
    
    tVertex(lVtxIndex).tPosition = tVertex(6).tPosition: tVertex(lVtxIndex).tNormal = tNormal
    tVertex(lVtxIndex + 1).tPosition = tVertex(4).tPosition: tVertex(lVtxIndex + 1).tNormal = tNormal
    tVertex(lVtxIndex + 2).tPosition = tVertex(0).tPosition: tVertex(lVtxIndex + 2).tNormal = tNormal
    tVertex(lVtxIndex + 3).tPosition = tVertex(12).tPosition: tVertex(lVtxIndex + 3).tNormal = tNormal
    tVertex(lVtxIndex + 4).tPosition = tVertex(4).tPosition: tVertex(lVtxIndex + 4).tNormal = tNormal
    tVertex(lVtxIndex + 5).tPosition = tVertex(6).tPosition: tVertex(lVtxIndex + 5).tNormal = tNormal
    
    lVtxIndex = lVtxIndex + 6
    
    ' // Down cap
    tNormal = vec3(0, -1, 0)
    
    tVertex(lVtxIndex).tPosition = tVertex(1).tPosition: tVertex(lVtxIndex).tNormal = tNormal
    tVertex(lVtxIndex + 1).tPosition = tVertex(8).tPosition: tVertex(lVtxIndex + 1).tNormal = tNormal
    tVertex(lVtxIndex + 2).tPosition = tVertex(7).tPosition: tVertex(lVtxIndex + 2).tNormal = tNormal
    tVertex(lVtxIndex + 3).tPosition = tVertex(14).tPosition: tVertex(lVtxIndex + 3).tNormal = tNormal
    tVertex(lVtxIndex + 4).tPosition = tVertex(8).tPosition: tVertex(lVtxIndex + 4).tNormal = tNormal
    tVertex(lVtxIndex + 5).tPosition = tVertex(1).tPosition: tVertex(lVtxIndex + 5).tNormal = tNormal
    
    ' // 2. Change scales corresponding parameters
    D3DXMatrixScaling tMtx, fWidth, fHeight, fDepth
    
    For lVtxIndex = 0 To UBound(tVertex)
        D3DXVec3TransformCoord tVertex(lVtxIndex).tPosition, tVertex(lVtxIndex).tPosition, tMtx
    Next
    
End Sub

' // Generate cone vertices
Private Sub GenerateCone( _
            ByVal lNumOfSegments As Long, _
            ByVal fHeight As Single, _
            ByVal fRadius1 As Single, _
            ByVal fRadius2 As Single, _
            ByRef tVertex() As tVertexFormat, _
            Optional ByVal bCap As Boolean = True)
    Dim lNumOfTriangles As Long
    Dim lSegIndex       As Long
    Dim tVec(3)         As D3DVECTOR
    Dim tYAxis          As D3DVECTOR
    Dim tNormals(1)     As D3DVECTOR
    Dim tConeNormal     As D3DVECTOR
    Dim lVtxIndex       As Long
    Dim tMtx            As D3DMATRIX
    
    lNumOfTriangles = lNumOfSegments * 2
    
    If bCap Then
        If fRadius1 = 0! Or fRadius2 = 0! Then
            lNumOfTriangles = lNumOfTriangles + lNumOfSegments
        Else
            lNumOfTriangles = lNumOfTriangles + lNumOfSegments * 2
        End If
    End If
    
    ReDim tVertex(lNumOfTriangles * 3 - 1)
    
    ' // Create initial vectors
    tYAxis = vec3(0, 1, 0)
    tVec(0) = vec3(fRadius1, 0, 0)
    tVec(1) = vec3(fRadius2, fHeight, 0)
    tConeNormal = vec3(fHeight, fRadius1 - fRadius2, 0)

    D3DXVec3Normalize tConeNormal, tConeNormal
    
    D3DXMatrixRotationY tMtx, PI * 2 / lNumOfSegments
    
    tNormals(0) = tConeNormal

    For lSegIndex = 0 To lNumOfSegments - 1
        
        tNormals(1) = tNormals(0)
        tVec(2) = tVec(0)
        tVec(3) = tVec(1)
        
        ' // Rotate depending on segment angle
        D3DXVec3TransformCoord tVec(0), tVec(0), tMtx
        D3DXVec3TransformCoord tVec(1), tVec(1), tMtx
        D3DXVec3TransformCoord tNormals(0), tNormals(0), tMtx

        tVertex(lVtxIndex).tPosition = tVec(1):     tVertex(lVtxIndex).tNormal = tNormals(0)
        tVertex(lVtxIndex + 1).tPosition = tVec(2): tVertex(lVtxIndex + 1).tNormal = tNormals(1)
        tVertex(lVtxIndex + 2).tPosition = tVec(0): tVertex(lVtxIndex + 2).tNormal = tNormals(0)

        tVertex(lVtxIndex + 3).tPosition = tVec(1): tVertex(lVtxIndex + 3).tNormal = tNormals(0)
        tVertex(lVtxIndex + 4).tPosition = tVec(3): tVertex(lVtxIndex + 4).tNormal = tNormals(1)
        tVertex(lVtxIndex + 5).tPosition = tVec(2): tVertex(lVtxIndex + 5).tNormal = tNormals(1)
        
        lVtxIndex = lVtxIndex + 6
        
        If bCap Then
        
            If fRadius1 <> 0! Then
            
                tVertex(lVtxIndex).tPosition = vec3(0, 0, 0):   D3DXVec3Scale tVertex(lVtxIndex).tNormal, tYAxis, -1
                tVertex(lVtxIndex + 1).tPosition = tVec(0):     tVertex(lVtxIndex + 1).tNormal = tVertex(lVtxIndex).tNormal
                tVertex(lVtxIndex + 2).tPosition = tVec(2):     tVertex(lVtxIndex + 2).tNormal = tVertex(lVtxIndex).tNormal
                
                lVtxIndex = lVtxIndex + 3
                
            End If
            
            If fRadius2 <> 0! Then
            
                tVertex(lVtxIndex).tPosition = vec3(0, fHeight, 0): tVertex(lVtxIndex).tNormal = tYAxis
                tVertex(lVtxIndex + 1).tPosition = tVec(3):     tVertex(lVtxIndex + 1).tNormal = tVertex(lVtxIndex).tNormal
                tVertex(lVtxIndex + 2).tPosition = tVec(1):     tVertex(lVtxIndex + 2).tNormal = tVertex(lVtxIndex).tNormal
                
                lVtxIndex = lVtxIndex + 3
                
            End If
            
        Else
            
        End If
        
    Next

End Sub

' // Generate sphere vertices
Private Sub GenerateSphere( _
            ByVal lNumOfSegments As Long, _
            ByVal fSize As Single, _
            ByRef tVertex() As tVertexFormat)
    Dim tVec(3)     As D3DVECTOR
    Dim tMtx        As D3DMATRIX
    Dim lSegIndex   As Long
    Dim lSideIndex  As Long
    Dim lVtxIndex   As Long
    Dim lTriIndex   As Long
    
    tVec(0) = vec3(0, 1, 0)
    
    lSegIndex = 1

    D3DXMatrixRotationY tMtx, PI * 2 / lNumOfSegments

    ReDim tVertex(lNumOfSegments * lNumOfSegments * 6 - 1)
       
    Do While lSegIndex <= lNumOfSegments
        
        tVec(1) = vec3(0, Cos(lSegIndex / lNumOfSegments * PI), Sin(lSegIndex / lNumOfSegments * PI))
        lSideIndex = 0

        Do While lSideIndex < lNumOfSegments
        
            D3DXVec3TransformCoord tVec(2), tVec(0), tMtx
            D3DXVec3TransformCoord tVec(3), tVec(1), tMtx
            
            For lTriIndex = 0 To 2
            
                tVertex(lVtxIndex).tPosition.X = tVec(lTriIndex).X * fSize
                tVertex(lVtxIndex).tPosition.Y = tVec(lTriIndex).Y * fSize
                tVertex(lVtxIndex).tPosition.Z = tVec(lTriIndex).Z * fSize
                tVertex(lVtxIndex).tNormal = tVec(lTriIndex)
                
                lVtxIndex = lVtxIndex + 1
            
            Next
            
            For lTriIndex = 3 To 1 Step -1
            
                tVertex(lVtxIndex).tPosition.X = tVec(lTriIndex).X * fSize
                tVertex(lVtxIndex).tPosition.Y = tVec(lTriIndex).Y * fSize
                tVertex(lVtxIndex).tPosition.Z = tVec(lTriIndex).Z * fSize
                tVertex(lVtxIndex).tNormal = tVec(lTriIndex)

                lVtxIndex = lVtxIndex + 1
            
            Next
    
            tVec(0) = tVec(2)
            tVec(1) = tVec(3)
            
            lSideIndex = lSideIndex + 1
            
        Loop
        
        tVec(0) = tVec(1)
        lSegIndex = lSegIndex + 1

    Loop

End Sub

Private Sub Class_Initialize()

    mtOrientation.w = 1
    Set mcMaterial = New CMaterial
    mbChanged = True
    
End Sub
