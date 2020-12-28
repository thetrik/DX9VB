VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Simple scene class 2 lights, 1 camera, grid
' //

Option Explicit

' // Vertex format for transfom gizmo
Private Type tColorVertex
    tPos    As D3DVECTOR
    lColor  As Long
End Type

Private Type RECT
    iLeft   As Long
    iTop    As Long
    iRight  As Long
    iBottom As Long
End Type

Private Declare Function GetClientRect Lib "user32" ( _
                         ByVal hwnd As Long, _
                         ByRef lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" ( _
                         ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
                         ByVal hwnd As Long, _
                         ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" _
                         Alias "TextOutW" ( _
                         ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal lpString As Long, _
                         ByVal nCount As Long) As Long

Private mlBackColor As Long ' // Foreground color

Private WithEvents mcCamera As CCamera      ' // Camera
Attribute mcCamera.VB_VarHelpID = -1
Private mcObjects           As Collection   ' // Objects collection
Private mcSelectedObject    As CMesh        ' // Selected mesh
Private mbDrawNormals       As Boolean      ' // Determine whether draw normals or not
Private mbDrawCameraTarget  As Boolean      ' // ... camera target

Dim cD3d9           As IDirect3D9               ' // Main Direct3d object
Dim cD3dev          As IDirect3DDevice9         ' // Device
Dim cGrid           As IDirect3DVertexBuffer9   ' // Grid vertex buffer
Dim cSelectionBox   As IDirect3DVertexBuffer9   ' // Selection box vertex buffer
Dim cTransformGizmo As IDirect3DVertexBuffer9   ' // Transform gizmo vertex buffer
Dim tSelectionWorld As D3DMATRIX                ' // Selection box transformation
Dim cCameraTarget   As CMesh                    ' // Camera target mesh (you can show it)

Dim hWindow As Long ' // Window handle

' // Determine whether draw camera target or not
Public Property Get DrawCameraTarget() As Boolean
    DrawCameraTarget = mbDrawCameraTarget
End Property
Public Property Let DrawCameraTarget( _
                    ByVal bValue As Boolean)
    mbDrawCameraTarget = bValue
    Render
End Property

' // Determine whether draw normals or not
Public Property Get DrawNormals() As Boolean
    DrawNormals = mbDrawNormals
End Property
Public Property Let DrawNormals( _
                    ByVal bValue As Boolean)
    mbDrawNormals = bValue
    Render
End Property

' // Get Direct3D device
Public Property Get Device() As IDirect3DDevice9
    Set Device = cD3dev
End Property

' // Get objects collection
Public Property Get Objects() As Collection
    Set Objects = mcObjects
End Property

' // Get camera
Public Property Get Camera() As CCamera
    Set Camera = mcCamera
End Property

' // Pick mesh using screenspace coordinates
' // !!! WARNING !!! NO SORTING!!! It can return meshes behind each other
Public Function Pick( _
                ByVal lX As Long, _
                ByVal lY As Long) As CMesh
    Dim tNear   As D3DVECTOR
    Dim tFar    As D3DVECTOR
    Dim cObject As CMesh
    
    ' // Get ray from screen coordinates
    RayFromScreenPos lX, lY, tNear, tFar

    For Each cObject In mcObjects
        
        ' // Check each object
        If cObject.HitTest(tNear, tFar) Then
            Set Pick = cObject
        End If

    Next
    
    ' // Check previous selection
    If Not mcSelectedObject Is Pick Then
        
        Set mcSelectedObject = Pick
        UpdateSelectionBox
        Render
        
    End If
    
End Function

' // Get ray from screen coordinates
' // From near plane to far plane
Public Sub RayFromScreenPos( _
           ByVal lX As Long, _
           ByVal lY As Long, _
           ByRef tRayFrom As D3DVECTOR, _
           ByRef tRayTo As D3DVECTOR)
    Dim tViewport   As D3DVIEWPORT9
    Dim tView       As D3DMATRIX
    Dim tWorld      As D3DMATRIX
    Dim tProjection As D3DMATRIX
    
    cD3dev.GetViewport tViewport
    cD3dev.GetTransform D3DTS_PROJECTION, tProjection
    cD3dev.GetTransform D3DTS_VIEW, tView
    
    D3DXMatrixIdentity tWorld
    
    tRayFrom = vec3(lX, lY, tViewport.MinZ)
    tRayTo = vec3(lX, lY, tViewport.MaxZ)
    
    D3DXVec3Unproject tRayFrom, tRayFrom, tViewport, tProjection, tView, tWorld
    D3DXVec3Unproject tRayTo, tRayTo, tViewport, tProjection, tView, tWorld
              
End Sub

' // Initialization procedure
Public Sub InitializeScene( _
           ByVal hwnd As Long)
    
    InitDevice hwnd
    InitLight
    InitCamera
    InitGrid
    InitSelectionBox
    InitTransformGizmo
    
    mlBackColor = RGB(&H30, &H3F, &H3F) ' // R B are swapped
    
End Sub

' // Render only specified mesh
Public Sub RenderMesh( _
           ByVal cObject As CMesh)
    Dim tMtx    As D3DMATRIX
    
    D3DXMatrixIdentity tMtx
    
    ' // Reset transform
    cD3dev.SetTransform D3DTS_WORLD, tMtx
    
    cD3dev.BeginScene
    
    cObject.Render
    
    cD3dev.EndScene
    
    cD3dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    
End Sub

' // Render scene. Use bFlip to control updation
Public Sub Render( _
           Optional ByVal bFlip As Boolean = True)
    Dim tMtx    As D3DMATRIX
    Dim cObject As CMesh
    
    D3DXMatrixIdentity tMtx
    
    cD3dev.SetTransform D3DTS_WORLD, tMtx
    
    ' // Clear background and Z-Buffer
    cD3dev.Clear 0, ByVal 0&, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, mlBackColor, 1, 0
    
    cD3dev.BeginScene
    
    ' // Draw grid
    RenderGrid
    
    For Each cObject In mcObjects
        
        ' // Render mesh
        cObject.Render
        
        If mbDrawNormals Then
            cObject.RenderNormals
        End If
        
    Next
    
    ' // Draw camera target
    If mbDrawCameraTarget Then
        cCameraTarget.Render
    End If
    
    ' // Selection box
    RenderSelectionBox
    
    ' // Transform gizmo
    RenderTransformGizmo

    cD3dev.EndScene
    
    If bFlip Then
        cD3dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    End If
    
End Sub

' // Transform specified point to screenspace
Public Function ToScreenPos( _
                ByRef tVec As D3DVECTOR, _
                ByRef tWorld As D3DMATRIX) As D3DVECTOR
    Dim tVpt    As D3DVIEWPORT9
    Dim tProj   As D3DMATRIX
    
    cD3dev.GetTransform D3DTS_PROJECTION, tProj
    cD3dev.GetViewport tVpt
    
    D3DXVec3Project ToScreenPos, tVec, tVpt, tProj, mcCamera.ViewMatrix, tWorld
    
End Function

' // Update selection box transformation
Private Sub UpdateSelectionBox()
    Dim tBounds(1)  As D3DVECTOR
    Dim fWidth      As Single
    Dim fHeight     As Single
    Dim fDepth      As Single
    Dim tPoints(7)  As D3DVECTOR
    Dim tMtx        As D3DMATRIX
    
    If mcSelectedObject Is Nothing Then Exit Sub
    
    mcSelectedObject.GetBounds tBounds(0), tBounds(1)
    
    fWidth = tBounds(1).X - tBounds(0).X
    fHeight = tBounds(1).Y - tBounds(0).Y
    fDepth = tBounds(1).Z - tBounds(0).Z
    
    D3DXMatrixTranslation tSelectionWorld, tBounds(0).X, tBounds(0).Y, tBounds(0).Z
    D3DXMatrixScaling tMtx, fWidth, fHeight, fDepth
    D3DXMatrixMultiply tSelectionWorld, tMtx, tSelectionWorld
    
End Sub

' // Draw transform gizmo
Private Sub RenderTransformGizmo()
    Dim tDesc           As D3DVERTEXBUFFER_DESC
    Dim tMat            As D3DMATERIAL9
    Dim lOldLightState  As Long
    Dim tMtx            As D3DMATRIX
    Dim tPivot          As D3DVECTOR
    
    If mcSelectedObject Is Nothing Then Exit Sub
    
    cD3dev.Clear 0, ByVal 0&, D3DCLEAR_ZBUFFER, 0, 1, 0
    
    tPivot = mcSelectedObject.PivotPoint
    
    ' // Transparency enable
    cD3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    cD3dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    cD3dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    
    D3DXVec3Scale tPivot, tPivot, -1
    
    D3DXMatrixTranslation tMtx, tPivot.X, tPivot.Y, tPivot.Z
    
    cD3dev.SetTransform D3DTS_WORLD, mcSelectedObject.WorldTransform
    cD3dev.MultiplyTransform D3DTS_WORLD, tMtx
    
    cTransformGizmo.GetDesc tDesc
    
    cD3dev.SetStreamSource 0, cTransformGizmo, 0, 16
    cD3dev.SetFVF tDesc.FVF
    
    lOldLightState = cD3dev.GetRenderState(D3DRS_LIGHTING)
    
    ' // Disable lighting
    cD3dev.SetRenderState D3DRS_LIGHTING, 0
    
    cD3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, tDesc.Size / 16
    
    cD3dev.SetRenderState D3DRS_LIGHTING, lOldLightState
    cD3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
End Sub

' // Draw selection box
Private Sub RenderSelectionBox()
    
    If mcSelectedObject Is Nothing Then Exit Sub
    
    cD3dev.SetTransform D3DTS_WORLD, mcSelectedObject.WorldTransform
    cD3dev.MultiplyTransform D3DTS_WORLD, tSelectionWorld
    
    RenderLines cSelectionBox, color(1, 1, 1)
    
End Sub

' // Draw grid
Private Sub RenderGrid()
    RenderLines cGrid, color(1, 1, 1)
End Sub

' // Render lines from vertex buffer
Private Sub RenderLines( _
            ByVal cBuffer As IDirect3DVertexBuffer9, _
            ByRef tColor As D3DCOLORVALUE)
    Dim tDesc           As D3DVERTEXBUFFER_DESC
    Dim tMat            As D3DMATERIAL9
    Dim lOldLightState  As Long
    
    cBuffer.GetDesc tDesc
    
    cD3dev.SetStreamSource 0, cBuffer, 0, 12
    cD3dev.SetFVF tDesc.FVF
    
    tMat.Diffuse = tColor
    
    cD3dev.SetMaterial tMat
    
    lOldLightState = cD3dev.GetRenderState(D3DRS_LIGHTING)
    
    ' // Disable lighting
    cD3dev.SetRenderState D3DRS_LIGHTING, 0
    
    cD3dev.DrawPrimitive D3DPT_LINELIST, 0, tDesc.Size / 12
    
    cD3dev.SetRenderState D3DRS_LIGHTING, lOldLightState
        
End Sub

' // Initialize Direct3D device
Private Sub InitDevice( _
            ByVal hwnd As Long)
    Dim tPP As D3DPRESENT_PARAMETERS

    If cD3d9 Is Nothing Then
        Set cD3d9 = Direct3DCreate9()
    End If
    
    With tPP
    
        .BackBufferCount = 1
        .Windowed = 1
        .BackBufferFormat = D3DFMT_X8R8G8B8
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        
    End With
    
    ' // Create device
    Set cD3dev = cD3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, _
                                    D3DCREATE_HARDWARE_VERTEXPROCESSING, tPP)
    
    ' // Enable Z_buffer
    cD3dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    cD3dev.SetRenderState D3DRS_LIGHTING, 1
    
    hWindow = hwnd
    
End Sub

' // Create 2 omni-lights
Private Sub InitLight()
    Dim tLight  As D3DLIGHT9

    With tLight
    
        .Type = D3DLIGHT_POINT
        .Position = vec3(-10, 50, 20)
        .Attenuation1 = 0.03
        .Range = 1000
        .Diffuse = color(1, 1, 1)
        .Specular = color(1, 1, 1)
        .Ambient = color(0.5, 0.5, 0.5)

        cD3dev.SetRenderState D3DRS_LIGHTING, 1
        cD3dev.SetLight 0, tLight
        cD3dev.LightEnable 0, 1
    
        .Position = vec3(10, -5, -40)
        .Range = 100000
        .Diffuse = color(0.3, 0.3, 0.3)
        .Ambient = color(0.3, 0.3, 0.3)
        
        cD3dev.SetLight 1, tLight
        cD3dev.LightEnable 1, 1

    End With

End Sub

' // Initialize camera
Private Sub InitCamera()
    Dim cCam    As CCamera
    
    Set cCam = New CCamera
    
    cCam.LookAt vec3(0, 3, -20), vec3(0, 0, 0), vec3(0, 1, 0)
    
    Set mcCamera = cCam
    
    Set cCameraTarget = New CMesh
    
    cCameraTarget.CreateSphere cD3dev, 16, 0.5, vbCyan
    cCameraTarget.Material.SpecularPower = 0
    cCameraTarget.Material.SpecularColor = 0
    
    UpdateCamera

End Sub

' // Update view and projection transformations
Private Sub UpdateCamera()
    Dim tRC     As RECT
    Dim tMtx    As D3DMATRIX
    
    On Error GoTo error_handler

    cD3dev.SetTransform D3DTS_VIEW, mcCamera.ViewMatrix
    
    GetClientRect hWindow, tRC
    
    D3DXMatrixPerspectiveFovLH tMtx, mcCamera.FOV, tRC.iRight / tRC.iBottom, 0.1, 100
    
    cD3dev.SetTransform D3DTS_PROJECTION, tMtx
    
    cCameraTarget.Position = mcCamera.Target

error_handler:

End Sub

' // Create transform gizmo (3-arrows with directions)
Private Sub InitTransformGizmo()
    Dim tPoint()    As D3DVECTOR
    Dim tVertex()   As tColorVertex
    Dim tDirection  As D3DVECTOR
    Dim tMtx        As D3DMATRIX
    Dim lIndex      As Long
    Dim lSideIndex  As Long
    Dim lPointIndex As Long
    Dim lPrevIndex  As Long
    Dim lBufferSize As Long
    Dim lArrowPart  As Long
    Dim pData       As Long
    
    ' // Arrow is cone and cylinder
    Const ARROW_LINE_RADIUS = 0.05
    Const ARROW_CONE_RADIUS = 0.1
    Const ARROW_LINE_HEIGHT = 1
    Const ARROW_CONE_HEIGHT = 0.3
    Const ARROW_SIDES = 8
    
    ReDim tPoint(ARROW_SIDES * 4 - 1)
    ReDim tVertex(ARROW_SIDES * 5 * 3 - 1)
    
    For lSideIndex = 0 To ARROW_SIDES
        
        lPointIndex = lSideIndex * 4
        lPrevIndex = lPointIndex - 4
        
        If lSideIndex < ARROW_SIDES Then
            
            tDirection = vec3(Cos(PI * 2 * lSideIndex / ARROW_SIDES), 0, Sin(PI * 2 * lSideIndex / ARROW_SIDES))
            
            tPoint(lPointIndex) = vec3(tDirection.X * ARROW_LINE_RADIUS, 0, tDirection.Z * ARROW_LINE_RADIUS)
            tPoint(lPointIndex + 1) = vec3(tDirection.X * ARROW_LINE_RADIUS, _
                                        ARROW_LINE_HEIGHT, tDirection.Z * ARROW_LINE_RADIUS)
            tPoint(lPointIndex + 2) = vec3(tDirection.X * ARROW_CONE_RADIUS, _
                                        ARROW_LINE_HEIGHT, tDirection.Z * ARROW_CONE_RADIUS)
            tPoint(lPointIndex + 3) = vec3(0, ARROW_LINE_HEIGHT + ARROW_CONE_HEIGHT, 0)

        Else
            lPointIndex = 0
        End If
        
        If lSideIndex Then
            
            For lArrowPart = 0 To 1
            
                tVertex(lIndex).tPos = tPoint(lPointIndex):      lIndex = lIndex + 1
                tVertex(lIndex).tPos = tPoint(lPrevIndex):       lIndex = lIndex + 1
                tVertex(lIndex).tPos = tPoint(lPointIndex + 1):  lIndex = lIndex + 1
                tVertex(lIndex).tPos = tVertex(lIndex - 1).tPos: lIndex = lIndex + 1
                tVertex(lIndex).tPos = tVertex(lIndex - 3).tPos: lIndex = lIndex + 1
                tVertex(lIndex).tPos = tPoint(lPrevIndex + 1):   lIndex = lIndex + 1
                
                lPointIndex = lPointIndex + 1
                lPrevIndex = lPrevIndex + 1
                
            Next
            
            tVertex(lIndex).tPos = tPoint(lPointIndex + 1):  lIndex = lIndex + 1
            tVertex(lIndex).tPos = tPoint(lPointIndex):      lIndex = lIndex + 1
            tVertex(lIndex).tPos = tPoint(lPrevIndex):       lIndex = lIndex + 1
            
        End If
        
    Next

    lBufferSize = (UBound(tVertex) + 1) * Len(tVertex(0)) * 3
    
    cD3dev.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_DIFFUSE, D3DPOOL_DEFAULT, cTransformGizmo
    
    cTransformGizmo.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    
    For lIndex = 0 To UBound(tVertex)
        tVertex(lIndex).lColor = &HFF
    Next
    
    memcpy ByVal pData, tVertex(0), lBufferSize \ 3
    
    D3DXMatrixRotationX tMtx, PI / 2

    For lIndex = 0 To UBound(tVertex)
    
        tVertex(lIndex).lColor = &HFF00&
        D3DXVec3TransformCoord tVertex(lIndex).tPos, tVertex(lIndex).tPos, tMtx
        
    Next
    
    memcpy ByVal pData + ((UBound(tVertex) + 1) * Len(tVertex(0))), tVertex(0), lBufferSize \ 3
    
    D3DXMatrixRotationY tMtx, PI / 2

    For lIndex = 0 To UBound(tVertex)
    
        tVertex(lIndex).lColor = &HFF0000
        D3DXVec3TransformCoord tVertex(lIndex).tPos, tVertex(lIndex).tPos, tMtx
        
    Next
    
    memcpy ByVal pData + ((UBound(tVertex) + 1) * Len(tVertex(0)) * 2), tVertex(0), lBufferSize \ 3
    
    cTransformGizmo.Unlock
    
End Sub

' // Initialize selection box
Private Sub InitSelectionBox()
    Dim lNumOfPoints    As Long
    Dim tAxis(2)        As D3DVECTOR
    Dim tLocalAxis(2)   As D3DVECTOR
    Dim tOutput(47)     As D3DVECTOR
    Dim tMtx1           As D3DMATRIX
    Dim tMtx2           As D3DMATRIX
    Dim lIndex          As Long
    Dim lSideIndex      As Long
    Dim pData           As Long
    Dim lBufferSize     As Long
    
    ' // Create initialization vectors axis
    tAxis(0) = vec3(0.1, 0, 0):  tAxis(1) = vec3(0, 0.1, 0):  tAxis(2) = vec3(0, 0, 0.1)
    
    D3DXMatrixIdentity tMtx1
    
    For lSideIndex = 0 To 7
        
        ' // Modular arithmetic to use proper transformation
        D3DXMatrixTranslation tMtx1, (lSideIndex And 4) \ 4, (lSideIndex And 2) \ 2, lSideIndex And 1
        D3DXMatrixRotationY tMtx2, PI * ((lSideIndex And 4) / (4 / 1.5) + _
                                  (lSideIndex And 1) * 0.5 * (-(lSideIndex And 4) \ 2 + 1))
        D3DXMatrixMultiply tMtx1, tMtx2, tMtx1
        D3DXMatrixRotationX tMtx2, (PI * ((lSideIndex And 2) \ 2) * 0.5)
        D3DXMatrixMultiply tMtx1, tMtx2, tMtx1
        
        TransformVectors tLocalAxis, tAxis, tMtx1
        
        tOutput(lIndex) = vec3(tMtx1.m41, tMtx1.m42, tMtx1.m43):    lIndex = lIndex + 1
        tOutput(lIndex) = tLocalAxis(0):                            lIndex = lIndex + 1
        tOutput(lIndex) = tOutput(lIndex - 2):                      lIndex = lIndex + 1
        tOutput(lIndex) = tLocalAxis(1):                            lIndex = lIndex + 1
        tOutput(lIndex) = tOutput(lIndex - 4):                      lIndex = lIndex + 1
        tOutput(lIndex) = tLocalAxis(2):                            lIndex = lIndex + 1
        
    Next
    
    lNumOfPoints = 6 * 8
    lBufferSize = lNumOfPoints * 12
    
    cD3dev.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ, D3DPOOL_DEFAULT, cSelectionBox
    
    cSelectionBox.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tOutput(0), lBufferSize
    cSelectionBox.Unlock
    
End Sub

' // Initialize grid (just crossed lines)
Private Sub InitGrid()
    Dim tGridPoints()   As D3DVECTOR
    Dim lNumOfPoints    As Long
    Dim lIndex          As Long
    Dim lBufferSize     As Long
    Dim pData           As Long
    Dim fOffset         As Single
    
    lNumOfPoints = 10
    fOffset = (lNumOfPoints - 1) / 2
    
    ReDim tGridPoints(lNumOfPoints * 2 * 2 - 1)
    
    lBufferSize = Len(tGridPoints(0)) * (UBound(tGridPoints) + 1)
    
    For lIndex = 0 To lNumOfPoints - 1

        tGridPoints(lIndex * 4) = vec3(-fOffset + lIndex, 0, -fOffset)
        tGridPoints(lIndex * 4 + 1) = vec3(-fOffset + lIndex, 0, fOffset)
        tGridPoints(lIndex * 4 + 2) = vec3(-fOffset, 0, -fOffset + lIndex)
        tGridPoints(lIndex * 4 + 3) = vec3(fOffset, 0, -fOffset + lIndex)

    Next
    
    cD3dev.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ, D3DPOOL_DEFAULT, cGrid
    
    cGrid.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tGridPoints(0), lBufferSize
    cGrid.Unlock
    
End Sub

' // Transform vectors using specified matrix
Private Sub TransformVectors( _
            ByRef tOut() As D3DVECTOR, _
            ByRef tIn() As D3DVECTOR, _
            ByRef tMtx As D3DMATRIX)
    Dim lIndex  As Long
    
    For lIndex = 0 To UBound(tIn)
        D3DXVec3TransformCoord tOut(lIndex), tIn(lIndex), tMtx
    Next
    
End Sub

Private Sub Class_Initialize()
    Set mcObjects = New Collection
End Sub

Private Sub mcCamera_Changed()
    UpdateCamera
End Sub
