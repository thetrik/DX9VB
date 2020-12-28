VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fire demo by The trick"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFrame 
      Interval        =   1000
      Left            =   3360
      Top             =   4920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" ( _
                         ByVal aHDC As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         lpBits As Any, _
                         lpBI As BITMAPINFO, _
                         ByVal wUsage As Long) As Long

Private Type Vertex
    position    As D3DVECTOR
    tu          As Single
    tv          As Single
End Type

Private Type Particle
    quad(5)     As Vertex
    birth       As Single
    dir         As D3DVECTOR
    transform   As D3DMATRIX
End Type

Private Const MAX_PARTICLES As Long = 100

Dim vFlag   As D3DFVF
Dim d3d9    As IDirect3D9
Dim d3dev   As IDirect3DDevice9
Dim vtxBuf  As IDirect3DVertexBuffer9
Dim texture As IDirect3DTexture9
Dim IsStop  As Boolean
Dim FPS     As Long
Dim part()  As Particle
Dim partCt  As Long

Private Sub Form_Load()
    ' // Create IDirect3D9 object
    Set d3d9 = Direct3DCreate9()
    
    Dim pP  As D3DPRESENT_PARAMETERS
    ' // Set vertex format
    vFlag = D3DFVF_XYZ Or D3DFVF_TEX1
    
    pP.BackBufferCount = 1
    pP.Windowed = 1
    pP.BackBufferFormat = D3DFMT_A8R8G8B8
    pP.SwapEffect = D3DSWAPEFFECT_DISCARD
    pP.EnableAutoDepthStencil = 1
    pP.AutoDepthStencilFormat = D3DFMT_D16
    pP.PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    ' // Create device
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, pP)
    ' // Enable Z_buffer
    d3dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    d3dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    d3dev.SetRenderState D3DRS_LIGHTING, 0
    d3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    d3dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    d3dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    d3dev.SetRenderState D3DRS_BLENDOP, D3DBLENDOP_ADD
  
    ' // Create vertex buffer
    d3dev.CreateVertexBuffer MAX_PARTICLES * 6 * 5 * 4, D3DUSAGE_DYNAMIC, vFlag, D3DPOOL_DEFAULT, vtxBuf
    
    ' // Init matrices
    Dim mtx As D3DMATRIX
    ' // Create view matrix
    D3DXMatrixLookAtLH mtx, vec3(0, 5, -10), vec3(0, 2, 0), vec3(0, 1, 0)
    d3dev.SetTransform D3DTS_VIEW, mtx
    ' // Create projection matrix
    D3DXMatrixPerspectiveFovLH mtx, PI / 3, ScaleWidth / ScaleHeight, 0.1, 100
    d3dev.SetTransform D3DTS_PROJECTION, mtx
    ' // Select vertex buffer
    d3dev.SetStreamSource 0, vtxBuf, 0, 5 * 4
    ' // Set format
    d3dev.SetFVF vFlag
    
    ' // Create texture
    Set texture = LoadTextureFromFile(App.Path & "\Texture.bmp")
    d3dev.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    d3dev.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR

    ' // Apply texture
    d3dev.SetTexture 0, texture
    
    Dim index   As Long
    Dim prev    As Single
    
    ReDim part(MAX_PARTICLES - 1)
    
    Me.Show
    prev = Timer
        
    Do
        
        If partCt < 50 And Timer - prev > 0.03 Then
            prev = Timer
            AddParticle partCt
        End If
        
        ProcessParticle
        
        ' // Clear background
        d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
        
        d3dev.BeginScene
        
        d3dev.SetRenderState D3DRS_ZWRITEENABLE, 0
        
        For index = 0 To partCt - 1
            d3dev.SetTransform D3DTS_WORLD, part(index).transform
            d3dev.DrawPrimitive D3DPT_TRIANGLELIST, index * 6, 2
        Next
        
        d3dev.SetRenderState D3DRS_ZWRITEENABLE, 1
        
        d3dev.EndScene
        
        d3dev.Present ByVal 0, ByVal 0, 0, ByVal 0

        FPS = FPS + 1
        
        DoEvents
        
    Loop Until IsStop
    
    ' // Free resources
    Set texture = Nothing
    Set vtxBuf = Nothing
    Set d3dev = Nothing
    Set d3d9 = Nothing
    
    Unload Me
    
End Sub

' // Load texture from file
Private Function LoadTextureFromFile(FileName As String) As IDirect3DTexture9
    Dim tex     As StdPicture
    Dim bi      As BITMAPINFO
    Dim RECT    As D3DLOCKED_RECT
    
    Set tex = LoadPicture(FileName)
    
    bi.bmiHeader.biSize = Len(bi.bmiHeader)
    GetDIBits Me.hDC, tex.Handle, 0, 0, ByVal 0&, bi, 0
    ' // Fix values
    bi.bmiHeader.biBitCount = 32
    bi.bmiHeader.biCompression = 0
    If bi.bmiHeader.biHeight > 0 Then bi.bmiHeader.biHeight = -bi.bmiHeader.biHeight
    ' // Create texture
    d3dev.CreateTexture bi.bmiHeader.biWidth, -bi.bmiHeader.biHeight, 1, D3DUSAGE_DYNAMIC, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, LoadTextureFromFile
    ' // Lock texture
    LoadTextureFromFile.LockRect 0, RECT, ByVal 0, 0
    ' // Get picture data to texture directly
    GetDIBits Me.hDC, tex.Handle, 0, -bi.bmiHeader.biHeight, ByVal RECT.pBits, bi, 0
    ' // Update
    LoadTextureFromFile.UnlockRect 0
    ' // Free
    Set tex = Nothing
End Function

' // Add a particle into buffer
Private Function AddParticle(ByVal index As Long) As Long
    Dim obj As Particle
    Dim mtx As D3DMATRIX
    Dim idx As Long
    Dim ptr As Long
    Dim dX  As Long
    Dim dY  As Long
    
    Randomize
    
    If index >= MAX_PARTICLES Then Stop
    
    obj.birth = Timer
    obj.dir = vec3(Rnd * 2 - 1, Rnd * 2 - 1, Rnd * 2 - 1)
    dX = Rnd * 4
    dY = Rnd * 4
    
    nPlan vec3(-1, 1, 0), vec3(-1, -1, 0), vec3(1, -1, 0), vec3(1, 1, 0), obj.quad(), 0.25 * dX, 0.25 * dY, 0.25 * dX + 0.25, 0.25 * dY + 0.25
    ' // Random rotation
    D3DXMatrixRotationY mtx, PI * Rnd * 2
    
    For idx = 0 To UBound(obj.quad)
        D3DXVec3TransformCoord obj.quad(idx).position, obj.quad(idx).position, mtx
    Next
    
    part(index) = obj
    
    vtxBuf.Lock index * 6 * 5 * 4, 6 * 5 * 4, ptr, 0
    memcpy ByVal ptr, obj.quad(0), 6 * 5 * 4
    vtxBuf.Unlock
    
    If index = partCt Then partCt = partCt + 1
    
End Function

' // Process
Private Sub ProcessParticle()
    Dim idx As Long
    Dim m1  As D3DMATRIX
    Dim m2  As D3DMATRIX
    Dim scl As Single
    Dim pos As Single
    Dim liv As Single
    
    For idx = 0 To partCt - 1
        
        liv = (Timer - part(idx).birth) / 1.5
        
        If liv > 1 Then
            AddParticle idx
        ElseIf liv < 0 Then
            liv = 0
        End If
        
        scl = Sin((liv * 9) ^ 0.7) * 3
        If scl < 0 Then scl = 0
        pos = (1 - Sin(Cos(liv * 4) * 1.5)) * 4

        D3DXMatrixTranslation m1, 0, pos, 0
        D3DXMatrixRotationY m2, (Timer - part(idx).birth) / 2
        D3DXMatrixMultiply m1, m2, m1
        D3DXMatrixTranslation m2, part(idx).dir.X * liv * 2, part(idx).dir.Y * liv, part(idx).dir.z * liv
        D3DXMatrixMultiply m1, m2, m1
        D3DXMatrixScaling m2, scl, scl * (liv + 1), scl
        D3DXMatrixMultiply m1, m2, m1
        
        part(idx).transform = m1
        
    Next
    
End Sub

' // Add quad to buffer
Private Sub nPlan(p1 As D3DVECTOR, _
                  p2 As D3DVECTOR, _
                  p3 As D3DVECTOR, _
                  p4 As D3DVECTOR, _
                  ret() As Vertex, _
                  ByVal u1 As Single, _
                  ByVal v1 As Single, _
                  ByVal u2 As Single, _
                  ByVal v2 As Single)
    Dim i   As Long
    
    ret(i).position = p3: ret(i).tu = u2: ret(i).tv = v2: i = i + 1
    ret(i).position = p1: ret(i).tu = u1: ret(i).tv = v1: i = i + 1
    ret(i).position = p2: ret(i).tu = u1: ret(i).tv = v2: i = i + 1
    ret(i).position = p3: ret(i).tu = u2: ret(i).tv = v2: i = i + 1
    ret(i).position = p4: ret(i).tu = u2: ret(i).tv = v1: i = i + 1
    ret(i).position = p1: ret(i).tu = u1: ret(i).tv = v1: i = i + 1
    
End Sub

' // Fast vector creation
Private Function vec3(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.z = z
End Function

Private Sub Form_Unload(Cancel As Integer)
    IsStop = True
End Sub

Private Sub tmrFrame_Timer()
    Caption = "Fire demo by The trick. FPS:" & FPS
    FPS = 0
End Sub
