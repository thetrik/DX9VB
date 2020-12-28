VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Landscape by The trick"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFrame 
      Interval        =   1000
      Left            =   1560
      Top             =   2880
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
    normal      As D3DVECTOR
    tu          As Single
    tV          As Single
End Type

Dim vFlag   As D3DFVF
Dim d3d9    As IDirect3D9
Dim d3dev   As IDirect3DDevice9
Dim vtxBuf  As IDirect3DVertexBuffer9
Dim idxBuf  As IDirect3DIndexBuffer9
Dim texture As IDirect3DTexture9
Dim IsStop  As Boolean
Dim FPS     As Long
Dim vtxCt   As Long
Dim idxCt   As Long

Private Sub Form_Load()
    ' // Create IDirect3D9 object
    Set d3d9 = Direct3DCreate9()
    
    Dim pP  As D3DPRESENT_PARAMETERS
    ' // Set vertex format
    vFlag = D3DFVF_XYZ Or D3DFVF_TEX1 Or D3DFVF_NORMAL
    
    pP.BackBufferCount = 1
    pP.Windowed = 1
    pP.BackBufferFormat = D3DFMT_A8R8G8B8
    pP.SwapEffect = D3DSWAPEFFECT_DISCARD
    pP.EnableAutoDepthStencil = 1
    pP.AutoDepthStencilFormat = D3DFMT_D16
    pP.PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    ' // Create device
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, pP)
    ' // Enable Z_buffer
    d3dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    ' // Enable Light
    Dim Light As D3DLIGHT9
    
    Light.Type = D3DLIGHT_POINT
    Light.position = vec3(0, 10, 0)
    Light.Attenuation1 = 0.1
    Light.Range = 100
    Light.Diffuse.r = 1:    Light.Ambient.r = 1
    Light.Diffuse.g = 1:    Light.Ambient.g = 1
    Light.Diffuse.b = 1:    Light.Ambient.b = 1
    
    d3dev.SetRenderState D3DRS_LIGHTING, 1
    d3dev.SetLight 0, Light
    d3dev.LightEnable 0, 1
    
    ' // Create landscape
    LoadLandscape App.Path & "\HeightMap.jpg", 15
    
    ' // Init matrices
    Dim Mtx As D3DMATRIX
    ' // Create view matrix
    D3DXMatrixLookAtLH Mtx, vec3(0, 1, -8), vec3(0, -4, 0), vec3(0, 1, 0)
    d3dev.SetTransform D3DTS_VIEW, Mtx
    ' // Create projection matrix
    D3DXMatrixPerspectiveFovLH Mtx, PI / 3, ScaleWidth / ScaleHeight, 0.1, 100
    d3dev.SetTransform D3DTS_PROJECTION, Mtx
    ' // Select vertex buffer
    d3dev.SetStreamSource 0, vtxBuf, 0, 8 * 4
    d3dev.SetIndices idxBuf
    
    ' // Set format
    d3dev.SetFVF vFlag
    
    ' // Create texture
    Set texture = LoadTextureFromFile(App.Path & "\Texture.jpg")
    
    d3dev.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    d3dev.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    d3dev.SetSamplerState 0, D3DSAMP_MIPFILTER, D3DTEXF_LINEAR
    
    ' // Apply texture
    d3dev.SetTexture 0, texture
    
    d3dev.SetTextureStageState 0, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_COUNT2
    
    ' // Resize texture
    D3DXMatrixScaling Mtx, 10, 10, 10
    d3dev.SetTransform D3DTS_TEXTURE0, Mtx
    
    ' // Create material
    Dim Mat As D3DMATERIAL9
    
    Mat.Diffuse.r = 1:    Mat.Ambient.r = 0
    Mat.Diffuse.g = 1:    Mat.Ambient.g = 0
    Mat.Diffuse.b = 1:    Mat.Ambient.b = 0
    
    d3dev.SetMaterial Mat
    
    ' // Fog enable
    d3dev.SetRenderState D3DRS_FOGENABLE, D3D_TRUE
    d3dev.SetRenderState D3DRS_FOGCOLOR, vbBlack
    d3dev.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
    d3dev.SetRenderState D3DRS_FOGSTART, &H40A00000
    d3dev.SetRenderState D3DRS_FOGEND, &H41A00000
    
    ' // Main cycle
    Dim ph  As Single
    
    Me.Show
    
    Do
    
        ' // Create transformation for a landscape
        D3DXMatrixRotationYawPitchRoll Mtx, Timer / 4, 0, 0
        d3dev.SetTransform D3DTS_WORLD, Mtx
        D3DXMatrixTranslation Mtx, 0, -2, 0
        d3dev.MultiplyTransform D3DTS_WORLD, Mtx
        
        ' // Clear background
        d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1, 0
        
        d3dev.BeginScene
        
        d3dev.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, vtxCt, 0, idxCt / 3
        
        d3dev.EndScene
        
        d3dev.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        FPS = FPS + 1

        DoEvents
        
    Loop Until IsStop
    
    ' // Free resources
    Set texture = Nothing
    Set vtxBuf = Nothing
    Set idxBuf = Nothing
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

' // Create a cube with specified size.
Private Function LoadLandscape(HeightMapFileName As String, ByVal ScaleFactor As Single) As Boolean
    Dim vert()  As Vertex
    Dim index() As Long
    Dim tex     As StdPicture
    Dim bi      As BITMAPINFO
    Dim dat()   As Long
    Dim X       As Single
    Dim Y       As Single
    
    Set tex = LoadPicture(HeightMapFileName)
    
    bi.bmiHeader.biSize = Len(bi.bmiHeader)
    GetDIBits Me.hDC, tex.Handle, 0, 0, ByVal 0&, bi, 0
    ' // Fix values
    bi.bmiHeader.biBitCount = 32
    bi.bmiHeader.biCompression = 0
    If bi.bmiHeader.biHeight > 0 Then bi.bmiHeader.biHeight = -bi.bmiHeader.biHeight
    ' // Alloc memory
    ReDim dat(bi.bmiHeader.biWidth - 1, Abs(bi.bmiHeader.biHeight) - 1)
    ' // Get picture data
    GetDIBits Me.hDC, tex.Handle, 0, Abs(bi.bmiHeader.biHeight), dat(0, 0), bi, 0
    ' // Alloc memory for landscape mesh
    ReDim vert(bi.bmiHeader.biWidth - 3, Abs(bi.bmiHeader.biHeight) - 3)
    ReDim index(5, (bi.bmiHeader.biWidth - 3) * (Abs(bi.bmiHeader.biHeight) - 3) - 1)
    
    Dim lr  As D3DVECTOR
    Dim tb  As D3DVECTOR
    Dim i1  As Long
    Dim i2  As Long
    
    ' // Get points (use RED channel)
    For Y = 1 To Abs(bi.bmiHeader.biHeight) - 2: For X = 1 To bi.bmiHeader.biWidth - 2
        ' // Y (height) dependent from R-value
        vert(X - 1, Y - 1).position = vec3(X / bi.bmiHeader.biWidth * ScaleFactor - ScaleFactor / 2, _
                                           (dat(X, Y) And &HFF) / 255 * ScaleFactor / 8, _
                                           Y / Abs(bi.bmiHeader.biHeight) * ScaleFactor - ScaleFactor / 2)

        vert(X - 1, Y - 1).tu = X / bi.bmiHeader.biWidth
        vert(X - 1, Y - 1).tV = Y / Abs(bi.bmiHeader.biHeight)
        ' // Calculae normal
        lr = vec3(1, ((dat(X - 1, Y) And &HFF) - (dat(X + 1, Y) And &HFF)) / 255 * ScaleFactor, 0)
        tb = vec3(0, ((dat(X, Y - 1) And &HFF) - (dat(X, Y + 1) And &HFF)) / 255 * ScaleFactor, -1)
        
        D3DXVec3Cross lr, lr, tb
        D3DXVec3Normalize vert(X - 1, Y - 1).normal, lr

        ' // Calculate index
        If Y < Abs(bi.bmiHeader.biHeight) - 2 And X < bi.bmiHeader.biWidth - 2 Then
            
            index(0, i1) = i2
            index(1, i1) = i2 + bi.bmiHeader.biWidth - 2
            index(2, i1) = i2 + 1
            index(3, i1) = index(1, i1)
            index(4, i1) = index(1, i1) + 1
            index(5, i1) = index(2, i1)
            
            i1 = i1 + 1
            
        End If
        
        i2 = i2 + 1
        
    Next: Next
    
    Dim ptr     As Long
    
    vtxCt = (bi.bmiHeader.biWidth - 2) * (Abs(bi.bmiHeader.biHeight) - 2)
    
    d3dev.CreateVertexBuffer Len(vert(0, 0)) * vtxCt, D3DUSAGE_NONE, vFlag, D3DPOOL_DEFAULT, vtxBuf
    ' // Fill values to vertex buffer
    vtxBuf.Lock 0, 0, ptr, 0
    memcpy ByVal ptr, vert(0, 0), Len(vert(0, 0)) * vtxCt
    vtxBuf.Unlock
    
    idxCt = i1 * 6
    
    d3dev.CreateIndexBuffer idxCt * Len(index(0, 0)), D3DUSAGE_DYNAMIC, D3DFMT_INDEX32, D3DPOOL_DEFAULT, idxBuf
    ' // Fill values to indexes buffer
    idxBuf.Lock 0, 0, ptr, D3DLOCK_DISCARD
    memcpy ByVal ptr, index(0, 0), idxCt * Len(index(0, 0))
    idxBuf.Unlock
    
End Function

' // Fast vector creation
Private Function vec3(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.z = z
End Function

Private Sub Form_Unload(Cancel As Integer)
    IsStop = True
End Sub

Private Sub tmrFrame_Timer()
    Caption = "Landscape demo by The trick. FPS:" & FPS
    FPS = 0
End Sub
