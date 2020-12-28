VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cube demo by The trick."
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
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

Private Type tVertex
    tPosition   As D3DVECTOR
    tNormal     As D3DVECTOR
    fU          As Single
    fv          As Single
End Type

Private Declare Function GetDIBits Lib "gdi32" ( _
                         ByVal aHDC As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO, _
                         ByVal wUsage As Long) As Long

Private m_lFVFFlags As D3DFVF
Private m_cD3D9     As IDirect3D9
Private m_cDevice   As IDirect3DDevice9
Private m_cCubeMesh As IDirect3DVertexBuffer9
Private m_cTexture  As IDirect3DTexture9
Private m_bActive   As Boolean
Private m_lFPS      As Long

Private Sub Form_Load()
    Dim tPP     As D3DPRESENT_PARAMETERS
    Dim tLight  As D3DLIGHT9
    Dim tMtx    As D3DMATRIX
    Dim tMat    As D3DMATERIAL9
    
    ' // Create IDirect3D9 object
    Set m_cD3D9 = Direct3DCreate9()
    
    ' // Set vertex format
    m_lFVFFlags = D3DFVF_XYZ Or D3DFVF_TEX1 Or D3DFVF_NORMAL
    
    tPP.BackBufferCount = 1
    tPP.Windowed = 1
    tPP.BackBufferFormat = D3DFMT_A8R8G8B8
    tPP.SwapEffect = D3DSWAPEFFECT_DISCARD
    tPP.EnableAutoDepthStencil = 1
    tPP.AutoDepthStencilFormat = D3DFMT_D16
    tPP.PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    ' // Create device
    Set m_cDevice = m_cD3D9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, tPP)
    
    ' // Enable Z_buffer
    m_cDevice.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    
    ' // Enable Light
    tLight.Type = D3DLIGHT_POINT
    tLight.position = vec3(0, 1, -8)
    tLight.Attenuation1 = 0.1
    tLight.Range = 100
    tLight.Diffuse.r = 1:    tLight.Ambient.r = 1
    tLight.Diffuse.g = 1:    tLight.Ambient.g = 1
    tLight.Diffuse.b = 1:    tLight.Ambient.b = 1
    tLight.Attenuation1 = 0.1
    
    m_cDevice.SetRenderState D3DRS_LIGHTING, 1
    m_cDevice.SetLight 0, tLight
    m_cDevice.LightEnable 0, 1
    
    ' // Create cube
    Set m_cCubeMesh = CreateCube(2)
    
    ' // Init matrices

    ' // Create view matrix
    D3DXMatrixLookAtLH tMtx, vec3(0, 0, -5), vec3(0, 0, 0), vec3(0, 1, 0)
    m_cDevice.SetTransform D3DTS_VIEW, tMtx
    ' // Create projection matrix
    D3DXMatrixPerspectiveFovLH tMtx, PI / 3, ScaleWidth / ScaleHeight, 0.1, 10
    m_cDevice.SetTransform D3DTS_PROJECTION, tMtx
    ' // Select vertex buffer
    m_cDevice.SetStreamSource 0, m_cCubeMesh, 0, 8 * 4
    ' // Set format
    m_cDevice.SetFVF m_lFVFFlags
    
    ' // Create texture
    Set m_cTexture = LoadTextureFromFile(App.Path & "\Texture.jpg")
    
    m_cDevice.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    m_cDevice.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    
    ' // Apply texture
    m_cDevice.SetTexture 0, m_cTexture
    
    ' // Create material
    tMat.Diffuse.r = 1:    tMat.Ambient.r = 0
    tMat.Diffuse.g = 1:    tMat.Ambient.g = 0
    tMat.Diffuse.b = 1:    tMat.Ambient.b = 0.5
    
    m_cDevice.SetMaterial tMat

    Me.Show
    
    MainLoop
    
End Sub

Private Sub MainLoop()
    Dim tMtx    As D3DMATRIX
    
    m_bActive = True
    
    Do While m_bActive
    
        ' // Create transformation for a cube
        D3DXMatrixRotationYawPitchRoll tMtx, Timer, Timer / 3, Timer / 7
        
        m_cDevice.SetTransform D3DTS_WORLD, tMtx
        
        ' // Clear background
        m_cDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbRed, 1, 0
        
        m_cDevice.BeginScene
        
        m_cDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
        
        m_cDevice.EndScene
        
        m_cDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&

        m_lFPS = m_lFPS + 1

        DoEvents
        
    Loop
    
End Sub

' // Load texture from file
Private Function LoadTextureFromFile( _
                 ByRef sFileName As String) As IDirect3DTexture9
    Dim cPicture    As StdPicture
    Dim tBI         As BITMAPINFO
    Dim tRect       As D3DLOCKED_RECT
    
    Set cPicture = LoadPicture(sFileName)
    
    tBI.bmiHeader.biSize = Len(tBI.bmiHeader)
    
    GetDIBits Me.hDC, cPicture.Handle, 0, 0, ByVal 0&, tBI, 0
    
    tBI.bmiHeader.biBitCount = 32
    tBI.bmiHeader.biCompression = 0
    
    If tBI.bmiHeader.biHeight > 0 Then
        tBI.bmiHeader.biHeight = -tBI.bmiHeader.biHeight
    End If
    
    ' // Create texture
    m_cDevice.CreateTexture tBI.bmiHeader.biWidth, -tBI.bmiHeader.biHeight, 1, D3DUSAGE_DYNAMIC, _
                            D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, LoadTextureFromFile
    
    ' // Lock texture
    LoadTextureFromFile.LockRect 0, tRect, ByVal 0&, 0
    
    ' // Get picture data to texture directly
    GetDIBits Me.hDC, cPicture.Handle, 0, -tBI.bmiHeader.biHeight, ByVal tRect.pBits, tBI, 0
    
    ' // Update
    LoadTextureFromFile.UnlockRect 0

End Function

' // Create a cube with specified size.
Private Function CreateCube( _
                 ByVal fSize As Single) As IDirect3DVertexBuffer9
    Dim tVert() As tVertex
    Dim lIdx()  As Long
    Dim fH      As Single
    Dim lI      As Long
    
    fH = fSize / 2
    
    ReDim tVert(35)
    
    nPlan vec3(-fH, fH, fH), vec3(-fH, fH, -fH), vec3(fH, fH, -fH), vec3(fH, fH, fH), vec3(0, 1, 0), lI, tVert(), 0.5, 0, 1, 0.5
    nPlan vec3(fH, -fH, fH), vec3(fH, -fH, -fH), vec3(-fH, -fH, -fH), vec3(-fH, -fH, fH), vec3(0, -1, 0), lI, tVert(), 0, 0, 0.5, 0.5
    nPlan vec3(fH, fH, fH), vec3(fH, fH, -fH), vec3(fH, -fH, -fH), vec3(fH, -fH, fH), vec3(1, 0, 0), lI, tVert(), 0, 0.5, 0.5, 1
    nPlan vec3(-fH, -fH, fH), vec3(-fH, -fH, -fH), vec3(-fH, fH, -fH), vec3(-fH, fH, fH), vec3(-1, 0, 0), lI, tVert(), 0, 0.5, 0.5, 1
    nPlan vec3(-fH, fH, -fH), vec3(-fH, -fH, -fH), vec3(fH, -fH, -fH), vec3(fH, fH, -fH), vec3(0, 0, -1), lI, tVert(), 0.5, 0.5, 1, 1
    nPlan vec3(fH, -fH, fH), vec3(-fH, -fH, fH), vec3(-fH, fH, fH), vec3(fH, fH, fH), vec3(0, 0, 1), lI, tVert(), 0.5, 0.5, 1, 1
    
    m_cDevice.CreateVertexBuffer Len(tVert(0)) * (UBound(tVert) + 1), D3DUSAGE_NONE, m_lFVFFlags, D3DPOOL_MANAGED, CreateCube
    
    CreateCube.Lock 0, 0, lI, 0
    memcpy ByVal lI, tVert(0), Len(tVert(0)) * (UBound(tVert) + 1)
    CreateCube.Unlock
    
End Function

' // Add quad to buffer
Private Sub nPlan( _
            ByRef fP1 As D3DVECTOR, _
            ByRef fP2 As D3DVECTOR, _
            ByRef fP3 As D3DVECTOR, _
            ByRef fP4 As D3DVECTOR, _
            ByRef tN As D3DVECTOR, _
            ByRef lI As Long, _
            ByRef tRet() As tVertex, _
            ByVal fU1 As Single, _
            ByVal fV1 As Single, _
            ByVal fU2 As Single, _
            ByVal fV2 As Single)
                       
    tRet(lI).tPosition = fP3: tRet(lI).tNormal = tN: tRet(lI).fU = fU1: tRet(lI).fv = fV2: lI = lI + 1
    tRet(lI).tPosition = fP2: tRet(lI).tNormal = tN: tRet(lI).fU = fU1: tRet(lI).fv = fV1: lI = lI + 1
    tRet(lI).tPosition = fP1: tRet(lI).tNormal = tN: tRet(lI).fU = fU2: tRet(lI).fv = fV1: lI = lI + 1
    tRet(lI).tPosition = fP3: tRet(lI).tNormal = tN: tRet(lI).fU = fU1: tRet(lI).fv = fV2: lI = lI + 1
    tRet(lI).tPosition = fP1: tRet(lI).tNormal = tN: tRet(lI).fU = fU2: tRet(lI).fv = fV1: lI = lI + 1
    tRet(lI).tPosition = fP4: tRet(lI).tNormal = tN: tRet(lI).fU = fU2: tRet(lI).fv = fV2: lI = lI + 1
    
End Sub

' // Fast vector creation
Private Function vec3( _
                 ByVal fX As Single, _
                 ByVal fY As Single, _
                 ByVal fz As Single) As D3DVECTOR
    vec3.X = fX: vec3.Y = fY: vec3.z = fz
End Function

Private Sub Form_QueryUnload( _
            ByRef Cancel As Integer, _
            ByRef UnloadMode As Integer)
    m_bActive = False
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
            
    Set m_cCubeMesh = Nothing
    Set m_cTexture = Nothing
    Set m_cDevice = Nothing
    Set m_cD3D9 = Nothing
    
End Sub

Private Sub tmrFrame_Timer()
    Caption = "Cube demo by The trick. FPS: " & m_lFPS
    m_lFPS = 0
End Sub
