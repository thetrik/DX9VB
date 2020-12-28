VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct3D multithreading by The trick"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // Render in different thread
' // By The trick 2018 (c)
' //

Option Explicit

Dim hThread         As Long         ' // Thread handle
Dim cD3d9           As IDirect3D9
Dim sCaptionPostfix As String

' // Startup
Private Sub Form_Load()
    Dim bIsInIDE    As Boolean
    
    Debug.Assert MakeTrue(bIsInIDE) ' // Check if we are in IDE
    
    Me.Show
    
    If Not Init Then
        MsgBox "Init failed"
        Unload Me
    End If
    
    ' // If we are in IDE all works in the single thread, therefore we shouldn't call that function
    ' // because it's meaningless. In that case render procedure is called from CalcPass
    If Not bIsInIDE Then
    
        hThread = vbCreateThread(0, 0, AddressOf ThreadProc, 0, 0, 0)
        sCaptionPostfix = "(multithread)"
        
    Else
        sCaptionPostfix = "(singlethread)"
    End If
    
    With gtSharedResources
    
    ' // Main cycle
    Do Until .bEndFlag
    
        CalcPass
        Me.Caption = "Render: " & .tRenderFPS.lFPS & _
                     " Calc: " & .tCalcFPS.lFPS & " " & sCaptionPostfix
        FastDoEvents
        TickFPS .tCalcFPS
        Sleep 0
        
    Loop
    
    End With
    
    Unload Me
    
End Sub

' // Initialization
Private Function Init() As Boolean
    Dim tPP     As D3DPRESENT_PARAMETERS
    Dim cDev    As IDirect3DDevice9
    Dim tLight  As D3DLIGHT9
    Dim tMtx    As D3DMATRIX
    Dim tMat    As D3DMATERIAL9
     
    On Error GoTo error_handler
    
    ' // Initialize shared resources
    modSharedResources.Init
    
    Set cD3d9 = Direct3DCreate9()
    
    With tPP
    
    .BackBufferCount = 1
    .Windowed = 1
    .BackBufferFormat = D3DFMT_X8R8G8B8
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    .EnableAutoDepthStencil = 1
    .AutoDepthStencilFormat = D3DFMT_D16
    .PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    End With
    
    ' // Create Direct3D device
    Set cDev = cD3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hwnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, tPP)
    
    cDev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE   ' // Enable Z-Buffer
    cDev.SetRenderState D3DRS_LIGHTING, 1           ' // Enable lighting
    
    ' // Lighting setting
    With tLight
    
    .Type = D3DLIGHT_POINT
    .Position = vec3(0, 1, -8)
    .Attenuation1 = 0.1
    .Range = 100
    .Diffuse.r = 1:    .Ambient.r = 1
    .Diffuse.g = 1:    .Ambient.g = 1
    .Diffuse.b = 1:    .Ambient.b = 1
    .Attenuation1 = 0.1
    
    End With
    
    cDev.SetRenderState D3DRS_LIGHTING, 1
    cDev.SetLight 0, tLight
    cDev.LightEnable 0, 1
    
    ' // Transformations
    D3DXMatrixLookAtLH tMtx, vec3(0, 5, -8), vec3(0, 0, 0), vec3(0, 1, 0)
    cDev.SetTransform D3DTS_VIEW, tMtx
    
    D3DXMatrixPerspectiveFovLH tMtx, PI / 3, ScaleWidth / ScaleHeight, 0.1, 100
    cDev.SetTransform D3DTS_PROJECTION, tMtx
    
    D3DXMatrixIdentity tMtx
    cDev.SetTransform D3DTS_WORLD, tMtx
    
    ' // Material
    tMat.Diffuse.r = 1:    tMat.Ambient.r = 0
    tMat.Diffuse.g = 1:    tMat.Ambient.g = 0
    tMat.Diffuse.b = 1:    tMat.Ambient.b = 0
    
    cDev.SetMaterial tMat
    cDev.SetFVF D3DFVF_NORMAL Or D3DFVF_XYZ  ' // Fixed format, but for real scene you should use appropriate format for mesh
    
    Set gtSharedResources.cDevice = cDev
    
    gtSharedResources.hEvent = CreateEvent(ByVal 0&, 0, 0, 0)
    
    Init = True
    
error_handler:
    
End Function

Private Sub Form_Unload( _
            ByRef Cancel As Integer)

    gtSharedResources.bEndFlag = True   ' // Synchronization isn't required
    
    PulseEvent gtSharedResources.hEvent   ' // If render thread sleeps we should wake it up to avoid deadlock
    
    If hThread Then
        
        ' // Wait for render thread termination
        WaitForSingleObject hThread, INFINITE
        CloseHandle hThread
        
    End If
    
    modSharedResources.Uninit
    
    CloseHandle gtSharedResources.hEvent
    
End Sub
