VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Triangle demo by The trick"
   ClientHeight    =   5475
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFrame 
      Interval        =   1000
      Left            =   1680
      Top             =   3165
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type Vertex
    position    As D3DVECTOR
    rhw         As Single
    color       As Long
End Type

Dim vFlag   As D3DFVF
Dim d3d9    As IDirect3D9
Dim d3dev   As IDirect3DDevice9
Dim vtxBuf  As IDirect3DVertexBuffer9
Dim IsStop  As Boolean
Dim FPS     As Long

Private Sub Form_Load()
    
    Set d3d9 = Direct3DCreate9()
    
    Dim pp  As D3DPRESENT_PARAMETERS
    ' // Set vertex format
    vFlag = D3DFVF_DIFFUSE Or D3DFVF_XYZRHW
    
    pp.BackBufferCount = 1
    pp.Windowed = 1
    pp.BackBufferFormat = D3DFMT_A8R8G8B8
    pp.SwapEffect = D3DSWAPEFFECT_DISCARD
    pp.PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    ' // Create device
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, pp)
    ' // Create vertex buffer
    d3dev.CreateVertexBuffer 5 * 4 * 3, 0, vFlag, D3DPOOL_DEFAULT, vtxBuf

    Dim alpha   As Single
    Dim vtx(2)  As Vertex
    Dim ptr     As Long
    Dim ca      As Single
    Dim sa      As Single
    
    ' // Main cycle
    
    Me.Show
    
    Do

        alpha = Timer
        
        ca = Cos(alpha):    sa = Sin(alpha)
        
        vtx(0).position = vec3(0 * ca + 100 * sa + 200, -100 * ca + 0 * sa + 200, 0):       vtx(0).color = vbCyan:  vtx(0).rhw = 10
        vtx(1).position = vec3(100 * ca - 100 * sa + 200, 100 * ca + 100 * sa + 200, 0):    vtx(1).color = vbGreen: vtx(1).rhw = 10
        vtx(2).position = vec3(-100 * ca - 100 * sa + 200, 100 * ca - 100 * sa + 200, 0):   vtx(2).color = vbBlue:  vtx(2).rhw = 10
    
        vtxBuf.Lock 0, Len(vtx(0)) * (UBound(vtx) + 1), ptr, 0
        memcpy ByVal ptr, vtx(0), Len(vtx(0)) * (UBound(vtx) + 1)
        vtxBuf.Unlock
    
        d3dev.SetStreamSource 0, vtxBuf, 0, (5) * 4
        d3dev.SetFVF vFlag
        d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET, vbRed, 1, 0
        
        d3dev.BeginScene
        
        d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
        
        d3dev.EndScene
        
        d3dev.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        FPS = FPS + 1
        
        DoEvents
        
    Loop Until IsStop
    
    ' // Free resources
    Set vtxBuf = Nothing
    Set d3dev = Nothing
    Set d3d9 = Nothing
    
    Unload Me
    
End Sub

Private Function vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    vec3.x = x: vec3.y = y: vec3.z = z
End Function

Private Sub Form_Unload(Cancel As Integer)
    IsStop = True
End Sub

Private Sub tmrFrame_Timer()
    Caption = "Triangle demo by The trick. FPS:" & FPS
    FPS = 0
End Sub
