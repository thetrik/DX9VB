VERSION 5.00
Begin VB.Form frmLaserLines 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LaserLines"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   344
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbColor 
      Height          =   315
      Left            =   180
      Max             =   15
      TabIndex        =   0
      Top             =   4020
      Value           =   1
      Width           =   4755
   End
End
Attribute VB_Name = "frmLaserLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X               As Long
    Y               As Long
End Type
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
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO, _
                         ByVal wUsage As Long) As Long
Private Declare Function memcpy Lib "kernel32" _
                         Alias "RtlMoveMemory" ( _
                         ByRef Destination As Any, _
                         ByRef Source As Any, _
                         ByVal length As Long) As Long
Private Declare Function GetCursorPos Lib "user32" ( _
                         ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" ( _
                         ByVal hwnd As Long, _
                         ByRef lpPoint As POINTAPI) As Long

Private Const LaserWidth As Long = 20

Private Type Vertex
    position    As D3DVECTOR
    rhw         As Single
    color       As Long
    tu          As Single
    tv          As Single
End Type

Dim vFlag       As D3DFVF
Dim d3d9        As IDirect3D9
Dim d3dev       As IDirect3DDevice9
Dim vtxBuf      As IDirect3DVertexBuffer9
Dim texture     As IDirect3DTexture9
Dim vert(29)    As Vertex

Private Sub Form_Load()

    If Not Initialize() Then
        MsgBox "error"
        Unload Me
    End If
    
    ' // Load background texture
    Set texture = LoadTextureFromFile(App.Path & "\space.jpg")
    
    ' // Create vertex buffer
    d3dev.CreateVertexBuffer (UBound(vert) + 1) * Len(vert(0)), D3DUSAGE_DYNAMIC, vFlag, D3DPOOL_DEFAULT, vtxBuf
    
    ' // Select vertex buffer
    d3dev.SetStreamSource 0, vtxBuf, 0, Len(vert(0))
    
    CreateBackgroundSprite
    
End Sub

Private Sub Render()

    ' // Clear background
    d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbRed, 1, 0
    
    d3dev.BeginScene
    
    ' // Apply texture
    d3dev.SetTexture 0, texture
    
    d3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    ' // Draw background
    d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    
    ' // Remove texture
    d3dev.SetTexture 0, Nothing
    
    CreateLasersSprites
    
    d3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    ' // Draw lasers
    d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 6, 8
    
    d3dev.EndScene
    
    d3dev.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub CreateLasersSprites()
    Dim pos     As POINTAPI
    Dim pos2    As D3DVECTOR2
    Dim ptr     As Long
    Dim tmp     As Single
    Dim length  As Single
    Dim col     As Long
    
    col = QBColor(hsbColor.Value)
    
    ' // Get cursor pos
    GetCursorPos pos
    ScreenToClient Me.hwnd, pos
    
    ' // Calculate perpendicular
    pos2.X = pos.X
    pos2.Y = pos.Y - ScaleHeight \ 2
    
    length = Sqr(pos2.X * pos2.X + pos2.Y * pos2.Y)
    
    tmp = pos2.X
    pos2.X = pos2.Y / length
    pos2.Y = -tmp / length
    
    vert(6) = vtx2D(pos2.X * LaserWidth, pos2.Y * LaserWidth + ScaleHeight \ 2, 0.5, vbBlack, 0, 0)
    vert(7) = vtx2D(pos2.X * LaserWidth \ 2 + pos.X, pos2.Y * LaserWidth \ 2 + pos.Y, 0.5, vbBlack, 0, 0)
    vert(8) = vtx2D(0, ScaleHeight \ 2, 0.5, col, 0, 0)
    
    vert(9) = vert(8)
    vert(10) = vert(7)
    vert(11) = vtx2D(pos.X, pos.Y, 0.5, col, 0, 0)

    vert(12) = vert(9)
    vert(13) = vert(11)
    vert(14) = vtx2D(-pos2.X * LaserWidth, -pos2.Y * LaserWidth + ScaleHeight \ 2, 0.5, vbBlack, 0, 0)
    
    vert(15) = vert(14)
    vert(16) = vert(13)
    vert(17) = vtx2D(-pos2.X * LaserWidth \ 2 + pos.X, -pos2.Y * LaserWidth \ 2 + pos.Y, 0.5, vbBlack, 0, 0)

    ' // Calculate perpendicular
    pos2.X = pos.X - ScaleWidth
    pos2.Y = pos.Y - ScaleHeight \ 2
    
    length = Sqr(pos2.X * pos2.X + pos2.Y * pos2.Y)
    
    tmp = pos2.X
    pos2.X = pos2.Y / length
    pos2.Y = -tmp / length
    
    vert(18) = vtx2D(pos2.X * LaserWidth + ScaleWidth, pos2.Y * LaserWidth + ScaleHeight \ 2, 0.5, vbBlack, 0, 0)
    vert(19) = vtx2D(pos2.X * LaserWidth \ 2 + pos.X, pos2.Y * LaserWidth \ 2 + pos.Y, 0.5, vbBlack, 0, 0)
    vert(20) = vtx2D(ScaleWidth, ScaleHeight \ 2, 0.5, col, 0, 0)
    
    vert(21) = vert(20)
    vert(22) = vert(19)
    vert(23) = vtx2D(pos.X, pos.Y, 0.5, col, 0, 0)

    vert(24) = vert(21)
    vert(25) = vert(23)
    vert(26) = vtx2D(-pos2.X * LaserWidth + ScaleWidth, -pos2.Y * LaserWidth + ScaleHeight \ 2, 0.5, vbBlack, 0, 0)
    
    vert(27) = vert(26)
    vert(28) = vert(25)
    vert(29) = vtx2D(-pos2.X * LaserWidth \ 2 + pos.X, -pos2.Y * LaserWidth \ 2 + pos.Y, 0.5, vbBlack, 0, 0)
    
    vtxBuf.Lock 6 * Len(vert(0)), 24 * Len(vert(0)), ptr, 0&
    memcpy ByVal ptr, vert(6), 24 * Len(vert(0))
    vtxBuf.Unlock
    
End Sub

Private Sub CreateBackgroundSprite()
    Dim ptr As Long
    
    ' // Create background sprite
    vert(0) = vtx2D(0, 0, 0.1, vbWhite, 0, 0)
    vert(1) = vtx2D(ScaleWidth, 0, 0.1, vbWhite, 1, 0)
    vert(2) = vtx2D(ScaleWidth, ScaleHeight, 0.1, vbWhite, 1, 1)
    vert(3) = vtx2D(0, 0, 0.1, vbWhite, 0, 0)
    vert(4) = vtx2D(ScaleWidth, ScaleHeight, 0.1, vbWhite, 1, 1)
    vert(5) = vtx2D(0, ScaleHeight, 0.1, vbWhite, 0, 1)
    
    vtxBuf.Lock 0, 6 * Len(vert(0)), ptr, 0&
    memcpy ByVal ptr, vert(0), 6 * Len(vert(0))
    vtxBuf.Unlock
    
End Sub

Private Function Initialize() As Boolean
    
    On Error GoTo error_handler
    
    ' // Create IDirect3D9 object
    Set d3d9 = Direct3DCreate9()
    
    Dim pP  As D3DPRESENT_PARAMETERS
    ' // Set vertex format
    vFlag = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
    
    pP.BackBufferCount = 1
    pP.Windowed = 1
    pP.BackBufferFormat = D3DFMT_A8R8G8B8
    pP.SwapEffect = D3DSWAPEFFECT_DISCARD
    pP.EnableAutoDepthStencil = 1
    pP.AutoDepthStencilFormat = D3DFMT_D16
    
    ' // Create device
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, pP)
    
    ' // Set format
    d3dev.SetFVF vFlag
    d3dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    d3dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR
    d3dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    d3dev.SetRenderState D3DRS_BLENDOP, D3DBLENDOP_ADD
    
    Initialize = True
    
error_handler:

End Function

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

' // Fast vector creation
Private Function vtx2D(ByVal X As Single, _
                       ByVal Y As Single, _
                       ByVal rhw As Single, _
                       ByVal col As Long, _
                       ByVal tu As Single, _
                       ByVal tv As Single) As Vertex
    vtx2D.position.X = X:   vtx2D.position.Y = Y:   vtx2D.rhw = rhw
    vtx2D.color = col:      vtx2D.tu = tu:          vtx2D.tv = tv
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Render
End Sub

Private Sub Form_Paint()
    Render
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set vtxBuf = Nothing
    Set texture = Nothing
    Set d3dev = Nothing
    Set d3d9 = Nothing
    
End Sub

