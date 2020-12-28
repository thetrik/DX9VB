VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw text by The trick"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDisp 
      BackColor       =   &H00FFFFFF&
      Height          =   4305
      Left            =   135
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   682
      TabIndex        =   1
      Top             =   105
      Width           =   10290
      Begin VB.Timer tmrFrame 
         Interval        =   1000
         Left            =   2880
         Top             =   3705
      End
   End
   Begin VB.TextBox txtText 
      Height          =   855
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   10335
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

Dim vFlag       As D3DFVF
Dim d3d9        As IDirect3D9
Dim d3dev       As IDirect3DDevice9
Dim vtxBuf      As IDirect3DVertexBuffer9
Dim texture     As IDirect3DTexture9
Dim IsStop      As Boolean
Dim FPS         As Long
Dim triCount    As Long
Dim maxLine     As Single
Dim ctLine      As Long

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
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, picDisp.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, pP)
    ' // Set states
    d3dev.SetRenderState D3DRS_ZENABLE, D3DZB_TRUE
    d3dev.SetRenderState D3DRS_LIGHTING, 0
    d3dev.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    ' // Alpha blending
    d3dev.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    d3dev.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    d3dev.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    d3dev.SetRenderState D3DRS_BLENDOP, D3DBLENDOP_ADD
    ' // Set format
    d3dev.SetFVF vFlag
    ' // Init matrices
    Dim Mtx As D3DMATRIX
    ' // Create view matrix
    D3DXMatrixLookAtLH Mtx, vec3(0, 0, -5), vec3(0, 0, 0), vec3(0, 1, 0)
    d3dev.SetTransform D3DTS_VIEW, Mtx
    ' // Create projection matrix
    D3DXMatrixPerspectiveFovLH Mtx, PI / 3, picDisp.ScaleWidth / picDisp.ScaleHeight, 0.1, 100
    d3dev.SetTransform D3DTS_PROJECTION, Mtx
    ' // Create texture
    Set texture = LoadTextureFromFile(App.Path & "\Font.bmp")
    d3dev.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    d3dev.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    ' // Apply texture
    d3dev.SetTexture 0, texture
    ' // Create text
    txtText.Text = "directx9 text" & vbNewLine & "demonstration" & vbNewLine & "visual basic6" & vbNewLine & "by the trick" & vbNewLine & "    2015"

    Me.Show
    
    Do

        ' // Create transformation for a text (center)
        D3DXMatrixTranslation Mtx, -maxLine / 2, ctLine / 2 - 1, 0

        d3dev.SetTransform D3DTS_WORLD, Mtx
        ' // Clear background
        d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbWhite, 1, 0
        
        d3dev.BeginScene
        
        d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, triCount
        
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

' // Create quads from text
Private Function CreateTextSurface(Text As String) As IDirect3DVertexBuffer9
    Dim tmpStr  As String
    Dim dx      As Single
    Dim dy      As Single
    Dim fx      As Single
    Dim fy      As Single
    Dim index   As Long
    Dim texIdx  As Long
    Dim vtxIdx  As Long
    Dim vert()  As Vertex
    Dim curLine As Single
    
    ' // Because we have only upper case symbols on texture
    tmpStr = UCase(Text)
    maxLine = 0
    triCount = 0
    ctLine = 0
    
    If Len(tmpStr) = 0 Then Exit Function
    
    ReDim vert(6 * Len(tmpStr) - 1)
    
    ctLine = 1
    
    For index = 1 To Len(tmpStr)
            
        texIdx = Asc(Mid$(tmpStr, index, 1))
        
        Select Case texIdx
        Case &HD
            ' // New line
            dy = dy - 1: dx = -1: texIdx = -1
            index = index + 1
            If curLine > maxLine Then maxLine = curLine
            curLine = -1
            ctLine = ctLine + 1
        Case &H30 To &H39
            ' // Digits
            texIdx = texIdx - 48
        Case &H41 To &H5A
            ' // Letters
            texIdx = texIdx - 55
        Case Else
            ' // Skip
            texIdx = -1
        End Select
        
        If texIdx >= 0 And texIdx < 36 Then
            ' // Get texture coordinates
            fx = (texIdx Mod 6) / 6
            fy = (texIdx \ 6) / 6
            ' // Create quad
            nPlan vec3(dx, dy, 0), vec3(dx + 1, dy, 0), vec3(dx + 1, dy + 1, 0), vec3(dx, dy + 1, 0), vtxIdx, vert(), fx, fy, fx + 1 / 6, fy + 1 / 6
        
            triCount = triCount + 2
        
        End If
            
        dx = dx + 1
        curLine = curLine + 1
        
    Next
    
    If curLine > maxLine Then maxLine = curLine
    
    ReDim Preserve vert(vtxIdx - 1)
    
    d3dev.CreateVertexBuffer Len(vert(0)) * (UBound(vert) + 1), D3DUSAGE_NONE, vFlag, D3DPOOL_DEFAULT, CreateTextSurface
    
    CreateTextSurface.Lock 0, 0, index, 0
    memcpy ByVal index, vert(0), LenB(vert(0)) * (UBound(vert) + 1)
    CreateTextSurface.Unlock
    
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

' // Add quad to buffer
Private Sub nPlan(p1 As D3DVECTOR, _
                  p2 As D3DVECTOR, _
                  p3 As D3DVECTOR, _
                  p4 As D3DVECTOR, _
                  i As Long, _
                  ret() As Vertex, _
                  ByVal u1 As Single, _
                  ByVal v1 As Single, _
                  ByVal u2 As Single, _
                  ByVal v2 As Single)
                       
    ret(i).position = p3: ret(i).tu = u2: ret(i).tv = v1: i = i + 1
    ret(i).position = p2: ret(i).tu = u2: ret(i).tv = v2: i = i + 1
    ret(i).position = p1: ret(i).tu = u1: ret(i).tv = v2: i = i + 1
    ret(i).position = p4: ret(i).tu = u1: ret(i).tv = v1: i = i + 1
    ret(i).position = p3: ret(i).tu = u2: ret(i).tv = v1: i = i + 1
    ret(i).position = p1: ret(i).tu = u1: ret(i).tv = v2: i = i + 1
    
End Sub

' // Fast vector creation
Private Function vec3(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.z = z
End Function

Private Sub Form_Unload(Cancel As Integer)
    IsStop = True
End Sub

Private Sub tmrFrame_Timer()
    Caption = "Draw text demo by The trick. FPS:" & FPS
    FPS = 0
End Sub

Private Sub txtText_Change()
    
    ' // Create cube
    Set vtxBuf = CreateTextSurface(txtText)
    ' // Select vertex buffer
    If Not vtxBuf Is Nothing Then d3dev.SetStreamSource 0, vtxBuf, 0, 5 * 4
    
End Sub
