VERSION 5.00
Begin VB.Form frm3D 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D form by the trick."
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm3D.frx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3480
      Width           =   4515
   End
   Begin VB.Timer tmrReturn 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   2040
      Top             =   780
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Simple button"
      Height          =   615
      Left            =   2460
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Simple text:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   2235
   End
End
Attribute VB_Name = "frm3D"
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

Private Type Size
    cx              As Long
    cy              As Long
End Type

Private Type Vertex
    position        As D3DVECTOR
    tu              As Single
    tv              As Single
End Type

Private Declare Function UpdateLayeredWindow Lib "User32.dll" ( _
                         ByVal hWnd As Long, _
                         ByVal hdcDst As Long, _
                         ByRef pptDst As Any, _
                         ByRef psize As Any, _
                         ByVal hdcSrc As Long, _
                         ByRef pptSrc As Any, _
                         ByVal crKey As Long, _
                         ByRef pblend As Long, _
                         ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
                         Alias "SetWindowLongA" ( _
                         ByVal hWnd As Long, _
                         ByVal nIndex As Long, _
                         ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "User32.dll" _
                         Alias "GetWindowLongA" ( _
                         ByVal hWnd As Long, _
                         ByVal nIndex As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByRef pBitmapInfo As BITMAPINFO, _
                         ByVal un As Long, _
                         ByRef lplpVoid As Long, _
                         ByVal handle As Long, _
                         ByVal dw As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" ( _
                         ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" ( _
                         ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" ( _
                         ByVal hWnd As Long, _
                         ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
                         ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByVal hObject As Long) As Long
Private Declare Function SetWindowPos Lib "user32" ( _
                         ByVal hWnd As Long, _
                         ByVal hWndInsertAfter As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal cx As Long, _
                         ByVal cy As Long, _
                         ByVal wFlags As Long) As Long
Private Declare Function PrintWindow Lib "User32.dll" ( _
                         ByVal hWnd As Long, _
                         ByVal hdcBlt As Long, _
                         ByVal nFlags As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" ( _
                         ByVal hdc As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal dx As Long, _
                         ByVal dy As Long, _
                         ByVal SrcX As Long, _
                         ByVal SrcY As Long, _
                         ByVal Scan As Long, _
                         ByVal NumScans As Long, _
                         ByRef Bits As Any, _
                         ByRef BitsInfo As BITMAPINFO, _
                         ByVal wUsage As Long) As Long
Private Declare Function GetWindowRect Lib "user32" ( _
                         ByVal hWnd As Long, _
                         ByRef lpRect As D3DRECT) As Long
Private Declare Function RedrawWindow Lib "user32" ( _
                         ByVal hWnd As Long, _
                         ByRef lprcUpdate As Any, _
                         ByVal hrgnUpdate As Long, _
                         ByVal fuRedraw As Long) As Long

Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private Const PI                As Double = 3.14159275180032
Private Const WS_EX_LAYERED     As Long = &H80000
Private Const GWL_EXSTYLE       As Long = -20
Private Const WM_MOVE           As Long = &H3
Private Const WM_NCCALCSIZE     As Long = &H83
Private Const WM_EXITSIZEMOVE   As Long = &H232
Private Const RDW_ALLCHILDREN   As Long = &H80
Private Const RDW_INVALIDATE    As Long = &H1
Private Const RDW_FRAME         As Long = &H400
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOSIZE        As Long = &H1
Private Const ULW_ALPHA         As Long = &H2
Private Const AB_32Bpp255       As Long = 33488896

Dim WithEvents WndProcHandler   As clsTrickSubclass2    ' // For subclassing
Attribute WndProcHandler.VB_VarHelpID = -1

Dim d3d9        As IDirect3D9                           ' // Direct3D9 object
Dim d3dev       As IDirect3DDevice9                     ' // Direct3D device
Dim surf        As IDirect3DSurface9                    ' // Surface for render to texture
Dim sysSurf     As IDirect3DSurface9                    ' // Surface that allow to get the render data
Dim vtxBuf      As IDirect3DVertexBuffer9               ' // Billboard
Dim texture     As IDirect3DTexture9                    ' // Texture of window
Dim backTex     As IDirect3DTexture9                    ' // Texture of render target

Dim isModify    As Boolean                  ' // If this flag set then window is rotated
Dim triggerXMin As Long                     ' // Minimum positions when the rotate is enabled
Dim triggerXMax As Long
Dim triggerYMin As Long
Dim triggerYMax As Long
Dim biWnd       As BITMAPINFO
Dim bmpShadow   As Long                     ' // Bitmap in memory, which represent the window
Dim lpBmpData   As Long                     ' // Pointer to the bmpShadow bits
Dim isInit      As Boolean                  ' // If initialized then set true

Private Sub Form_Load()
    Dim sWidth  As Long
    Dim sHeight As Long
    Dim pP      As D3DPRESENT_PARAMETERS
    
    On Error GoTo ErrorHandler
    
    ' // Subclass main window
    Set WndProcHandler = New clsTrickSubclass2
    WndProcHandler.Hook Me.hWnd
    ' // Create bitmap in memory
    biWnd.bmiHeader.biSize = Len(biWnd.bmiHeader)
    biWnd.bmiHeader.biBitCount = 32
    biWnd.bmiHeader.biHeight = -Me.Height / Screen.TwipsPerPixelY
    biWnd.bmiHeader.biWidth = Me.Width / Screen.TwipsPerPixelY
    biWnd.bmiHeader.biPlanes = 1
    
    bmpShadow = CreateDIBSection(Me.hdc, biWnd, 0, lpBmpData, 0, 0)
    
    ' // Set triggers
    sWidth = Screen.Width / Screen.TwipsPerPixelX - biWnd.bmiHeader.biWidth
    sHeight = Screen.Height / Screen.TwipsPerPixelY + biWnd.bmiHeader.biHeight
    
    triggerXMin = sWidth * (1 / 5)
    triggerXMax = sWidth - triggerXMin
    triggerYMin = sHeight * (1 / 5)
    triggerYMax = sHeight - triggerYMin

    Set d3d9 = Direct3DCreate9()

    pP.BackBufferCount = 1
    pP.Windowed = 1
    pP.BackBufferFormat = D3DFMT_A8R8G8B8
    pP.SwapEffect = D3DSWAPEFFECT_DISCARD
    pP.EnableAutoDepthStencil = 1
    pP.AutoDepthStencilFormat = D3DFMT_D16
    
    ' // Firstly we should remove the non-client area, because Direct3D should redraws the entire window.
    isModify = True
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
    ' // Create device
    Set d3dev = d3d9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, pP)
    ' // Secondly we restore the areas.
    isModify = False
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
    
    Dim vtx()   As Vertex
    Dim lpDat   As Long
    
    ' // Create the billboard which represent window
    ReDim vtx(5)
    
    nPlan vec3(-biWnd.bmiHeader.biWidth, biWnd.bmiHeader.biHeight, 0), _
          vec3(biWnd.bmiHeader.biWidth, biWnd.bmiHeader.biHeight, 0), _
          vec3(biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 0), _
          vec3(-biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 0), 0, vtx(), 0, 0, 1, 1
    
    d3dev.CreateVertexBuffer Len(vtx(0)) * (UBound(vtx) + 1), D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_TEX1, D3DPOOL_DEFAULT, vtxBuf
    
    vtxBuf.Lock 0, 0, lpDat, 0
    memcpy ByVal lpDat, vtx(0), Len(vtx(0)) * (UBound(vtx) + 1)
    vtxBuf.Unlock
    
    d3dev.SetFVF D3DFVF_XYZ Or D3DFVF_TEX1
    d3dev.SetStreamSource 0, vtxBuf, 0, 5 * 4
    
    Dim Mtx As D3DMATRIX
    Dim fov As Single
    Dim l   As Single
    
    fov = PI / 3
    ' // Calculate distance to billboard in order to fit window to the render area
    l = -biWnd.bmiHeader.biHeight * Tan(fov)
    
    D3DXMatrixLookAtLH Mtx, vec3(0, 0, -l), vec3(0, 0, 0), vec3(0, 1, 0)
    d3dev.SetTransform D3DTS_VIEW, Mtx
    D3DXMatrixPerspectiveFovLH Mtx, fov, Width / Height, 1, 10000
    d3dev.SetTransform D3DTS_PROJECTION, Mtx
    
    ' // Create the window-texture
    d3dev.CreateTexture biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 1, D3DUSAGE_DYNAMIC, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, texture
    ' // Create the render-target texture
    d3dev.CreateTexture biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, 1, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, backTex
    ' // Create the lockable surface
    d3dev.CreateOffscreenPlainSurface biWnd.bmiHeader.biWidth, -biWnd.bmiHeader.biHeight, D3DFMT_A8R8G8B8, D3DPOOL_SYSTEMMEM, sysSurf
    
    Set surf = backTex.GetSurfaceLevel(0)
    d3dev.SetRenderTarget 0, surf
    
    d3dev.SetTexture 0, texture
    ' // Disable lighting
    d3dev.SetRenderState D3DRS_LIGHTING, 0
    ' // Apply filtering
    d3dev.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    d3dev.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    ' // Remove alpha-blending
    d3dev.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
    d3dev.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_CONSTANT
    d3dev.SetTextureStageState 0, D3DTSS_CONSTANT, &HFF000000
    
    isInit = True
    
    Exit Sub
    
ErrorHandler:
    
    MsgBox "Error occurred: " & Err.Description
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' // Clean
    Set sysSurf = Nothing
    Set surf = Nothing
    Set backTex = Nothing
    Set texture = Nothing
    Set vtxBuf = Nothing
    Set d3dev = Nothing
    Set d3d9 = Nothing
    DeleteObject bmpShadow
End Sub

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

' // The smooth return to normal state
Private Sub tmrReturn_Timer()
    Dim X   As Long
    Dim Y   As Long
    Dim dx  As Single
    Dim dy  As Single
    Dim nx  As Long
    Dim ny  As Long
    
    X = Me.Left / Screen.TwipsPerPixelX
    Y = Me.Top / Screen.TwipsPerPixelY
    
    If X < triggerXMin Then
        dx = (triggerXMin - X) / triggerXMin
        If dx > 1 Then dx = 1
        nx = Sin(dx * PI / 2) * 30 + 1
    ElseIf X > triggerXMax Then
        dx = -(X - triggerXMax) / triggerXMin
        If dx < -1 Then dx = -1
        nx = Sin(dx * PI / 2) * 30 - 1
    End If
    
    If Y < triggerYMin Then
        dy = (triggerYMin - Y) / triggerYMin * (PI / 2)
        If dy > 1 Then dy = 1
        ny = Sin(dy * PI / 2) * 30 + 1
    ElseIf Y > triggerYMax Then
        dy = -(Y - triggerYMax) / triggerYMin * (PI / 2)
        If dy < -1 Then dy = -1
        ny = Sin(dy * PI / 2) * 30 - 1
    End If
    
    Me.Move Me.Left + nx * Screen.TwipsPerPixelX, Me.Top + ny * Screen.TwipsPerPixelY
    
End Sub

Private Sub WndProcHandler_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ret As Long, DefCall As Boolean)
    
    DefCall = False
    
    Select Case Msg
    Case WM_MOVE    ' // When the window was moved
        Dim X   As Integer
        Dim Y   As Integer
        Dim dx  As Single
        Dim dy  As Single
        Dim rc  As D3DRECT
        
        If (Not isInit) Then Exit Sub
        
        GetWindowRect hWnd, rc
        
        X = rc.X1
        Y = rc.Y1
        
        ' // Check "hot"-zones
        If X < triggerXMin Then
            dx = (triggerXMin - X) / triggerXMin * (PI / 2)
        ElseIf X > triggerXMax Then
            dx = -(X - triggerXMax) / triggerXMin * (PI / 2)
        End If
        
        If Y < triggerYMin Then
            dy = (triggerYMin - Y) / triggerYMin * (PI / 2)
        ElseIf Y > triggerYMax Then
            dy = -(Y - triggerYMax) / triggerYMin * (PI / 2)
        End If
        
        ' // If the window is at the "hot" zone.
        If dx <> 0 Or dy <> 0 Then
            
            If Not isModify Then
                ' // 1-st time
                ' // Set the layered style
                isModify = True
                SetWindowLong Me.hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED

                Dim tmpDC   As Long
                Dim oldBmp  As Long
                Dim rect    As D3DLOCKED_RECT
                Dim delta   As Long
                Dim lpDat   As Long
                ' // Create the bitmap with the contents of window
                tmpDC = CreateCompatibleDC(Me.hdc)
                oldBmp = SelectObject(tmpDC, bmpShadow)
                PrintWindow hWnd, tmpDC, 0
                SelectObject tmpDC, oldBmp
                DeleteDC tmpDC
                ' // Move data from the bitmap to the texture
                texture.LockRect 0, rect, ByVal 0&, 0
                           
                lpDat = lpBmpData
                
                ' // Copy each scan-line
                For delta = 0 To (-biWnd.bmiHeader.biHeight) - 1
                
                    memcpy ByVal rect.pBits, ByVal lpDat, biWnd.bmiHeader.biWidth * 4
                    rect.pBits = rect.pBits + rect.pitch
                    lpDat = lpDat + biWnd.bmiHeader.biWidth * 4
                    
                Next
                
                texture.UnlockRect 0
                
                ' // Remove non-client area
                SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
                
            End If
            
            RotateWindow dx, dy
            
        Else
            
            If isModify Then
            
                isModify = False
                ' // Restore the client area
                SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
                SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(Me.hWnd, GWL_EXSTYLE) And (Not WS_EX_LAYERED)
                
                RedrawWindow hWnd, ByVal 0&, 0, RDW_ALLCHILDREN Or RDW_INVALIDATE Or RDW_FRAME
                Me.Cls
                
                tmrReturn.Enabled = False
                
            End If
            
        End If
    
    Case WM_NCCALCSIZE
        ' // Non-client size handler
        If Not isModify Then
            DefCall = True
        End If
      
    Case WM_EXITSIZEMOVE
        ' // When the moving is finished
        If isModify Then
            tmrReturn.Enabled = True
        End If
        
    Case Else
        DefCall = True
    End Select
    
End Sub

' // Make transformation
Private Sub RotateWindow(ByVal dx As Single, ByVal dy As Single)
    Dim Mtx As D3DMATRIX
    Dim off As Single
    
    ' // Calc maximum offset
    If Abs(dx) > Abs(dy) Then off = Abs(dx) Else off = Abs(dy)
    ' // Move aside the window
    D3DXMatrixTranslation Mtx, 0, 0, off * 300
    d3dev.SetTransform D3DTS_WORLD, Mtx
    ' // Rotation the window
    D3DXMatrixRotationY Mtx, dx
    d3dev.MultiplyTransform D3DTS_WORLD, Mtx
    D3DXMatrixRotationX Mtx, dy
    d3dev.MultiplyTransform D3DTS_WORLD, Mtx
    
    ' // Clear the window background
    d3dev.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbBlack, 1, 0
    ' // Draw window to buffer
    d3dev.BeginScene
    
    d3dev.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
    
    d3dev.EndScene
    
    d3dev.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    
    Dim pt      As Size
    Dim sz      As Size
    Dim pos     As Size
    Dim rect    As D3DLOCKED_RECT

    pt.cx = Me.Left / Screen.TwipsPerPixelX
    pt.cy = Me.Top / Screen.TwipsPerPixelY

    sz.cx = biWnd.bmiHeader.biWidth
    sz.cy = -biWnd.bmiHeader.biHeight
    ' // Copy bitmap to the system memory surface
    d3dev.GetRenderTargetData surf, sysSurf

    sysSurf.LockRect rect, ByVal 0&, D3DLOCK_DISCARD
    ' // Copy to form
    SetDIBitsToDevice Me.hdc, 0, 0, sz.cx, sz.cy, 0, 0, 0, sz.cy, ByVal rect.pBits, biWnd, 0

    UpdateLayeredWindow Me.hWnd, Me.hdc, pt, sz, Me.hdc, pos, 0, AB_32Bpp255, ULW_ALPHA
    
    sysSurf.UnlockRect

End Sub

