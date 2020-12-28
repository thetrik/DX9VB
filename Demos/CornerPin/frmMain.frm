VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CornerPin"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   FillColor       =   &H00000040&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00404000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Corner pin by The trick
' // 2020
' //

Option Explicit

Private Const PixelFormat32bppARGB        As Long = &H26200A
Private Const PixelFormat32bppPARGB       As Long = &HE200B
Private Const ImageLockModeRead           As Long = &H1
Private Const ImageLockModeWrite          As Long = &H2
Private Const ImageLockModeUserInputBuf   As Long = &H4
Private Const UnitPixel                   As Long = 2

Private Type BitmapData
    Width                       As Long
    Height                      As Long
    stride                      As Long
    PixelFormat                 As Long
    scan0                       As Long
    reserved                    As Long
End Type
Private Type GdiplusStartupInput
    GdiplusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Declare Function GdipLoadImageFromFile Lib "gdiplus" ( _
                         ByVal pfilename As Long, _
                         ByRef image As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" ( _
                         ByRef token As Long, _
                         ByRef inputbuf As GdiplusStartupInput, _
                         Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
                    ByVal token As Long)
Private Declare Function GdipGetImageWidth Lib "gdiplus" ( _
                         ByVal image As Long, _
                         ByRef Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" ( _
                         ByVal image As Long, _
                         ByRef Height As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" ( _
                         ByVal BITMAP As Long, _
                         ByRef rc As RECT, _
                         ByVal flags As Long, _
                         ByVal PixelFormat As Long, _
                         ByRef lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" ( _
                         ByVal BITMAP As Long, _
                         ByRef lockedBitmapData As BitmapData) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
                         ByVal image As Long) As Long
Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private Type tVertex
    tPos    As D3DVECTOR
    fRHW    As Single
    fU      As Single
    fV      As Single
End Type

Private Const VERTEX_SIZE   As Long = 6 * 4

Private m_tCorners(3)   As tVertex
Private m_cD3D9         As IDirect3D9
Private m_cDevice       As IDirect3DDevice9
Private m_cVtxBuf       As IDirect3DVertexBuffer9
Private m_cIndexBuf     As IDirect3DIndexBuffer9
Private m_cTexture      As IDirect3DTexture9
Private m_hGpToken      As Long
Private m_lSelCorner    As Long
Private m_bIsDrag       As Boolean
Private m_eFmtFlags     As D3DFVF

Private Sub Form_Load()
    Dim tPP         As D3DPRESENT_PARAMETERS
    Dim pVtxData    As Long
    Dim pIdxData    As Long
    Dim tGpInput    As GdiplusStartupInput
    Dim iIndices(5) As Integer
    
    m_lSelCorner = -1
    
    tGpInput.GdiplusVersion = 1
    
    If GdiplusStartup(m_hGpToken, tGpInput) <> 0 Then
        MsgBox "Unable to initialize GDI+", vbCritical
        Exit Sub
    End If
    
    Set m_cD3D9 = Direct3DCreate9()
    
    ' // Set vertex format
    m_eFmtFlags = D3DFVF_TEX1 Or D3DFVF_XYZRHW
    
    tPP.BackBufferCount = 1
    tPP.Windowed = 1
    tPP.BackBufferFormat = D3DFMT_A8R8G8B8
    tPP.SwapEffect = D3DSWAPEFFECT_DISCARD
    
    ' // Create device
    Set m_cDevice = m_cD3D9.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, tPP)
    
    ' // Create vertex buffer which will contain corners
    m_cDevice.CreateVertexBuffer VERTEX_SIZE * 4, 0, m_eFmtFlags, D3DPOOL_DEFAULT, m_cVtxBuf
    
    ' // Create corners
    m_tCorners(0) = vtx(10, 10, 0, 1, 0, 0)
    m_tCorners(1) = vtx(Me.ScaleWidth - 10, 10, 0, 1, 1, 0)
    m_tCorners(2) = vtx(10, Me.ScaleHeight - 10, 0, 1, 0, 1)
    m_tCorners(3) = vtx(Me.ScaleWidth - 10, Me.ScaleHeight - 10, 0, 1, 1, 1)
    
    ' // Put corners to buffer
    m_cVtxBuf.Lock 0, Len(m_tCorners(0)) * (UBound(m_tCorners) + 1), pVtxData, 0
    memcpy ByVal pVtxData, m_tCorners(0), Len(m_tCorners(0)) * (UBound(m_tCorners) + 1)
    m_cVtxBuf.Unlock
    
    m_cDevice.SetStreamSource 0, m_cVtxBuf, 0, VERTEX_SIZE
    
    m_cDevice.SetFVF m_eFmtFlags
    
    ' // Create index buffer which will specify triangles corners
    m_cDevice.CreateIndexBuffer 6 * 2, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT, m_cIndexBuf
    
    ' // Specify corners to create 2 triangle
    iIndices(0) = 0:    iIndices(1) = 1:    iIndices(2) = 3
    iIndices(3) = 0:    iIndices(4) = 3:    iIndices(5) = 2
    
    ' // Put indices
    m_cIndexBuf.Lock 0, Len(iIndices(0)) * (UBound(iIndices) + 1), pIdxData, 0
    memcpy ByVal pIdxData, iIndices(0), Len(iIndices(0)) * (UBound(iIndices) + 1)
    m_cIndexBuf.Unlock
    
    m_cDevice.SetIndices m_cIndexBuf
    
    m_cDevice.SetRenderState D3DRS_LIGHTING, 0
    
    m_cDevice.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
    m_cDevice.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
    
    Set m_cTexture = LoadTextureFromFile(App.Path & "\test.jpg")
    
    If m_cTexture Is Nothing Then
        Exit Sub
    End If

    m_cDevice.SetTexture 0, m_cTexture
    
End Sub

' // Load texture from file
Private Function LoadTextureFromFile( _
                 ByRef sFileName As String) As IDirect3DTexture9
    Dim hImage      As Long
    Dim tRC         As RECT
    Dim tLockRC     As D3DLOCKED_RECT
    Dim tBmpData    As BitmapData
    Dim cRet        As IDirect3DTexture9
    
    On Error GoTo CleanUp
    
    If GdipLoadImageFromFile(StrPtr(sFileName), hImage) <> 0 Then
        MsgBox "Unable to load picture", vbCritical
        Exit Function
    End If
    
    If GdipGetImageWidth(hImage, tBmpData.Width) <> 0 Then
        MsgBox "Unable to get picture width", vbCritical
        GoTo CleanUp
    End If
    
    If GdipGetImageHeight(hImage, tBmpData.Height) <> 0 Then
        MsgBox "Unable to get picture height", vbCritical
        GoTo CleanUp
    End If
    
    m_cDevice.CreateTexture tBmpData.Width, tBmpData.Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, cRet
    
    cRet.LockRect 0, tLockRC, ByVal 0&, 0
    
    tBmpData.scan0 = tLockRC.pBits
    tBmpData.stride = tLockRC.Pitch
    tBmpData.PixelFormat = PixelFormat32bppARGB
    
    tRC.Right = tBmpData.Width
    tRC.bottom = tBmpData.Height
    
    If GdipBitmapLockBits(hImage, tRC, ImageLockModeRead Or ImageLockModeWrite Or ImageLockModeUserInputBuf, _
                            PixelFormat32bppARGB, tBmpData) <> 0 Then
        cRet.UnlockRect 0
        MsgBox "Unable to get image bits", vbCritical
        GoTo CleanUp
    End If
    
    GdipBitmapUnlockBits hImage, tBmpData
    
    cRet.UnlockRect 0
    
    Set LoadTextureFromFile = cRet
    
CleanUp:
    
    If Err.Number Then
        MsgBox "An error occured 0x" & Hex$(Err.Number)
    End If
    
    GdipDisposeImage hImage
    
End Function

Private Function vtx( _
                 ByVal fX As Single, _
                 ByVal fY As Single, _
                 ByVal fZ As Single, _
                 ByVal fRHW As Single, _
                 ByVal fU As Single, _
                 ByVal fV As Single) As tVertex
                 
    vtx.tPos.X = fX
    vtx.tPos.Y = fY
    vtx.tPos.Z = fZ
    vtx.fRHW = fRHW
    vtx.fU = fU
    vtx.fV = fV
    
End Function

Private Function mtxInverse3x3( _
                 ByRef m() As Single, _
                 ByRef m_ret() As Single) As Boolean
    Dim fDet    As Single
    Dim fRet()  As Single
    
    fDet = m(0, 0) * (m(1, 1) * m(2, 2) - m(2, 1) * m(1, 2)) - _
           m(0, 1) * (m(1, 0) * m(2, 2) - m(1, 2) * m(2, 0)) + _
           m(0, 2) * (m(1, 0) * m(2, 1) - m(1, 1) * m(2, 0))
             
    If Abs(fDet) < 0.00001 Then Exit Function
             
    ReDim fRet(2, 2)
    
    fRet(0, 0) = (m(1, 1) * m(2, 2) - m(2, 1) * m(1, 2)) / fDet
    fRet(0, 1) = (m(0, 2) * m(2, 1) - m(0, 1) * m(2, 2)) / fDet
    fRet(0, 2) = (m(0, 1) * m(1, 2) - m(0, 2) * m(1, 1)) / fDet
    fRet(1, 0) = (m(1, 2) * m(2, 0) - m(1, 0) * m(2, 2)) / fDet
    fRet(1, 1) = (m(0, 0) * m(2, 2) - m(0, 2) * m(2, 0)) / fDet
    fRet(1, 2) = (m(1, 0) * m(0, 2) - m(0, 0) * m(1, 2)) / fDet
    fRet(2, 0) = (m(1, 0) * m(2, 1) - m(2, 0) * m(1, 1)) / fDet
    fRet(2, 1) = (m(2, 0) * m(0, 1) - m(0, 0) * m(2, 1)) / fDet
    fRet(2, 2) = (m(0, 0) * m(1, 1) - m(1, 0) * m(0, 1)) / fDet
    
    m_ret = fRet
    
    mtxInverse3x3 = True
    
End Function

Private Function CalcPerspective() As Boolean
    Dim s() As Single
    Dim d() As Single
    Dim m() As Single
    Dim v() As Single
    Dim lV  As Long
    Dim fW  As Single
    
    If Not IsConvex() Then Exit Function
    
    ReDim s(2, 2)
    ReDim m(2, 2)
    ReDim v(2)
    
    s(0, 0) = -1: s(0, 1) = -1: s(0, 2) = 1
    s(1, 0) = -1: s(1, 1) = 0: s(1, 2) = 0
    s(2, 0) = 0: s(2, 1) = -1: s(2, 2) = 0
    
    m(0, 0) = m_tCorners(0).tPos.X: m(0, 1) = m_tCorners(1).tPos.X: m(0, 2) = m_tCorners(2).tPos.X
    m(1, 0) = m_tCorners(0).tPos.Y: m(1, 1) = m_tCorners(1).tPos.Y: m(1, 2) = m_tCorners(2).tPos.Y
    m(2, 0) = 1: m(2, 1) = 1: m(2, 2) = 1
    
    If Not mtxInverse3x3(m(), m()) Then Exit Function
    
    v(0) = m(0, 0) * m_tCorners(3).tPos.X + m(0, 1) * m_tCorners(3).tPos.Y + m(0, 2)
    v(1) = m(1, 0) * m_tCorners(3).tPos.X + m(1, 1) * m_tCorners(3).tPos.Y + m(1, 2)
    v(2) = m(2, 0) * m_tCorners(3).tPos.X + m(2, 1) * m_tCorners(3).tPos.Y + m(2, 2)
    
    m(0, 0) = v(0) * m_tCorners(0).tPos.X: m(0, 1) = v(1) * m_tCorners(1).tPos.X: m(0, 2) = v(2) * m_tCorners(2).tPos.X
    m(1, 0) = v(0) * m_tCorners(0).tPos.Y: m(1, 1) = v(1) * m_tCorners(1).tPos.Y: m(1, 2) = v(2) * m_tCorners(2).tPos.Y
    m(2, 0) = v(0): m(2, 1) = v(1): m(2, 2) = v(2)

    m = mtxMul3x3(m, s)

    For lV = 0 To 3
    
        v(0) = lV And 1: v(1) = (lV And 2) \ 2: v(2) = 1
        fW = m(2, 0) * v(0) + m(2, 1) * v(1) + m(2, 2) * v(2)
        
        If fW = 0 Then Exit Function
        
        m_tCorners(lV).fRHW = 1 / fW
        
    Next
    
    CalcPerspective = True
    
End Function

Private Function IsConvex() As Boolean
    Dim lIndex      As Long
    Dim lCW()       As Long
    Dim fPoints()   As Single
    Dim bIsNeg      As Boolean
    Dim bOriginDir  As Boolean
    
    ReDim lCW(3)

    lCW(0) = 2: lCW(1) = 0: lCW(2) = 1: lCW(3) = 3
    
    For lIndex = 0 To 3

        bIsNeg = PerpDot(m_tCorners(lCW((lIndex + 1) And &H3)).tPos.X - m_tCorners(lCW(lIndex)).tPos.X, _
                         m_tCorners(lCW((lIndex + 1) And &H3)).tPos.Y - m_tCorners(lCW(lIndex)).tPos.Y, _
                         m_tCorners(lCW((lIndex + 2) And &H3)).tPos.X - m_tCorners(lCW((lIndex + 1) And &H3)).tPos.X, _
                         m_tCorners(lCW((lIndex + 2) And &H3)).tPos.Y - m_tCorners(lCW((lIndex + 1) And &H3)).tPos.Y) < 0
        
        If lIndex Then
            If bOriginDir <> bIsNeg Then
                IsConvex = False
                Exit Function
            End If
        Else
            bOriginDir = bIsNeg
        End If
        
    Next

    IsConvex = True
    
End Function

Private Function PerpDot( _
                 ByVal x1 As Single, _
                 ByVal y1 As Single, _
                 ByVal x2 As Single, _
                 ByVal y2 As Single) As Single
    PerpDot = x1 * y2 - y1 * x2
End Function

Private Function mtxMul3x3( _
                 ByRef m1() As Single, _
                 ByRef m2() As Single) As Single()
    Dim lI      As Long
    Dim lJ      As Long
    Dim lU      As Long
    Dim fRet()  As Single

    ReDim fRet(2, 2)

    For lI = 0 To 2
        For lJ = 0 To 2
            For lU = 0 To 2
                fRet(lI, lJ) = fRet(lI, lJ) + m1(lI, lU) * m2(lU, lJ)
            Next
        Next
    Next

    mtxMul3x3 = fRet

End Function

Private Sub Form_MouseDown( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef X As Single, _
            ByRef Y As Single)
    
    If m_lSelCorner = -1 Then Exit Sub
    
    m_bIsDrag = True
    
End Sub

Private Sub Form_MouseMove( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef X As Single, _
            ByRef Y As Single)
    Dim lIndex      As Long
    Dim pVtxData    As Long
    Dim fOrigin(1)  As Single
    
    If m_bIsDrag Then
        
        With m_tCorners(m_lSelCorner).tPos
            
            If X < 10 Then
                X = 10
            ElseIf X > Me.ScaleWidth - 10 Then
                X = Me.ScaleWidth - 10
            End If
            
            If Y < 10 Then
                Y = 10
            ElseIf Y > Me.ScaleHeight - 10 Then
                Y = Me.ScaleHeight - 10
            End If
            
            fOrigin(0) = .X:    fOrigin(1) = .Y
            
            .X = X: .Y = Y
            
            If Not CalcPerspective() Then
                .X = fOrigin(0):    .Y = fOrigin(1)
            End If

            m_cVtxBuf.Lock 0, Len(m_tCorners(0)) * (UBound(m_tCorners) + 1), pVtxData, 0
            memcpy ByVal pVtxData, m_tCorners(0), Len(m_tCorners(0)) * (UBound(m_tCorners) + 1)
            m_cVtxBuf.Unlock
            
            Form_Paint
            
            Exit Sub
            
        End With
        
    Else
    
        For lIndex = 0 To UBound(m_tCorners)
        
            With m_tCorners(lIndex).tPos
                
                If (.X - X) ^ 2 + (.Y - Y) ^ 2 <= 25 Then
                    
                    If m_lSelCorner <> lIndex Then
                        m_lSelCorner = lIndex
                        Form_Paint
                    End If
                    
                    Exit Sub
                    
                End If
                
            End With
        
        Next
        
        If m_lSelCorner <> -1 Then
            m_lSelCorner = -1
            Form_Paint
        End If
    
    End If
    
End Sub

Private Sub Form_MouseUp( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef X As Single, _
            ByRef Y As Single)
    
    m_bIsDrag = False
    
End Sub

Private Sub Form_Paint()
    Dim lIndex  As Long
    
    m_cDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, &H504030, 1, 0
    
    m_cDevice.BeginScene
    
    ' // Draw image
    m_cDevice.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, 0, 6, 0, 2
    
    For lIndex = 0 To UBound(m_tCorners)
           
        With m_tCorners(lIndex).tPos
            
            DrawCircle .X, .Y, 5, lIndex = m_lSelCorner
        
        End With
           
    Next
    
    m_cDevice.EndScene
        
    m_cDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    
End Sub

Private Sub DrawCircle( _
            ByVal lX As Long, _
            ByVal lY As Long, _
            ByVal lRadius As Long, _
            ByVal bFill As Boolean)
    Dim tVtx()  As tVertex
    Dim lIndex  As Long

    ReDim tVtx(31)
    
    m_cDevice.SetTexture 0, Nothing
    
    If bFill Then
        
        tVtx(0).tPos.X = lX:    tVtx(0).tPos.Y = lY
        
        For lIndex = 0 To UBound(tVtx) - 1
    
            tVtx(lIndex + 1).tPos.X = Cos(6.2831 * lIndex / (UBound(tVtx) - 1)) * lRadius + lX
            tVtx(lIndex + 1).tPos.Y = Sin(6.2831 * lIndex / (UBound(tVtx) - 1)) * lRadius + lY

        Next
        
        m_cDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 31, tVtx(0), Len(tVtx(0))
        
    Else
    
        For lIndex = 0 To UBound(tVtx)
    
            tVtx(lIndex).tPos.X = Cos(6.2831 * lIndex / (UBound(tVtx))) * lRadius + lX
            tVtx(lIndex).tPos.Y = Sin(6.2831 * lIndex / (UBound(tVtx))) * lRadius + lY

        Next
        
        m_cDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 31, tVtx(0), Len(tVtx(0))
        
    End If
    
    m_cDevice.SetStreamSource 0, m_cVtxBuf, 0, VERTEX_SIZE
    m_cDevice.SetTexture 0, m_cTexture
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)

    ' // Free resources
    Set m_cTexture = Nothing
    Set m_cVtxBuf = Nothing
    Set m_cDevice = Nothing
    Set m_cD3D9 = Nothing
    Set m_cIndexBuf = Nothing
    
    If m_hGpToken Then
        GdiplusShutdown m_hGpToken
    End If
    
End Sub


