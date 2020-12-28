VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Raymatching using Direct3D9 shaders"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFPS 
      Interval        =   1000
      Left            =   4320
      Top             =   3540
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const COMPILE_SHADERS = False   ' // Set to true to compile HLSL to bin. Otherwise use precompiled code

' // Input of vertex shader
Private Type tVertex
    fX  As Single
    fY  As Single
    fU  As Single
    fV  As Single
End Type

Private Declare Function D3DXGetVertexShaderProfile Lib "d3dx9_43" ( _
                         ByVal pDevice As IDirect3DDevice9) As Long
Private Declare Function D3DXGetPixelShaderProfile Lib "d3dx9_43" ( _
                         ByVal pDevice As IDirect3DDevice9) As Long
Private Declare Function D3DXCreateBuffer Lib "d3dx9_43" ( _
                         ByVal NumBytes As Long, _
                         ByRef ppBuffer As ID3DXBuffer) As Long
Private Declare Function D3DXCompileShaderFromFile Lib "d3dx9_43" _
                         Alias "D3DXCompileShaderFromFileW" ( _
                         ByVal pSrcFile As Long, _
                         ByRef pDefines As Any, _
                         ByVal pInclude As ID3DXInclude, _
                         ByVal pFunctionName As String, _
                         ByVal pProfile As Long, _
                         ByVal Flags As Long, _
                         ByRef ppShader As ID3DXBuffer, _
                         ByRef ppErrorMsgs As ID3DXBuffer, _
                         ByRef ppConstantTable As ID3DXConstantTable) As Long
Private Declare Function D3DXGetShaderConstantTable Lib "d3dx9_43" ( _
                         ByRef pFunction As Any, _
                         ByRef ppConstantTable As ID3DXConstantTable) As Long

Private Declare Sub memcpy Lib "kernel32" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)

Private m_cD3D      As IDirect3D9
Private m_cDevice   As IDirect3DDevice9
Private m_cQuad     As IDirect3DVertexBuffer9
Private m_lTime1Reg As Long
Private m_lFPS      As Long
Private m_bActive   As Boolean

Private Sub Form_Load()
    Dim tPP         As D3DPRESENT_PARAMETERS
    Dim cErrMsg     As ID3DXBuffer
    Dim cPSCode     As ID3DXBuffer
    Dim cVSCode     As ID3DXBuffer
    Dim cPSConstTbl As ID3DXConstantTable
    Dim cVShader    As IDirect3DVertexShader9
    Dim cPShader    As IDirect3DPixelShader9
    Dim cVtxDecl    As IDirect3DVertexDeclaration9
    Dim tVertex(5)  As tVertex
    Dim tVtxDecl(2) As D3DVERTEXELEMENT9
    Dim pData       As Long
    Dim hConst      As Long
    Dim fAspect     As Single
    Dim hr          As Long
    
    Set m_cD3D = Direct3DCreate9()
    
    tPP.BackBufferCount = 1
    tPP.Windowed = 1
    tPP.BackBufferFormat = D3DFMT_A8R8G8B8
    tPP.SwapEffect = D3DSWAPEFFECT_DISCARD
    tPP.PresentationInterval = D3DPRESENT_INTERVAL_IMMEDIATE
    
    Set m_cDevice = m_cD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, tPP)
    
#If COMPILE_SHADERS Then

    ' // Compile vertex shader code
    hr = D3DXCompileShaderFromFile(StrPtr(App.Path & "\vs.txt"), ByVal 0&, Nothing, "vs_main", _
                                   StrPtr(StrConv("vs_1_1", vbFromUnicode)), 0, cVSCode, cErrMsg, Nothing)
    If hr < 0 Then
        ShowCompError cErrMsg
        Exit Sub
    Else
        Set cErrMsg = Nothing
    End If

    SaveShaderToFile cVSCode, App.Path & "\vs.bin"
    
    ' // Compile pixel shader code
    hr = D3DXCompileShaderFromFile(StrPtr(App.Path & "\ps.txt"), ByVal 0&, Nothing, "ps_main", _
                                   StrPtr(StrConv("ps_3_0", vbFromUnicode)), 0, cPSCode, cErrMsg, cPSConstTbl)
    If hr < 0 Then
        ShowCompError cErrMsg
        Exit Sub
    End If
    
    SaveShaderToFile cPSCode, App.Path & "\ps.bin"
    
#Else
    
    Set cVSCode = LoadShaderFromFile(App.Path & "\vs.bin")
    Set cPSCode = LoadShaderFromFile(App.Path & "\ps.bin")
    
    hr = D3DXGetShaderConstantTable(ByVal cPSCode.GetBufferPointer, cPSConstTbl)
    
    If hr < 0 Then
        Err.Raise hr
    End If
    
#End If

    m_lTime1Reg = GetShaderConstantRegister(cPSConstTbl, "TIME1")
    
    ' // Create shaders
    Set cVShader = m_cDevice.CreateVertexShader(ByVal cVSCode.GetBufferPointer)
    Set cPShader = m_cDevice.CreatePixelShader(ByVal cPSCode.GetBufferPointer)
    
    ' // Create vertex declaration
    tVtxDecl(0) = vtx_element(0, 0, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_POSITION, 0)
    tVtxDecl(1) = vtx_element(0, 8, D3DDECLTYPE_FLOAT2, D3DDECLMETHOD_DEFAULT, D3DDECLUSAGE_TEXCOORD, 0)
    tVtxDecl(2) = D3DDECL_END
    
    Set cVtxDecl = m_cDevice.CreateVertexDeclaration(tVtxDecl(0))
    
    m_cDevice.SetVertexDeclaration cVtxDecl
    
    ' // Create full-screen quad based on screen aspect ration
    fAspect = Me.ScaleHeight / Me.ScaleWidth
    
    tVertex(0) = vtx(-1, 1, -1, fAspect)
    tVertex(1) = vtx(1, -1, 1, -fAspect)
    tVertex(2) = vtx(1, 1, 1, fAspect)
    
    tVertex(3) = vtx(-1, 1, -1, fAspect)
    tVertex(4) = vtx(1, -1, 1, -fAspect)
    tVertex(5) = vtx(-1, -1, -1, -fAspect)
    
    ' // Create vertex buffer with quad data
    m_cDevice.CreateVertexBuffer Len(tVertex(0)) * (UBound(tVertex) + 1), 0, 0, 0, m_cQuad
    
    m_cQuad.Lock 0, Len(tVertex(0)) * (UBound(tVertex) + 1), pData, 0
    memcpy ByVal pData, tVertex(0), Len(tVertex(0)) * (UBound(tVertex) + 1)
    m_cQuad.Unlock
    
    m_cDevice.SetStreamSource 0, m_cQuad, 0, Len(tVertex(0))
    
    ' // Disable culling
    m_cDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    
    ' // Set shaders to device
    m_cDevice.SetPixelShader cPShader
    m_cDevice.SetVertexShader cVShader

    Me.Show
    
    MainLoop
    
End Sub

Private Sub MainLoop()
        
    m_bActive = True
    
    Do While m_bActive
        
        m_cDevice.SetPixelShaderConstantF m_lTime1Reg, CSng(2 * Timer * 0.6), 1
        
        m_cDevice.BeginScene
        
        ' // Draw full-scree quad
        m_cDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2

        m_cDevice.EndScene
    
        m_cDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
        
        m_lFPS = m_lFPS + 1
        
        DoEvents
        
    Loop
    
End Sub

' // Load binary shader from file
Private Function LoadShaderFromFile( _
                 ByRef sFileName As String) As ID3DXBuffer
    Dim iFile   As Integer
    Dim bData() As Byte
    Dim lSize   As Long
    Dim cRet    As ID3DXBuffer
    Dim hr      As Long
    
    iFile = FreeFile
    
    Open sFileName For Binary As iFile
    
    lSize = LOF(iFile)
    
    If lSize <= 0 Then
        Err.Raise 5
    End If
    
    ReDim bData(lSize - 1)
    
    Get iFile, , bData
    
    Close iFile
    
    hr = D3DXCreateBuffer(lSize, cRet)
    
    If hr < 0 Then
        Err.Raise hr
    End If
    
    memcpy ByVal cRet.GetBufferPointer, bData(0), UBound(bData) + 1
    
    Set LoadShaderFromFile = cRet
    
End Function

' // Save binary shader to file
Private Sub SaveShaderToFile( _
            ByVal cShader As ID3DXBuffer, _
            ByRef sFileName As String)
    Dim iFile   As Integer
    Dim bData() As Byte
    
    If Len(Dir(sFileName)) Then
        Kill sFileName
    End If
    
    iFile = FreeFile
    
    Open sFileName For Binary As iFile
    
    If cShader.GetBufferSize > 0 Then
    
        ReDim bData(cShader.GetBufferSize - 1)
        memcpy bData(0), ByVal cShader.GetBufferPointer, UBound(bData) + 1
        
        Put iFile, , bData
        
    End If
    
    Close iFile
    
End Sub

' // Get register index of shader constant
Private Function GetShaderConstantRegister( _
                 ByVal cTable As ID3DXConstantTable, _
                 ByVal sName As String) As Long
    Dim hConst  As Long
    Dim tDesc   As D3DXCONSTANT_DESC
    
    hConst = cTable.GetConstantByName(0, sName)
    If hConst = 0 Then
        Err.Raise 5
    End If
    
    cTable.GetConstantDesc hConst, tDesc, 1
    
    GetShaderConstantRegister = tDesc.RegisterIndex
                     
End Function

Private Function D3DDECL_END() As D3DVERTEXELEMENT9
    D3DDECL_END = vtx_element(255, 0, D3DDECLTYPE_UNUSED, 0, 0, 0)
End Function

' // Create D3DVERTEXELEMENT9 ittem
Private Function vtx_element( _
                 ByVal lStream As Long, _
                 ByVal lOffset As Long, _
                 ByVal eType As D3DDECLTYPE, _
                 ByVal eMethod As D3DDECLMETHOD, _
                 ByVal eUsage As D3DDECLUSAGE, _
                 ByVal lUsageIndex As Long) As D3DVERTEXELEMENT9
                 
    With vtx_element
        .Stream = lStream
        .Offset = lOffset
        .Type = eType
        .Method = eMethod
        .Usage = eUsage
        .UsageIndex = lUsageIndex
    End With
    
End Function

' // Create vertex
Private Function vtx( _
                 ByVal fX As Single, _
                 ByVal fY As Single, _
                 ByVal fU As Single, _
                 ByVal fV As Single) As tVertex
    vtx.fX = fX
    vtx.fY = fY
    vtx.fU = fU
    vtx.fV = fV
End Function

' // Show error message storred to ID3DXBuffer buffer
Private Sub ShowCompError( _
            ByVal cErrMsg As ID3DXBuffer)
    Dim bAnsiInfo() As Byte
    Dim sMsgUnicode As String
    
    If cErrMsg.GetBufferSize > 0 Then
        
        ReDim bAnsiInfo(cErrMsg.GetBufferSize - 1)
        
        memcpy bAnsiInfo(0), ByVal cErrMsg.GetBufferPointer, UBound(bAnsiInfo) + 1
        
        sMsgUnicode = StrConv(bAnsiInfo, vbUnicode)
        
        MsgBox sMsgUnicode, vbCritical
     
    Else
    
        MsgBox "Unknown error", vbCritical
        
    End If
                
End Sub

Private Sub Form_QueryUnload( _
            ByRef Cancel As Integer, _
            ByRef UnloadMode As Integer)
    m_bActive = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_cQuad = Nothing
    Set m_cDevice = Nothing
    Set m_cD3D = Nothing
    
End Sub

Private Sub tmrFPS_Timer()
    
    Me.Caption = "FPS: " & m_lFPS
    m_lFPS = 0
    
End Sub
