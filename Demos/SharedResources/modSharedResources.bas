Attribute VB_Name = "modSharedResources"
' //
' // Shared resources module
' //

Option Explicit

' // Type for FPS calculation
Public Type tFPS
    lFPS        As Long
    lFPSCounter As Long
    dPrevTime   As Double
End Type

' // Vertex format
Public Type tVertexFormat
    tPosition   As D3DVECTOR
    tNormal     As D3DVECTOR
End Type

' // Shared resources that requires atomic access
Public Type tSharedResources
    cDevice         As IDirect3DDevice9
    cVertexBuffer   As IDirect3DVertexBuffer9
    lVertexCount    As Long
    tRenderFPS      As tFPS
    tCalcFPS        As tFPS
    lFailCounter    As Long     ' // Failed captures counter
    bEndFlag        As Boolean  ' // If True - end the rendering thread
    hEvent          As Long     ' // Event that allows stop rendering thread until the main thread get access
End Type

Public gtSharedResources    As tSharedResources

Dim mtCriticalSection   As CRITICAL_SECTION ' // Critical section

Public Sub Init()
    InitializeCriticalSection mtCriticalSection
End Sub

Public Sub Uninit()
    DeleteCriticalSection mtCriticalSection
End Sub

Public Function Lock_Resources() As Long
    EnterCriticalSection mtCriticalSection
End Function

Public Function Lock_Resources_Unblock() As Long
    Lock_Resources_Unblock = TryEnterCriticalSection(mtCriticalSection)
End Function

Public Function Unlock_Resources() As Long
    LeaveCriticalSection mtCriticalSection
End Function
