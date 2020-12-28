Attribute VB_Name = "modRender"
' //
' // Render/Calculation procedures
' //

Option Explicit

Public Const ATTEMPTS_COUNT As Long = 10    ' // Failed captures threshold

' // Rendering thread proc
Public Function ThreadProc( _
                ByVal void As Long) As Long
    
    With gtSharedResources
    
    Do

        RenderPass
        
        If .lFailCounter > ATTEMPTS_COUNT Then
            
            ' // Stop the thread until main thread code (PulseEvent) has been performed
            ' // That logic depends on application, in current case we force to perfom the main thread

            ResetEvent .hEvent
            WaitForSingleObject .hEvent, INFINITE
            
            ' // Reset counter
            .lFailCounter = 0
            
        End If
        
        Sleep 0
        
    Loop Until .bEndFlag
    
    End With
    
End Function

' // Render
Public Sub RenderPass()
    Dim bIsInIDE    As Boolean
    Dim bLocked     As Boolean
    
    On Error GoTo error_handler
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Capture resource
    If Not bIsInIDE Then Lock_Resources: bLocked = True
    
    ' // We can access to any fields of shared resources because only current thread has access to shared data
    With gtSharedResources
    
    .cDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, vbRed, 1, 0
    
    .cDevice.BeginScene
    
    If Not .cVertexBuffer Is Nothing And .lVertexCount > 0 Then
    
        .cDevice.SetStreamSource 0, .cVertexBuffer, 0, 24 ' // sizeof(vertex)
        .cDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, .lVertexCount \ 3
        
    End If
    
    .cDevice.EndScene
    
    .cDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    
    ' // Calculate Rendering FPS
    TickFPS .tRenderFPS
    
    End With
    
error_handler:
    
    ' // Release resource ALWAYS otherwise main thread will never get access
    If bLocked Then Unlock_Resources
    
End Sub

' // The "Load" of main thread. It calculates mesh and passes it to shared resource
Public Sub CalcPass()
    Dim tVtx()      As tVertexFormat
    Dim tPoints()   As D3DVECTOR
    Dim bIsInIDE    As Boolean
    Dim fX          As Single
    Dim fY          As Single
    Dim fZ          As Single
    Dim lX          As Long
    Dim lY          As Long
    Dim lSize       As Long
    Dim lBufferSize As Long
    Dim lPtIndex    As Long
    Dim lVtxIndex   As Long
    Dim bLocked     As Boolean
    Dim pData       As Long
    Dim fTheta      As Single
    Static fPhase   As Single
    
    On Error GoTo error_handler
    
    Debug.Assert MakeTrue(bIsInIDE)
    
    ' // Waves mesh
    
    lSize = 40
    
    ReDim tPoints(lSize * lSize - 1)
    ReDim tVtx(((lSize - 1) ^ 2) * 6 - 1)
    
    For lX = 0 To lSize - 1
    
        fX = (lX / lSize - 0.5) * 9
        
        For lY = 0 To lSize - 1
            
            fTheta = Sqr((lX / lSize - 0.5) ^ 2 + (lY / lSize - 0.5) ^ 2)
            fZ = Sin(fPhase + fTheta * 20)
            fY = (lY / lSize - 0.5) * 9
            tPoints(lPtIndex) = vec3(fX, fZ, fY)
            lPtIndex = lPtIndex + 1
            
        Next
        
    Next
    
    fPhase = fPhase + 0.01
    lPtIndex = lSize

    ' // Collect triangles by points
    ' // It's more optimal to make it through index buffer but for simplification it uses simple vertex buffer
    Do Until lPtIndex > UBound(tPoints)

        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex - lSize + 1)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        
        lVtxIndex = lVtxIndex + 1
        
        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        lVtxIndex = lVtxIndex + 1
        
        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex - lSize)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        lVtxIndex = lVtxIndex + 1
        
        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex - lSize + 1)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        lVtxIndex = lVtxIndex + 1
        
        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex + 1)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        lVtxIndex = lVtxIndex + 1
        
        tVtx(lVtxIndex).tPosition = tPoints(lPtIndex)
        tVtx(lVtxIndex).tNormal = vec3(0, 1, 0)
        lVtxIndex = lVtxIndex + 1
        
        lPtIndex = lPtIndex + 1
        
        If (lPtIndex + 1) Mod lSize = 0 Then
            lPtIndex = lPtIndex + 1
        End If
        
    Loop
    
    ' // Capture the shared vertex buffer to update vertices
    With gtSharedResources
    
    If Not bIsInIDE Then

        ' // In current implementation we try to capture the shared resource, if one is busy (the rendering thread already
        ' // captured the one) we increment the counter of the failed captures and end the procedure (for example, optionally
        ' // we can make physics or sounds calcualtion in a game or smth. like that) but don't block the main thread.
        ' // When the counter will have the threshold value we can block the main thread (for waiting) and wait until
        ' // the render thread release the resource.
        ' // We can just call Lock_Resources then the main thread will wait always until the resource has been released
        ' // in the render thread. In current case we can use the calculated data, for example, in physics calculation
        ' // since the data is more detailed in time.
        
        ' // Check counter of failed captures
        If .lFailCounter > ATTEMPTS_COUNT Then
            
            ' // Call the blocked function because the render thread anyway will be stoped and we'll get access as soon
            ' // as render thread call WaitForSingleObject (even maybe earlier) since the same condition is in the rendering
            ' // thread and it force to switch to the main thread.
            Call Lock_Resources
            bLocked = True
            
            ' // Release the rendering thread. The rendering thread ready now to reset the counter of failed captures.
            PulseEvent .hEvent
            
        Else

            ' // Try to capture resource, if failed then end the procedure (optionally make something other if need)
            bLocked = Lock_Resources_Unblock()

            If Not bLocked Then
            
                .lFailCounter = .lFailCounter + 1

                ' // We get out of here and all data is being lost. We can cache it in a real application (for example,
                ' // shadow calculation or physics). It isn't required in current example.
                Exit Sub
                
            End If

        End If

    End If
    
    ' // Atomic access
    lBufferSize = (UBound(tVtx) + 1) * Len(tVtx(0))
    
    If .cVertexBuffer Is Nothing Then

        ' // Create in first time
        .cDevice.CreateVertexBuffer lBufferSize, D3DUSAGE_NONE, D3DFVF_XYZ Or D3DFVF_NORMAL, _
                                    D3DPOOL_DEFAULT, .cVertexBuffer
        
    End If

    ' // Place data to vertex buffer
    .cVertexBuffer.Lock 0, lBufferSize, pData, D3DLOCK_DISCARD
    memcpy ByVal pData, tVtx(0), lBufferSize
    .cVertexBuffer.Unlock
    
    ' // Update vertex counter
    .lVertexCount = UBound(tVtx) + 1
    
    End With

    ' // Make rendering in IDE because all is preformed in the main thread
    If bIsInIDE Then
        RenderPass
    End If
    
error_handler:
    
    ' // Release resource ALWAYS otherwise the rendering thread will never get access

    If bLocked Then
        Unlock_Resources
    End If
    
End Sub

' // "Rough" FPS calclulation
Public Sub TickFPS( _
           ByRef tFPS As tFPS)
           
    With tFPS
    
    .lFPSCounter = .lFPSCounter + 1
    
    If Abs(Timer - .dPrevTime) >= 1 Then
        
        .dPrevTime = Timer
        .lFPS = .lFPSCounter
        .lFPSCounter = 0
        
    End If
    
    End With
    
End Sub

' // DoEvents fast analog
Public Sub FastDoEvents()
    Dim Msg(6) As Long
    
    Do While PeekMessage(Msg(0), 0, 0, 0, 1)
        TranslateMessage Msg(0):  DispatchMessage Msg(0)
    Loop
        
End Sub

' // Fast vector creation
Public Function vec3( _
                ByVal X As Single, _
                ByVal Y As Single, _
                ByVal z As Single) As D3DVECTOR
    vec3.X = X: vec3.Y = Y: vec3.z = z
End Function

Public Function MakeTrue( _
                ByRef bValue As Boolean) As Boolean
    bValue = True
    MakeTrue = True
End Function
                 

