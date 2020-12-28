Attribute VB_Name = "modMultiThreading"
' // modMultiThreading.bas - The module provides support for multi-threading.
' // Version 1.1
' // © Krivous Anatoly Anatolevich (The trick), 2015

Option Explicit

Private Type uuid
    data1       As Long
    data2       As Integer
    data3       As Integer
    data4(7)    As Byte
End Type

Private Type ThreadData
    lpParameter As Long
    lpAddress   As Long
End Type

Private tlsIndex    As Long  ' Index of the item in the TLS. There will be data specific to the thread.
Private lpVBHeader  As Long  ' Pointer to VBHeader structure.
Private hModule     As Long  ' Base address.
Private lpAsm       As Long  ' Pointer to a binary code.

' // Create a new thread
Public Function vbCreateThread(ByVal lpThreadAttributes As Long, _
                               ByVal dwStackSize As Long, _
                               ByVal lpStartAddress As Long, _
                               ByVal lpParameter As Long, _
                               ByVal dwCreationFlags As Long, _
                               lpThreadId As Long) As Long
    Dim InIDE   As Boolean
    
    Debug.Assert MakeTrue(InIDE)
    
    If InIDE Then
        Dim ret As Long
        
        ret = MsgBox("Multithreading does not work in IDE." & vbNewLine & "Run it in the same thread?", vbQuestion Or vbYesNo)
        If ret = vbYes Then
            ' Run function in main thread
            ret = EXEInitialize.DispCallFunc(ByVal 0&, lpStartAddress, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(lpParameter)), CVar(0))
            If ret Then
                Err.Raise ret
            End If
        End If
        
        Exit Function
    End If
    
    ' Alloc new index from thread local storage
    If tlsIndex = 0 Then
        
        tlsIndex = EXEInitialize.TlsAlloc()
        
        If tlsIndex = 0 Then Exit Function
        
    End If
    ' Get module handle
    If hModule = 0 Then
        
        hModule = App.hInstance
        
    End If
    ' Create assembler code
    If lpAsm = 0 Then
        
        lpAsm = CreateAsm()
        If lpAsm = 0 Then Exit Function
        
    End If
    ' Get pointer to VBHeader and modify
    If lpVBHeader = 0 Then
    
        lpVBHeader = GetVBHeader()
        If lpVBHeader = 0 Then Exit Function
        
        ModifyVBHeader lpAsm
        
    End If
    
    Dim lpThreadData    As Long
    Dim tmpData         As ThreadData
    ' Alloc thread-specific memory for threadData structure
    lpThreadData = EXEInitialize.HeapAlloc(GetProcessHeap(), 0, Len(tmpData))
    
    If lpThreadData = 0 Then Exit Function
    ' Set parameters
    tmpData.lpAddress = lpStartAddress
    tmpData.lpParameter = lpParameter
    ' Copy parameters to thread-specific memory
    EXEInitialize.GetMem8 tmpData, ByVal lpThreadData
    ' Create thread
    vbCreateThread = EXEInitialize.CreateThread(ByVal lpThreadAttributes, _
                                                dwStackSize, _
                                                AddressOf ThreadProc, _
                                                ByVal lpThreadData, _
                                                dwCreationFlags, _
                                                lpThreadId)
    
End Function

' // Initialize runtime for new thread and run procedure
Private Function ThreadProc(lpParameter As ThreadData) As Long
    Dim iid         As uuid
    Dim clsid       As uuid
    Dim lpNewHdr    As Long
    Dim hHeap       As Long
    ' Initialize COM
    EXEInitialize.vbCoInitialize ByVal 0&
    ' IID_IUnknown
    iid.data4(0) = &HC0: iid.data4(7) = &H46
    ' Store parameter to thread local storage
    EXEInitialize.TlsSetValue tlsIndex, lpParameter
    ' Create the copy of VBHeader
    hHeap = EXEInitialize.GetProcessHeap()
    lpNewHdr = EXEInitialize.HeapAlloc(hHeap, 0, &H6A)
    EXEInitialize.CopyMemory ByVal lpNewHdr, ByVal lpVBHeader, &H6A
    ' Adjust offsets
    Dim names()     As Long
    Dim diff        As Long
    Dim Index       As Long
    
    ReDim names(3)
    diff = lpNewHdr - lpVBHeader
    EXEInitialize.CopyMemory names(0), ByVal lpVBHeader + &H58, &H10
    
    For Index = 0 To 3
        names(Index) = names(Index) - diff
    Next
    
    EXEInitialize.CopyMemory ByVal lpNewHdr + &H58, names(0), &H10
    ' This line calls the binary code that runs the asm function.
    EXEInitialize.VBDllGetClassObject VarPtr(hModule), 0, lpNewHdr, clsid, iid, 0
    ' Free memeory
    EXEInitialize.HeapFree hHeap, 0, ByVal lpNewHdr
    EXEInitialize.HeapFree hHeap, 0, lpParameter
    
End Function

' // Get VBHeader structure
Private Function GetVBHeader() As Long
    Dim ptr     As Long
   
    ' Get e_lfanew
    EXEInitialize.GetMem4 ByVal hModule + &H3C, ptr
    ' Get AddressOfEntryPoint
    EXEInitialize.GetMem4 ByVal ptr + &H28 + hModule, ptr
    ' Get VBHeader
    EXEInitialize.GetMem4 ByVal ptr + hModule + 1, GetVBHeader
    
End Function

' // Modify VBHeader to replace Sub Main
Private Sub ModifyVBHeader(ByVal newAddress As Long)
    Dim ptr     As Long
    Dim old     As Long
    Dim flag    As Long
    Dim count   As Long
    Dim size    As Long
    
    ptr = lpVBHeader + &H2C
    ' Are allowed to write in the page
    EXEInitialize.VirtualProtect ByVal ptr, 4, PAGE_READWRITE, old
    ' Set a new address of Sub Main
    EXEInitialize.GetMem4 newAddress, ByVal ptr
    EXEInitialize.VirtualProtect ByVal ptr, 4, old, 0
    
    ' Remove startup form
    EXEInitialize.GetMem4 ByVal lpVBHeader + &H4C, ptr
    ' Get forms count
    EXEInitialize.GetMem2 ByVal lpVBHeader + &H44, count
    
    Do While count > 0
        ' Get structure size
        EXEInitialize.GetMem4 ByVal ptr, size
        ' Get flag (unknown5) from current form
        EXEInitialize.GetMem4 ByVal ptr + &H28, flag
        ' When set, bit 5,
        If flag And &H10 Then
            ' Unset bit 5
            flag = flag And &HFFFFFFEF
            ' Are allowed to write in the page
            EXEInitialize.VirtualProtect ByVal ptr, 4, PAGE_READWRITE, old
            ' Write changet flag
            EXEInitialize.GetMem4 flag, ByVal ptr + &H28
            ' Restoring the memory attributes
            EXEInitialize.VirtualProtect ByVal ptr, 4, old, 0
            
        End If
        
        count = count - 1
        ptr = ptr + size
        
    Loop
    
End Sub

' // Create binary code.
Private Function CreateAsm() As Long
    Dim hMod    As Long
    Dim lpProc  As Long
    Dim ptr     As Long
    
    hMod = EXEInitialize.GetModuleHandle(ByVal StrPtr("kernel32"))
    lpProc = EXEInitialize.GetProcAddress(hMod, "TlsGetValue")
    
    If lpProc = 0 Then Exit Function
    
    ptr = EXEInitialize.VirtualAlloc(ByVal 0&, &HF, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    
    If ptr = 0 Then Exit Function
    
    ' push  tlsIndex
    ' call  TLSGetValue
    ' pop   ecx
    ' push  DWORD [eax]
    ' push  ecx
    ' jmp   DWORD [eax + 4]
    
    EXEInitialize.GetMem4 &H68, ByVal ptr + &H0:            EXEInitialize.GetMem4 &HE800, ByVal ptr + &H4
    EXEInitialize.GetMem4 &HFF590000, ByVal ptr + &H8:      EXEInitialize.GetMem4 &H60FF5130, ByVal ptr + &HC
    EXEInitialize.GetMem4 &H4, ByVal ptr + &H10:            EXEInitialize.GetMem4 tlsIndex, ByVal ptr + 1
    EXEInitialize.GetMem4 lpProc - ptr - 10, ByVal ptr + 6
    
    CreateAsm = ptr
    
End Function

Private Function MakeTrue(value As Boolean) As Boolean
    MakeTrue = True: value = True
End Function
