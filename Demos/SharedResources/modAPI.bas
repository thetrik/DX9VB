Attribute VB_Name = "modAPI"
Option Explicit

Public Const INFINITE As Long = -1&

Public Type CRITICAL_SECTION
    pDebugInfo      As Long
    LockCount       As Long
    RecursionCount  As Long
    OwningThread    As Long
    LockSemaphore   As Long
    SpinCount       As Long
End Type

Public Declare Function WaitForSingleObject Lib "kernel32" ( _
                        ByVal hHandle As Long, _
                        ByVal dwMilliseconds As Long) As Long
Public Declare Function InitializeCriticalSection Lib "kernel32" ( _
                        ByRef lpCriticalSection As CRITICAL_SECTION) As Long
Public Declare Sub EnterCriticalSection Lib "kernel32" ( _
                   ByRef lpCriticalSection As CRITICAL_SECTION)
Public Declare Function TryEnterCriticalSection Lib "kernel32" ( _
                        ByRef lpCriticalSection As CRITICAL_SECTION) As Long
Public Declare Sub LeaveCriticalSection Lib "kernel32" ( _
                   ByRef lpCriticalSection As CRITICAL_SECTION)
Public Declare Sub DeleteCriticalSection Lib "kernel32" ( _
                   ByRef lpCriticalSection As CRITICAL_SECTION)
Public Declare Function PeekMessage Lib "user32" _
                        Alias "PeekMessageW" ( _
                        ByRef lpMsg As Any, _
                        ByVal hwnd As Long, _
                        ByVal wMsgFilterMin As Long, _
                        ByVal wMsgFilterMax As Long, _
                        ByVal wRemoveMsg As Long) As Long
Public Declare Function TranslateMessage Lib "user32" ( _
                        ByRef lpMsg As Any) As Long
Public Declare Function DispatchMessage Lib "user32" _
                        Alias "DispatchMessageW" ( _
                        ByRef lpMsg As Any) As Long
Public Declare Function Sleep Lib "kernel32" ( _
                        ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateEvent Lib "kernel32" _
                        Alias "CreateEventW" ( _
                        ByRef lpEventAttributes As Any, _
                        ByVal bManualReset As Long, _
                        ByVal bInitialState As Long, _
                        ByVal lpName As Long) As Long
Public Declare Function PulseEvent Lib "kernel32" ( _
                        ByVal hEvent As Long) As Long
Public Declare Function ResetEvent Lib "kernel32" ( _
                        ByVal hEvent As Long) As Long


