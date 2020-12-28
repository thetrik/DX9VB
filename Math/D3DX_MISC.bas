Attribute VB_Name = "D3DX_MISC"
Option Explicit

Public Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub GetMem4 Lib "MSVBVM60" (Src As Any, Dst As Any)

Public Const PI As Single = 3.14159265358979

' // Determines whether a value is an illegal number.
Public Function sngIsNaN(ByVal Value As Single) As Boolean
    Dim dat As Long
    
    GetMem4 Value, dat
    sngIsNaN = (dat And &H7F800000) = &H7F800000 And (dat And &H7FFFFF) > 0
       
End Function

' // Determines whether a value is an infinite.
Public Function sngIsInf(ByVal Value As Single) As Boolean
    Dim dat As Long
    
    GetMem4 Value, dat
    sngIsInf = (dat And &H7F800000) = &H7F800000 And (dat And &H7FFFFF) = 0
       
End Function

' // Converts an array of 16-bit floats to 32-bit floats.
Public Sub D3DXFloat16To32Array(pOut() As Single, _
                                pIn() As Integer, _
                                ByVal n As Long)
    Dim s   As Long
    Dim e   As Long
    Dim m   As Long
    Dim sn  As Long
    Dim v   As Integer
    Dim i   As Long
    
    For i = 0 To n - 1
    
        v = pIn(i): s = v And &H8000&:  e = (v And &H7C00&) \ 1024: m = v And &H3FF
        sn = IIf(s, -1, 1)
    
        If e = 0 Then
            If m = 0 Then
                pOut(i) = 0!
            Else
                pOut(i) = sn * 6.103516E-05! * (m / 1024!)
            End If
        Else
            pOut(i) = sn * (2 ^ (e - 15)) * (1! + m / 1024!)
        End If
    Next
    
End Sub

' // Converts an array of 32-bit floats to 16-bit floats.
Public Sub D3DXFloat32To16Array(pOut() As Integer, _
                                 pIn() As Single, _
                                 ByVal n As Long)
    Dim exp_        As Long
    Dim origexp     As Long
    Dim tmp         As Single
    Dim mantissa    As Long
    Dim sign        As Long
    Dim ret         As Integer
    Dim i           As Long
    Dim v           As Single
    
    For i = 0 To n - 1
        
        v = pIn(i)
        tmp = Abs(v):    sign = IIf(v >= 0, 0, 1)
        If sngIsInf(v) Or sngIsNaN(v) Then
            pOut(i) = IIf(sign, &HFFFF, &H7FFF)
        ElseIf v = 0! Then
            pOut(i) = IIf(sign, &H8000, &H0)
        Else
            If tmp < 1024! Then
                Do
                    tmp = tmp * 2!
                    exp_ = exp_ - 1
                Loop While tmp < 1024!
            ElseIf tmp >= 2048 Then
                Do
                    tmp = tmp / 2
                    exp_ = exp_ + 1
                Loop While tmp >= 2048!
            End If
            
            exp_ = exp_ + 10 + 15
            origexp = exp_
            
            If (tmp = 2018.5! Or tmp = 2016.5! Or tmp = 2014.5! Or tmp = 2012.5! Or _
                tmp = 1954.5! Or tmp = 1952.5! Or tmp = 1950.5! Or tmp = 1948.5!) Then
                mantissa = tmp
            ElseIf (tmp = 2019.5! Or tmp = 2017.5! Or tmp = 2015.5! Or tmp = 2013.5! Or _
                    tmp = 1955.5! Or tmp = 1953.5! Or tmp = 1951.5! Or tmp = 1949.5!) Then
                mantissa = tmp + 1
            Else
                mantissa = tmp
                If tmp - mantissa >= 0.5! Then mantissa = mantissa + 1
            End If
            If mantissa = 2048 Then
                mantissa = 1024
                exp_ = exp_ + 1
            End If
        
            If exp_ > 31 Then
                ret = &H7FFF
            ElseIf exp_ <= 0 Then
                Dim rounding    As Long
                
                exp_ = origexp
                mantissa = tmp
                mantissa = mantissa And &H3FF Or &H400
                
                Do While exp_ <= 0
                    rounding = mantissa And 1
                    mantissa = mantissa \ 2
                    exp_ = exp_ + 1
                Loop
                
                ret = mantissa + rounding
                
            Else
                ret = (exp_ * 1024) Or (mantissa And &H3FF)
            End If
            
            If sign Then
                ret = ret Or &H8000
            End If
            
            pOut(i) = ret
        End If
        
    Next

End Sub

' // Calculate the Fresnel term.
Public Function D3DXFresnelTerm(ByVal CosTheta As Single, _
                                ByVal RefractionIndex As Single) As Single
    Dim a   As Single
    Dim d   As Single
    Dim g   As Single
    Dim ret As Single
    
    If CosTheta < 1.175494E-38! Then D3DXFresnelTerm = 1!
    
    g = Sqr(Abs(RefractionIndex * RefractionIndex + CosTheta * CosTheta - 1!))
    a = g + CosTheta
    d = g - CosTheta
    ret = (CosTheta * a - 1!) * (CosTheta * a - 1!) / ((CosTheta * d + 1!) * (CosTheta * d + 1!)) + 1!
    D3DXFresnelTerm = ret * 0.5! * d * d / (a * a)

End Function

' // Returns the angle whose tangent is the ratio of the two numbers
Public Function Atan2(ByVal Y As Double, ByVal X As Double) As Double
    If Y > 0 Then
        If X >= Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= -Y Then
            Atan2 = Atn(Y / X) + PI
        Else
            Atan2 = PI / 2 - Atn(X / Y)
        End If
    Else
        If X >= -Y Then
            Atan2 = Atn(Y / X)
        ElseIf X <= Y Then
            Atan2 = Atn(Y / X) - PI
        Else
            Atan2 = -Atn(X / Y) - PI / 2
        End If
    End If
End Function

