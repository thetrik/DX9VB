VERSION 5.00
Begin VB.UserControl ctlVector 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ScaleHeight     =   1470
   ScaleWidth      =   2745
   Begin VB.TextBox txtPos 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   300
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   2355
   End
   Begin VB.TextBox txtPos 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Text            =   "0"
      Top             =   540
      Width           =   2355
   End
   Begin VB.TextBox txtPos 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Text            =   "0"
      Top             =   120
      Width           =   2355
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   150
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   150
   End
   Begin VB.Label lblPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   150
   End
End
Attribute VB_Name = "ctlVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Control that represents vector in recatangular space
' //

Option Explicit

Private Declare Function VarR4FromStr Lib "oleaut32" ( _
                         ByVal lpstrValue As Long, _
                         ByVal lcid As Long, _
                         ByVal lFlags As Long, _
                         ByRef pF4 As Single) As Long
Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

' // Raised when user change any property in textbox
Public Event Changed()

' // Internal copy
Private mtVector    As D3DVECTOR

Public Property Get X() As Single
    X = mtVector.X
End Property

Public Property Let X( _
                    ByVal fValue As Single)
    
    mtVector.X = fValue
    txtPos(0).Text = Format$(fValue, "0.0000")

End Property

Public Property Get Y() As Single
    Y = mtVector.Y
End Property

Public Property Let Y( _
                    ByVal fValue As Single)
    
    mtVector.Y = fValue
    txtPos(1).Text = Format$(fValue, "0.0000")

End Property

Public Property Get Z() As Single
    Z = mtVector.Z
End Property

Public Property Let Z( _
                    ByVal fValue As Single)
    
    mtVector.Z = fValue
    txtPos(2).Text = Format$(fValue, "0.0000")

End Property

Private Sub txtPos_KeyDown( _
            ByRef Index As Integer, _
            ByRef KeyCode As Integer, _
            ByRef Shift As Integer)
            
    If KeyCode = vbKeyReturn Then
        txtPos_Validate Index, False
    End If
    
End Sub

Private Sub txtPos_Validate( _
            ByRef Index As Integer, _
            ByRef Cancel As Boolean)
    Dim fOut    As Single
    
    If VarR4FromStr(StrPtr(txtPos(Index).Text), GetUserDefaultLCID, 0, fOut) < 0 Then
    
        Select Case Index
        Case 0: txtPos(Index).Text = Format$(mtVector.X)
        Case 1: txtPos(Index).Text = Format$(mtVector.Y)
        Case 2: txtPos(Index).Text = Format$(mtVector.Z)
        End Select
        
    Else
    
        Select Case Index
        Case 0: mtVector.X = fOut
        Case 1: mtVector.Y = fOut
        Case 2: mtVector.Z = fOut
        End Select
        
        RaiseEvent Changed
        
    End If
    
End Sub

Private Sub UserControl_Initialize()
    X = 0: Y = 0: Z = 0
End Sub

Private Sub UserControl_Resize()
    Dim lIndex  As Long
    
    For lIndex = 0 To 2
        txtPos(lIndex).Width = UserControl.ScaleWidth - txtPos(lIndex).Left
    Next
    
End Sub
