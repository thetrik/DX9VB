VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // Material class
' //

Option Explicit

Private mlDiffuseColor  As Long
Private mlSpecularColor As Long
Private mlAmbientColor  As Long
Private mfSpecularPower As Single
Private mtMaterial      As D3DMATERIAL9

Dim mbChanged   As Boolean

' // Get material
Public Property Get Material() As D3DMATERIAL9
    
    If mbChanged Then
        
        With mtMaterial.Diffuse
        
            .r = (mlDiffuseColor And &HFF) / &HFF
            .g = (mlDiffuseColor And &HFF00&) / &HFF00&
            .b = (mlDiffuseColor And &HFF0000) / &HFF0000
            .a = 1
        
        End With
        
        With mtMaterial.Ambient
        
            .r = (mlAmbientColor And &HFF) / &HFF
            .g = (mlAmbientColor And &HFF00&) / &HFF00&
            .b = (mlAmbientColor And &HFF0000) / &HFF0000
            .a = 1
        
        End With
        
        With mtMaterial.Specular
        
            .r = (mlSpecularColor And &HFF) / &HFF
            .g = (mlSpecularColor And &HFF00&) / &HFF00&
            .b = (mlSpecularColor And &HFF0000) / &HFF0000
            .a = 1
        
        End With
        
        mtMaterial.Power = mfSpecularPower
        
    End If
    
    Material = mtMaterial
    
End Property

Public Property Let DiffuseColor( _
                    ByVal lValue As Long)
                    
    mlDiffuseColor = lValue
    mbChanged = True
    
End Property
Public Property Get DiffuseColor() As Long
    DiffuseColor = mlDiffuseColor
End Property

Public Property Let SpecularColor( _
                    ByVal lValue As Long)
                    
    mlSpecularColor = lValue
    mbChanged = True
    
End Property
Public Property Get SpecularColor() As Long
    SpecularColor = mlSpecularColor
End Property

Public Property Let AmbientColor( _
                    ByVal lValue As Long)
                    
    mlAmbientColor = lValue
    mbChanged = True
    
End Property
Public Property Get AmbientColor() As Long
    AmbientColor = mlAmbientColor
End Property

Public Property Let SpecularPower( _
                    ByVal fValue As Single)
                    
    mfSpecularPower = fValue
    mbChanged = True
    
End Property
Public Property Get SpecularPower() As Single
    SpecularPower = mfSpecularPower
End Property

Private Sub Class_Initialize()
    
    mlDiffuseColor = vbWhite
    mlSpecularColor = vbWhite
    mlAmbientColor = &H303030
    mfSpecularPower = 30
    mbChanged = True
    
End Sub
