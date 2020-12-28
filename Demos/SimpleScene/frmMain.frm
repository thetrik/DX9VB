VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct3D9 VB6 scene demo by The trick"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9585.001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   506
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDrawTarget 
      BackColor       =   &H00404040&
      Caption         =   "Draw camera target"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7680
      TabIndex        =   9
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CheckBox chkNormals 
      BackColor       =   &H00404040&
      Caption         =   "Draw normals"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7680
      TabIndex        =   8
      Top             =   6660
      Width           =   1815
   End
   Begin VB.Frame fraOrientation 
      BackColor       =   &H00404040&
      Caption         =   "Orientation"
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   7680
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
      Begin SceneDemo.ctlVector ctlOrientation 
         Height          =   1395
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2461
      End
   End
   Begin VB.Frame fraPivot 
      BackColor       =   &H00404040&
      Caption         =   "Pivot point"
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   7680
      TabIndex        =   11
      Top             =   2940
      Width           =   1815
      Begin SceneDemo.ctlVector ctlPivot 
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
      End
   End
   Begin VB.Frame fraPos 
      BackColor       =   &H00404040&
      Caption         =   "Position"
      ForeColor       =   &H00FFFFFF&
      Height          =   1755
      Left            =   7680
      TabIndex        =   10
      Top             =   1140
      Width           =   1815
      Begin SceneDemo.ctlVector ctlPositon 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
      End
   End
   Begin VB.Frame fraCreate 
      BackColor       =   &H00404040&
      Caption         =   "Create"
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton optCreate 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   2
         Left            =   1200
         Picture         =   "frmMain.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Create a box"
         Top             =   300
         Width           =   495
      End
      Begin VB.OptionButton optCreate 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   1
         Left            =   660
         Picture         =   "frmMain.frx":07DB
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Create a cone"
         Top             =   300
         Width           =   495
      End
      Begin VB.OptionButton optCreate 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":0A3E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Create a sphere"
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.PictureBox picViewport 
      BackColor       =   &H00404040&
      Height          =   7275
      Left            =   120
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // +================================================================+
' // |                  Direct3D9 example includes:                   |
' // |                                                                |
' // | 1. Orbit camera movement;                                      |
' // | 2. Procedural generation of meshes (sphere, cone, box)         |
' // | 3. Mouse picking 3D object on scene                            |
' // | 4. Basic transformations (orientation/translation/pivot point) |
' // +----------------------------------------------------------------+
' // |                                          by The trick 2018 (c) |
' // +================================================================+
' //

Option Explicit

' // Mouse mode
' // User can select objects in the scene and create these objects as well. To control that
' // behavior program uses different mouse modes.
Public Enum eMouseMode
    
    MM_SELECT           ' // User click to select any mesh on scene
    
    MM_CREATE_SPHERE    ' // User start spehere creation
    MM_CREATE_CONE      ' // ... cone
    MM_CREATE_BOX       ' // ... box
        
    MM_CREATION         ' // Creation stages. For example when user want to create a sphere
                        ' // he click on the panel and mode changes to MM_CREATE_SPHERE, then
                        ' // user click first time to specify position of that object - mode
                        ' // is being changed to MM_CREATION to specify object-specific properies
                        ' // radius/ segments. When the sphere has been created mode is changed
                        ' // back to MM_SELECT.
    
End Enum

' // New object creation settings
' // In real application it's better to use an interface to creation
Private Type tNewObject
    tCreationPoint  As D3DVECTOR        ' // Creation point
    tLastValidPoint As D3DVECTOR2       ' // Last valid mouse point (screenspace)
    tStartPointScr  As D3DVECTOR2       ' // Start screen point
    lLevel          As Long             ' // Object level creation
    cObject         As CMesh            ' // Mesh object
    eObject         As eMouseMode       ' // Object type
    lColor          As Long             ' // Color of mesh
    fParameters()   As Single           ' // Parameters
End Type
    
Private mcScene         As CScene       ' // Scene
Private meCurrentMode   As eMouseMode   ' // Current mode
Private mcSelectedMesh  As CMesh        ' // Selected mesh

Dim tObjectCreation As tNewObject       ' // Creation object setting variable

' // Specified current selected mesh on scene and updates all fields
Public Property Set SelectedMesh( _
                    ByVal cValue As CMesh)
    Dim tRotation   As D3DVECTOR
    Dim tMtx        As D3DMATRIX
    
    Set mcSelectedMesh = cValue
    
    If cValue Is Nothing Then
        
        ' // Disable all controls if nothing select
        fraPos.Enabled = False
        fraPivot.Enabled = False
        fraOrientation.Enabled = False
        
        Exit Property
        
    Else
    
        fraPos.Enabled = True
        fraPivot.Enabled = True
        fraOrientation.Enabled = True
        
    End If
    
    ctlPositon.X = cValue.Position.X
    ctlPositon.Y = cValue.Position.Y
    ctlPositon.Z = cValue.Position.Z
    
    ctlPivot.X = cValue.PivotPoint.X
    ctlPivot.Y = cValue.PivotPoint.Y
    ctlPivot.Z = cValue.PivotPoint.Z
    
    ' // Convert quaternion to matrix
    D3DXMatrixRotationQuaternion tMtx, cValue.Orientation
    
    ' // Convert matrix to Euler angles
    tRotation = MatrixToEuler(tMtx)
    
    ctlOrientation.X = tRotation.X
    ctlOrientation.Y = tRotation.Y
    ctlOrientation.Z = tRotation.Z
    
End Property

' // Retrieve current selected object
Public Property Get SelectedMesh() As CMesh
    Set SelectedMesh = mcSelectedMesh
End Property

' // Specifiy current mode/updates mouse pointer and controls
Public Property Let CurrentMode( _
                    ByVal eValue As eMouseMode)
    Dim cOpt    As OptionButton
    
    meCurrentMode = eValue
    
    Select Case meCurrentMode
    Case MM_SELECT, MM_CREATION
    
        picViewport.MousePointer = vbArrow
        
        For Each cOpt In optCreate
            cOpt.Value = False
        Next
        
    Case MM_CREATE_SPHERE, MM_CREATE_CONE, MM_CREATE_BOX
        picViewport.MousePointer = vbCrosshair
    End Select
    
End Property

' // Retrieve current mode
Public Property Get CurrentMode() As eMouseMode
    CurrentMode = meCurrentMode
End Property

' // Determine whether draw camera target or not
Private Sub chkDrawTarget_Click()
    mcScene.DrawCameraTarget = chkDrawTarget.Value = vbChecked
End Sub

' // Determine whether draw normals on meshes or not
Private Sub chkNormals_Click()
    mcScene.DrawNormals = chkNormals.Value = vbChecked
End Sub

' // User changed orientation ;)
Private Sub ctlOrientation_Changed()
    Dim tQ  As D3DQUATERNION
    
    ' // Convert Euler angles to quaternion
    D3DXQuaternionRotationYawPitchRoll tQ, ctlOrientation.Y, ctlOrientation.X, ctlOrientation.Z
    
    SelectedMesh.Orientation = tQ
    
    mcScene.Render
    
End Sub

' // User changed pivot point
Private Sub ctlPivot_Changed()
    Dim tOffset As D3DVECTOR
    Dim tPivot  As D3DVECTOR
    
    ' // We need to change position according to new pivot point to keep position of mesh
    ' // It requires to find offset between old pivot position and new one and move the object
    ' // by this distance
    
    tPivot = vec3(ctlPivot.X, ctlPivot.Y, ctlPivot.Z)
    
    D3DXVec3Subtract tOffset, tPivot, SelectedMesh.PivotPoint
    
    SelectedMesh.PivotPoint = tPivot
    
    D3DXVec3Subtract tOffset, SelectedMesh.Position, tOffset
    
    SelectedMesh.Position = tOffset
    
    ctlPositon.X = tOffset.X
    ctlPositon.Y = tOffset.Y
    ctlPositon.Z = tOffset.Z
    
    mcScene.Render
    
End Sub

' // User changed position
Private Sub ctlPositon_Changed()
    
    SelectedMesh.Position = vec3(ctlPositon.X, ctlPositon.Y, ctlPositon.Z)
    
    mcScene.Render
    
End Sub

' // Initialization
Private Sub Form_Load()
    
    Set mcScene = New CScene
    
    ' // Initialize scene to draw on picturebox
    mcScene.InitializeScene picViewport.hwnd

    CurrentMode = MM_SELECT
    Set SelectedMesh = Nothing
    
End Sub

' // User want to create a mesh
Private Sub optCreate_Click( _
            ByRef Index As Integer)
    
    Select Case Index
    Case 0: CurrentMode = MM_CREATE_SPHERE
    Case 1: CurrentMode = MM_CREATE_CONE
    Case 2: CurrentMode = MM_CREATE_BOX
    End Select
    
End Sub

' // Click on viewport
Private Sub picViewport_MouseDown( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef X As Single, _
            ByRef Y As Single)
    
    ' // Check current mode
    If meCurrentMode = MM_SELECT Then
        
        If Button = vbLeftButton Then
            
            ' // If user click left button on viewport - select object
            Set SelectedMesh = mcScene.Pick(X, Y)

        End If
        
    Else
        
        If Button = vbRightButton Then
            
            ' // Cancellation if user click right button during object creation
            Select Case meCurrentMode
            Case MM_CREATE_SPHERE, MM_CREATE_CONE, MM_CREATE_BOX, MM_CREATION
                CurrentMode = MM_SELECT
            End Select
            
        ElseIf Button = vbLeftButton Then
            
            If meCurrentMode = MM_CREATION Then
                ' // Second and further creation stages
                NextCreationStage
            Else
                ' // First creation stage
                StartCreation X, Y
            End If
            
        End If
        
    End If
    
End Sub

' // First creation stage
Private Sub StartCreation( _
            ByVal fX As Single, _
            ByVal fY As Single)
    Dim tPoint  As D3DVECTOR

    ' // Find the intersection point between plane and screen ray
    If GetMouseIntersectionPlane(fX, fY, vec3(0, 0, 0), vec3(1, 0, 0), vec3(0, 0, 1), tPoint) Then
        
        With tObjectCreation
            
            ' // Create new mesh
            Set .cObject = New CMesh
            
            .lLevel = 0
            .tCreationPoint = tPoint            ' // Origin
            .tLastValidPoint = vec2(fX, fY)
            .tStartPointScr = .tLastValidPoint
            .eObject = meCurrentMode
            .cObject.Position = .tCreationPoint
            .lColor = Rnd * &HFFFFFF
            
            ' // Each object has own parameters
            ' // In real application it's better to make an interface that provides
            ' // individual approach of creation
            
            Select Case .eObject
            Case MM_CREATE_SPHERE
                
                ' // 0. Segments
                ' // 1. Radius
                
                ReDim .fParameters(1)
                
                .fParameters(0) = 32
                
            Case MM_CREATE_CONE
                
                ' // 0. Segments
                ' // 1. Down radius
                ' // 2. Top radius
                ' // 3. Height
                
                ReDim .fParameters(3)
                
                .fParameters(0) = 32
                
            Case MM_CREATE_BOX
                
                ' // 0. Width
                ' // 1. Depth
                ' // 2. Height
                
                ReDim .fParameters(2)
                
            End Select
            
            ' // Add the mesh to scene
            mcScene.Objects.Add .cObject
            
        End With
        
        ' // Change mode
        CurrentMode = MM_CREATION

    End If
    
End Sub

' // Select the next stage
Private Sub NextCreationStage()
    
    With tObjectCreation
    
    Select Case .eObject
    Case MM_CREATE_SPHERE
        
        .lLevel = .lLevel + 1
            
        If .lLevel >= 2 Then
            CurrentMode = MM_SELECT
        End If
            
    Case MM_CREATE_CONE
    
        .lLevel = .lLevel + 1
            
        If .lLevel >= 4 Then
            CurrentMode = MM_SELECT
        End If
        
    Case MM_CREATE_BOX
    
        .lLevel = .lLevel + 1
        
        If .lLevel >= 2 Then
            CurrentMode = MM_SELECT
        End If
        
    End Select
    
    End With
    
End Sub

' // Cone creation stages
Private Sub CreatingCone( _
            ByVal fX As Single, _
            ByVal fY As Single)
    Dim tPoint      As D3DVECTOR
    Dim tDelta      As D3DVECTOR2
    Dim fParameter  As Single
    
    With tObjectCreation
        
        If .lLevel = 1 Then
            ' // Find intersection point between wall and screen ray
            If Not GetMouseIntersectionPlane(fX, fY, .tCreationPoint, _
                        vec3(.tCreationPoint.X + 1, .tCreationPoint.Y, .tCreationPoint.Z), _
                        vec3(.tCreationPoint.X, .tCreationPoint.Y + 1, .tCreationPoint.Z), tPoint) Then Exit Sub
        Else
            ' // Find intersection point between floor and screen ray
            If Not GetMouseIntersectionPlane(fX, fY, vec3(0, 0, 0), vec3(1, 0, 0), vec3(0, 0, 1), tPoint) Then Exit Sub
        End If
        
        ' // Distance from start point
        D3DXVec3Subtract tPoint, tPoint, .tCreationPoint

        Select Case .lLevel
        Case 0      ' // Raduis 1
            
            ' // Radius
            fParameter = D3DXVec3Length(tPoint)
            .fParameters(1) = fParameter
            .fParameters(2) = fParameter
            
        Case 1      ' // Height
            
            fParameter = D3DXVec3Length(tPoint)
            .fParameters(3) = fParameter
            
        Case 2      ' // Radius 2
        
            fParameter = D3DXVec3Length(tPoint)
            .fParameters(2) = fParameter

        Case 3
            
            ' // Segments
            tDelta = vec2(fX - .tStartPointScr.X, fY - .tStartPointScr.X)
            
            fParameter = D3DXVec2Length(tDelta) / 5
            
            .fParameters(0) = Int(fParameter)
            
            If .fParameters(0) < 3 Then
                .fParameters(0) = 3
            ElseIf .fParameters(0) > 64 Then
                .fParameters(0) = 64
            End If

        End Select
        
        ' // Generate cone
        .cObject.CreateCone mcScene.Device, .fParameters(0), .fParameters(3), .fParameters(1), .fParameters(2), .lColor
        
        mcScene.Render

        .tLastValidPoint = vec2(fX, fY)
            
    End With

End Sub

' // Sphere creation stages
Private Sub CreatingSphere( _
            ByVal fX As Single, _
            ByVal fY As Single)
    Dim tPoint      As D3DVECTOR
    Dim tDelta      As D3DVECTOR2
    Dim fParameter  As Single
    
    ' // Find intersection point between floor and screen ray
    If Not GetMouseIntersectionPlane(fX, fY, vec3(0, 0, 0), vec3(1, 0, 0), vec3(0, 0, 1), tPoint) Then Exit Sub
    
    With tObjectCreation
        
        ' // Distance from start point
        D3DXVec3Subtract tPoint, tPoint, .tCreationPoint
        
        Select Case .lLevel
        Case 0
            
            ' // Radius
            fParameter = D3DXVec3Length(tPoint)
            .fParameters(1) = fParameter

        Case 1
            
            ' // Segments
            tDelta = vec2(fX - .tStartPointScr.X, fY - .tStartPointScr.X)
            
            fParameter = D3DXVec2Length(tDelta) / 5
            
            .fParameters(0) = Int(fParameter)
            
            If .fParameters(0) < 3 Then
                .fParameters(0) = 3
            ElseIf .fParameters(0) > 64 Then
                .fParameters(0) = 64
            End If

        End Select
        
        .cObject.CreateSphere mcScene.Device, .fParameters(0), .fParameters(1), .lColor
        
        mcScene.Render

        .tLastValidPoint = vec2(fX, fY)
            
    End With

End Sub

' // Box creation stages
Private Sub CreatingBox( _
            ByVal fX As Single, _
            ByVal fY As Single)
    Dim tPoint      As D3DVECTOR
    Dim tDelta      As D3DVECTOR2

    With tObjectCreation
        
        If .lLevel = 1 Then
            ' // Find intersection point between wall and screen ray
            If Not GetMouseIntersectionPlane(fX, fY, .tCreationPoint, _
                        vec3(.tCreationPoint.X + 1, .tCreationPoint.Y, .tCreationPoint.Z), _
                        vec3(.tCreationPoint.X, .tCreationPoint.Y + 1, .tCreationPoint.Z), tPoint) Then Exit Sub
        Else
            ' // Find intersection point between floor and screen ray
            If Not GetMouseIntersectionPlane(fX, fY, vec3(0, 0, 0), vec3(1, 0, 0), vec3(0, 0, 1), tPoint) Then Exit Sub
        End If
        
        D3DXVec3Subtract tPoint, tPoint, .tCreationPoint
        
        Select Case .lLevel
        Case 0
            
            ' // Width and depth
            .fParameters(0) = Abs(tPoint.X * 2)
            .fParameters(1) = Abs(tPoint.Z * 2)

        Case 1
            
            ' // Height
            .fParameters(2) = D3DXVec3Length(tPoint)
            
        End Select
        
        .cObject.CreateBox mcScene.Device, .fParameters(0), .fParameters(2), .fParameters(1), .lColor
        
        mcScene.Render

        .tLastValidPoint = vec2(fX, fY)
            
    End With

End Sub

' // Get intersection point between screen space ray and plane. If exists return true
Private Function GetMouseIntersectionPlane( _
                 ByVal fX As Single, _
                 ByVal fY As Single, _
                 ByRef tplane1 As D3DVECTOR, _
                 ByRef tplane2 As D3DVECTOR, _
                 ByRef tplane3 As D3DVECTOR, _
                 ByRef tPoint As D3DVECTOR) As Boolean
    Dim tPlane      As D3DPLANE
    Dim tRayFrom    As D3DVECTOR
    Dim tRayTo      As D3DVECTOR
    
    mcScene.RayFromScreenPos fX, fY, tRayFrom, tRayTo
    
    D3DXPlaneFromPoints tPlane, tplane1, tplane2, tplane3
    
    GetMouseIntersectionPlane = D3DXPlaneIntersectLine(tPoint, tPlane, tRayFrom, tRayTo)

End Function

' // User moves mouse in viewport
Private Sub picViewport_MouseMove( _
            ByRef Button As Integer, _
            ByRef Shift As Integer, _
            ByRef X As Single, _
            ByRef Y As Single)
    Static fOx  As Single
    Static fOy  As Single
    
    Select Case meCurrentMode
    Case MM_CREATION
        
        ' // Creation stages
        Select Case tObjectCreation.eObject
        Case MM_CREATE_SPHERE
            CreatingSphere X, Y
        Case MM_CREATE_CONE
            CreatingCone X, Y
        Case MM_CREATE_BOX
            CreatingBox X, Y
        End Select
        
    Case Else
        
        ' // Camera movement
        If Button = vbRightButton Then
        
            mcScene.Camera.RotateRel vec3(-(Y - fOy) / 100, -(X - fOx) / 100, 0)
            Render
            
        ElseIf Button = vbLeftButton Then
        
            mcScene.Camera.Zoom -(fOy - Y) / 10
            Render
            
        ElseIf Button = vbMiddleButton Then
        
            mcScene.Camera.Pan -(X - fOx) / 10, (Y - fOy) / 10
            Render
            
        End If
    
    End Select
    
    fOx = X: fOy = Y
    
End Sub

Private Sub Render()
    mcScene.Render
End Sub

Private Sub picViewport_Paint()
    Render
End Sub
