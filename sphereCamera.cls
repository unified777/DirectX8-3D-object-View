VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoolCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'Main camera parametres:
Private m_AOV               As Single               ' Angle Of View, usually PI/4
Private m_MaxVisualRange    As Single               ' Maximum distance from the camera, at which an object is visible
Private m_MinVisualRange    As Single               ' Minimum distance from the camera, at which an object is visible
Private m_Zoom              As Single               ' Zoom coefficient (not used in this smaple)
Private m_AspectRatio       As Single               ' Aspect ratio...
Private m_mGenerators       As D3DMATRIX            ' The generator matrix for this camera
Private m_mProjection       As D3DMATRIX            ' Projection transformation matrix for this camera
Private m_mView             As D3DMATRIX            ' View transformation matrix for this cmera

'Pivot point and lever:
Private m_vPivotPt          As D3DVECTOR            ' Position vector indicating the point (in scene coordinates) around which the camera will rotate
Private m_mPivToOrigin      As D3DMATRIX            ' A matrix describing the translation from the pivot point to the origin of the coordinate system
Private m_mOriginToPiv      As D3DMATRIX            ' A matrix describing the translation from the origin of the coordinate system to the pivot point

'Angular and linear velocities:
Private m_SpinLclX          As Single               ' The value of spin around a local (camera's) X-axis
Private m_SpinGlblY         As Single               ' The value of spin around a global (scene's) Y_axis
Private m_vVel              As D3DVECTOR            ' Vector of linear velocity.

'Transition variables:
Private m_vDestEye          As D3DVECTOR            ' The position of the camera at the end of a transition loop
Private m_vDestUp           As D3DVECTOR            ' The vertical axis of the camera at the end of a transition loop
Private m_vDestAt           As D3DVECTOR            ' The point observed by the camera at the end of a transition loop
Private m_vSrcEye           As D3DVECTOR            ' The position of the camera at the start of a transition loop
Private m_vSrcUp            As D3DVECTOR            ' The vertical axis of the camera at the start of a transition loop
Private m_vSrcAt            As D3DVECTOR            ' The point observed by the camera at the start of a transition loop
Private m_numFramesLeft     As Long                 ' The number of frames left in a transition loop
Private m_numFramesTotal    As Long                 ' The total number of frames in a transition loop
Private m_bInTransition     As Boolean              ' "True" if a transition is in progress

'Helpers:
Private i As Long
Private det As Single
Private vVctr1 As D3DVECTOR, vVctr2 As D3DVECTOR
Private vVctr3 As D3DVECTOR, vVctr4 As D3DVECTOR

'---------------------------------------
' Name: Setup()
' Desc:
'---------------------------------------
Friend Sub Setup(vEye As D3DVECTOR, vAt As D3DVECTOR, vUp As D3DVECTOR, _
                 ByVal MaxZ As Single, ByVal MinZ As Single, ByVal AspectRatio As Single)
     
    D3DXMatrixLookAtLH m_mView, vEye, vAt, vUp
    D3DXMatrixInverse m_mGenerators, det, m_mView
    m_MaxVisualRange = MaxZ
    m_MinVisualRange = MinZ
    m_Zoom = 1  ' by default
    m_AspectRatio = AspectRatio
    m_AOV = g_PI / (4 * m_Zoom)
    D3DXMatrixPerspectiveFovLH m_mProjection, m_AOV, m_AspectRatio, m_MinVisualRange, m_MaxVisualRange
    D3DXMatrixIdentity m_mOriginToPiv
    D3DXMatrixIdentity m_mPivToOrigin
End Sub

'----------------------------------------------
' Name: ChangeView()
' Desc: Enables unconditional and rapid change
'       of the camera's position
'----------------------------------------------
Friend Sub ChangeView(vEye As D3DVECTOR, vAt As D3DVECTOR, vUp As D3DVECTOR)
    D3DXMatrixLookAtLH m_mView, vEye, vAt, vUp
    D3DXMatrixInverse m_mGenerators, det, m_mView
End Sub

'-------------------------------------------
' Name: Pivot()
' Desc: Enables pivotting the camera around
'       a point in space.
'-------------------------------------------
Friend Sub Pivot(glbl_vPivotPt As D3DVECTOR, ByVal HorizRot As Single, ByVal VertRot As Single)
    ' Ignore the Pivot call if the camera is performing a transition:
    If m_bInTransition Then Exit Sub
    
    m_vPivotPt = glbl_vPivotPt
    ' Update the spins:
    m_SpinLclX = m_SpinLclX + HorizRot / g_dt
    m_SpinGlblY = m_SpinGlblY + VertRot / g_dt
    ' Matrix of translation from pivot point to the origin:
    D3DXMatrixTranslation m_mPivToOrigin, -m_vPivotPt.x, -m_vPivotPt.y, -m_vPivotPt.z
    ' Matrix of translation from the origin to the pivot point:
    D3DXMatrixTranslation m_mOriginToPiv, m_vPivotPt.x, m_vPivotPt.y, m_vPivotPt.z
End Sub

'---------------------------------------
' Name: Move()
' Desc: Moves the camera along a vector.
'---------------------------------------
Friend Sub Move(glbl_vShift As D3DVECTOR)
    ' If the camera is not performing a transition, update its linear velocity vector
    If Not m_bInTransition Then D3DXVec3Add m_vVel, m_vVel, ScaleVec3(glbl_vShift, 1 / g_dt)
End Sub

'---------------------------------------
' Name: StopSpins()
' Desc: Stops spins...
'---------------------------------------
Friend Sub StopSpins()
    m_SpinLclX = 0
    m_SpinGlblY = 0
End Sub

'----------------------------------------------
' Name: StopMovement()
' Desc: Stops movement along a specified axis
'----------------------------------------------
Friend Sub StopMovement(ByVal bX As Boolean, bY As Boolean, bZ As Boolean)
    If bX Then m_vVel.x = 0
    If bY Then m_vVel.y = 0
    If bZ Then m_vVel.z = 0
End Sub

'---------------------------------------
' Name: Transit()
' Desc: Starts a transition loop
'---------------------------------------
Friend Sub Transit(vDestEye As D3DVECTOR, vDestAt As D3DVECTOR, vDestUp As D3DVECTOR, ByVal numFrames As Long)
    ' First stop all camera motions:
    StopMovement True, True, True
    StopSpins
    ' Now set the transition variables:
    ' 1) The initial camera vectors:
    m_vSrcEye = vec3(m_mGenerators.m41, m_mGenerators.m42, m_mGenerators.m43)
    m_vSrcAt = vec3(m_mGenerators.m41 + m_mGenerators.m31, m_mGenerators.m42 + m_mGenerators.m32, m_mGenerators.m43 + m_mGenerators.m33)
    m_vSrcUp = vec3(m_mGenerators.m21, m_mGenerators.m22, m_mGenerators.m23)
    ' 2) The final camera vectors:
    m_vDestEye = vDestEye
    m_vDestAt = vDestAt
    m_vDestUp = vDestUp
    ' 3) The total number of frames for this transition:
    m_numFramesTotal = numFrames
    ' 4) The remaining number of frames (currently equall to the total number).
    m_numFramesLeft = numFrames
    ' 5) The boolean indicating, that the camera is in transition mode:
    m_bInTransition = True
End Sub

'---------------------------------------
' Name: Update()
' Desc:
'---------------------------------------
Friend Sub Update()
    Dim RotLclX         As Single           ' Rotation around a local X-Axis
    Dim PitchAngle      As Single           ' Pitch angle...
    Dim CosPitchAngle   As Single           ' Cosine of the pitch angle
    Dim mRotationGlblY  As D3DMATRIX        ' Matrix of rotation around a global Y-Axis
    Dim mRotationLclX   As D3DMATRIX        ' Matrix of rotation around a local X-Axis
    Dim mTranslation    As D3DMATRIX        ' Matrix of translation
    Dim mFinalTrans     As D3DMATRIX        ' The final (summed) transformation matrix
    Dim BlendWeight     As Single           ' Used for transition
            
    If m_bInTransition Then
        ' The process of transition is very similiar to a technique called tweening, used for
        ' animating flexible bodies. It is build around Linear Interpolation (Lerp) of vectors.
        BlendWeight = Sin((1 - m_numFramesLeft / m_numFramesTotal) * g_PI / 2)
        BlendWeight = BlendWeight * BlendWeight
        D3DXVec3Lerp vVctr1, m_vSrcEye, m_vDestEye, BlendWeight
        D3DXVec3Lerp vVctr2, m_vSrcAt, m_vDestAt, BlendWeight
        D3DXVec3Lerp vVctr3, m_vSrcUp, m_vDestUp, BlendWeight
        D3DXMatrixLookAtLH m_mView, vVctr1, vVctr2, vVctr3
        D3DXMatrixInverse m_mGenerators, det, m_mView
        m_numFramesLeft = m_numFramesLeft - 1
        If m_numFramesLeft = 0 Then m_bInTransition = False
    Else
        ' The angle between camera's line of sight and the vetrical axis:
        CosPitchAngle = m_mGenerators.m32
        If CosPitchAngle > -1 Then
            PitchAngle = Atn(-CosPitchAngle / Sqr(-CosPitchAngle * CosPitchAngle + 1)) + g_PI / 2
        Else
            PitchAngle = 0
        End If
            
        ' The PitchAngle of rotation around local X axis:
        RotLclX = m_SpinLclX * g_dt
        ' The rotation is limited to angles between pi/2 and a little less than pi
        ' (measured form the vertical axis):
        If PitchAngle + RotLclX <= g_PI / 2 + 0.05 Then RotLclX = 0: m_SpinLclX = 0
        If PitchAngle + RotLclX >= g_PI - 0.5 Then
            If RotLclX > 0 Then RotLclX = 0: m_SpinLclX = 0
        End If
        
        ' If the camera's elevation above the table should fall below a certain value,
        ' stop all linear movement:
        If m_mGenerators.m42 + m_vVel.y * g_dt < 0.1 Then m_vVel = vec3(0, 0, 0)
                                
        ' Matrix of rotation around local X axis:
        D3DXMatrixRotationAxis mRotationLclX, vec3(m_mGenerators.m11, m_mGenerators.m12, m_mGenerators.m13), RotLclX
        ' Matrix of rotation around the global Y axis:
        D3DXMatrixRotationY mRotationGlblY, m_SpinGlblY * g_dt
        ' Matrix of translation along the linear velocity vector:
        D3DXMatrixTranslation mTranslation, m_vVel.x * g_dt, m_vVel.y * g_dt, m_vVel.z * g_dt
        
        ' Create the transformation matrix:
        ' 1) Translate from pivot point to the origin:
        D3DXMatrixMultiply mFinalTrans, m_mGenerators, m_mPivToOrigin
        ' 2) Rotate around the local X axis:
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mRotationLclX
        ' 3) Rotate around the global Y axis:
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mRotationGlblY
        ' 4) Return to the pivot point:
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, m_mOriginToPiv
        ' 5) Move the camera along integrated m_vVel vector:
        D3DXMatrixMultiply mFinalTrans, mFinalTrans, mTranslation
        
        ' The new generator matrix:
        m_mGenerators = mFinalTrans
        ' And the new view transformation matrix:
        D3DXMatrixInverse m_mView, det, m_mGenerators
        
        ' Reduce the spins and the linear velocity, by a constant coefficient:
        m_SpinLclX = m_SpinLclX * 0.7
        m_SpinGlblY = m_SpinGlblY * 0.7
        D3DXVec3Scale m_vVel, m_vVel, 0.7
    End If
End Sub



'       +--------------------------+
'       |       PROPERTIES:        |
'       +--------------------------+

Friend Property Get Generators() As D3DMATRIX
     Generators = m_mGenerators
End Property

Friend Property Get Position() As D3DVECTOR
     With Position
          .x = m_mGenerators.m41
          .y = m_mGenerators.m42
          .z = m_mGenerators.m43
     End With
End Property

Friend Property Get MaxVisualRange() As Single
     MaxVisualRange = m_MaxVisualRange
End Property

Friend Property Get MinVisualRange() As Single
     MinVisualRange = m_MinVisualRange
End Property

Friend Property Get ViewMatrix() As D3DMATRIX
     ViewMatrix = m_mView
End Property

Friend Property Get ProjectionMatrix() As D3DMATRIX
     ProjectionMatrix = m_mProjection
End Property

