VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PoolTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'           +=====================================================+
'           | WARNING:                                            |
'           | This class module is only partially commented.      |
'           | Most of its functions deal with procedural creation |
'           | of the pool table. Usually you construct the table  |
'           | using a 3D modelling tool.                          |
'           | Another words, don't waste time on reading this     |
'           +=====================================================+
'

' Flexible Vertex Format for the table:
Private Type TABLEVERTEX
    XYZ As D3DVECTOR
    Normal As D3DVECTOR
    Tex1 As D3DVECTOR2
End Type
Private Const D3DFVF_TABLEVERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

' Flexible Vertex Format for table shadows:
Private Type SHADOWVERTEX
    XYZ As D3DVECTOR
    Color As Long
End Type
Private Const D3DFVF_SHADOWVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)


' Direct3D device reference, vertex and index buffers:
Private m_D3DDevice8        As Direct3DDevice8
Private m_D3DX8             As D3DX8

Private m_BaseVB            As Direct3DVertexBuffer8
Private m_BorderVB          As Direct3DVertexBuffer8
Private m_TopVB             As Direct3DVertexBuffer8
Private m_SideVB            As Direct3DVertexBuffer8
Private m_BumpersVB         As Direct3DVertexBuffer8
Private m_BumpShadowVB      As Direct3DVertexBuffer8
Private m_PocketVB          As Direct3DVertexBuffer8

Private m_BaseIB            As Direct3DIndexBuffer8
Private m_BorderIB          As Direct3DIndexBuffer8
Private m_TopIB             As Direct3DIndexBuffer8
Private m_SideIB            As Direct3DIndexBuffer8
Private m_BumpersIB         As Direct3DIndexBuffer8

' Textures:
Private m_TopTexture        As Direct3DTexture8
Private m_BorderTexture     As Direct3DTexture8
Private m_EmptyTexture      As Direct3DTexture8
Private m_PocketTexture     As Direct3DTexture8

' Vertex size and buffer counters:
Private m_sizeVertex        As Long
Private m_sizeShadowVertex  As Long
Private m_numBaseVB         As Long
Private m_numBorderVB       As Long
Private m_numTopVB          As Long
Private m_numSideVB         As Long
Private m_numBumpersVB      As Long
Private m_numBaseIB         As Long
Private m_numBorderIB       As Long
Private m_numTopIB          As Long
Private m_numSideIB         As Long
Private m_numBumpersIB      As Long

' Materials:
Private m_matBase           As D3DMATERIAL8
Private m_matBorder         As D3DMATERIAL8
Private m_matTop            As D3DMATERIAL8
Private m_matSides          As D3DMATERIAL8
Private m_matBumpers        As D3DMATERIAL8

' Table's dimensions:
Private m_MaxX              As Single
Private m_MinX              As Single
Private m_MaxZ              As Single
Private m_MinZ              As Single
Private m_BoundaryHght      As Single
Private m_BoundaryWdth      As Single

' Physics data:
Private m_BumperElasticity  As Single
Private m_Friction          As Single

' Bumpers array:
Private m_Bumpers(1 To 6, 1 To 4)   As D3DVECTOR

' Pockets:
Private m_Pockets(1 To 6)  As D3DVECTOR
Private m_PocketRadius     As Single

' Shadow Plane
Private m_plnShadow         As D3DPLANE

' Light direction:
Private m_vLightDir         As D3DVECTOR4

'Helpers and iterators:
Private i As Long, j As Long, k As Long
Private mIdentity As D3DMATRIX
Private v1 As D3DVECTOR, v2 As D3DVECTOR, v3 As D3DVECTOR, v4 As D3DVECTOR
Private VertexList() As TABLEVERTEX
Private IndexList() As Integer

'------------------------
' Name: Render()
' Desc: Draws the table
'------------------------
Friend Sub Render()
    Dim matBlack        As D3DMATERIAL8
    Dim mShadowTrans    As D3DMATRIX
    Dim mPocketGen      As D3DMATRIX
    
    If m_D3DDevice8 Is Nothing Then Exit Sub
    
    With m_D3DDevice8
        .SetTransform D3DTS_WORLD, mIdentity
        .SetVertexShader D3DFVF_TABLEVERTEX
         
         ' The border:
        .SetVertexShader D3DFVF_TABLEVERTEX
        .SetTransform D3DTS_WORLD, mIdentity
        .SetTexture 0, m_BorderTexture
        .SetMaterial m_matBorder
        .SetStreamSource 0, m_BorderVB, m_sizeVertex
        .SetIndices m_BorderIB, 0
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numBorderVB, 0, m_numBorderIB / 3
        
        ' Table top:
        .SetTexture 0, m_TopTexture
        .SetMaterial m_matTop
        .SetStreamSource 0, m_TopVB, m_sizeVertex
        .SetIndices m_TopIB, 0
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numTopVB, 0, m_numTopIB / 3
         
        ' Pockets:
        .SetMaterial matBlack
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
        .SetStreamSource 0, m_PocketVB, m_sizeVertex
        .SetTexture 0, m_PocketTexture
        For i = 1 To 6
            D3DXMatrixTranslation mPocketGen, m_Pockets(i).x, 0.005, m_Pockets(i).z
            .SetTransform D3DTS_WORLD, mPocketGen
            .DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
        Next i
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
          
        ' The base:
        .SetTexture 0, m_EmptyTexture
        .SetMaterial m_matBase
        .SetTransform D3DTS_WORLD, mIdentity
        .SetStreamSource 0, m_BaseVB, m_sizeVertex
        .SetIndices m_BaseIB, 0
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numBaseVB, 0, m_numBaseIB / 3
               
        ' Sides:
        .SetMaterial m_matSides
        .SetStreamSource 0, m_SideVB, m_sizeVertex
        .SetIndices m_SideIB, 0
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numSideVB, 0, m_numSideIB / 3
                
        ' Bumpers:
        .SetMaterial m_matBumpers
        .SetStreamSource 0, m_BumpersVB, m_sizeVertex
        .SetIndices m_BumpersIB, 0
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numBumpersVB, 0, m_numBumpersIB / 3
        ' Bumpers' shadows:
        D3DXMatrixShadow mShadowTrans, m_vLightDir, m_plnShadow
        .SetTransform D3DTS_WORLD, mShadowTrans
        .SetVertexShader D3DFVF_SHADOWVERTEX
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CW
        .SetRenderState D3DRS_SPECULARENABLE, 0
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
        .SetRenderState D3DRS_LIGHTING, 0
        .SetMaterial matBlack
        .SetStreamSource 0, m_BumpShadowVB, m_sizeShadowVertex
        ' Draw the shadow:
        .DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, m_numBumpersVB, 0, m_numBumpersIB / 3
        .SetRenderState D3DRS_LIGHTING, 1
        .SetRenderState D3DRS_SPECULARENABLE, 1
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
    
    End With
End Sub

'------------------------------------------------------
' Name: Create()
' Desc: Creates a table according to some input data.
'------------------------------------------------------
Friend Sub Create(ByVal inTableWidth As Single, ByVal inTableLength As Single, _
                  ByVal inD3DDevice8 As Direct3DDevice8, ByVal inD3DX8 As D3DX8)
    Dim Accomplished As Boolean
    
    ' Link the table object with a valid Direct3DDevice8 object:
    Set m_D3DDevice8 = inD3DDevice8
    Set m_D3DX8 = inD3DX8
    
    ' Set elasticity and friction factors:
    m_Friction = 0
    m_BumperElasticity = 0.75
    
    'Compute Table 's edges:
    m_MinX = -inTableWidth / 2: m_MaxX = inTableWidth / 2
    m_MinZ = -inTableLength / 2: m_MaxZ = inTableLength / 2
    
    'Boundary height and width:
    m_BoundaryHght = 0.1
    m_BoundaryWdth = 0.2
    
    'Set pockets' radius and centre points:
    m_PocketRadius = 0.08
    'Right side:
    m_Pockets(1) = vec3(m_MaxX - m_BoundaryWdth, 0, m_MinZ + m_BoundaryWdth)
    m_Pockets(2) = vec3(m_MaxX - m_BoundaryWdth, 0, 0)
    m_Pockets(3) = vec3(m_MaxX - m_BoundaryWdth, 0, m_MaxZ - m_BoundaryWdth)
    'Left side:
    m_Pockets(4) = vec3(m_MinX + m_BoundaryWdth, 0, m_MaxZ - m_BoundaryWdth)
    m_Pockets(5) = vec3(m_MinX + m_BoundaryWdth, 0, 0)
    m_Pockets(6) = vec3(m_MinX + m_BoundaryWdth, 0, m_MinZ + m_BoundaryWdth)
                      
    '--------------------
    ' Create pockets:
    '--------------------
    Accomplished = CreatePocket
    If Not Accomplished Then GoTo RaiseError
    
    '--------------------
    ' Table's base:
    '--------------------
    Accomplished = CreateTableBase
    If Not Accomplished Then GoTo RaiseError
   
   '--------------------
    ' Table border:
    '--------------------
    Accomplished = CreateTableBorder
    If Not Accomplished Then GoTo RaiseError
   
    '--------------------
    ' Table top:
    '--------------------
    Accomplished = CreateTableTop
    If Not Accomplished Then GoTo RaiseError
       
    '--------------------
    ' Table sides:
    '--------------------
    Accomplished = CreateTableSides
    If Not Accomplished Then GoTo RaiseError
        
    '--------------------
    ' Table bumpers:
    '--------------------
    Accomplished = CreateTableBumpers
    If Not Accomplished Then GoTo RaiseError
    
    '------------------------------------
    ' Light direction and shadow plane:
    '------------------------------------
    m_vLightDir = GetLightDir
    D3DXPlaneFromPoints m_plnShadow, vec3(0, 0.01, 0), vec3(1, 0.01, 0), vec3(0, 0.01, 1)
    
    '------------------------------------
    ' Erase the vertex and index arrays:
    '------------------------------------
    Erase VertexList, IndexList
           
    Exit Sub

RaiseError:
    Err.Raise vbObjectError + 4, , "Unable to create the Table"
End Sub

'----------------------------------------------------------
' Name: GetBumpers()
' Desc: Fills an array (6 x 5) provided by the caller
'       with vector data that make up the 6 bumpers.
'----------------------------------------------------------
Friend Sub GetBumpers(ByRef retBumperArray() As D3DVECTOR)
    For i = 1 To 6
        retBumperArray(i, 1) = m_Bumpers(i, 1)
        For j = 2 To 5
            retBumperArray(i, j) = m_Bumpers(i, j - 1)
        Next j
    Next i
End Sub

'-----------------------------------------------------------
' Name: GetGameArea()
' Desc: Returns the minimum and maximum Z and X coordinates
'       that fit into the table's game area.
'-----------------------------------------------------------
Friend Sub GetGameArea(ByRef retMinZ As Single, ByRef retMaxZ As Single, ByRef retMinX As Single, ByRef retMaxX As Single)
    retMinZ = m_Bumpers(4, 3).z
    retMaxZ = m_Bumpers(1, 3).z
    retMinX = m_Bumpers(5, 3).x
    retMaxX = m_Bumpers(2, 3).x
End Sub

'----------------------------------------------------------------------
' Name: GetPhysicalConstants()
' Desc: Returns elasticity coefficient (used for collision responce)
'       and friction coefficient (should be used for things like
'       collisions with a vertical spin vector, but since the current
'       collision detection and responce "engine" does not support
'       spins at all, the coefficient is quite useless...)
'-----------------------------------------------------------------------
Friend Sub GetPhysicalConstants(ByRef retElasticity As Single, ByRef retFriction As Single)
    retFriction = m_Friction
    retElasticity = m_BumperElasticity
End Sub

'----------------------------------------------------------------------
' Name: GetPockets()
' Desc: Returns a six element array containing position vectors for
'       the six pockets and a single containing the pocket radius.
'----------------------------------------------------------------------
Friend Sub GetPockets(ByRef retPocketCentres() As D3DVECTOR, ByRef retPocketRadius As Single)
    For i = 1 To 6
        retPocketCentres(i) = m_Pockets(i)
    Next i
    retPocketRadius = m_PocketRadius
End Sub

'======================================== PRIVATE PROCEDURES ===========================================

Private Sub Class_Initialize()
    D3DXMatrixIdentity mIdentity
End Sub

Private Sub Class_Terminate()
    Set m_D3DDevice8 = Nothing
    Set m_D3DX8 = Nothing
    
    Set m_BaseVB = Nothing
    Set m_BorderVB = Nothing
    Set m_TopVB = Nothing
    Set m_SideVB = Nothing
    Set m_BumpersVB = Nothing
    Set m_BumpShadowVB = Nothing
    Set m_PocketVB = Nothing
    
    Set m_BaseIB = Nothing
    Set m_BorderIB = Nothing
    Set m_TopIB = Nothing
    Set m_SideIB = Nothing
    Set m_BumpersIB = Nothing
    
    Set m_BorderTexture = Nothing
    Set m_TopTexture = Nothing
    Set m_EmptyTexture = Nothing
End Sub

Private Function CreateTableBase() As Boolean
    Dim XYZList() As D3DVECTOR
    ReDim VertexList(35)
    ReDim IndexList(53)
    
    'Initail counter setting:
    m_numBaseVB = 0: m_numBaseIB = 0
    
    With m_matBase
        .Ambient = MakeD3DCOLORVALUE(153 / 255, 113 / 255, 84 / 255, 0)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(0, 0, 0, 0)
    End With
    
    '----------------------------------------------------------------------------------------
    ' Base of the table (4 faces, 16 vertices):
    '----------------------------------------------------------------------------------------
        'Bottom:
        ReDim XYZList(1 To 8)
        XYZList(1) = vec3(m_MinX + 0.3, -1, m_MinZ + 0.3)
        XYZList(2) = vec3(m_MaxX - 0.3, -1, m_MinZ + 0.3)
        XYZList(3) = vec3(m_MaxX - 0.3, -1, m_MaxZ - 0.3)
        XYZList(4) = vec3(m_MinX + 0.3, -1, m_MaxZ - 0.3)
        'Top:
        XYZList(5) = vec3(m_MinX + 0.1, -0.4, m_MinZ + 0.1)
        XYZList(6) = vec3(m_MaxX - 0.1, -0.4, m_MinZ + 0.1)
        XYZList(7) = vec3(m_MaxX - 0.1, -0.4, m_MaxZ - 0.1)
        XYZList(8) = vec3(m_MinX + 0.1, -0.4, m_MaxZ - 0.1)
        
        'Rewrite these values to the VertexList:
        j = m_numBaseVB
        For i = 1 To 4
            If i < 4 Then k = i Else k = 0
            VertexList(j).XYZ = XYZList(i + 0): j = j + 1
            VertexList(j).XYZ = XYZList(i + 4): j = j + 1
            VertexList(j).XYZ = XYZList(k + 5): j = j + 1
            VertexList(j).XYZ = XYZList(k + 1): j = j + 1
        Next i
        m_numBaseVB = j
             
        'Now, the indexes:
        j = m_numBaseIB
        For i = 0 To 3
            IndexList(j) = i * 4 + 0: j = j + 1
            IndexList(j) = i * 4 + 2: j = j + 1
            IndexList(j) = i * 4 + 1: j = j + 1
            IndexList(j) = i * 4 + 0: j = j + 1
            IndexList(j) = i * 4 + 3: j = j + 1
            IndexList(j) = i * 4 + 2: j = j + 1
        Next i
        m_numBaseIB = j
        
        'The normal vectors:
        For i = 0 To m_numBaseIB - 1 Step 3
            D3DXVec3Subtract v1, VertexList(IndexList(i + 1)).XYZ, VertexList(IndexList(i + 2)).XYZ
            D3DXVec3Subtract v2, VertexList(IndexList(i)).XYZ, VertexList(IndexList(i + 1)).XYZ
            D3DXVec3Cross v3, v1, v2
            D3DXVec3Normalize v3, v3
            VertexList(IndexList(i + 0)).Normal = v3
            VertexList(IndexList(i + 1)).Normal = v3
            VertexList(IndexList(i + 2)).Normal = v3
        Next i
    
    '----------------------------------------------------------------------------------------
    ' Bottom of the table (1 face, 4 vertices):
    '----------------------------------------------------------------------------------------
        'Vertices:
        VertexList(m_numBaseVB + 0).XYZ = vec3(m_MinX, -0.4, m_MinZ)
        VertexList(m_numBaseVB + 1).XYZ = vec3(m_MaxX, -0.4, m_MinZ)
        VertexList(m_numBaseVB + 2).XYZ = vec3(m_MaxX, -0.4, m_MaxZ)
        VertexList(m_numBaseVB + 3).XYZ = vec3(m_MinX, -0.4, m_MaxZ)
        For i = 0 To 3
            VertexList(m_numBaseVB + i).Normal = vec3(0, -1, 0)
        Next i
        
        'Indices:
        IndexList(m_numBaseIB + 0) = m_numBaseVB + 0
        IndexList(m_numBaseIB + 1) = m_numBaseVB + 2
        IndexList(m_numBaseIB + 2) = m_numBaseVB + 1
        IndexList(m_numBaseIB + 3) = m_numBaseVB + 0
        IndexList(m_numBaseIB + 4) = m_numBaseVB + 3
        IndexList(m_numBaseIB + 5) = m_numBaseVB + 2
    '----------------------------------------------------------------------------------------
    
    'Update the counters:
    m_numBaseVB = m_numBaseVB + 4
    m_numBaseIB = m_numBaseIB + 6
    
         
    '----------------------------------------------------------------------------------------
    ' Table'outer side surfaces (4 faces, 16 vertices):
    '----------------------------------------------------------------------------------------
        'Vertices:
        'First wall: facing -Z
        VertexList(m_numBaseVB + 0).XYZ = vec3(m_MaxX, m_BoundaryHght, m_MinZ)
        VertexList(m_numBaseVB + 1).XYZ = vec3(m_MinX, m_BoundaryHght, m_MinZ)
        VertexList(m_numBaseVB + 2).XYZ = vec3(m_MinX, -0.4, m_MinZ)
        VertexList(m_numBaseVB + 3).XYZ = vec3(m_MaxX, -0.4, m_MinZ)
        For i = 0 To 3
            VertexList(m_numBaseVB + i).Normal = vec3(0, 0, -1)
        Next i
        'Second wall: facing +X
        VertexList(m_numBaseVB + 4).XYZ = vec3(m_MaxX, m_BoundaryHght, m_MaxZ)
        VertexList(m_numBaseVB + 5).XYZ = vec3(m_MaxX, m_BoundaryHght, m_MinZ)
        VertexList(m_numBaseVB + 6).XYZ = vec3(m_MaxX, -0.4, m_MinZ)
        VertexList(m_numBaseVB + 7).XYZ = vec3(m_MaxX, -0.4, m_MaxZ)
        For i = 4 To 7
            VertexList(m_numBaseVB + i).Normal = vec3(1, 0, 0)
        Next i
        'Third wall: facing +Z
        VertexList(m_numBaseVB + 8).XYZ = vec3(m_MinX, m_BoundaryHght, m_MaxZ)
        VertexList(m_numBaseVB + 9).XYZ = vec3(m_MaxX, m_BoundaryHght, m_MaxZ)
        VertexList(m_numBaseVB + 10).XYZ = vec3(m_MaxX, -0.4, m_MaxZ)
        VertexList(m_numBaseVB + 11).XYZ = vec3(m_MinX, -0.4, m_MaxZ)
        For i = 8 To 11
            VertexList(m_numBaseVB + i).Normal = vec3(0, 0, 1)
        Next i
        'Fourth wall: facing -X
        VertexList(m_numBaseVB + 12).XYZ = vec3(m_MinX, m_BoundaryHght, m_MinZ)
        VertexList(m_numBaseVB + 13).XYZ = vec3(m_MinX, m_BoundaryHght, m_MaxZ)
        VertexList(m_numBaseVB + 14).XYZ = vec3(m_MinX, -0.4, m_MaxZ)
        VertexList(m_numBaseVB + 15).XYZ = vec3(m_MinX, -0.4, m_MinZ)
        For i = 12 To 15
            VertexList(m_numBaseVB + i).Normal = vec3(-1, 0, 0)
        Next i
        
        'Indices:
        For i = 0 To 3
            IndexList(m_numBaseIB + i * 6 + 0) = m_numBaseVB + i * 4 + 0
            IndexList(m_numBaseIB + i * 6 + 1) = m_numBaseVB + i * 4 + 1
            IndexList(m_numBaseIB + i * 6 + 2) = m_numBaseVB + i * 4 + 2
            IndexList(m_numBaseIB + i * 6 + 3) = m_numBaseVB + i * 4 + 0
            IndexList(m_numBaseIB + i * 6 + 4) = m_numBaseVB + i * 4 + 2
            IndexList(m_numBaseIB + i * 6 + 5) = m_numBaseVB + i * 4 + 3
        Next i
    '-----------------------------------------------------------------------------------
        
    'Update the counters:
    m_numBaseVB = m_numBaseVB + 16
    m_numBaseIB = m_numBaseIB + 24
    
     'The following is valid only for left-hand coordinate sytem:
    For i = 0 To m_numBaseVB - 1
        VertexList(i).XYZ.z = -VertexList(i).XYZ.z
        VertexList(i).Normal.z = -VertexList(i).Normal.z
    Next i
    
    '-----------------------------
    ' Create the vertex buffer:
    '-----------------------------
    m_sizeVertex = Len(VertexList(0))
    Set m_BaseVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numBaseVB, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_BaseVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_BaseVB, 0, m_sizeVertex * m_numBaseVB, 0, VertexList(0)
    
    '--------------------------
    'Create the index buffer:
    '--------------------------
    Set m_BaseIB = m_D3DDevice8.CreateIndexBuffer(2 * m_numBaseIB, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    ' If the buffer was not created, exit function:
    If m_BaseVB Is Nothing Then Exit Function
    ' Fill the index buffer from our array
    D3DIndexBuffer8SetData m_BaseIB, 0, 2 * m_numBaseIB, 0, IndexList(0)
    
    
    CreateTableBase = True
End Function

Private Function CreateTableBorder() As Boolean
    Dim Angle As Single
    ReDim VertexList(83)
    ReDim IndexList(227)
    
    Angle = 2 * g_PI / 12
    
    'Initial counter setting:
    m_numBorderVB = 0: m_numBorderIB = 0
    
    'Material setting:
    With m_matBorder
        .Ambient = MakeD3DCOLORVALUE(1, 1, 1, 0)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(0.8, 0.8, 0.8, 0)
    End With
    
        ' Top level (0 to 20):
        ' We will create only one quarter of the table's upper part and then using symetry transformations
        ' we will build the rest.
        VertexList(0).XYZ = vec3(0, m_BoundaryHght, m_MinZ)
        VertexList(1).XYZ = vec3(m_Pockets(1).x - m_PocketRadius, m_BoundaryHght, m_MinZ)
        VertexList(2).XYZ = vec3(m_MaxX, m_BoundaryHght, m_MinZ)
        VertexList(3).XYZ = vec3(m_MaxX, m_BoundaryHght, m_Pockets(1).z + m_PocketRadius)
        VertexList(4).XYZ = vec3(m_MaxX, m_BoundaryHght, m_Pockets(2).z - m_PocketRadius)
        VertexList(5).XYZ = vec3(m_MaxX, m_BoundaryHght, m_Pockets(2).z)
        'A quarter of pocket #2:
        For i = 0 To 3
            VertexList(6 + i).XYZ = vec3(m_Pockets(2).x + m_PocketRadius * Sin(i * Angle + g_PI / 2), _
                                                          m_BoundaryHght, _
                                                          m_Pockets(2).z + m_PocketRadius * Cos(i * Angle + g_PI / 2))
        Next i
        'Three quarters of pocket #1:
        For i = 0 To 9
            VertexList(10 + i).XYZ = vec3(m_Pockets(1).x + m_PocketRadius * Sin(i * Angle), _
                                                           m_BoundaryHght, _
                                                           m_Pockets(1).z + m_PocketRadius * Cos(i * Angle))
        Next i
        VertexList(20).XYZ = vec3(0, m_BoundaryHght, m_MinZ + m_BoundaryWdth)
        
        'Normals (all facing up):
        For i = 0 To 20
            VertexList(i).Normal = vec3(0, 1, 0)
        Next i
    
        ' ------------ Indices: --------------------------
        ' triangle 1:
        IndexList(0) = 0
        IndexList(1) = 19
        IndexList(2) = 20
        ' triangle 2:
        IndexList(3) = 0
        IndexList(4) = 1
        IndexList(5) = 19
        ' triangle 3:
        IndexList(6) = 1
        IndexList(7) = 18
        IndexList(8) = 19
        ' triangle 4:
        IndexList(9) = 1
        IndexList(10) = 17
        IndexList(11) = 18
        ' triangle 5:
        IndexList(12) = 1
        IndexList(13) = 16
        IndexList(14) = 17
        ' triangle 6:
        IndexList(15) = 1
        IndexList(16) = 2
        IndexList(17) = 16
        ' triangle 7:
        IndexList(18) = 2
        IndexList(19) = 15
        IndexList(20) = 16
        ' triangle 8:
        IndexList(21) = 2
        IndexList(22) = 14
        IndexList(23) = 15
        ' triangle 9:
        IndexList(24) = 2
        IndexList(25) = 13
        IndexList(26) = 14
        ' triangle 10:
        IndexList(27) = 2
        IndexList(28) = 3
        IndexList(29) = 13
        ' triangle 11:
        IndexList(30) = 3
        IndexList(31) = 12
        IndexList(32) = 13
        ' triangle 12:
        IndexList(33) = 3
        IndexList(34) = 11
        IndexList(35) = 12
        ' triangle 13:
        IndexList(36) = 3
        IndexList(37) = 10
        IndexList(38) = 11
        ' triangle 14:
        IndexList(39) = 3
        IndexList(40) = 4
        IndexList(41) = 10
        ' triangle 15:
        IndexList(42) = 4
        IndexList(43) = 9
        IndexList(44) = 10
        ' triangle 16:
        IndexList(45) = 4
        IndexList(46) = 8
        IndexList(47) = 9
        ' triangle 17:
        IndexList(48) = 4
        IndexList(49) = 7
        IndexList(50) = 8
        ' triangle 18:
        IndexList(51) = 4
        IndexList(52) = 6
        IndexList(53) = 7
        ' triangle 19:
        IndexList(54) = 4
        IndexList(55) = 5
        IndexList(56) = 6
        '-------------------------------------------------
        
    ' Using symetry:
    ' First, transform through the ZY plane:
    For i = 0 To 20
        VertexList(i + 21).XYZ = VertexList(i).XYZ
        VertexList(i + 21).Normal = VertexList(i).Normal
        VertexList(i + 21).XYZ.x = -VertexList(i).XYZ.x
    Next i
    For i = 0 To 56 Step 3
        IndexList(i + 57 + 0) = IndexList(i + 0) + 21
        IndexList(i + 57 + 1) = IndexList(i + 2) + 21
        IndexList(i + 57 + 2) = IndexList(i + 1) + 21
    Next i
    ' Now lets, flip this whole part around the XY plane:
    For i = 0 To 0 + 41
        VertexList(i + 42).XYZ = VertexList(i).XYZ
        VertexList(i + 42).Normal = VertexList(i).Normal
        VertexList(i + 42).XYZ.z = -VertexList(i).XYZ.z
    Next i
    For i = 0 To 113 Step 3
        IndexList(i + 114 + 0) = IndexList(i + 0) + 42
        IndexList(i + 114 + 1) = IndexList(i + 2) + 42
        IndexList(i + 114 + 2) = IndexList(i + 1) + 42
    Next i
    
    'Update the counters:
    m_numBorderVB = 4 * 21
    m_numBorderIB = 4 * 57
    
    
    '------------------------------
    ' Texture coordinates:
    '------------------------------
    ' NOTE: Although the Flexible Vertex Format is the same for all four vertex buffers
    ' only the m_BorderVB and m_BorderVB actually use the Tex1 texture coordinates from TABLEVERTEX.
    
    For i = 0 To m_numBorderVB - 1
        With VertexList(i)
            .Tex1.x = (.XYZ.x - m_MinX) / (m_MaxX - m_MinX) * 10
            .Tex1.y = (.XYZ.z - m_MinZ) / (m_MaxZ - m_MinZ) * 20
        End With
    Next i
               
    'The following is valid only for left-hand coordinate sytem:
    For i = 0 To m_numBorderVB - 1
        VertexList(i).XYZ.z = -VertexList(i).XYZ.z
        VertexList(i).Normal.z = -VertexList(i).Normal.z
    Next i
               
    '-----------------------------
    ' Create the vertex buffer:
    '-----------------------------
    m_sizeVertex = Len(VertexList(0))
    Set m_BorderVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numBorderVB, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_BorderVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_BorderVB, 0, m_sizeVertex * m_numBorderVB, 0, VertexList(0)
    
    '--------------------------
    'Create the index buffer:
    '--------------------------
    Set m_BorderIB = m_D3DDevice8.CreateIndexBuffer(2 * m_numBorderIB, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    ' If the buffer was not created, exit function:
    If m_BorderVB Is Nothing Then Exit Function
    ' Fill the index buffer from our array
    D3DIndexBuffer8SetData m_BorderIB, 0, 2 * m_numBorderIB, 0, IndexList(0)
    
    '---------------------------
    ' Create texture from file:
    '---------------------------
    Set m_BorderTexture = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, g_AppPath & "\TexWood.bmp")
    
    
    CreateTableBorder = True
End Function

Private Function CreateTableTop() As Boolean
    ReDim VertexList(3)
    ReDim IndexList(5)
       
    'Initial counter setting:
    m_numTopVB = 0: m_numTopIB = 0
    
    'Material setting:
    With m_matTop
        .Ambient = MakeD3DCOLORVALUE(1, 1, 1, 0)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(0, 0, 0, 0)
    End With
   
   
    'The "game area" (1 face, 4 vertices):
        'Vertices:
        VertexList(0).XYZ = vec3(m_MinX + 0.02, 0, m_MinZ + 0.02)
        VertexList(1).XYZ = vec3(m_MaxX - 0.02, 0, m_MinZ + 0.02)
        VertexList(2).XYZ = vec3(m_MaxX - 0.02, 0, m_MaxZ - 0.02)
        VertexList(3).XYZ = vec3(m_MinX + 0.02, 0, m_MaxZ - 0.02)
        For i = 0 To 3
            VertexList(i).Normal = vec3(0, 1, 0)
        Next i
        
        'Indexes:
        ' triangle 1:
        IndexList(0) = 0
        IndexList(1) = 1
        IndexList(2) = 2
        ' triangle 2:
        IndexList(3) = 0
        IndexList(4) = 2
        IndexList(5) = 3
    
    'Update the counters:
    m_numTopVB = 4
    m_numTopIB = 6
    
    '------------------------------
    ' Texture coordinates:
    '------------------------------
    ' NOTE: Although the Flexible Vertex Format is the same for all four vertex buffers
    ' only the m_TopVB and m_TopVB actually use the Tex1 texture coordinates from TABLEVERTEX.
    
    For i = 0 To m_numTopVB - 1
        With VertexList(i)
            .Tex1.x = (.XYZ.x - m_MinX) / (m_MaxX - m_MinX) * 10
            .Tex1.y = (.XYZ.z - m_MinZ) / (m_MaxZ - m_MinZ) * 20
        End With
    Next i
    
    'The following is valid only for left-hand coordinate sytem:
    For i = 0 To m_numTopVB - 1
        VertexList(i).XYZ.z = -VertexList(i).XYZ.z
        VertexList(i).Tex1.y = -VertexList(i).Tex1.y
    Next i
                  
    '-----------------------------
    ' Create the vertex buffer:
    '-----------------------------
    m_sizeVertex = Len(VertexList(0))
    Set m_TopVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numTopVB, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_TopVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_TopVB, 0, m_sizeVertex * m_numTopVB, 0, VertexList(0)
    
    '--------------------------
    'Create the index buffer:
    '--------------------------
    Set m_TopIB = m_D3DDevice8.CreateIndexBuffer(2 * m_numTopIB, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    ' If the buffer was not created, exit function:
    If m_TopVB Is Nothing Then Exit Function
    ' Fill the index buffer from our array
    D3DIndexBuffer8SetData m_TopIB, 0, 2 * m_numTopIB, 0, IndexList(0)
    
    '---------------------------
    ' Create texture from file:
    '---------------------------
    Set m_TopTexture = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, g_AppPath & "\TableTop.bmp")
    
    
    CreateTableTop = True
End Function

Private Function CreateTableSides() As Boolean
    Dim Angle As Single
    ReDim VertexList(143)
    ReDim IndexList(335)
    
    Angle = 2 * g_PI / 12
    
    'Initial counter setting:
    m_numSideVB = 0: m_numSideIB = 0
    
    'Material setting:
    With m_matSides
        .Ambient = MakeD3DCOLORVALUE(0.4, 0.4, 0.4, 0)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(0, 0, 0, 0)
    End With
    
    ' ----------------- Vertices: -----------------------------------------------------------------
    ' The top (0 to 17):
    ' A quarter of pocket #2:
    For i = 0 To 3
        VertexList(i).XYZ = vec3(m_Pockets(2).x + m_PocketRadius * Sin(i * Angle + g_PI / 2), _
                                 m_BoundaryHght, _
                                 m_Pockets(2).z + m_PocketRadius * Cos(i * Angle + g_PI / 2))
        VertexList(i).Normal = vec3(-Sin(i * Angle + g_PI / 2), 0, -Cos(i * Angle + g_PI / 2))
    Next i
    VertexList(4).XYZ = vec3(m_Pockets(2).x, m_BoundaryHght, m_Pockets(2).z - m_PocketRadius)
    VertexList(4).Normal = vec3(-1, 0, 0)
    VertexList(5).XYZ = vec3(m_Pockets(1).x, m_BoundaryHght, m_Pockets(1).z + m_PocketRadius)
    VertexList(5).Normal = vec3(-1, 0, 0)
    'Three quarters of pocket #1:
    For i = 0 To 9
        VertexList(6 + i).XYZ = vec3(m_Pockets(1).x + m_PocketRadius * Sin(i * Angle), _
                                     m_BoundaryHght, _
                                     m_Pockets(1).z + m_PocketRadius * Cos(i * Angle))
        VertexList(6 + i).Normal = vec3(-Sin(i * Angle), 0, -Cos(i * Angle))
    Next i
    VertexList(16).XYZ = vec3(m_Pockets(1).x - m_PocketRadius, m_BoundaryHght, m_Pockets(1).z)
    VertexList(16).Normal = vec3(0, 0, 1)
    VertexList(17).XYZ = vec3(0, m_BoundaryHght, m_Pockets(1).z)
    VertexList(17).Normal = vec3(0, 0, 1)
    
    ' The bottom (18 to 25):
    For i = 0 To 17
        VertexList(18 + i).XYZ = VertexList(m_numSideVB + i).XYZ
        VertexList(18 + i).Normal = VertexList(m_numSideVB + i).Normal
        VertexList(18 + i).XYZ.y = -0.1
    Next i
    
    
    ' ------------ Indices: ------------------------------------------------
    For i = 0 To 3
        ' triangle 2i:
        IndexList(i * 6 + 0) = m_numSideVB + 18 + i + 1
        IndexList(i * 6 + 1) = m_numSideVB + i + 1
        IndexList(i * 6 + 2) = m_numSideVB + i
        ' triangle 2i+1:
        IndexList(i * 6 + 3) = m_numSideVB + 18 + i + 1
        IndexList(i * 6 + 4) = m_numSideVB + i
        IndexList(i * 6 + 5) = m_numSideVB + 18 + i
    Next i
                
    ' triangle 7:
    IndexList(18) = m_numSideVB + 5 + 18
    IndexList(19) = m_numSideVB + 5
    IndexList(20) = m_numSideVB + 4
    ' triangle 8:
    IndexList(21) = m_numSideVB + 5 + 18
    IndexList(22) = m_numSideVB + 4
    IndexList(23) = m_numSideVB + 4 + 18
    
    For i = 0 To 8
        ' triangle 9+2i:
        IndexList(24 + i * 6 + 0) = m_numSideVB + 6 + 18 + i + 1
        IndexList(24 + i * 6 + 1) = m_numSideVB + 6 + i + 1
        IndexList(24 + i * 6 + 2) = m_numSideVB + 6 + i
        ' triangle 9+2i+1:
        IndexList(24 + i * 6 + 3) = m_numSideVB + 6 + 18 + i + 1
        IndexList(24 + i * 6 + 4) = m_numSideVB + 6 + i
        IndexList(24 + i * 6 + 5) = m_numSideVB + 6 + 18 + i
    Next i
    
    ' triangle 27:
    IndexList(78) = m_numSideVB + 16 + 18 + 1
    IndexList(79) = m_numSideVB + 16 + 1
    IndexList(80) = m_numSideVB + 16
    ' triangle 28:
    IndexList(81) = m_numSideVB + 16 + 18 + 1
    IndexList(82) = m_numSideVB + 16
    IndexList(83) = m_numSideVB + 16 + 18
    '--------------------------------------------------------------------------
    
    ' Again we use symetry:
    ' First, transform through the ZY plane:
    For i = 0 To 35
        VertexList(i + 36).XYZ = VertexList(i).XYZ
        VertexList(i + 36).Normal = VertexList(i).Normal
        VertexList(i + 36).XYZ.x = -VertexList(i).XYZ.x
        VertexList(i + 36).Normal.x = -VertexList(i).Normal.x       'NOTE: this time we transform the normal vector as well.
    Next i
    For i = 0 To 83 Step 3
        IndexList(i + 84 + 0) = IndexList(i + 0) + 36
        IndexList(i + 84 + 1) = IndexList(i + 2) + 36
        IndexList(i + 84 + 2) = IndexList(i + 1) + 36
    Next i
    ' Now lets, flip this whole part around the XY plane:
    For i = 0 To 71
        VertexList(i + 72).XYZ = VertexList(i).XYZ
        VertexList(i + 72).Normal = VertexList(i).Normal
        VertexList(i + 72).XYZ.z = -VertexList(i).XYZ.z
        VertexList(i + 72).Normal.z = -VertexList(i).Normal.z
    Next i
    For i = 0 To 167 Step 3
        IndexList(i + 168 + 0) = IndexList(i + 0) + 72
        IndexList(i + 168 + 1) = IndexList(i + 2) + 72
        IndexList(i + 168 + 2) = IndexList(i + 1) + 72
    Next i
    
    'Update the counters:
    m_numSideVB = 4 * 36
    m_numSideIB = 4 * 84
   
    'The following is valid only for left-hand coordinate sytem:
    For i = 0 To m_numSideVB - 1
        VertexList(i).XYZ.z = -VertexList(i).XYZ.z
    Next i

    '-----------------------------
    ' Create the vertex buffer:
    '-----------------------------
    m_sizeVertex = Len(VertexList(0))
    Set m_SideVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numSideVB, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_SideVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_SideVB, 0, m_sizeVertex * m_numSideVB, 0, VertexList(0)
    
    '--------------------------
    'Create the index buffer:
    '--------------------------
    Set m_SideIB = m_D3DDevice8.CreateIndexBuffer(2 * m_numSideIB, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    ' If the buffer was not created, exit function:
    If m_SideVB Is Nothing Then Exit Function
    ' Fill the index buffer from our array
    D3DIndexBuffer8SetData m_SideIB, 0, 2 * m_numSideIB, 0, IndexList(0)
    
    
    CreateTableSides = True
End Function

Private Function CreateTableBumpers() As Boolean
    Dim XYZList() As D3DVECTOR
    Dim ShadowVertices(6 * 14 - 1) As SHADOWVERTEX
    
    ReDim VertexList(6 * 14 - 1)
    ReDim IndexList(6 * 18 - 1)
       
    'Initial counter setting:
    m_numBumpersVB = 0: m_numBumpersIB = 0
    
    'Material setting:
    With m_matBumpers
        .Ambient = MakeD3DCOLORVALUE(0.2, 0.6, 0.2, 0)
        .diffuse = .Ambient
        .power = 10
        .specular = MakeD3DCOLORVALUE(0, 0, 0, 0)
    End With
     
    'Auxiliary vertex list:
    ReDim XYZList(1 To 6)
      
    For i = 1 To 6
        If i > 1 Then j = i - 1 Else j = 6
        D3DXVec3Subtract v1, m_Pockets(i), m_Pockets(j)
        D3DXVec3Normalize v1, v1
        D3DXVec3Cross v2, v1, vec3(0, 1, 0)
        D3DXVec3Scale v2, v2, 0.06
        D3DXVec3Scale v3, v1, m_PocketRadius
        D3DXVec3Scale v4, v3, 2.5
        
        ' Fill the auxiliary vertex list:
        ' j'th pocket:
        D3DXVec3Add XYZList(1), m_Pockets(j), v3: XYZList(1).y = 0
        D3DXVec3Add XYZList(2), m_Pockets(j), v3: XYZList(2).y = m_BoundaryHght - 0.01
        D3DXVec3Add XYZList(3), m_Pockets(j), v4
        D3DXVec3Add XYZList(3), XYZList(3), v2: XYZList(3).y = m_BoundaryHght - 0.02
        ' i'th pocket
        D3DXVec3Subtract XYZList(4), m_Pockets(i), v3: XYZList(4).y = 0
        D3DXVec3Subtract XYZList(5), m_Pockets(i), v3: XYZList(5).y = m_BoundaryHght - 0.01
        D3DXVec3Subtract XYZList(6), m_Pockets(i), v4
        D3DXVec3Add XYZList(6), XYZList(6), v2: XYZList(6).y = m_BoundaryHght - 0.02
        
                               
        ' ------------------ Vertices: --------------------------------
        VertexList(m_numBumpersVB + 0).XYZ = XYZList(1)
        VertexList(m_numBumpersVB + 1).XYZ = XYZList(2)
        VertexList(m_numBumpersVB + 2).XYZ = XYZList(3)
        
        VertexList(m_numBumpersVB + 3).XYZ = XYZList(2)
        VertexList(m_numBumpersVB + 4).XYZ = XYZList(5)
        VertexList(m_numBumpersVB + 5).XYZ = XYZList(6)
        VertexList(m_numBumpersVB + 6).XYZ = XYZList(3)
        
        VertexList(m_numBumpersVB + 7).XYZ = XYZList(4)
        VertexList(m_numBumpersVB + 8).XYZ = XYZList(6)
        VertexList(m_numBumpersVB + 9).XYZ = XYZList(5)
        
        VertexList(m_numBumpersVB + 10).XYZ = XYZList(1)
        VertexList(m_numBumpersVB + 11).XYZ = XYZList(3)
        VertexList(m_numBumpersVB + 12).XYZ = XYZList(6)
        VertexList(m_numBumpersVB + 13).XYZ = XYZList(4)
        
        ' ------------------ Indices: -----------------------------------
        ' Triangle 1
        IndexList(m_numBumpersIB + 0) = m_numBumpersVB + 0
        IndexList(m_numBumpersIB + 1) = m_numBumpersVB + 1
        IndexList(m_numBumpersIB + 2) = m_numBumpersVB + 2
            
        ' Triangle 2
        IndexList(m_numBumpersIB + 3) = m_numBumpersVB + 3
        IndexList(m_numBumpersIB + 4) = m_numBumpersVB + 4
        IndexList(m_numBumpersIB + 5) = m_numBumpersVB + 5
            
        ' Triangle 3
        IndexList(m_numBumpersIB + 6) = m_numBumpersVB + 3
        IndexList(m_numBumpersIB + 7) = m_numBumpersVB + 5
        IndexList(m_numBumpersIB + 8) = m_numBumpersVB + 6
            
        ' Triangle 4
        IndexList(m_numBumpersIB + 9) = m_numBumpersVB + 7
        IndexList(m_numBumpersIB + 10) = m_numBumpersVB + 8
        IndexList(m_numBumpersIB + 11) = m_numBumpersVB + 9
            
        ' Triangle 5
        IndexList(m_numBumpersIB + 12) = m_numBumpersVB + 10
        IndexList(m_numBumpersIB + 13) = m_numBumpersVB + 11
        IndexList(m_numBumpersIB + 14) = m_numBumpersVB + 12
            
        ' Triangle 6
        IndexList(m_numBumpersIB + 15) = m_numBumpersVB + 10
        IndexList(m_numBumpersIB + 16) = m_numBumpersVB + 12
        IndexList(m_numBumpersIB + 17) = m_numBumpersVB + 13
                                        
        m_numBumpersVB = m_numBumpersVB + 14
        m_numBumpersIB = m_numBumpersIB + 18
    Next i
                     
    ' ------------------ Normals: -----------------------------------------------------------------------
    For i = 0 To m_numBumpersIB - 1 Step 3
        D3DXVec3Subtract v1, VertexList(IndexList(i + 1)).XYZ, VertexList(IndexList(i + 2)).XYZ
        D3DXVec3Subtract v2, VertexList(IndexList(i)).XYZ, VertexList(IndexList(i + 1)).XYZ
        D3DXVec3Cross v3, v1, v2
        D3DXVec3Normalize v3, v3
        VertexList(IndexList(i + 0)).Normal = v3
        VertexList(IndexList(i + 1)).Normal = v3
        VertexList(IndexList(i + 2)).Normal = v3
    Next i
    ' ---------------------------------------------------------------------------------------------------
    
    'The following is valid only for left-hand coordinate sytem:
    For i = 0 To m_numBumpersVB - 1
        VertexList(i).XYZ.z = -VertexList(i).XYZ.z
        VertexList(i).Normal.z = -VertexList(i).Normal.z
    Next i
    '---------------------------------------------------------------
    
    'Shadow vertices:
    For i = 0 To m_numBumpersVB - 1
        ShadowVertices(i).XYZ = VertexList(i).XYZ
        ShadowVertices(i).Color = D3DColorMake(0, 0, 0, 0.3)
    Next i
    
    '-----------------------------
    ' Create the vertex buffers:
    '-----------------------------
    Set m_BumpersVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numBumpersVB, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_BumpersVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_BumpersVB, 0, m_sizeVertex * m_numBumpersVB, 0, VertexList(0)
    
    m_sizeShadowVertex = Len(ShadowVertices(0))
    Set m_BumpShadowVB = m_D3DDevice8.CreateVertexBuffer(m_sizeVertex * m_numBumpersVB, 0, D3DFVF_SHADOWVERTEX, D3DPOOL_DEFAULT)
    ' If the vertex buffer was not created, exit function:
    If m_BumpShadowVB Is Nothing Then Exit Function
    ' Fill the vertex buffer from our array
    D3DVertexBuffer8SetData m_BumpShadowVB, 0, m_sizeShadowVertex * m_numBumpersVB, 0, ShadowVertices(0)
    
    
    '--------------------------
    'Create the index buffer:
    '--------------------------
    Set m_BumpersIB = m_D3DDevice8.CreateIndexBuffer(2 * m_numBumpersIB, 0, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
    ' If the buffer was not created, exit function:
    If m_BumpersVB Is Nothing Then Exit Function
    ' Fill the index buffer from our array
    D3DIndexBuffer8SetData m_BumpersIB, 0, 2 * m_numBumpersIB, 0, IndexList(0)
         
    ' Finally, update the bumper array, which will be used for collision detection:
    ' The first index indicates the bumper, while the second index refers to
    ' a vertex of this bumper.
    ' NOTE: For collision detection we use only four (out of six) bumper vertices.
    For i = 1 To 6
        m_Bumpers(i, 1) = VertexList((i - 1) * 14 + 4).XYZ
        m_Bumpers(i, 2) = VertexList((i - 1) * 14 + 5).XYZ
        m_Bumpers(i, 3) = VertexList((i - 1) * 14 + 6).XYZ
        m_Bumpers(i, 4) = VertexList((i - 1) * 14 + 3).XYZ
    Next i
         
    Erase XYZList, ShadowVertices
    
    CreateTableBumpers = True
End Function

Function CreatePocket() As Boolean
    ReDim VertexList(3)
    
    ' Vertices are numbered according to this
    ' diagram:
    '       0 ----- 1     z
    '       |       |     |
    '       |   +   |     |
    '       |       |     |
    '       2 ----- 3     +------ x
       
    With VertexList(0)
        .XYZ = vec3(-m_PocketRadius, 0, m_PocketRadius)
        .Normal = vec3(0, 1, 0)
        .Tex1.x = 0: .Tex1.y = 0
    End With
    
    With VertexList(1)
        .XYZ = vec3(m_PocketRadius, 0, m_PocketRadius)
        .Normal = vec3(0, 1, 0)
        .Tex1.x = 1: .Tex1.y = 0
    End With
  
    With VertexList(2)
        .XYZ = vec3(-m_PocketRadius, 0, -m_PocketRadius)
        .Normal = vec3(0, 1, 0)
        .Tex1.x = 0: .Tex1.y = 1
    End With
      
    With VertexList(3)
        .XYZ = vec3(m_PocketRadius, 0, -m_PocketRadius)
        .Normal = vec3(0, 1, 0)
        .Tex1.x = 1: .Tex1.y = 1
    End With
    
    '------------------------------------
    ' Create and fill the vertex buffer
    '------------------------------------
    m_sizeVertex = Len(VertexList(0))
    Set m_PocketVB = m_D3DDevice8.CreateVertexBuffer(4 * m_sizeVertex, 0, D3DFVF_TABLEVERTEX, D3DPOOL_DEFAULT)
    If m_PocketVB Is Nothing Then Exit Function
    D3DVertexBuffer8SetData m_PocketVB, 0, 4 * m_sizeVertex, 0, VertexList(0)
    
    '---------------------------
    ' Create the pocket texture
    '---------------------------
    Set m_PocketTexture = m_D3DX8.CreateTextureFromFile(m_D3DDevice8, g_AppPath & "\BlackHole.dds")
        
    CreatePocket = True
End Function
