Attribute VB_Name = "modDXFunctions"
Option Explicit

Public mdX As New DirectX7
Public mdD As DirectDraw7

Public Sub Initialize(Form As Form, ScreenWidth As Long, ScreenHeight As Long, ScreenDepth As Long)
    On Local Error Resume Next
    
    Dim ddsdMain As DDSURFACEDESC2
    Dim ddsdFlip As DDSURFACEDESC2
    
    '// Create DirectX Components
    Set mdD = mdX.DirectDrawCreate("")
    
    '// Set Display Resolution and Cooperative Levels
    mdD.SetDisplayMode ScreenWidth, ScreenHeight, ScreenDepth, 0, DDSDM_DEFAULT
    mdD.SetCooperativeLevel Form.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE

    '// Create Flipping Structure
    ddsdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdMain.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
    ddsdMain.lBackBufferCount = 1
    
    '// Create Primary Surface
    Set msFront = mdD.CreateSurface(ddsdMain)
    
    '// Create Backbuffer
    ddsdFlip.ddsCaps.lCaps = DDSCAPS_BACKBUFFER
    Set msBack = msFront.GetAttachedSurface(ddsdFlip.ddsCaps)
End Sub

Public Sub Terminate()
    On Local Error Resume Next
    
    '// Restore Settings
    mdD.RestoreDisplayMode
    mdD.SetCooperativeLevel 0, DDSCL_NORMAL
    
    '// Kill DirectX
    Set mdD = Nothing
End Sub

Public Function LostSurfaces() As Boolean
    '// Check if we should reload our bitmaps or not
    LostSurfaces = False
    Do Until ExclusiveMode
        DoEvents
        LostSurfaces = True
    Loop
    
    '// Lost bitmaps, restore the surfaces and return 'true'
    DoEvents
    If LostSurfaces Then
        mdD.RestoreAllSurfaces
    End If
End Function



Public Function ExclusiveMode() As Boolean
    Dim lTestExMode As Long
    
    '// Test if we're still in exclusive mode
    lTestExMode = mdD.TestCooperativeLevel
    
    If (lTestExMode = DD_OK) Then
        ExclusiveMode = True
    Else
        ExclusiveMode = False
    End If
End Function


Public Sub LoadSurface(Surface As DirectDrawSurface7, Filename As String, Width As Long, Height As Long)
    Dim ddGeneric As DDSURFACEDESC2
    
    '// Set up surface description
    ddGeneric.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddGeneric.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddGeneric.lWidth = Width
    ddGeneric.lHeight = Height
    
    '// Load Surface
    Set Surface = mdD.CreateSurfaceFromFile(Filename, ddGeneric)
End Sub
