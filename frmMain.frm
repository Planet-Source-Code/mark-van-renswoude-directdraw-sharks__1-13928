VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sharks"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cShark() As New clsShark
Dim cBubble() As New clsBubble
Dim cDiddie() As New clsDiddie
Dim bRunning As Boolean

Dim lSharkCount As Long
Dim lBubbleCount As Long
Dim lDiddieCount As Long
Sub LoadSurfaces()
    '// Load Shark Surfaces
    LoadSurface msSwim, App.Path & "\images\shark_swim.bmp", 315, 74
    LoadSurface msTurn, App.Path & "\images\shark_turn.bmp", 416, 74
    
    '// Load Bubble Surfaces
    LoadSurface msBubbleSmall, App.Path & "\images\bubble_small.bmp", 4, 6
    LoadSurface msBubbleLarge, App.Path & "\images\bubble_large.bmp", 8, 8
    
    '// Load Diddie Surfaces
    LoadSurface msDiddie, App.Path & "\images\diddie.bmp", 270, 78
End Sub

Sub MainLoop()
    Dim mrScreen As RECT
    Dim lK As Long
    
    With mrScreen
        .Top = 0
        .Left = 0
        .Bottom = lScreenHeight
        .Right = lScreenWidth
    End With
    
    bRunning = True
    
    Do While bRunning
        If LostSurfaces Then LoadSurfaces
        msBack.BltColorFill mrScreen, 0
        
        For lK = 0 To lSharkCount - 1
            '// Move Shark
            cShark(lK).DoMovement
            
            '// Draw Shark
            cShark(lK).Draw
        Next lK
        
        For lK = 0 To lBubbleCount - 1
            '// Move Bubble
            cBubble(lK).DoMovement
            
            '// Draw Bubble
            cBubble(lK).Draw
        Next lK
        
        For lK = 0 To lDiddieCount - 1
            '// Move Diddie
            cDiddie(lK).DoMovement
            
            '// Draw Diddie
            cDiddie(lK).Draw
        Next lK
        
        msFront.Flip Nothing, 0
        DoEvents
    Loop
    
    Terminate
    Unload frmMain
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    bRunning = False
End Sub

Private Sub Form_Load()
    Dim lK As Long
    
    '// These variables reflect the number of
    '// sprites to be used...
    lSharkCount = 15
    lBubbleCount = 50
    lDiddieCount = 1
    
    ReDim cShark(lSharkCount - 1)
    ReDim cBubble(lBubbleCount - 1)
    ReDim cDiddie(lDiddieCount - 1)
    
    '// Initialize DirectDraw
    lScreenWidth = 640
    lScreenHeight = 480
    lScreenDepth = 32
    Call modDXFunctions.Initialize(frmMain, lScreenWidth, lScreenHeight, lScreenDepth)
    
    '// Initialize Sharks
    For lK = 0 To lSharkCount - 1
        cShark(lK).Initialize
    Next lK
    
    '// Initialize Bubbles
    For lK = 0 To lBubbleCount - 1
        cBubble(lK).Initialize
    Next lK
    
    '// Initialize Diddies
    For lK = 0 To lDiddieCount - 1
        cDiddie(lK).Initialize
    Next lK
    
    '// Load Surfaces
    LoadSurfaces
    
    '// Start Running
    MainLoop
End Sub


