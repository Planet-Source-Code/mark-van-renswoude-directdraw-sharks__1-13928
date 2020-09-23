VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lSharkX As Long
Private lSharkY As Long

Private bSharkLeft As Boolean
Private bSharkTurning As Boolean
Private bSharkTurnFrameLeft As Boolean

Private lFrameTurn As Long
Private lFrame As Long
Private lFrameTop As Long

Private bVMoving As Boolean
Private lVMoving As Long
Private lVTotal As Long
Public Sub DoMovement()
    Static lTimer As Long
    Static lTimerMove As Long
    Dim lDirect As Long
    
    If mdX.TickCount() - lTimerMove >= 25 Then
        lTimerMove = mdX.TickCount()
        
        '// Check if we're turning
        If bSharkTurning Then
            If mdX.TickCount() - lTimer >= 50 Then
                lTimer = mdX.TickCount()
                '// Next frame
                lFrameTurn = lFrameTurn + 1
                
                '// Check for last frame
                If lFrameTurn = 7 Then
                    bSharkTurning = False
                    bSharkLeft = Not bSharkLeft
                    lFrame = 0
                    bSharkTurnFrameLeft = False
                    
                    If bSharkLeft Then
                        lSharkX = lSharkX + 6
                    Else
                        lSharkX = lSharkX - 6
                    End If
                End If
            End If
        Else
            '// Decide whether or not to Turn
            If CInt(Rnd(1) * 100) = 50 Then
                bSharkTurning = True
                lFrameTurn = 0
            Else
                If bSharkLeft Then
                    '// Shark goes left, check for edge
                    If lSharkX <= 25 Then
                        bSharkTurning = True
                        lFrameTurn = 0
                    End If
                Else
                    '// Shark goes right, check for edge
                    If lSharkX >= lScreenWidth - 75 Then
                        bSharkTurning = True
                        lFrameTurn = 0
                    End If
                End If
            End If
        End If
        
        If mdX.TickCount() - lTimer >= 100 Then
            lTimer = mdX.TickCount()
            
            If bSharkTurning = False Then
                '// Change Frame
                If bSharkTurnFrameLeft Then
                    '// One Frame back
                    lFrame = lFrame - 1
                    
                    '// Check for first frame
                    If lFrame = 0 Then
                        bSharkTurnFrameLeft = False
                    End If
                Else
                    '// One Frame forward
                    lFrame = lFrame + 1
                    
                    '// Check for last frame
                    If lFrame = 4 Then
                        bSharkTurnFrameLeft = True
                    End If
                End If
            End If
            
            If bVMoving Then
                '// Move
                lVTotal = lVTotal + 1
                If lVTotal = 10 Then bVMoving = False
                
                lSharkY = lSharkY + lVMoving
            Else
                '// Decide Movements
                lDirect = CLng(Rnd(1) * 2)
                If lDirect = 1 And lSharkY >= 50 Then
                    lVMoving = -1
                    lVTotal = 0
                    bVMoving = True
                End If
                
                If lDirect = 2 And lSharkY <= lScreenHeight - 50 Then
                    lVMoving = 1
                    lVTotal = 0
                    bVMoving = True
                End If
            End If
        End If
        
        If bSharkTurning = False Then
            '// Move Shark
            If bSharkLeft Then
                lSharkX = lSharkX - 3
            Else
                lSharkX = lSharkX + 3
            End If
        End If
    End If
    
    '// Set Frame Top
    If bSharkLeft Then
        lFrameTop = 0
    Else
        lFrameTop = 37
    End If
End Sub


Public Sub Draw()
    Dim cKey As DDCOLORKEY
    Dim rSwim As RECT
    Dim rTurn As RECT
    Dim lX As Long
    Dim lY As Long
    
    '// Set Transparent Color
    cKey.high = mdX.CreateColorRGB(255, 0, 255)
    cKey.low = cKey.high
    msSwim.SetColorKey DDCKEY_SRCBLT, cKey
    msTurn.SetColorKey DDCKEY_SRCBLT, cKey
    
    '// Draw Shark onto backbuffer
    If bSharkTurning Then
        '// Set Rectangle
        With rTurn
            .Top = lFrameTop
            .Left = lFrameTurn * 52
            .Bottom = .Top + 37
            .Right = .Left + 52
            
            '// Fix Bottom
            If lSharkY + (.Bottom - .Top) > lScreenHeight Then
                .Bottom = .Top + (lScreenHeight - lSharkY)
            End If
            
            '// Fix Right
            If lSharkX + (.Right - .Left) > lScreenWidth Then
                .Right = .Left + (lScreenWidth - lSharkY)
            End If
            
            '// Fix Top
            If lSharkY < 0 Then
                .Top = .Top - lSharkY
                lY = 0
            Else
                lY = lSharkY
            End If
            
            '// Fix Left
            If lSharkX < 0 Then
                .Left = .Left - lSharkX
                lX = 0
            Else
                lX = lSharkX
            End If
        End With
        
        msBack.BltFast lX, lY, msTurn, rTurn, DDBLTFAST_SRCCOLORKEY
    Else
        '// Set Rectangle
        With rSwim
            .Top = lFrameTop
            .Left = lFrame * 63
            .Bottom = .Top + 37
            .Right = .Left + 63
        
            '// Fix Bottom
            If lSharkY + (.Bottom - .Top) > lScreenHeight Then
                .Bottom = .Top + (lScreenHeight - lSharkY)
            End If
            
            '// Fix Right
            If lSharkX + (.Right - .Left) > lScreenWidth Then
                .Right = .Left + (lScreenWidth - lSharkY)
            End If
            
            '// Fix Top
            If lSharkY < 0 Then
                .Top = .Top - lSharkY
                lY = 0
            Else
                lY = lSharkY
            End If
            
            '// Fix Left
            If lSharkX < 0 Then
                .Left = .Left - lSharkX
                lX = 0
            Else
                lX = lSharkX
            End If
        End With
        
        msBack.BltFast lX, lY, msSwim, rSwim, DDBLTFAST_SRCCOLORKEY
    End If
End Sub


Public Sub Initialize()
    '// Set Random Position
    Randomize Timer
    lSharkY = CLng(Rnd(1) * (lScreenHeight - 100) + 25)
    lSharkX = CLng(Rnd(1) * (lScreenWidth - 100) + 25)
    bSharkLeft = IIf(CInt(Rnd(1)) = 1, True, False)
End Sub

