VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBubble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lBubbleX As Long
Private lBubbleY As Long

Private bBubbleLeft As Boolean
Private lBubbleHSpeed As Long
Private lBubbleVSpeed As Long

Private bBubbleLarge As Boolean
Public Sub DoMovement()
    Static lTimer As Long
    Static lHTimer As Long
    Randomize Timer
    
    If mdX.TickCount() - lTimer >= 10 Then
        lTimer = mdX.TickCount()
        
        '// Move Bubble Vertically
        lBubbleY = lBubbleY - lBubbleVSpeed
        
        '// Check for Edges
        If lBubbleY <= -CLng(Rnd(1) * 25) Then
            lBubbleY = lScreenHeight
            lBubbleX = CLng(Rnd(1) * (lScreenWidth - 100) + 25)
        End If
    End If
    
    If mdX.TickCount() - lHTimer >= 50 Then
        lHTimer = mdX.TickCount()
        
        '// Move Bubble Horizontally
        If bBubbleLeft Then
            lBubbleHSpeed = lBubbleHSpeed - 1
            If lBubbleHSpeed <= -2 Then bBubbleLeft = False
        Else
            lBubbleHSpeed = lBubbleHSpeed + 1
            If lBubbleHSpeed >= 2 Then bBubbleLeft = True
        End If
        
        lBubbleX = lBubbleX + lBubbleHSpeed
    End If
End Sub


Public Sub Draw()
    Dim cKey As DDCOLORKEY
    Dim rBubble As RECT
    Dim lX As Long
    Dim lY As Long
    
    '// Set Transparent Color
    cKey.high = mdX.CreateColorRGB(255, 0, 255)
    cKey.low = cKey.high
    msBubbleSmall.SetColorKey DDCKEY_SRCBLT, cKey
    msBubbleLarge.SetColorKey DDCKEY_SRCBLT, cKey
    
    If bBubbleLarge Then
        '// Set Rectangle
        With rBubble
            .Top = 0
            .Left = 0
            .Bottom = 8
            .Right = 8
        
            '// Fix Bottom
            If lBubbleY + (.Bottom - .Top) > lScreenHeight Then
                .Bottom = .Top + (lScreenHeight - lBubbleY)
            End If
            
            '// Fix Right
            If lBubbleX + (.Right - .Left) > lScreenWidth Then
                .Right = .Left + (lScreenWidth - lBubbleY)
            End If
            
            '// Fix Top
            If lBubbleY < 0 Then
                .Top = .Top - lBubbleY
                lY = 0
            Else
                lY = lBubbleY
            End If
            
            '// Fix Left
            If lBubbleX < 0 Then
                .Left = .Left - lBubbleX
                lX = 0
            Else
                lX = lBubbleX
            End If
        End With
        
        '// Draw Large Bubble
        msBack.BltFast lX, lY, msBubbleLarge, rBubble, DDBLTFAST_SRCCOLORKEY
    Else
        '// Set Rectangle
        With rBubble
            .Top = 0
            .Left = 0
            .Bottom = 6
            .Right = 4
        
            '// Fix Bottom
            If lBubbleY + (.Bottom - .Top) > lScreenHeight Then
                .Bottom = .Top + (lScreenHeight - lBubbleY)
            End If
            
            '// Fix Right
            If lBubbleX + (.Right - .Left) > lScreenWidth Then
                .Right = .Left + (lScreenWidth - lBubbleY)
            End If
            
            '// Fix Top
            If lBubbleY < 0 Then
                .Top = .Top - lBubbleY
                lY = 0
            Else
                lY = lBubbleY
            End If
            
            '// Fix Left
            If lBubbleX < 0 Then
                .Left = .Left - lBubbleX
                lX = 0
            Else
                lX = lBubbleX
            End If
        End With
        
        '// Draw Small Bubble
        msBack.BltFast lX, lY, msBubbleSmall, rBubble, DDBLTFAST_SRCCOLORKEY
    End If
End Sub


Public Sub Initialize()
    '// Set Random Position
    Randomize Timer
    lBubbleY = CLng(Rnd(1) * (lScreenHeight - 100) + 25)
    lBubbleX = CLng(Rnd(1) * (lScreenWidth - 100) + 25)
    bBubbleLeft = IIf(CInt(Rnd(1)) = 1, True, False)
    lBubbleVSpeed = CLng(Rnd(1) * 5 + 1)
    lBubbleHSpeed = IIf(bBubbleLeft, -5, 5)
    bBubbleLarge = IIf(CInt(Rnd(1)) = 1, True, False)
End Sub


