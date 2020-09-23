VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiddie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lDiddieX As Long
Private lDiddieY As Long

Private bDiddieSwim As Boolean
Private bDiddieLeft As Boolean

Private lFrame As Long
Private lFrameTop As Long

Private bVMoving As Boolean
Private lVMoving As Long
Private lVTotal As Long
Public Sub DoMovement()
    Static lTimer As Long
    Static lVTimer As Long
    Dim lDirect As Long
    
    If bDiddieSwim Then
        If mdX.TickCount() - lTimer >= 10 Then
            lTimer = mdX.TickCount()
            
            '// Move Horizontally
            If bDiddieLeft Then
                lDiddieX = lDiddieX - 2
                
                If lDiddieX <= -75 Then bDiddieSwim = False
            Else
                lDiddieX = lDiddieX + 2
            
                If lDiddieX >= lScreenWidth + 50 Then bDiddieSwim = False
            End If
        End If
        
        If mdX.TickCount() - lVTimer >= 100 Then
            lVTimer = mdX.TickCount()
                       
            '// Move Frame
            lFrame = lFrame + 1
            If lFrame = 6 Then lFrame = 0
        
            If bVMoving Then
                '// Move
                lVTotal = lVTotal + 1
                If lVTotal = 10 Then bVMoving = False
                
                lDiddieY = lDiddieY + lVMoving
            Else
                '// Decide Movements
                lDirect = CLng(Rnd(1) * 2)
                If lDirect = 1 And lDiddieY >= 50 Then
                    lVMoving = -1
                    lVTotal = 0
                    bVMoving = True
                End If
                
                If lDirect = 2 And lDiddieY <= lScreenHeight - 50 Then
                    lVMoving = 1
                    lVTotal = 0
                    bVMoving = True
                End If
            End If
        End If
        
        '// Set Frame Top
        If bDiddieLeft Then lFrameTop = 0 Else lFrameTop = 39
    Else
        '// Check if swim should start
        If CInt(Rnd(1) * 200) = 100 Then
            '// Set Random Position
            Randomize Timer
            lDiddieY = CLng(Rnd(1) * (lScreenHeight - 100) + 25)
            bDiddieLeft = Not bDiddieLeft
            If bDiddieLeft Then lDiddieX = lScreenWidth Else lDiddieX = -45
            bDiddieSwim = True
        End If
    End If
End Sub


Public Sub Draw()
    If bDiddieSwim Then
        Dim cKey As DDCOLORKEY
        Dim rDiddie As RECT
        Dim lX As Long
        Dim lY As Long
        
        '// Set Transparent Color
        cKey.high = mdX.CreateColorRGB(255, 0, 255)
        cKey.low = cKey.high
        msDiddie.SetColorKey DDCKEY_SRCBLT, cKey
        
        '// Set Rectangle
        With rDiddie
            .Top = lFrameTop
            .Left = lFrame * 45
            .Bottom = .Top + 39
            .Right = .Left + 45
            
            '// Fix Bottom
            If lDiddieY + (.Bottom - .Top) > lScreenHeight Then
                .Bottom = .Top + (lScreenHeight - lDiddieY)
            End If
            
            '// Fix Right
            If lDiddieX + (.Right - .Left) > lScreenWidth Then
                .Right = .Left + (lScreenWidth - lDiddieX)
            End If
            
            '// Fix Top
            If lDiddieY < 0 Then
                .Top = .Top - lDiddieY
                lY = 0
            Else
                lY = lDiddieY
            End If
            
            '// Fix Left
            If lDiddieX < 0 Then
                .Left = .Left - lDiddieX
                lX = 0
            Else
                lX = lDiddieX
            End If
        End With
        
        msBack.BltFast lX, lY, msDiddie, rDiddie, DDBLTFAST_SRCCOLORKEY
    End If
End Sub


Public Sub Initialize()
    '// Set Random Position
    Randomize Timer
    lDiddieY = CLng(Rnd(1) * (lScreenHeight - 100) + 25)
    bDiddieLeft = IIf(CInt(Rnd(1)) = 1, True, False)
    If bDiddieLeft Then lDiddieX = lScreenWidth Else lDiddieX = -45
    bDiddieSwim = True
End Sub


