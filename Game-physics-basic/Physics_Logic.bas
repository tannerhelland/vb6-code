Attribute VB_Name = "Logic_Module"
'Simple Game Physics Demo - www.tannerhelland.com

Option Explicit

'API CALLS:
'BitBlt is used to draw the ship and bullets onto the screen
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'SetPixelV is used to draw the background stars
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte

'CONSTANTS (feel free to set these to whatever you'd like):
Public Const NumOfStars As Long = 500
Public Const BulletSpeed As Long = 20
Public Const HAcc As Single = 1
Public Const HDel As Single = 1
Public Const VAcc As Single = 1
Public Const VDel As Single = 1
Public Const KeySpeed As Long = 10
Public Const NumOfBullets As Long = 75

'TYPES:
Public Type Star
    X As Single
    Y As Single
    Bright As Byte
    Speed As Single
End Type

Public Type Bullet
    X As Long
    Y As Long
    Velocity As Long
    Activated As Boolean
End Type

'VARIABLES:

'Ship coordinates
Public ShipX As Long, ShipY As Long

'The ship's velocity in each direction
Public sLeft As Long
Public sRight As Long
Public sUp As Long
Public sDown As Long
Public sVelVert As Single
Public sVelHoriz As Single

'Whether or not the ship is firing
Public Firing As Boolean

'The number of active bullets on the screen
Public BulletsActivated As Long
'An array of all the bullets currently being fired
Public Bullets(0 To NumOfBullets) As Bullet

'An array of all the background stars
Public StarArray(0 To NumOfStars) As Star

'Whether or not to start the game
Public GameActive As Boolean

'Plain ol' looping variables
Dim X As Long, Y As Long

'The width and height of the game screen
Public BufferWidth As Long, BufferHeight As Long


'When the program is started, this routine is run first
Public Sub InitializeGameEngine()
    
    'Randomize the random number generator
    Randomize Timer
    
    'Store the width and height of the game area
    BufferWidth = frmMain.PicMain.ScaleWidth
    BufferHeight = frmMain.PicMain.ScaleHeight
    
    'Build all of the background stars
    For X = 0 To NumOfStars
        BuildStar X, True
    Next X
    
    'Set the ship in the middle of the picture box
    ShipX = frmMain.PicMain.ScaleWidth / 2 - 32
    ShipY = frmMain.PicMain.ScaleHeight - 64
    
    'No active bullets (...yet!)
    BulletsActivated = 0

    'Show the main form
    frmMain.Show
    
End Sub

'This routine draws the background stars onto the buffer
Public Sub DrawStars()
    
    If (Not GameActive) Then Exit Sub
    
    'First, clear the buffer of any previous stars
    frmMain.PicScreenBuffer.Cls
    
    'Run a loop through every star in the array
    For X = 0 To NumOfStars
        'Move the star down the screen according to its speed
        StarArray(X).Y = StarArray(X).Y + StarArray(X).Speed
        
        'If the star moves off the screen, make a new one in its place
        If StarArray(X).Y > BufferHeight Then BuildStar X
        
        'Draw a pixel at location (x,y) using the star's brightness value
        SetPixelV frmMain.PicScreenBuffer.hDC, StarArray(X).X, StarArray(X).Y, RGB(StarArray(X).Bright, StarArray(X).Bright, StarArray(X).Bright)
        
    Next X

End Sub

'This routine tracks and processes the movement of our little spaceship
Public Sub VelocityCode()
    
    If (Not GameActive) Then Exit Sub
    
    'Movement Up
    If sUp = 1 Then
        'Increase velocity as per the acceleration variable
        sVelVert = sVelVert - VAcc
        'Increase position as per the new velocity variable
        ShipY = ShipY + sVelVert
    End If
    
    'Movement Down (see detailed comments under "Movement Up" above)
    If sDown = 1 Then
        sVelVert = sVelVert + VAcc
        ShipY = ShipY + sVelVert
    End If
    
    'Vertical Deceleration (applied only when no keys have been pressed AND the ship
    ' is already in motion)
    If sUp = 0 And sDown = 0 And sVelVert <> 0 Then
        If sVelVert > 0 Then
            sVelVert = sVelVert - VDel
            If sVelVert <= 0 Then sVelVert = 0
        Else
            sVelVert = sVelVert + VDel
            If sVelVert >= 0 Then sVelVert = 0
        End If
        
        'Apply the deceleration to the ship's velocity
        ShipY = ShipY + sVelVert
    End If
    
    'Movement Left
    If sLeft = 1 Then
        sVelHoriz = sVelHoriz - HAcc
        ShipX = ShipX + sVelHoriz
    End If
    
    'Movement Right
    If sRight = 1 Then
        sVelHoriz = sVelHoriz + HAcc
        ShipX = ShipX + sVelHoriz
    End If
    
    'Horizontal Deceleration
    If sRight = 0 And sLeft = 0 And sVelHoriz <> 0 Then
        If sVelHoriz > 0 Then
            sVelHoriz = sVelHoriz - HDel
            If sVelHoriz <= 0 Then sVelHoriz = 0
        Else
            sVelHoriz = sVelHoriz + HDel
            If sVelHoriz >= 0 Then sVelHoriz = 0
        End If
        ShipX = ShipX + sVelHoriz
    End If
    
    'If the ship scrolls completely off-screen, move it to the other side
    If ShipX > BufferWidth Then ShipX = 0
    If ShipX < -64 Then ShipX = BufferWidth
    If ShipY > BufferHeight Then ShipY = 0
    If ShipY < -64 Then ShipY = BufferHeight
    
    'If the ship isn't moving left or right, reset the ship picture to the default orientation
    If sVelHoriz = 0 Then
        frmMain.picShip.Picture = frmMain.PicT.Picture
        frmMain.picShipMask.Picture = frmMain.PicTM.Picture
    End If
    
    'Last but not least, draw the ship onto the buffer
    BitBlt frmMain.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, frmMain.picShipMask.hDC, 0, 0, vbMergePaint
    BitBlt frmMain.PicScreenBuffer.hDC, ShipX, ShipY, 64, 64, frmMain.picShip.hDC, 0, 0, vbSrcAnd

End Sub

'This routine fires bullets/laser blasts from the ship's nose
Public Sub FireBullets()
    
    If (Not GameActive) Then Exit Sub
    
    'Run a loop through every bullet
    For X = 0 To NumOfBullets
    
        'If the user is firing, and this bullet is inactive, make it active
        If Firing = True And BulletsActivated < 1 And Bullets(X).Activated = False Then
            
            'Mark this bullet as active
            Bullets(X).Activated = True
            
            'Let the engine know that we've activated a bullet this round (so as to not activate more)
            BulletsActivated = BulletsActivated + 1
            
            'Set the bullet in the center of the ship
            Bullets(X).X = ShipX + 30
            
            'Set the bullet to start from either side of the ship's nose
            Bullets(X).Y = ShipY + 10
            
        End If
        
        'If the bullet is active, update its information
        If Bullets(X).Activated = True Then
        
            'First, move the bullet up according to its speed
            Bullets(X).Y = Bullets(X).Y - BulletSpeed
            
            'If the bullet leaves the screen, deactivate it
            If Bullets(X).Y < -4 Then Bullets(X).Activated = False
            
            'Draw our bullet(s) to the screen
            BitBlt frmMain.PicScreenBuffer.hDC, Bullets(X).X - 6, Bullets(X).Y, 4, 4, frmMain.PicBulletM.hDC, 0, 0, vbMergePaint
            BitBlt frmMain.PicScreenBuffer.hDC, Bullets(X).X - 6, Bullets(X).Y, 4, 4, frmMain.PicBullet.hDC, 0, 0, vbSrcAnd
            BitBlt frmMain.PicScreenBuffer.hDC, Bullets(X).X + 6, Bullets(X).Y, 4, 4, frmMain.PicBulletM.hDC, 0, 0, vbMergePaint
            BitBlt frmMain.PicScreenBuffer.hDC, Bullets(X).X + 6, Bullets(X).Y, 4, 4, frmMain.PicBullet.hDC, 0, 0, vbSrcAnd
            
        End If
        
    'Now do the same for the next bullet
    Next X
    
    'Reset the number of activated bullets to zero
    BulletsActivated = 0
    
    'Since this is the last step in the logic process, copy the entire buffer to the screen
    BitBlt frmMain.PicMain.hDC, 0, 0, BufferWidth, BufferHeight, frmMain.PicScreenBuffer.hDC, 0, 0, vbSrcCopy
        
End Sub

'Use this to create a star; if firstTime is true, the star can appear anywhere - if false,
' it is created only along the top edge of the screen
Public Sub BuildStar(ByVal ArrayVal As Long, Optional ByVal firstTime As Boolean = False)
    
    'Random x location
    StarArray(ArrayVal).X = Rnd * BufferWidth
    
    'Calculate the y location according to firstTime
    If firstTime = True Then
        StarArray(ArrayVal).Y = Rnd * frmMain.PicMain.ScaleHeight
    Else
        StarArray(ArrayVal).Y = 0
    End If
    
    'Brightness is somewhere between 5 and 255
    StarArray(ArrayVal).Bright = (Rnd * 250) + 5
    
    'Star velocity is somewhere between .1 and 8
    StarArray(ArrayVal).Speed = Rnd * 8 + 0.1
    
End Sub
