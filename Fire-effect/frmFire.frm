VERSION 5.00
Begin VB.Form frmFire 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Even Faster Real-Time Fire Effect - www.tannerhelland.com"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFire.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hsBlue 
      Height          =   255
      Left            =   840
      Max             =   300
      TabIndex        =   16
      Top             =   1800
      Value           =   50
      Width           =   6975
   End
   Begin VB.HScrollBar hsGreen 
      Height          =   255
      Left            =   840
      Max             =   300
      TabIndex        =   15
      Top             =   1440
      Value           =   100
      Width           =   6975
   End
   Begin VB.HScrollBar hsRed 
      Height          =   255
      Left            =   840
      Max             =   300
      TabIndex        =   14
      Top             =   1080
      Value           =   200
      Width           =   6975
   End
   Begin VB.HScrollBar scrUniformity 
      Height          =   255
      Left            =   3840
      Max             =   3
      TabIndex        =   10
      Top             =   120
      Value           =   2
      Width           =   1695
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "128"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "512"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      Appearance      =   0  'Flat
      Caption         =   "Stop Fire"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdStart 
      Appearance      =   0  'Flat
      Caption         =   "Start Fire"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   120
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      ToolTipText     =   "Here's your flame!"
      Top             =   2280
      Width           =   7710
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: changes will not take place until the fire is restarted."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label lblFrameRate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Frame Rate..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Uniformity:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Copyright 2018 by Tanner Helland
' www.tannerhelland.com
'
'Documentation for this project can be found at https://tannerhelland.com/code/
'
'The source code in this project is licensed under a Simplified BSD license.
' For more information, please review LICENSE.md at https://github.com/tannerhelland/thdc-code/
'
'If you find this code useful, please consider a small donation to https://www.paypal.me/TannerHelland
'
'***************************************************************************

'API:
'BitBlt for blitting the fire from the buffer to the screen
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Pixel routines; faster then PSet and Point but not as fast as GetDIBits, StretchDIBits
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'High performance timer functions
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long

'Values for tracking function times
Private curFreq As Currency, curStart As Currency, curEnd As Currency

'Whether or not to run the flames
Private Engage As Boolean

'How many frames have elapsed since the program began
Private frameCount As Long

'Pressing the START button initializes the screen and calls the DrawFastFlames() routine
Private Sub cmdStart_Click()
    
    'First, clear the picture boxes
    picDest.Picture = LoadPicture(vbNullString)
    picDest.Refresh
    
    'Collect the width and height values of the flame effect
    Dim tempX As Long, tempY As Long
    tempX = Val(txtWidth)
    tempY = Val(txtHeight)
    
    'Resize the picture boxes to match (and add two pixels for the borders of the picture box (1 pixel each))
    picDest.Width = tempX + 2
    picDest.Height = tempY + 2
    
    'Determine the resolution of this system's high performance timer
    QueryPerformanceFrequency curFreq
    
    'Set the "engage" parameter to true and call the fire function
    frameCount = 0
    Engage = True
    DrawFastFlames picDest, tempX, tempY, CSng(hsRed) / 100!, CSng(hsGreen) / 100!, CSng(hsBlue) / 100!, scrUniformity.Value

End Sub

'When STOP is pressed, disengage the fire routine
Private Sub cmdStop_Click()
    Engage = False
End Sub

'This routine draws a cool fire effect using the specified parameters
Private Sub DrawFastFlames(dstPic As PictureBox, ByVal flameWidth As Long, ByVal flameHeight As Long, ByVal redModifier As Single, ByVal greenModifier As Single, ByVal blueModifier As Single, Optional ByVal Uniformity As Byte = 2)
    
    'This array holds all the individual flame pixels
    Dim FlameArray() As Long
    ReDim FlameArray(0 To flameWidth, 0 To flameHeight) As Long
    
    'How much of the flames to reheat after each drawing (typically the bottom 10% of the array)
    Dim FillVal As Long
    FillVal = flameHeight * 0.9
    
    'This will temporarily hold the value of each flame pixel
    Dim curPixel As Long
    
    'As much as I hate an arbitrary name like "temp", that's what we'll use this for
    Dim tmpValue As Long
    
    'Loop variables
    Dim x As Long, y As Long
    
    'This is used to generate an RGB() value for the flame's color
    Static Color As Long
    
    'Look-up tables for each color (faster than generating them on the fly)
    Dim Red(0 To 255) As Byte
    Dim Green(0 To 255) As Byte
    Dim Blue(0 To 255) As Byte
    For x = 0 To 255
        Red(x) = ByteMe(x * redModifier)
        Green(x) = ByteMe(x * greenModifier)
        Blue(x) = ByteMe(x * blueModifier)
    Next x

    'Used to calculate FPS
    Dim curTime As Single

    'NEW! DIB Section code
    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Depending on VB quirks, the DIB array may or may not be the same size
    ' as the scalewidth/scaleheight properties
    Dim iWidth As Long, iHeight As Long
    Dim tmpX As Long, tmpY As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(picDest)
    iHeight = fDraw.GetImageHeight(picDest)
    fDraw.GetImageData2D picDest, imageData()

    'Uniformity is a somewhat arbitrary measurement; basically, I use it to control how
    ' random the flames vertical growth is.  The initial number passed into this function
    ' should lie between 0 and 3
    Uniformity = 5 - Uniformity
    
    'Before starting, randomize the seed for VB's built-in Randomize function
    Randomize Timer

    'If engaged, draw the flames!
    Do While Engage
        
        'Remember the current time (for calculating FPS at the end of the loop)
        QueryPerformanceCounter curStart
        
        'Run a loop through the flames from the bottom up...
        For y = 2 To flameHeight
        
        '...and from right-to-left
        For x = 0 To flameWidth
        
            'Grab the current flame pixel
            curPixel = FlameArray(x, y)
            
            'If this pixel is dim enough, ignore it for sake of uniformity
            If curPixel < Uniformity Then GoTo 1
            
            'Give this pixel a random degree of cooldown
            tmpValue = Int(Rnd * Uniformity)
            FlameArray(x, y) = curPixel - tmpValue
            
            'In addition, move this pixel upward
            tmpY = y - tmpValue
            If tmpY < 0 Then tmpY = 0
            FlameArray(x, tmpY) = FlameArray(x, y)
            
            'Generate a color based on the "heat" value
            Color = (FlameArray(x, y) * 255) \ flameHeight
            
            'To improve speed, randomize the x movement of every 5th pixel
            ' (as opposed to doing it for every pixel)
            If ((x + y) And 3) = 0 Then
                tmpX = (x + Int(Rnd * 2)) * 3
            Else
                tmpX = x * 3
            End If
            
            'Draw this flame pixel using a cool flame-colored formula (based upon
            ' the values passed into this function)
            If tmpX <= (iWidth * 3) - 3 Then
                tmpY = flameHeight - y
                If tmpY < 0 Then tmpY = 0
                imageData(tmpX, tmpY) = Blue(Color)
                imageData(tmpX + 1, tmpY) = Green(Color)
                imageData(tmpX + 2, tmpY) = Red(Color)
            End If
            
1       Next x
        Next y

        'Once all flames are processed, make the bottom 4 rows hot again
        For y = FillVal To flameHeight
        For x = 0 To flameWidth
            FlameArray(x, y) = flameHeight
        Next x
        Next y
        
        'Copy the whole flame image from the buffer to the viewable picture box
        fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()
        
        'Calculate an exact framerate and display it
        QueryPerformanceCounter curEnd
        lblFrameRate.Caption = Int(1 / ((curEnd - curStart) / curFreq)) & " FPS"
    
        'Every 32 frames, halt for external events (keypresses, mouse movement, etc.)
        frameCount = frameCount + 1
        If (frameCount And 31) Then DoEvents
    
    'Do it all over again!
    Loop

End Sub

'When the form is closed, make sure the flame function is disengaged
Private Sub Form_Unload(Cancel As Integer)
    Engage = False
End Sub

'This simple little function ensure that the variable it receives falls within the
' proper byte range (0-255)
Private Function ByteMe(ByVal toByte As Long) As Byte
    If (toByte > 255) Then
        ByteMe = 255
    ElseIf (toByte < 0) Then
        ByteMe = 0
    Else
        ByteMe = toByte
    End If
End Function
