VERSION 5.00
Begin VB.Form frmFractal 
   BackColor       =   &H80000005&
   Caption         =   "Mandelbrot Fractal Demo - www.tannerhelland.com"
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   705
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset"
      Height          =   615
      Left            =   8280
      TabIndex        =   14
      Top             =   480
      Width           =   855
   End
   Begin VB.HScrollBar scrAccuracy 
      Height          =   255
      Left            =   1200
      Max             =   20
      Min             =   1
      TabIndex        =   10
      Top             =   90
      Value           =   1
      Width           =   6255
   End
   Begin VB.CommandButton CmdRedraw 
      Appearance      =   0  'Flat
      Caption         =   "Draw the Fractal"
      Height          =   615
      Left            =   6120
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox TxtYMax 
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
      Height          =   330
      Left            =   4560
      TabIndex        =   4
      Text            =   "1.5"
      Top             =   720
      Width           =   1365
   End
   Begin VB.TextBox TxtYMin 
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
      Height          =   330
      Left            =   3120
      TabIndex        =   3
      Text            =   "-1.5"
      Top             =   720
      Width           =   1365
   End
   Begin VB.TextBox TxtXMax 
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
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Text            =   ".75"
      Top             =   720
      Width           =   1365
   End
   Begin VB.TextBox TxtXMin 
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
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Text            =   "-1.75"
      Top             =   720
      Width           =   1365
   End
   Begin VB.PictureBox PicDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   599
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   599
      TabIndex        =   0
      Top             =   1440
      Width           =   9015
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Use the mouse to click-drag new bound values. (Press the ""draw"" button above after selecting a new frame of reference.)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   9015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accurate shading"
      Height          =   210
      Left            =   7560
      TabIndex        =   12
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fast shading"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Max:"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Y Min:"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "X Max:"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X Min:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmFractal"
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

'These dimensions relate to using the mouse to select new bounding areas
Dim iX As Double, iY As Double
Dim tX As Double, tY As Double
Dim iWidth As Double, iHeight As Double
Dim toDraw As Boolean

'This array stores the values of the fractal as they're calculated
Dim mFractal() As Double

'This number (set by the corresponding scrollbar) takes on values between 100 and 2000
'Higher numbers = more accurate/slow; lower mumbers = less accurate/fast
Dim IterationLimit As Long

'Simple API declaration for drawing pixels
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte


'DRAW THE FRACTAL
Private Sub CmdRedraw_Click()

    'Lots of variables are required for fractal-drawing
    Dim drawWidth As Long, drawHeight As Long
    Dim count As Double
    Dim red As Long, green As Long, blue As Long
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim calcX As Double, calcY As Double, stepX As Double, stepY As Double
    Dim color As Long
    Dim loValue As Double, hiValue As Double
    loValue = 2000000000#
    hiValue = -2000000000#
    
    'Reset the mouse values after each draw (since no box is present after drawing)
    iX = -1
    iY = -1
    tX = -1
    tY = -1
    
    'Prepare the picture box using hard-coded size values.  (Feel free to change these.)
    drawWidth = 600
    drawHeight = 600
    PicDraw.Width = drawWidth + 2
    PicDraw.Height = drawHeight + 2
    PicDraw.Refresh
    DoEvents
    
    'Set minimum and maximum view windows based on corresponding text box values
    minX = Val(TxtXMin)
    maxX = Val(TxtXMax)
    minY = Val(TxtYMin)
    maxY = Val(TxtYMax)
    
    'Create a conversion between the Mandlebrot set window and the actual pixels where the fractal will appear
    stepX = Abs(maxX - minX) / drawWidth
    stepY = Abs(maxY - minY) / drawHeight
    
    'Resize the fractal array
    ReDim mFractal(0 To drawWidth, 0 To drawHeight) As Double
    
    'Run a loop through each pixel, calculating Mandelbrot values as we go
    calcX = minX
    calcY = minY
    
    Dim x As Long, y As Long
    Dim rootCount As Double
    
    For y = 0 To drawHeight
    For x = 0 To drawWidth
    
        'This main recursive function is what actually calculates the fractal values
        count = Calculate(calcX, calcY, 0, 0, 0)
        rootCount = Sqr(count)
        
        'Keep track of the highest and lowest values encountered so far (to help with rendering)
        If rootCount < loValue Then
            loValue = rootCount
        Else
            If rootCount > hiValue Then hiValue = rootCount
        End If
        
        'Store the fractal value in the array
        mFractal(x, y) = rootCount
        
        calcX = calcX + stepX
        
    Next x
        calcX = minX
        calcY = calcY + stepY
        'Give the user some indication of how close we are to finishing
        CmdRedraw.Caption = Int((y / drawHeight) * 100) & "% Complete"
    Next y
    
    'Once the fractal values have been saved, calculate the difference between the lowest and highest values
    Dim lohiDif As Double
    lohiDif = hiValue - loValue
    If lohiDif = 0# Then lohiDif = 0.0000000001
    
    'Draw the fractal
    For y = 0 To drawHeight
    For x = 0 To drawWidth
    
        'Convert the fractal value to a 0-255 RGB value
        count = Int(((mFractal(x, y) - loValue) / lohiDif) * 255# + 0.5)
        
        'Make the highest values pure white (looks better, IMO)
        If mFractal(x, y) = hiValue Then
            red = 255
            green = 255
            blue = 255
        Else
            'I use this simple color formula to make the fractal shades of violet instead of just gray
            red = count
            green = count \ 2
            blue = count * 2
            If blue > 255 Then blue = 255
        End If
        
        'Draw this pixel
        SetPixelV PicDraw.hDC, x, y, RGB(red, green, blue)
        
    Next x
        'Refresh the image every 16 lines
        If y And 15 = 0 Then PicDraw.Refresh
    Next y
    
    'Final image refresh
    PicDraw.Refresh

    'Restore the command button text
    CmdRedraw.Caption = "Redraw the Fractal"
    
End Sub

Private Sub CmdReset_Click()
    toDraw = False
    TxtXMin.Text = "-1.75"
    TxtXMax.Text = ".75"
    TxtYMin.Text = "-1.5"
    TxtYMax.Text = "1.5"
    IterationLimit = 400
    iX = -1
    iY = -1
    tX = -1
    tY = -1
    CmdRedraw_Click
End Sub

'ON STARTUP: set various variables to proper default values
Private Sub Form_Load()
    toDraw = False
    IterationLimit = 400
    iX = -1
    iY = -1
    tX = -1
    tY = -1
End Sub

'MOUSE DOWN: on the main picture box
Private Sub PicDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Erase any previous boxes the user may have drawn
    PicDraw.DrawMode = 6
    PicDraw.Line (tX, tY)-(tX, iY)
    PicDraw.Line (tX, tY)-(iX, tY)
    PicDraw.Line (iX, iY)-(tX, iY)
    PicDraw.Line (iX, iY)-(iX, tY)
    PicDraw.DrawMode = 13
    
    'Note the current location and mark a mouse button as being "down"
    iX = x
    iY = y
    tX = x
    tY = y
    toDraw = True
    
End Sub

'MOUSE MOVE: on the main picture box
Private Sub PicDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If a mouse button has been pressed, draw a rectangle around the current selection area
    If toDraw Then
    
        'Set the draw style to "invert"
        PicDraw.DrawMode = 6
        
        'Erase the old box
        PicDraw.Line (tX, tY)-(tX, iY)
        PicDraw.Line (tX, tY)-(iX, tY)
        PicDraw.Line (iX, iY)-(tX, iY)
        PicDraw.Line (iX, iY)-(iX, tY)
        
        'Draw a new box
        PicDraw.Line (iX, iY)-(iX, y)
        PicDraw.Line (iX, iY)-(x, iY)
        PicDraw.Line (x, y)-(iX, y)
        PicDraw.Line (x, y)-(x, iY)
        
        'Remember where this box is located
        tX = x
        tY = y
        
        'Reset the draw style to "copy"
        PicDraw.DrawMode = 13
        
    End If
    
End Sub

'MOUSE UP: on the main picture box
Private Sub PicDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Mark the mouse button as being released
    toDraw = False
    
    'Figure out the width and height of the selected area
    iWidth = x - iX
    iHeight = y - iY
    
    'Use some fancy math to calculate the approximate reference frame that the user has selected.
    '(Basically, translate the pixel width on-screen to the coordinate system used by Mandelbrot.)
    Dim nX As Double, nY As Double
    Dim oX As Double, oY As Double
    Dim oSX As Double, oEX As Double
    Dim oSY As Double, oEY As Double
    
    nX = x
    nY = y
    
    oSX = TxtXMin
    oEX = TxtXMax
    oSY = TxtYMin
    oEY = TxtYMax
    
    oX = (iX / PicDraw.ScaleWidth) * (TxtXMax - TxtXMin)
    oX = oSX + oX
    
    oY = (iY / PicDraw.ScaleHeight) * (TxtYMax - TxtYMin)
    oY = oSY + oY
    
    nX = (nX / PicDraw.ScaleWidth) * (TxtXMax - TxtXMin)
    nX = oSX + nX
    
    nY = (nY / PicDraw.ScaleHeight) * (TxtYMax - TxtYMin)
    nY = oSY + nY
    
    'Change the text boxes to match our new frame of reference
    TxtXMin = oX
    TxtXMax = nX
    TxtYMin = oY
    TxtYMax = nY
    
End Sub

'SCROLLBAR: for changing the accuracy / speed of the fractal render
Private Sub scrAccuracy_Change()
    IterationLimit = scrAccuracy.Value * 100
End Sub

'Here is a standard recursive function for generating the Mandlebrot set. Technical details are available at the Wikipedia link referenced earlier.
Private Function Calculate(ByVal origX As Double, ByVal origY As Double, ByVal curX As Double, ByVal curY As Double, ByVal counter As Long) As Long
    
    Static NewX As Double, NewY As Double
    
    'Check for overflow (well, technically "underflow" :)
    If curX < -1000000 Then
        Calculate = counter
        Exit Function
    ElseIf curY < -1000000 Then
        Calculate = counter
        Exit Function
    End If
    
    'If you so desire, you can also remove the if/then statements above and let VB do the underflow checking (uncomment this next line to allow)
    'On Error GoTo 1
    
    NewX = (curX * curX) - (curY * curY) + origX
    NewY = (2 * curX * curY) + origY
    
    If ((NewX * NewX) - (NewY * NewY)) > 2 Then
1       Calculate = counter
    Else
        'This line checks to see if we've exceeded the "IterationLimit" (IterationLimit controls accuracy / speed of execution)
        If counter >= IterationLimit Then
            Calculate = IterationLimit
        Else
            counter = counter + 1
            Calculate = Calculate(origX, origY, NewX, NewY, counter)
        End If
    End If
    
End Function
