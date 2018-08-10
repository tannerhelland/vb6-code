VERSION 5.00
Begin VB.Form frmFill 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Fill Image Regions Demo - www.tannerhelland.com"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   523
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDemo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DrawWidth       =   3
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   399
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   1
      Top             =   480
      Width           =   7575
   End
   Begin VB.Label lblInstructions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Use the left mouse button to draw lines, and the right button to fill a region with color."
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmFill"
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

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'Location variables necessary for drawing connected lines
Dim xStart As Long, yStart As Long

'This variable remembers whether or not the left mouse button has been pressed (see below)
Dim toDraw As Boolean

'When the left mouse button is pressed...
Private Sub picDemo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Check the mouse button
    ' If the left button is pressed...
    If Button = vbLeftButton Then
        
        'Remember that the left-hand button is down
        toDraw = True
        
        'Remember the current mouse location
        xStart = x
        yStart = y
    
    ' If the right button...
    ElseIf Button = vbRightButton Then
    
        'Seed VB's random number generator
        Randomize Timer
        
        'Set the picture box's .FillColor property to a random number
        picDemo.FillColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        
        'Use the API to fill this region
        'Parameters: (picture box), (x coordinate), (y coordinate), (color being replaced), (fill method - always 1)
        ExtFloodFill picDemo.hDC, x, y, picDemo.Point(x, y), 1
        
    End If
    
End Sub

'When the mouse is moved...
Private Sub picDemo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If drawing is enabled (meaning the left mouse button has previously been pressed),
    ' draw a line between the old x coordinate and the current one
    If toDraw Then
    
        'Use VB's line function to draw the line
        picDemo.Line (xStart, yStart)-(x, y)
        
        'Replace the old xStart and yStart variables with the current location of the mouse
        xStart = x
        yStart = y
        
    End If
    
End Sub

'When a mouse button is released...
Private Sub picDemo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Disable drawing
    toDraw = False
    
End Sub
