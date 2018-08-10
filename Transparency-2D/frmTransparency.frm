VERSION 5.00
Begin VB.Form frmTransparency 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Real-time transparency - tannerhelland.com"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmTransparency.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hScroll1 
      Height          =   375
      Left            =   120
      Max             =   100
      TabIndex        =   4
      Top             =   5160
      Value           =   50
      Width           =   6015
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   2
      Top             =   120
      Width           =   6030
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "frmTransparency.frx":000C
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "frmTransparency.frx":B2C5
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Percent Transparency:"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   4800
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmTransparency"
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

Private Sub Form_Load()
    'When the program starts, combine the two images at 50% transparency
    DrawTransparency Picture1, Picture2, Picture3, 50
End Sub

'When the scrollbar is used, change transparency accordingly
Private Sub hScroll1_Change()
    DrawTransparency Picture1, Picture2, Picture3, hScroll1.Value
End Sub

Private Sub hScroll1_Scroll()
    DrawTransparency Picture1, Picture2, Picture3, hScroll1.Value
End Sub

'Draw an alpha blend from two source picture boxes into a destination picture box.
'The transparency value is a simple percentage from 1 to 100
Public Sub DrawTransparency(srcPic1 As PictureBox, srcPic2 As PictureBox, dstPic As PictureBox, ByVal lvlTransparency As Byte)
    
    'These arrays will hold both image's pixel data
    Dim imageDataSrc1() As Byte, imageDataSrc2() As Byte, imageDataDst() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the first source image's data
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic1)
    iHeight = fDraw.GetImageHeight(srcPic1)
    fDraw.GetImageData2D srcPic1, imageDataSrc1()
    
    'Now do it for the second source image
    iWidth = fDraw.GetImageWidth(srcPic2)
    iHeight = fDraw.GetImageHeight(srcPic2)
    fDraw.GetImageData2D srcPic2, imageDataSrc2()
    
    'Last but not least, do it for the destination picturebox
    iWidth = fDraw.GetImageWidth(srcPic2)
    iHeight = fDraw.GetImageHeight(srcPic2)
    fDraw.GetImageData2D dstPic, imageDataDst()
    
    'These variables will hold temporary pixel color values
    Dim r As Byte, g As Byte, b As Byte
    Dim r2 As Byte, g2 As Byte, b2 As Byte
    
    'Build a look-up table to increase speed
    Dim lookUp(0 To 255, 0 To 255) As Byte
    Dim invTransparency As Byte
    invTransparency = 100 - lvlTransparency
    For x = 0 To 255
    For y = 0 To 255
        'Mix all possible color values based on simple weighted averaging
        lookUp(x, y) = CByte(((invTransparency * x) + (lvlTransparency * y)) \ 100)
    Next y
    Next x
    
    'Now run a quick loop through the image, adjusting pixel values with the look-up tables
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        'Grab red, green, and blue from the source images
        r = imageDataSrc1(quickX + 2, y)
        g = imageDataSrc1(quickX + 1, y)
        b = imageDataSrc1(quickX, y)
        r2 = imageDataSrc2(quickX + 2, y)
        g2 = imageDataSrc2(quickX + 1, y)
        b2 = imageDataSrc2(quickX, y)
        'Use our source values to calculate a new, transparent color
        imageDataDst(quickX + 2, y) = lookUp(r, r2)
        imageDataDst(quickX + 1, y) = lookUp(g, g2)
        imageDataDst(quickX, y) = lookUp(b, b2)
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D Picture3, iWidth, iHeight, imageDataDst()
    
End Sub

