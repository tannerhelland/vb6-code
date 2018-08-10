VERSION 5.00
Begin VB.Form frmSepia 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sepia / ""Antique"" Effect - www.tannerhelland.com"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6270
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdSepia 
      Appearance      =   0  'Flat
      Caption         =   "Apply Sepia Effect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   600
      Width           =   6030
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmSepia"
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

Option Explicit


'Copy the original image to the foreground picture box
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'Apply a sepia effect to the image
Private Sub cmdSepia_Click()
    DrawSepia Me.picMain, Me.picMain, True
End Sub

Private Sub Form_Load()

    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    Me.Show
    
End Sub

Private Sub MnuOpenImage_Click()

    'Windows common dialog interface
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'VB6's built-in image loader (which wraps an ancient OLE loader) only supports a
    ' subset of modern image file types.
    Dim cdfStr As String
    cdfStr = "All Compatible Graphics|*.bmp;*.jpg;*.jpeg;*.gif;*.wmf;*.emf|"
    cdfStr = cdfStr & "BMP - Windows Bitmap|*.bmp|EMF - Enhanced Metafile|*.emf|GIF - Compuserve|*.gif|JPG/JPEG - JFIF Compliant|*.jpg;*.jpeg|WMF - Windows Meta File|*.wmf|All files|*.*"
    
    Dim sFile As String
    If cDialog.GetOpenFileName(sFile, , True, False, cdfStr, , , "Select image", , Me.hWnd) Then
        LoadImageAutosized sFile
    End If
    
End Sub

'If a null string is passed, simply paint an unmodified copy of the backbuffer into the foreground buffer
Private Sub LoadImageAutosized(Optional ByVal srcFilePath As String = vbNullString)

    If (LenB(srcFilePath) <> 0) Then picBack.Picture = LoadPicture(srcFilePath)
        
    'Copy the image, automatically resized, from the background picture box to the foreground one
    Dim fDraw As FastDrawing
    Set fDraw = New FastDrawing
    
    Dim ImageData() As Byte, imgWidth As Long, imgHeight As Long
    imgWidth = fDraw.GetImageWidth(Me.picBack)
    imgHeight = fDraw.GetImageHeight(Me.picBack)
    fDraw.GetImageData2D Me.picBack, ImageData()
    
    Me.picMain.Width = imgWidth + 2
    Me.picMain.Height = imgHeight + 2
    fDraw.SetImageData2D Me.picMain, imgWidth, imgHeight, ImageData()
    
    'Resize the form to automatically contain the new picture box dimensions.
    ' (This performs some sloppy checks to keep the form from resizing larger than the primary display,
    '  but it is *not* good code - a proper solution would involve the AdjustWindowRect() API or similar!)
    Dim newWidth As Long, minWidth As Long
    minWidth = (cmdReset.Left + cmdReset.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 2.5) * Screen.TwipsPerPixelY
    
End Sub

'Apply a sepia (or "antique") filter to an image
Public Sub DrawSepia(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal shiftRight As Boolean = True)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, ImageData()
    
    'These variables will hold original pixel color values
    Dim r As Long, g As Long, b As Long, gray As Long

    'These variables will hold transformed pixel color values
    Dim tR As Long, tG As Long, tB As Long

    'As part of this transformation, we're going to adjust gamma values.
    ' This look-up table will help us do that quickly
    ' (Without this, the image will be very dim and difficult to see)
    Dim LookUp(0 To 255) As Integer
    Dim TempVal As Single
    For x = 0 To 255
        TempVal = x / 255
        TempVal = TempVal ^ (1 / 1.6)  ' 1.6 is the gamma adjustment
        TempVal = TempVal * 255
        If TempVal > 255 Then TempVal = 255
        If TempVal < 0 Then TempVal = 0
        LookUp(x) = TempVal
    Next x

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = ImageData(quickX + 2, y)
        g = ImageData(quickX + 1, y)
        b = ImageData(quickX, y)
        'Calculate a gray value programmatically
        gray = (r + g + b) \ 3
        'Blend each color value with gray (reduces saturation)
        r = (r + gray) \ 2
        g = (g + gray) \ 2
        b = (b + gray) \ 2
        'Mix and match color channels to create the "sepia" coloring
        ' (This has the unintended side-effect of darkening the pixels as well)
        r = (g * b) \ 255
        g = (b * r) \ 255
        b = (r * g) \ 255
        'To fix the darkening, multiply everything by 1.75. Using multiplication
        ' instead of addition means that bright colors become brighter, but dark
        ' colors stay dark.
        tR = r * 1.75
        tG = g * 1.75
        tB = b * 1.75
        'Make sure all our values will fit inside the look-up table
        If tR > 255 Then tR = 255
        If tG > 255 Then tG = 255
        If tB > 255 Then tB = 255
        'Use our previously generated look-up table to apply a gamma ramp to
        ' the colors
        tR = LookUp(tR)
        tG = LookUp(tG)
        tB = LookUp(tB)
        'Set the new color values into the pixel array and continue with the next one
        ImageData(quickX + 2, y) = tR
        ImageData(quickX + 1, y) = tG
        ImageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()


End Sub
