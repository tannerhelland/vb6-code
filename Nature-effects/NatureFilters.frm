VERSION 5.00
Begin VB.Form frmNatureFilters 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Effects Inspired by Nature - www.tannerhelland.com"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   13125
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
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset the Image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.ListBox lstFilters 
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox picMain 
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
      Height          =   5700
      Left            =   2880
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   0
      Top             =   120
      Width           =   10110
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
      Height          =   5700
      Left            =   2880
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   672
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   10110
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmNatureFilters"
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

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
    lstFilters.ListIndex = -1
End Sub

'Populate the possible edge detection routines
Private Sub Form_Load()
    
    lstFilters.AddItem " Atmosphere"
    lstFilters.AddItem " Burn"
    lstFilters.AddItem " Fog"
    lstFilters.AddItem " Freeze"
    lstFilters.AddItem " Lava"
    lstFilters.AddItem " Metal"
    lstFilters.AddItem " Ocean"
    lstFilters.AddItem " Rainbow"
    lstFilters.AddItem " Water"
    lstFilters.ListIndex = -1
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'When an edge detection routine is selected, run it and display its corresponding matrix in the label on the main form
Private Sub lstFilters_Click()
    UpdateEffect
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
        
        'Load the image and refresh the effect
        LoadImageAutosized sFile
        lstFilters.ListIndex = -1
        UpdateEffect
        
    End If
    
End Sub

'If a null string is passed, simply paint an unmodified copy of the backbuffer into the foreground buffer
Private Sub LoadImageAutosized(Optional ByVal srcFilePath As String = vbNullString)

    If (LenB(srcFilePath) <> 0) Then picBack.Picture = LoadPicture(srcFilePath)
        
    'Copy the image, automatically resized, from the background picture box to the foreground one
    Dim fDraw As FastDrawing
    Set fDraw = New FastDrawing
    
    Dim imageData() As Byte, imgWidth As Long, imgHeight As Long
    imgWidth = fDraw.GetImageWidth(Me.picBack)
    imgHeight = fDraw.GetImageHeight(Me.picBack)
    fDraw.GetImageData2D Me.picBack, imageData()
    
    Me.picMain.Width = imgWidth + 2
    Me.picMain.Height = imgHeight + 2
    fDraw.SetImageData2D Me.picMain, imgWidth, imgHeight, imageData()
    
    'Resize the form to automatically contain the new picture box dimensions.
    ' (This performs some sloppy checks to keep the form from resizing larger than the primary display,
    '  but it is *not* good code - a proper solution would involve the AdjustWindowRect() API or similar!)
    Dim newHeight As Long, minHeight As Long
    minHeight = (cmdReset.Top + cmdReset.Height + 60) * Screen.TwipsPerPixelY
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then
        newHeight = (imgHeight + picMain.Top * 9) * Screen.TwipsPerPixelY
        If (newHeight < minHeight) Then newHeight = minHeight
        Me.Height = newHeight
    End If
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then Me.Width = (imgWidth + picMain.Left * 1.15) * Screen.TwipsPerPixelX
    
End Sub

Private Sub UpdateEffect()
    
    Select Case lstFilters.ListIndex
        'Atmosphere
        Case 0
            DrawAtmosphere Me.picBack, Me.picMain
        'Burn
        Case 1
            DrawBurn Me.picBack, Me.picMain
        'Fog
        Case 2
            DrawFog Me.picBack, Me.picMain
        'Freeze
        Case 3
            DrawFreeze Me.picBack, Me.picMain
        'Lava
        Case 4
            DrawLava Me.picBack, Me.picMain
        'Metal
        Case 5
            DrawMetal Me.picBack, Me.picMain
        'Ocean
        Case 6
            DrawOcean Me.picBack, Me.picMain
        'Rainbow
        Case 7
            DrawRainbow Me.picBack, Me.picMain
        'Water
        Case 8
            DrawWater Me.picBack, Me.picMain
    End Select
    
End Sub

Public Sub DrawAtmosphere(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        tR = (g + b) \ 2
        tG = (r + b) \ 2
        tB = (r + g) \ 2
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawBurn(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long, gray As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        gray = (r + g + b) \ 3
        tR = gray * 3
        tG = gray
        tB = gray \ 3
        ByteMe tR
        ByteMe tG
        ByteMe tB
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawFog(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long

    'Change this value to change the "thickness" of the fog
    Dim FogLimit As Long
    FogLimit = 40

    'We'll use a look-up table to speed this up dramatically
    Dim lookUp(0 To 255) As Byte
    For x = 0 To 255
        If x > 127 Then
            lookUp(x) = x - FogLimit
            If lookUp(x) < 127 Then lookUp(x) = 127
        Else
            lookUp(x) = x + FogLimit
            If lookUp(x) > 127 Then lookUp(x) = 127
        End If
    Next x

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        imageData(quickX + 2, y) = lookUp(r)
        imageData(quickX + 1, y) = lookUp(g)
        imageData(quickX, y) = lookUp(b)
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawFreeze(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        tR = Abs((r - g - b) * 1.5)
        tG = Abs((g - b - tR) * 1.5)
        tB = Abs((b - tR - tG) * 1.5)
        ByteMe tR
        ByteMe tG
        ByteMe tB
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawLava(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long, gray As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        gray = (r + g + b) \ 3
        tR = gray
        tG = Abs(b - 128)
        tB = Abs(b - 128)
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawMetal(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long, gray As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        'This is a somewhat convoluted formula, and certainly not the best way to go about it.
        'If you can find a more reasonable way to construct this function, please let me know!
        tR = imageData(quickX + 2, y)
        r = Abs(tR - 64)
        g = Abs(r - 64)
        b = Abs(g - 64)
        gray = Int((222 * r + 707 * g + 71 * b) \ 1000) 'This is a more accurate grayscale conversion (based off human perception of color)
        r = gray + 70
        r = r + (((r - 128) * 100) \ 100)
        g = gray + 65
        g = g + (((g - 128) * 100) \ 100)
        b = gray + 75
        b = b + (((b - 128) * 100) \ 100)
        ByteMe r
        ByteMe g
        ByteMe b
        imageData(quickX + 2, y) = r
        imageData(quickX + 1, y) = g
        imageData(quickX, y) = b
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawOcean(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long, gray As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        gray = (r + g + b) \ 3
        tR = gray \ 3
        tG = gray
        tB = gray * 3
        ByteMe tR
        ByteMe tG
        ByteMe tB
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawRainbow(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel, HSL, and location values
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    Dim hVal As Single

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
        'Based on our x-position, gradient a hue value between -1 and 5
        hVal = (x / iWidth) * 360
        hVal = (hVal - 60) / 60
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        tRGBToHSL r, g, b, HH, SS, LL
        'Convert back to RGB using our artificial hue value
        tHSLToRGB hVal, 0.5, LL, r, g, b
        imageData(quickX + 2, y) = r
        imageData(quickX + 1, y) = g
        imageData(quickX, y) = b
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

Public Sub DrawWater(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, tR As Long, tG As Long, tB As Long, gray As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        gray = (r + g + b) \ 3
        tR = gray - g - b
        tG = gray - tR - b
        tB = gray - tR - tG
        ByteMe tR
        ByteMe tG
        ByteMe tB
        imageData(quickX + 2, y) = tR
        imageData(quickX + 1, y) = tG
        imageData(quickX, y) = tB
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert to absolute byte values (Long-type)
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub

'HSL conversion routines
Public Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Single, s As Single, L As Single)
Dim Max As Single, Min As Single, delta As Single
Dim rR As Single, rG As Single, rB As Single
   rR = r / 255: rG = g / 255: rB = b / 255
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        L = (Max + Min) / 2
        If Max = Min Then
            s = 0
            h = 0
        Else
           If L <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            delta = Max - Min
           If rR = Max Then
                h = (rG - rB) / delta
           ElseIf rG = Max Then
                h = 2 + (rB - rR) / delta
           ElseIf rB = Max Then
                h = 4 + (rR - rG) / delta
           End If
      End If
End Sub

Public Sub tHSLToRGB(h As Single, s As Single, L As Single, r As Long, g As Long, b As Long)
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single
   If s = 0 Then
      rR = L: rG = L: rB = L
   Else
      If L <= 0.5 Then
         Min = L * (1 - s)
      Else
         Min = L - s * (1 - L)
      End If
      Max = 2 * L - Min
      If (h < 1) Then
         rR = Max
         If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
         Else
            rB = Min
            rG = h * (Max - Min) + rB
         End If
      ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
         Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
         Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
         End If
      End If
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

'Return the maximum of three variables
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

'Return the minimum of three variables
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

