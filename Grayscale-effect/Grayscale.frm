VERSION 5.00
Begin VB.Form frmGrayscale 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grayscale Conversion Algorithms - www.tannerhelland.com"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   15135
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
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1009
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameShades 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "How many shades of gray? "
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
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
      Begin VB.HScrollBar hscrShades 
         Height          =   255
         Left            =   120
         Max             =   256
         Min             =   2
         TabIndex        =   13
         Top             =   480
         Value           =   6
         Width           =   2535
      End
      Begin VB.Label lblShades 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   14
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame frameDecompose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Decompose according to... "
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
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Maximum values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optDecompose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Minimum values"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame frameChannel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select a color channel: "
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
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton optChannel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.ListBox lstFilters 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Show the Original Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   7080
      Width           =   2535
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
      Height          =   735
      Left            =   3000
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   0
      Top             =   120
      Width           =   1335
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
      Height          =   7530
      Left            =   3000
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   12030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grayscale algorithms: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1755
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmGrayscale"
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


'When the program is first loaded, several things needs to happen...
Private Sub Form_Load()

    'Populate the list box with the various available grayscale conversion algorithms
    lstFilters.AddItem "Average value [(R+G+B) / 3]", 0
    lstFilters.AddItem "Adjusted for the human eye", 1
    lstFilters.AddItem "Desaturate", 2
    lstFilters.AddItem "Decompose", 3
    lstFilters.AddItem "Single color channel", 4
    lstFilters.AddItem "Specific # of shades", 5
    lstFilters.AddItem "Specific # of shades (dithered)", 6
    
    'Start the program with nothing selected yet
    lstFilters.ListIndex = -1
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'When the # of shades scroll bar is changed, make the associated label display the current value (and redraw the image)
Private Sub hscrShades_Change()
    lblShades.Caption = hscrShades.Value
    UpdateEffect
End Sub

Private Sub hscrShades_Scroll()
    lblShades.Caption = hscrShades.Value
    UpdateEffect
End Sub

'When a different grayscale algorithm is selected, activate th erelevant options panel (if any) and deactivate all others
Private Sub lstFilters_Click()

    Select Case lstFilters.ListIndex
    
        'Average method
        Case 0
            frameChannel.Visible = False
            frameDecompose.Visible = False
            frameShades.Visible = False
            
        'Adjusted to match cone density of the human eye
        Case 1
            frameChannel.Visible = False
            frameDecompose.Visible = False
            frameShades.Visible = False
            
        'Desaturate
        Case 2
            frameChannel.Visible = False
            frameDecompose.Visible = False
            frameShades.Visible = False
            
        'Decompose (high or low)
        Case 3
            frameChannel.Visible = False
            frameDecompose.Visible = True
            frameDecompose.Top = 168
            frameShades.Visible = False
            
        'Single color channel (user-selectable between red, green, and blue)
        Case 4
            frameDecompose.Visible = False
            frameChannel.Visible = True
            frameChannel.Top = 168
            frameShades.Visible = False
            
        'Specific # of shades of gray (3-255)
        Case 5
            frameChannel.Visible = False
            frameDecompose.Visible = False
            frameShades.Visible = True
            frameShades.Top = 168
            
        'Specific # of shades of gray (3-255) WITH dithering
        Case 6
            frameChannel.Visible = False
            frameDecompose.Visible = False
            frameShades.Visible = True
            frameShades.Top = 168
            
    End Select
    
    UpdateEffect

End Sub

'To load a new image...
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
        
        'De-select all of the current grayscale methods
        lstFilters.ListIndex = -1
        frameChannel.Visible = False
        frameDecompose.Visible = False
        frameShades.Visible = False
        
        'Load the image and refresh the effect
        LoadImageAutosized sFile
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
    
        'Average method
        Case 0
            DrawGrayscaleAverageMethod picBack, picMain
            
        'Adjusted to match cone density of the human eye
        Case 1
            DrawGrayscaleHumanMethod picBack, picMain
            
        'Desaturate
        Case 2
            DrawDesaturate picBack, picMain
            
        'Decompose (high or low)
        Case 3
            DrawGrayscaleDecompose picBack, picMain, optDecompose(0).Value
            
        'Single color channel (user-selectable between red, green, and blue)
        Case 4
            If optChannel.Item(0) Then
                DrawGrayscaleSingleChannel picBack, picMain, 0
            ElseIf optChannel.Item(1) Then
                DrawGrayscaleSingleChannel picBack, picMain, 1
            Else
                DrawGrayscaleSingleChannel picBack, picMain, 2
            End If
            
        'Specific # of shades of gray (3-255)
        Case 5
            DrawGrayscaleCustomShades picBack, picMain, hscrShades.Value
            
        'Specific # of shades of gray (3-255) WITH dithering
        Case 6
            DrawGrayscaleCustomShadesDithered picBack, picMain, hscrShades.Value
            
    End Select
    
End Sub

'This function ensures that a long-type variable falls into the range of 0-255
Public Function ByteMeL(ByVal tempVar As Long) As Byte
    If tempVar > 255 Then
        ByteMeL = 255
    ElseIf tempVar < 0 Then
        ByteMeL = 0
    Else
        ByteMeL = CByte(tempVar)
    End If
End Function

'Return the maximum of three Long-type variables
Private Function Maximum(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
   If (r > g) Then
      If (r > b) Then
         Maximum = r
      Else
         Maximum = b
      End If
   Else
      If (b > g) Then
         Maximum = b
      Else
         Maximum = g
      End If
   End If
End Function

'Return the minimum of three Long-type variables
Private Function Minimum(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
   If (r < g) Then
      If (r < b) Then
         Minimum = r
      Else
         Minimum = b
      End If
   Else
      If (b < g) Then
         Minimum = b
      Else
         Minimum = g
      End If
   End If
End Function

'If one of the Color Channel option buttons are clicked, redraw the image to match
Private Sub optChannel_Click(Index As Integer)
    DrawGrayscaleSingleChannel picBack, picMain, Index
End Sub

'If one of the Decompose option buttons are clicked, redraw the image to match
Private Sub optDecompose_Click(Index As Integer)
    DrawGrayscaleDecompose picBack, picMain, optDecompose(0).Value
End Sub

'Convert an image to grayscale using the (R+G+B)/3 formula
Public Sub DrawGrayscaleAverageMethod(srcPic As PictureBox, dstPic As PictureBox)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Byte, g As Byte, b As Byte
    
    'This look-up table holds all possible totals of adding together the R, G, and B values of an image (0 to 255*3 - for pure white)
    Dim grayLookup(0 To 765) As Byte
    
    'Populate the look-up table
    For x = 0 To 765
        grayLookup(x) = x \ 3
    Next x
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Look up this pixel's value in the lookup table
        gray = grayLookup(CLng(r) + CLng(g) + CLng(b))
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub


'Convert an image to grayscale using a formula based on cone density of the human eye [ (222 * R + 707 * G + 71 * B) \ 1000 ]
Public Sub DrawGrayscaleHumanMethod(srcPic As PictureBox, dstPic As PictureBox)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Look up this pixel's value in the lookup table
        gray = ByteMeL((222 * r + 707 * g + 71 * b) \ 1000)
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert an image to grayscale using the standard luminance formula [ (Max(R,G,B) + Min(R,G,B)) \ 2 ]
Public Sub DrawDesaturate(srcPic As PictureBox, dstPic As PictureBox)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'These variables will hold the maximum and minimum channel values for each pixel
    Dim cMax As Long, cMin As Long
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Look up this pixel's value in the lookup table
        cMax = Maximum(r, g, b)
        cMin = Minimum(r, g, b)
        
        gray = ByteMeL((cMax + cMin) \ 2)
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert an image to grayscale using a channel decompose function (which uses either the minimum [true] or maximum [false] color channel as the gray value)
Public Sub DrawGrayscaleDecompose(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal minValue As Boolean = True)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Assign a gray value based on the low or high value, as determined by the minValue parameter
        If minValue = True Then
            gray = ByteMeL(Minimum(r, g, b))
        Else
            gray = ByteMeL(Maximum(r, g, b))
        End If
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert an image to grayscale using a channel decompose function (where red = 0, green = 1, blue = 2)
Public Sub DrawGrayscaleSingleChannel(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal cChannel As Long = 0)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Assign a gray value based on the red, green, or blue channel, as determined by the cChannel parameter
        If cChannel = 0 Then
            gray = r
        ElseIf cChannel = 1 Then
            gray = g
        Else
            gray = b
        End If
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert an image to a specific number of shades of gray; any value in the range [2,256] is acceptable
Public Sub DrawGrayscaleCustomShades(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Byte, g As Byte, b As Byte
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'This conversionFactor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'This algorithm is well-suited to using a look-up table, so let's build one and (obviously!) prepopulate it
    Dim grayLookup(0 To 255) As Byte
    Dim grayTempCalc As Long
    
    For x = 0 To 255
        grayTempCalc = Int((CDbl(x) / conversionFactor) + 0.5) * conversionFactor
        grayLookup(x) = ByteMeL(grayTempCalc)
    Next x
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Look up this pixel's value in the lookup table
        gray = grayLookup((CLng(r) + CLng(g) + CLng(b)) \ 3)
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

'Convert an image to a specific number of shades of gray (WITH error-diffusion dithering support!)
Public Sub DrawGrayscaleCustomShadesDithered(srcPic As PictureBox, dstPic As PictureBox, Optional ByVal numOfShades As Long = 256)

    'These arrays will hold the source and destination image's pixel data, respectively
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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Byte, g As Byte, b As Byte
    
    'This value will hold the grayscale value of each pixel
    Dim gray As Byte
    
    'This look-up table holds all possible totals of adding together the R, G, and B values of an image (0 to 255*3 - for pure white)
    Dim grayLookup(0 To 765) As Byte
    
    'Populate the look-up table
    For x = 0 To 765
        grayLookup(x) = x \ 3
    Next x
    
    'This conversionFactor is the value we need to turn grayscale values in the [0,255] range into a specific subset of values
    Dim conversionFactor As Single
    conversionFactor = (255 / (numOfShades - 1))
    
    'Unfortunately, this algorithm (unlike its non-dithering counterpart) is not well-suited to using a look-up table, so all calculations have been moved into the loop
    Dim grayTempCalc As Long
    
    'This value tracks the drifting error of our conversions, which allows us to dither
    Dim errorValue As Long
    errorValue = 0
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long
    
    'Note that I have reversed the loop order (now we go horizontally instead of vertically).
    ' This is because I want my dithering algorithm to work from left-to-right instead of top-to-bottom.
    For y = 0 To iHeight - 1
    For x = 0 To iWidth - 1
        quickX = x * 3

        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'First, generate a raw grayscale value
        gray = grayLookup(CLng(r) + CLng(g) + CLng(b))
        grayTempCalc = gray
        
        'Add the error value (a cumulative value of the difference between actual gray values and gray values we've selected) to the current gray value
        grayTempCalc = grayTempCalc + errorValue
        
        'Rebuild our temporary calculation variable using the shade reduction formula
        grayTempCalc = Int((CDbl(grayTempCalc) / conversionFactor) + 0.5) * conversionFactor
        
        'Adjust our error value to include this latest calculation
        errorValue = CLng(gray) + errorValue - grayTempCalc
        
        gray = ByteMeL(grayTempCalc)
        
        'Assign all color channels to the new gray value
        imageData(quickX + 2, y) = gray
        imageData(quickX + 1, y) = gray
        imageData(quickX, y) = gray
        
    Next x
        'Reset our error value after each row
        errorValue = 0
    Next y
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

