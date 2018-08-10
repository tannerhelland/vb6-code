VERSION 5.00
Begin VB.Form frmDiffuse 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diffuse Image Filter - www.tannerhelland.com"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12330
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
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   822
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDrawExplosion 
      Caption         =   "Animate ""Explosion"""
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
      Left            =   9600
      TabIndex        =   10
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CheckBox chkWrap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Wrap pixels around image edges?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3855
   End
   Begin VB.HScrollBar hScrollY 
      Height          =   255
      Left            =   2520
      Max             =   600
      TabIndex        =   5
      Top             =   720
      Width           =   6855
   End
   Begin VB.HScrollBar hScrollX 
      Height          =   255
      Left            =   2520
      Max             =   800
      TabIndex        =   4
      Top             =   360
      Value           =   20
      Width           =   6855
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
      Height          =   615
      Left            =   9600
      TabIndex        =   2
      Top             =   240
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
      Left            =   120
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   0
      Top             =   1680
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
      Left            =   120
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   12030
   End
   Begin VB.Label lblY 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 px"
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
      Height          =   210
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "20 px"
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
      Height          =   210
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical distance: "
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
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal distance: "
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
      Top             =   360
      Width           =   1650
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmDiffuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Diffuse/Spread Image Filter by Tanner Helland (published 2011)
'
'http://www.tannerhelland.com
'
'This project demonstrates a full implementation of a "Diffuse" (or "Spread" if you use GIMP) image filter.
' A diffuse filter works by randomly rearranging pixels in an image.  The maximum distance that pixels can
' be moved is specified by the user, and a wrap option is provided for image edges.  If wrap is set to true,
' any pixels near the edge of the image that are selected for randomization can be moved to the opposite
' side of the image, or "wrapped around."
'
'For fun, I've included an animation option that runs a loop of increasingly large diffusion distances.
' It makes for a fun kind of explosion effect, and it reminds me of how some old 8-bit games used to look.
'
'Enjoy the project, and be sure to let me know if you found the code useful.
'
'Also, if you enjoy free VB game and graphics code, subscribe to my RSS feed at
' http://www.tannerhelland.com/feed/

Option Explicit

Private Sub chkWrap_Click()
    ProcessDiffuse
End Sub

'This routine will draw an animated explosion (note: it may momentarily lock up an underpowered computer - so only do this after compilation, or consider yourself warned!)
Private Sub cmdDrawExplosion_Click()

    Dim x As Long
    
    For x = 0 To 200 Step 5
        DrawDiffuse picBack, picMain, x, x, False
    Next x

End Sub

'When the program is first loaded, several things needs to happen...
Private Sub Form_Load()

    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    ProcessDiffuse True
    Me.Show
    
End Sub

'Apply a diffuse filter to an image using user-specified distances and, optionally, a paramater that defines wrapping around image edges
Public Sub DrawDiffuse(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, ByVal xDist As Single, ByVal yDist As Single, Optional ByVal wrapPixels As Boolean = False)

    'These arrays will hold the source and destination image's pixel data, respectively
    Dim srcImageData() As Byte, dstImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, srcImageData()
    fDraw.GetImageData2D srcPic, dstImageData()
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Byte, g As Byte, b As Byte
    
    'Let's do simple checking to make sure the hue and saturation values we were passed don't exceed the
    ' width of the source image
    If xDist < 0 Then xDist = 0
    If xDist > iWidth Then xDist = iWidth
    If yDist < 0 Then yDist = 0
    If yDist > iHeight Then yDist = iHeight
    
    'These values will be used to calculate the diffused pixels
    Dim halfDX As Long, halfDY As Long
    halfDX = xDist / 2
    halfDY = yDist / 2
    
    Dim diffuseX As Long, diffuseY As Long
    Dim dstX As Long, dstY As Long
    
    'Seed the random number generator with a unique value
    Randomize Timer
    
    'Loop through the image, adjusting pixel values as we go
    Dim quickX As Long, quickX2 As Long
    
    For x = 0 To iWidth - 1
    For y = 0 To iHeight - 1
        
        'To calculate a new pixel position, we add a random value between 0 and the specified max distance.
        ' Then we subtract half that value, so the pixel can move to the right OR left (up AND down).
        diffuseX = Rnd * xDist - halfDX
        diffuseY = Rnd * yDist - halfDY
        dstX = x + diffuseX
        dstY = y + diffuseY
        
        'If wrapPixels is TRUE, wrap this pixel around to the other side (if necessary)
        If wrapPixels Then
            If dstX < 0 Then dstX = iWidth - 1 + dstX
            If dstY < 0 Then dstY = iHeight - 1 + dstY
            If dstX > iWidth - 1 Then dstX = dstX - iWidth + 1
            If dstY > iHeight - 1 Then dstY = dstY - iHeight + 1
        
        'If wrapPixels is FALSE, then check to make sure this pixel falls within image boundaries
        Else
            If dstX < 0 Then dstX = 0
            If dstY < 0 Then dstY = 0
            If dstX > iWidth - 1 Then dstX = iWidth - 1
            If dstY > iHeight - 1 Then dstY = iHeight - 1
        End If
    
        quickX = x * 3
        quickX2 = dstX * 3
        
        'Swap the pixels at these two locations
        b = srcImageData(quickX2, dstY)
        g = srcImageData(quickX2 + 1, dstY)
        r = srcImageData(quickX2 + 2, dstY)
        
        dstImageData(quickX2, dstY) = srcImageData(quickX, y)
        dstImageData(quickX2 + 1, dstY) = srcImageData(quickX + 1, y)
        dstImageData(quickX2 + 2, dstY) = srcImageData(quickX + 2, y)
        
        dstImageData(quickX, y) = b
        dstImageData(quickX + 1, y) = g
        dstImageData(quickX + 2, y) = r
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, dstImageData()

End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'When the scroll bars are changed, make the associated labels match and redraw the image
Private Sub hScrollX_Change()
    lblX.Caption = hScrollX.Value & " px"
    ProcessDiffuse
End Sub

Private Sub hScrollX_Scroll()
    lblX.Caption = hScrollX.Value & " px"
    ProcessDiffuse
End Sub

Private Sub hScrollY_Change()
    lblY.Caption = hScrollY.Value & " px"
    ProcessDiffuse
End Sub

Private Sub hScrollY_Scroll()
    lblY.Caption = hScrollY.Value & " px"
    ProcessDiffuse
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
        ProcessDiffuse True
        
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
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.TOp * 1.75) * Screen.TwipsPerPixelY
    
End Sub

'If the user changes any of the available options, this routine will be triggered.
' All it does is call the diffuse function with proper values.
Private Sub ProcessDiffuse(Optional ByVal updateScrollLimits As Boolean = False)

    If updateScrollLimits Then
        hScrollX.Min = 0
        hScrollX.Max = picBack.ScaleWidth
        hScrollY.Min = 0
        hScrollY.Max = picBack.ScaleHeight
    End If
    
    If (chkWrap.Value = vbChecked) Then
        DrawDiffuse picBack, picMain, hScrollX.Value, hScrollY.Value, True
    Else
        DrawDiffuse picBack, picMain, hScrollX.Value, hScrollY.Value, False
    End If
    
End Sub
