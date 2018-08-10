VERSION 5.00
Begin VB.Form frmThreshold 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Threshold Filter - www.tannerhelland.com"
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
   Begin VB.CheckBox chkWhite 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Force pixels above threshold to white?"
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
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.HScrollBar hscrThreshold 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   4
      Top             =   720
      Value           =   127
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
      Top             =   2520
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
   Begin VB.Label lblBlackout 
      BackStyle       =   0  'Transparent
      Caption         =   "Black-out pixels below this luminance:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblThreshold 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "127"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The above image (from The Secret of Monkey Island: Special Edition) is ©2009 LucasArts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   7680
      Width           =   12015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmThreshold"
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

'When the checkbox is changed, re-apply the filter
Private Sub chkWhite_Click()
    UpdateEffect
End Sub

'When the program is first loaded, several things needs to happen...
Private Sub Form_Load()

    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
        
End Sub

Private Sub UpdateEffect()
    DrawThreshold picBack, picMain, hscrThreshold.Value, (chkWhite.Value = vbChecked)
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'When the percentage scroll bar is changed, make the associated label display the current value (and redraw the image)
Private Sub hscrThreshold_Change()
    lblThreshold.Caption = hscrThreshold.Value
    UpdateEffect
End Sub

Private Sub hscrThreshold_Scroll()
    lblThreshold.Caption = hscrThreshold.Value
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

'Automatically black-out pixels below a certain threshold; parameters are source picture box, destination picture box, blackout threshold (value between 0 and 255)
Public Sub DrawThreshold(srcPic As PictureBox, dstPic As PictureBox, ByVal blackOutValue As Long, Optional forceToWhite As Boolean = False)

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
    
    'These variables will hold temporary pixel color values (red, green, blue)
    Dim r As Long, g As Long, b As Long
    
    'This value will be used to calculate a temporary grayscale value for each pixel
    Dim gray As Long
    
    'Initialize a grayscale look-up table
    Dim gLookup(0 To 765) As Long
    For x = 0 To 765
        gLookup(x) = x \ 3
    Next x
        
    'Now, loop through each pixel, blacking it out (or ignoring it) according to the user's specified threshold value
    Dim quickX As Long
    
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        
        'Get the source image pixels
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        
        'Calculate luminance
        gray = gLookup(r + g + b)
        
        'If below the threshold, assign all color channels to their new values based on the look-up table created above
        If gray <= blackOutValue Then
            imageData(quickX + 2, y) = 0
            imageData(quickX + 1, y) = 0
            imageData(quickX, y) = 0
        Else
            'If forceToWhite is "true", set this pixel to white
            If forceToWhite Then
                imageData(quickX + 2, y) = 255
                imageData(quickX + 1, y) = 255
                imageData(quickX, y) = 255
            End If
        End If
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, imageData()

End Sub

