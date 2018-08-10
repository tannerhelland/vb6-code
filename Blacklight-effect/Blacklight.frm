VERSION 5.00
Begin VB.Form frmBlacklight 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklight Effect - tannerhelland.com"
   ClientHeight    =   5820
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
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEnable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enable Blacklight Effect"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.HScrollBar hScroll 
      Height          =   255
      Left            =   960
      Max             =   10
      Min             =   1
      TabIndex        =   3
      Top             =   720
      Value           =   2
      Width           =   5175
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
      Top             =   1080
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
      Top             =   1080
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Intensity:"
      BeginProperty Font 
         Name            =   "Segoe UI"
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
      TabIndex        =   2
      Top             =   720
      Width           =   750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmBlacklight"
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

Private Sub chkEnable_Click()

    'Toggle the effect
    If (chkEnable.Value = vbChecked) Then UpdateEffect Else LoadImageAutosized vbNullString
    
End Sub

Private Sub Form_Load()
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'When the blacklight intensity is changed, redraw the image effect (if it's enabled)
Private Sub hScroll_Change()
    UpdateEffect
End Sub

Private Sub hScroll_Scroll()
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
        UpdateEffect
        
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
    minWidth = (hScroll.Left + hScroll.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 2) * Screen.TwipsPerPixelY
    
End Sub

Private Sub UpdateEffect()
    If (chkEnable.Value = vbChecked) Then DrawBlacklight picBack, picMain, hScroll.Value
End Sub

'This is the core Blacklight function
Public Sub DrawBlacklight(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, Optional ByVal fxWeight As Long = 2)

    'Coordinates, dimensions, and pixel data
    Dim x As Long, y As Long, imgWidth As Long, imgHeight As Long, imgPixels() As Byte
    
    'Instantiate a FastDrawing class and gather the image's data (into imgpixels())
    Dim fDraw As New FastDrawing
    imgWidth = fDraw.GetImageWidth(srcPic)
    imgHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imgPixels()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long, l As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim xStride As Long
    
    For y = 0 To imgHeight - 1
    For x = 0 To imgWidth - 1
        
        xStride = x * 3
        
        'Windows DIBs are always in BGR order
        b = imgPixels(xStride, y)
        g = imgPixels(xStride + 1, y)
        r = imgPixels(xStride + 2, y)
        
        'Calculate luminance.  (R + G + B) / 3 would also work.
        l = (222 * r + 707 * g + 71 * b) \ 1000
        
        'The blacklight effect is pretty simple!
        r = Abs(r - l) * fxWeight
        g = Abs(g - l) * fxWeight
        b = Abs(b - l) * fxWeight
        If (r > 255) Then r = 255
        If (g > 255) Then g = 255
        If (b > 255) Then b = 255
        
        'Assign the new pixel values
        imgPixels(xStride, y) = b
        imgPixels(xStride + 1, y) = g
        imgPixels(xStride + 2, y) = r
        
    Next x
    Next y
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, imgWidth, imgHeight, imgPixels()
    
End Sub
