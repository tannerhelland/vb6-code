VERSION 5.00
Begin VB.Form frmContrast 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Image contrast - tannerhelland.com"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hScrContrast 
      Height          =   255
      Left            =   120
      Max             =   200
      Min             =   -100
      TabIndex        =   2
      Top             =   480
      Width           =   6015
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   3
      Top             =   1200
      Width           =   6030
   End
   Begin VB.PictureBox PicBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "| - zero"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrast adjustment (-100% to 200%):"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpenImage 
         Caption         =   "&Open image..."
      End
   End
End
Attribute VB_Name = "frmContrast"
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
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
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
        
        'Load the image and refresh the effect
        LoadImageAutosized sFile
        UpdateEffect
        
    End If
    
End Sub

'If a null string is passed, simply paint an unmodified copy of the backbuffer into the foreground buffer
Private Sub LoadImageAutosized(Optional ByVal srcFilePath As String = vbNullString)

    If (LenB(srcFilePath) <> 0) Then PicBack.Picture = LoadPicture(srcFilePath)
        
    'Copy the image, automatically resized, from the background picture box to the foreground one
    Dim fDraw As FastDrawing
    Set fDraw = New FastDrawing
    
    Dim ImageData() As Byte, imgWidth As Long, imgHeight As Long
    imgWidth = fDraw.GetImageWidth(Me.PicBack)
    imgHeight = fDraw.GetImageHeight(Me.PicBack)
    fDraw.GetImageData2D Me.PicBack, ImageData()
    
    Me.PicMain.Width = imgWidth + 2
    Me.PicMain.Height = imgHeight + 2
    fDraw.SetImageData2D Me.PicMain, imgWidth, imgHeight, ImageData()
    
    'Resize the form to automatically contain the new picture box dimensions.
    ' (This performs some sloppy checks to keep the form from resizing larger than the primary display,
    '  but it is *not* good code - a proper solution would involve the AdjustWindowRect() API or similar!)
    Dim newWidth As Long, minWidth As Long
    minWidth = (hScrContrast.Left + hScrContrast.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + PicMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + PicMain.Top * 2) * Screen.TwipsPerPixelX
    
End Sub

Private Sub hScrContrast_Change()
    UpdateEffect
End Sub

Private Sub hScrContrast_Scroll()
    UpdateEffect
End Sub

Private Sub UpdateEffect()
    ModifyContrast PicBack, PicMain, CLng(hScrContrast.Value)
End Sub

'To keep this subroutine simple, the source and destination picture boxes are assumed to be the same size.
' If they are NOT the same size, you will need to use the minimum bounds (width and height) to avoid errors.
Public Sub ModifyContrast(ByRef srcPicture As PictureBox, ByRef dstPicture As PictureBox, ByVal contrast As Long)

    'Contrast uses a very simple formula:
    ' - Positive contrast pushes values away from the midtone (127).
    ' - Negative contrast pushes values toward the midtone (127)
    '
    'Because a universal adjustment is applied to each channel and pixel in the image, we can use a lookup table
    ' to greatly improve performance.
    Dim contrastTable() As Byte
    ReDim contrastTable(0 To 255) As Byte
    
    Dim x As Long, y As Long
    Dim colorCalculation As Long
    
    For x = 0 To 255
        
        'This line contains the formula for basic contrast correction.  We will calculate contrast for each possible input value
        ' (0 to 255) and store it to a table.  Then we can use this table to quickly modify the entire image.
        colorCalculation = x + (((x - 127) * contrast) \ 100)
        
        'Clamp values to 0, 255
        If colorCalculation > 255 Then
            colorCalculation = 255
        ElseIf colorCalculation < 0 Then
            colorCalculation = 0
        End If
        
        'Store this value in the lookup table
        contrastTable(x) = colorCalculation
        
    Next x
    
    'Use the FastDrawing class to retrieve the image's pixels into a standard VB array
    Dim fDraw As FastDrawing
    Set fDraw = New FastDrawing
    
    Dim imageWidth As Long, imageHeight As Long
    imageWidth = fDraw.GetImageWidth(srcPicture)
    imageHeight = fDraw.GetImageHeight(dstPicture)
    
    Dim imagePixels() As Byte
    fDraw.GetImageData2D srcPicture, imagePixels
    
    'Now all we have to do is pass each channel value in the image through our lookup table
    Dim xStride As Long
    
    For x = 0 To imageWidth - 1
        xStride = x * 3
    For y = 0 To imageHeight - 1
    
        'Modify each of red, green, and blue
        imagePixels(xStride, y) = contrastTable(imagePixels(xStride, y))
        imagePixels(xStride + 1, y) = contrastTable(imagePixels(xStride + 1, y))
        imagePixels(xStride + 2, y) = contrastTable(imagePixels(xStride + 2, y))
        
    Next y
    Next x

    'Paint the modified pixels onto the destination picture box
    fDraw.SetImageData2D dstPicture, srcPicture.ScaleWidth, srcPicture.ScaleHeight, imagePixels
    
End Sub
