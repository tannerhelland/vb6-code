VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Real-time drawing class for VB6
'Published in 2011 by Tanner Helland
'
'This class was built to simplify per-pixel access in VB6.  The attached project demonstrates how to use it.
' Note that its simplified design is primarily based around modifying 24-bpp RGB data.  Alpha channels are
' not currently supported (although the code could be easily modified for this).
'
'The source code in this project is licensed under a Simplified BSD license.
' For more information, please review LICENSE.md at https://github.com/tannerhelland/thdc-code/
'
'If you find this code useful, please consider a small donation to https://www.paypal.me/TannerHelland
'
'***************************************************************************

Option Explicit

'Stripped down bitmap information
Private Type GDI_Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Retrieve object properties directly from Windows; this lets us bypass some VB-specific annoyances
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'Standard pixel data (Windows BGRA order)
Private Type GDI_RGBQuad
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type

'Full-size bitmap header
Private Type GDI_BitmapInfoHeader
    bmSize As Long
    bmWidth As Long
    bmHeight As Long
    bmPlanes As Integer
    bmBitCount As Integer
    bmCompression As Long
    bmSizeImage As Long
    bmXPelsPerMeter As Long
    bmYPelsPerMeter As Long
    bmClrUsed As Long
    bmClrImportant As Long
End Type

'Extended header for 8-bit images
Private Type GDI_BitmapInfo
    bmHeader As GDI_BitmapInfoHeader
    bmColors(0 To 255) As GDI_RGBQuad
End Type

'Higher-quality image stretching must be activated manually
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

'DIB section interfaces
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As GDI_BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As GDI_BitmapInfo, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Return an image's width, in pixels
Public Function GetImageWidth(ByRef srcPictureBox As PictureBox) As Long
    Dim bm As GDI_Bitmap
    GetObject srcPictureBox.Image, Len(bm), bm
    GetImageWidth = bm.bmWidth
End Function

'Return an image's height, in pixels
Public Function GetImageHeight(ByRef srcPictureBox As PictureBox) As Long
    Dim bm As GDI_Bitmap
    GetObject srcPictureBox.Image, Len(bm), bm
    GetImageHeight = Abs(bm.bmHeight)
End Function

'Copy an 24-bpp image's pixel information into an array dimensioned (x * 3 , y).  Each row of the array represents
' Blue, Green, Red pixel values - e.g. BGRBGRBGR...
'
'Also, because Windows DIBs are stored upside-down (last line first), you can ask this function to reverse
' line-order for you.  This can make position-specific calculations more intuitive.
Public Sub GetImageData2D(ByRef srcPictureBox As PictureBox, ByRef dstPixelData() As Byte, Optional ByVal fixOrientation As Boolean = False)
    
    'To avoid dealing with .ScaleMode issues, retrieve the underlying image dimensions via WAPI
    Dim bm As GDI_Bitmap
    GetObject srcPictureBox.Image, Len(bm), bm
    
    'Calculate a size for the destination array.  For performance reasons, Windows requires
    ' 24-bpp DIB scanlines to be aligned on 4-byte boundaries - this means that the destination
    ' array's width may not be a multiple of (imageWidth * 3), by design.
    Dim imgStride As Long
    imgStride = (bm.bmWidth * 3) - 1 + (bm.bmWidth Mod 4)
    ReDim dstPixelData(0 To imgStride, 0 To bm.bmHeight - 1) As Byte
    
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As GDI_BitmapInfo
    
    With bmi.bmHeader
        .bmWidth = bm.bmWidth
        .bmHeight = bm.bmHeight
        If fixOrientation Then .bmHeight = -1 * .bmHeight
        .bmSize = 40                'Size, in bytes, of the header
        .bmPlanes = 1               'Number of planes (always one for this instance)
        .bmBitCount = 24            'Bits per pixel (always 24 for this instance)
        .bmCompression = 0          'Compression (always none for this instance)
    End With
    
    'Copy the pixel data into our array
    GetDIBits srcPictureBox.hDC, srcPictureBox.Image, 0, bm.bmHeight, dstPixelData(0, 0), bmi, 0
    
End Sub

'Set an image's pixel information from a raw byte array.  The byte array *must* have been created
' according to the rules defined in GetImageData2D(), above.  Similarly, if you reversed scanline
' order at retrieval time, you should also pass TRUE for fixOrientation when calling this function.
'
'The image will be stretched to fit the destination picture box (without regard to aspect ratio).
' You can manually modify this by changing the destination parameters of the StretchDIBits call.
Public Sub SetImageData2D(ByRef dstPictureBox As PictureBox, ByVal imgWidth As Long, ByVal imgHeight As Long, ByRef srcPixelData() As Byte, Optional ByVal fixOrientation As Boolean = False)
    
    'To avoid dealing with .ScaleMode issues, retrieve the underlying image dimensions via WAPI
    Dim bm As GDI_Bitmap
    GetObject dstPictureBox.Image, Len(bm), bm
    
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As GDI_BitmapInfo
    
    With bmi.bmHeader
        .bmWidth = imgWidth
        .bmHeight = imgHeight
        If fixOrientation Then .bmHeight = -1 * .bmHeight
        .bmSize = 40                'Size, in bytes, of the header
        .bmPlanes = 1               'Number of planes (always one for this instance)
        .bmBitCount = 24            'Bits per pixel (always 24 for this instance)
        .bmCompression = 0          'Compression (always none for this instance)
    End With
    
    'If stretching will be involved, request high-quality stretching from GDI
    SetStretchBltMode dstPictureBox.hDC, 3&
    
    'Paint the pixel data into the destination picture box.  By using bm.bmWidth and bm.bmHeight as
    ' the destination width/height, the pixel data will be automatically resized to fit the picture box.
    StretchDIBits dstPictureBox.hDC, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, imgWidth, imgHeight, srcPixelData(0, 0), bmi, 0, vbSrcCopy
    
    'VB's built-in backbuffer system ("AutoRedraw") won't detect changes made by WAPI functions,
    ' so we need to manually notify it of the need to redraw.
    If dstPictureBox.AutoRedraw Then
        dstPictureBox.Picture = dstPictureBox.Image
        dstPictureBox.Refresh
    End If
    
End Sub
