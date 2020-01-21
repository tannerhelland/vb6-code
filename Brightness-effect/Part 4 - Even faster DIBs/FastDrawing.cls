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
Option Explicit
'**************************************************************************************************************************************
'Real-time drawing class for VB6
'Published in 2011 by Tanner Helland
'
'This class is every graphics programmers dream - it does all the dirty API work required for real-time
' graphics, and all you have to do is call a few very simple routines.  The format is simple; the only
' variables required are the picture boxes you want to process, and an array to hold the DIB information.
' Many comments are included, though the routines should be self-explanatory.
'
'This source code has been released under a BSD license. You may read more about this license at:
' http://creativecommons.org/licenses/BSD/
'
'If you would like to make a donation to help offset the cost of maintaining this code and the site on which it resides, please visit:
' http://www.tannerhelland.com/donate/
'
'While no payment will ever be required to use or distribute this source code, donations are GREATLY appreciated.
'
'Comments and questions regarding this source code can be submitted via this contact form:
' http://www.tannerhelland.com/contact/
'
'**************************************************************************************************************************************

'Stripped down bitmap information
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Call to transfer an object's properties into a custom variable
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'Standard pixel data
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbAlpha As Byte
End Type

'Full-size bitmap header
Private Type BITMAPINFOHEADER
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
Private Type BITMAPINFO
        bmHeader As BITMAPINFOHEADER
        bmColors(0 To 255) As RGBQUAD
End Type

'Used to ensure quality stretching of color images
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'DIB section interfaces
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Get the image width (via API - always accurate, unlike PictureBox.ScaleWidth)
Public Function GetImageWidth(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageWidth = bm.bmWidth
End Function

'Get the image height (via API - always accurate)
Public Function GetImageHeight(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageHeight = bm.bmHeight
End Function

'Get the stream length of an image (via API - always accurate)
Public Function GetImageStreamLength(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageStreamLength = (bm.bmWidth * (bm.bmHeight + 1)) * 3
End Function

'Get an image's pixel information into an array dimensioned (x * 3 + bgr, y), with the option to get it in its true orientation
Public Sub GetImageData2D(SrcPictureBox As PictureBox, ImageData() As Byte, Optional ByVal CorrectOrientation As Boolean = False)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array
    Erase ImageData()
    'Generate a correctly-dimensioned array (for 2-dimensional access)
    Dim ArrayWidth As Long
    ArrayWidth = (bm.bmWidth * 3) - 1
    ArrayWidth = ArrayWidth + (bm.bmWidth Mod 4)  '4-bit alignment
    ReDim ImageData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    ReDim tmpData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    
    'Get the image data into our array
    If CorrectOrientation = False Then
        GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0
    Else
        GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, tmpData(0, 0), bmi, 0
    End If
    
    'This code is to orient the image data correctly in the array (i.e. (0,0) as top-left, (max,max) as bottom right)
    ' (if this option is enabled, we must set the DIB height to negative in the SetImageData routine below)
    If CorrectOrientation = True Then
    
        Dim x As Long, y As Long, z As Long
        Dim QuickVal As Long
        For x = 0 To bm.bmWidth - 1
            QuickVal = x * 3
         For y = 0 To bm.bmHeight - 1
          For z = 0 To 2
            ImageData(QuickVal + z, y) = tmpData(QuickVal + z, bm.bmHeight - y)
          Next z
         Next y
        Next x
        
    End If
    
    'Save memory...?
    Erase tmpData

End Sub

'Set an image's pixel information from an array dimensioned (x * 3 + bgr, y)
Public Sub SetImageData2D(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte, Optional ByVal CorrectOrientation As Boolean = False)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    If CorrectOrientation = False Then
        bmi.bmHeader.bmHeight = OriginalHeight
    Else
        bmi.bmHeader.bmHeight = -OriginalHeight
    End If
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Assume color images and set the corresponding best stretch mode
    SetStretchBltMode DstPictureBox.hdc, 3&
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

'Get an image's pixel information into an array dimensioned (r/g/b, x, y)
Public Sub GetImageData(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array
    Erase ImageData()
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Get the image data into our array
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub

'Set an image's pixel information from an array dimensioned (r/g/b, x, y)
Public Sub SetImageData(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Assume color images and set the corresponding best stretch mode
    SetStretchBltMode DstPictureBox.hdc, 3&
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

'Get an image's pixel data into a one-dimesional array (stream)
Public Sub GetImageDataStream(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array - in this case, designed as a stream
    Erase ImageData()
    ReDim ImageData(0 To GetImageStreamLength(SrcPictureBox))
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Get the image data into our array
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0), bmi, 0
End Sub

'Set an image's data from a one-dimensional array (stream)
Public Sub SetImageDataStream(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

