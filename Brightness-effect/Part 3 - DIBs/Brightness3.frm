VERSION 5.00
Begin VB.Form frmBrightness 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Brightness Example - DIB sections - www.tannerhelland.com"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
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
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBrightnessBB 
      Caption         =   "Change Brightness using BitmapBits (USE THIS ONLY IN 24/32-BIT COLOR MODES)"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   2895
   End
   Begin VB.CheckBox ChkAutoRedraw 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "AutoRedraw"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   3120
      TabIndex        =   4
      Top             =   5280
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox TxtBrightness 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Text            =   "150"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton CmdBrightness 
      Caption         =   "Change Brightness using DIB Sections"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "Brightness3.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6030
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Brightness Change (%):"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   4920
      Width           =   1815
   End
End
Attribute VB_Name = "frmBrightness"
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


'All of the DIB types
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type
 
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
 
Private Type BITMAPINFO
    bmHeader As BITMAPINFOHEADER
    bmColors(0 To 255) As RGBQUAD
End Type

'The GetObject API call gives us the bitmap variables we need for the other API calls
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

'The GetBitmapBits and SetBitmapBits API calls (use ONLY in 24/32-bit color mode!!)
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

'The magical API DIB function calls (they're long!)
Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long

'The array that will hold our pixel data
Dim ImageData() As Byte

'Temporary brightness variable
Dim tBrightness As Single

Private Sub ChkAutoRedraw_Click()
    'Change the AutoRedraw property of the picture box based on the check box's value
    If ChkAutoRedraw.Value = vbChecked Then Picture1.AutoRedraw = True Else Picture1.AutoRedraw = False
End Sub

Private Sub CmdBrightness_Click()
    'Get the text value, convert it to type 'Single,' and send it to the sub
    tBrightness = CSng(Val(TxtBrightness)) / 100
    DrawDIBBrightness Picture1, Picture1, tBrightness
End Sub

'A simple subroutine that will change the brightness of a picturebox using DIB sections.
Public Sub DrawDIBBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Single)
    'Coordinate variables
    Dim x As Long, y As Long
    'Build a look-up table for all possible brightness values
    Dim bTable(0 To 255) As Long
    Dim TempColor As Long
    For x = 0 To 255
        'Calculate the brightness for pixel value x
        TempColor = Int(CSng(x) * Brightness)
        'Make sure that the calculated value is between 0 and 255 (so we don't get an error)
        ByteMe TempColor
        'Place the corrected value into its array spot
        bTable(x) = TempColor
    Next x
    'Get the pixel data into our ImageData array
    GetImageData SrcPicture, ImageData()
    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim TempWidth As Long, TempHeight As Long
    TempWidth = DstPicture.ScaleWidth - 1
    TempHeight = DstPicture.ScaleHeight - 1
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        'Use the values in the look-up table to quickly change the brightness values
        'of each color.  The look-up table is much faster than doing the math
        'over and over for each individual pixel.
        ImageData(2, x, y) = bTable(ImageData(2, x, y))   'Change the red
        ImageData(1, x, y) = bTable(ImageData(1, x, y))   'Change the green
        ImageData(0, x, y) = bTable(ImageData(0, x, y))   'Change the blue
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then SetImageData DstPicture, ImageData()
    Next x
    'final picture refresh
    SetImageData DstPicture, ImageData()
End Sub

'Routine to get an image's pixel information into an array dimensioned (rgb, x, y)
Public Sub GetImageData(ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from SrcPictureBox and put it into our 'bm' variable
    GetObject SrcPictureBox.Image, bmLen, bm
    'Build a correctly sized array
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    'Finish building the 'bmi' variable we want to pass to the GetDIBits call (the same one we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've completely filled up the 'bmi' variable, we use GetDIBits to take the data from
    'SrcPictureBox and put it into the ImageData() array using the settings we specified in 'bmi'
    GetDIBits SrcPictureBox.hDC, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub

'Routine to set an image's pixel information from an array dimensioned (rgb, x, y)
Public Sub SetImageData(ByRef DstPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from DstPictureBox and put it into our 'bm' variable
    GetObject DstPictureBox.Image, bmLen, bm
    'Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call
    '(continuing to use the 'bmi' we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've built the temporary header, we use StretchDIBits to take the data from the
    'ImageData() array and put it into SrcPictureBox using the settings specified in 'bmi' (the
    'StretchDIBits call should be on one continuous line)
    StretchDIBits DstPictureBox.hDC, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    'Note: Always set AutoRedraw to true when using DIB sections; when AutoRedraw is false
    'you will get unpredictable results.
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
End Sub

'Standardized routine for converting to absolute byte values
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255: Exit Sub
    If TempVar < 0 Then TempVar = 0: Exit Sub
End Sub

Private Sub CmdBrightnessBB_Click()
    'Get the text value, convert it to type 'Single,' and send it to the sub
    tBrightness = CSng(Val(TxtBrightness)) / 100
    DrawBitmapBitsBrightness Picture1, Picture1, tBrightness
End Sub

'A subroutine for changing the brightness of a picturebox IN 24/32-BIT COLOR MODES ONLY!!
Public Sub DrawBitmapBitsBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Single)
    'Coordinate variables
    Dim x As Long, y As Long
    'Build a look-up table for all possible brightness values
    Dim bTable(0 To 255) As Long
    Dim TempColor As Long
    For x = 0 To 255
        'Calculate the brightness for pixel value x
        TempColor = Int(CSng(x) * Brightness)
        'Make sure that the calculated value is between 0 and 255 (so we don't get an error)
        ByteMe TempColor
        'Place the corrected value into its array spot
        bTable(x) = TempColor
    Next x
    'Create a bitmap variable and copy the basic information from 'PictureBox.Image' into it
    Dim bm As BITMAP
    GetObject DstPicture.Image, Len(bm), bm
    'Create an array of bytes and fill it with the information from 'bm' (i.e. PictureBox.image)
    Dim ImageData() As Byte
    ReDim ImageData(0 To (bm.bmBitsPixel \ 8) - 1, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    GetBitmapBits DstPicture.Image, bm.bmWidthBytes * bm.bmHeight, ImageData(0, 0, 0)

    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim TempWidth As Long, TempHeight As Long
    TempWidth = DstPicture.ScaleWidth - 1
    TempHeight = DstPicture.ScaleHeight - 1
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        'Use the values in the look-up table to quickly change the brightness values
        'of each color.  The look-up table is much faster than doing the math
        'over and over for each individual pixel.
        ImageData(2, x, y) = bTable(ImageData(2, x, y))   'Change the red
        ImageData(1, x, y) = bTable(ImageData(1, x, y))   'Change the green
        ImageData(0, x, y) = bTable(ImageData(0, x, y))   'Change the blue
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then
            SetBitmapBits DstPicture.Image, bm.bmWidthBytes * bm.bmHeight, ImageData(0, 0, 0)
            DstPicture.Picture = DstPicture.Image
            DstPicture.Refresh
        End If
    Next x
    'final picture refresh
    SetBitmapBits DstPicture.Image, bm.bmWidthBytes * bm.bmHeight, ImageData(0, 0, 0)
    DstPicture.Picture = DstPicture.Image
    DstPicture.Refresh
End Sub

