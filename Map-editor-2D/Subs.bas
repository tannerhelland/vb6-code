Attribute VB_Name = "Declaration_Module"
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


'The StretchBlt declaration (the GDI32 alternative to VB's PaintPicture)
Public Declare Function StretchBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal ClipX As Long, ByVal ClipY As Long, ByVal RasterOp As Long) As Long

'An array, which we'll use to track our map tile values
Public MapArray() As Byte

'Loop variables
Public x As Long, Y As Long

'The currently selected position
Public CurX As Long, CurY As Long

'The viewport position
Public ViewX As Long, ViewY As Long

'The size of the map
Public SizeX As Integer, SizeY As Integer

'Whether a mouse button is up or down
Public MouseState As Boolean

'The currently selected picture
Public CurPicIndex As Long, CurPicIndex2 As Long

'The zoom value
Public Zoom As Integer

'Whether or not to use StretchBlt in place of PaintPicture
Public UseStretchBlt As Boolean

