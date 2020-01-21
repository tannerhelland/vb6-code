VERSION 5.00
Begin VB.Form frmBrightness 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Brightness Example - Realtime - www.tannerhelland.com"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
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
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hScrollBrightness 
      Height          =   255
      Left            =   2640
      Max             =   255
      Min             =   -255
      TabIndex        =   3
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CheckBox ChkMethod 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Use Stream (Fastest)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   4
      Top             =   120
      Width           =   6030
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   120
      Picture         =   "Brightness.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
      Width           =   6030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Brightness Change (-255 to 255):"
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2415
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

'Temporary brightness variable
Dim tBrightness As Long

'Subroutine for rapidly altering an image's brightness value (includes progress bar effect)
Public Sub DrawBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Long)
    Dim NewColor As Long
    Dim x As Long, y As Long
    'Image data
    Dim iArray() As Byte
    Dim bTable(0 To 255) As Long
    'Build a look-up table of all 256 possible brightness values
    For x = 0 To 255
        NewColor = x + Brightness
        ByteMe NewColor
        bTable(x) = NewColor
    Next x
    'Instantiate a FastDrawing class
    Dim fDraw As New FastDrawing
    'Get the image information
    fDraw.GetImageData SrcPicture, iArray()
    'Temporary width and height variables are faster
    Dim TempWidth As Long, TempHeight As Long
    'I don't know why the width must always be (width - 1); this seems to be a
    'strange byproduct of using DIB sections in VB.  Go figure.
    TempWidth = fDraw.GetImageWidth(SrcPicture) - 1
    TempHeight = fDraw.GetImageHeight(SrcPicture) - 1
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        iArray(2, x, y) = bTable(iArray(2, x, y))
        iArray(1, x, y) = bTable(iArray(1, x, y))
        iArray(0, x, y) = bTable(iArray(0, x, y))
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect)
        If x Mod 25 = 0 Then fDraw.SetImageData DstPicture, fDraw.GetImageWidth(SrcPicture), fDraw.GetImageHeight(SrcPicture), iArray()
    Next x
    'final picture refresh
    fDraw.SetImageData DstPicture, fDraw.GetImageWidth(SrcPicture), fDraw.GetImageHeight(SrcPicture), iArray()
    'free up the memory we borrowed for our image array
    Erase iArray
End Sub

'Subroutine for altering an image's brightness as fast as is possible in VB (using data streams)
Public Sub DrawFastBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Long)
    Dim NewColor As Long
    Dim x As Long
    'Image data
    Dim iArray() As Byte
    Dim bTable(0 To 255) As Long
    'Build a look-up table of all 256 possible brightness values
    For x = 0 To 255
        NewColor = x + Brightness
        ByteMe NewColor
        bTable(x) = NewColor
    Next x
    'Instantiate a FastDrawing class
    Dim fDraw As New FastDrawing
    'Get the image data (in stream format)
    fDraw.GetImageDataStream SrcPicture, iArray()
    'Get the stream length from the image
    Dim StreamLength As Long
    StreamLength = fDraw.GetImageStreamLength(SrcPicture)
    'run a loop through the picture to change every pixel
    For x = 0 To StreamLength
        iArray(x) = bTable(iArray(x))
    Next x
    'Draw the picture
    fDraw.SetImageDataStream DstPicture, fDraw.GetImageWidth(SrcPicture), fDraw.GetImageHeight(SrcPicture), iArray()
    'free up the memory we borrowed for our image array
    Erase iArray
End Sub

'Standard sub for converting to absolute byte values
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255: Exit Sub
    If TempVar < 0 Then TempVar = 0: Exit Sub
End Sub

Private Sub Form_Load()
    Picture2.Picture = Picture1.Picture
End Sub

Private Sub hScrollBrightness_Change()
    'Get the brightness value and send it to the sub
    tBrightness = CLng(hScrollBrightness.Value)
    If ChkMethod.Value = vbChecked Then
        DrawFastBrightness Picture2, Picture1, tBrightness
    Else
        DrawBrightness Picture2, Picture1, tBrightness
    End If
End Sub

Private Sub hScrollBrightness_Scroll()
    'Get the brightness value and send it to the subroutine
    tBrightness = CLng(hScrollBrightness.Value)
    If ChkMethod.Value = vbChecked Then
        DrawFastBrightness Picture2, Picture1, tBrightness
    Else
        DrawBrightness Picture2, Picture1, tBrightness
    End If
End Sub
