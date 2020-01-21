VERSION 5.00
Begin VB.Form frmBrightness 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Brightness Example - GetPixel/SetPixel - www.tannerhelland.com"
   ClientHeight    =   5655
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
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   StartUpPosition =   2  'CenterScreen
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
      Caption         =   "Change Brightness"
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
      Picture         =   "Brightness2.frx":0000
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

'New API functions
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte

Private Sub ChkAutoRedraw_Click()
    
    'Change the AutoRedraw property of the picture box based on the check box's value
    Picture1.AutoRedraw = (ChkAutoRedraw.Value = vbChecked)
    
End Sub

Private Sub CmdBrightness_Click()
    
    'Get the text value, convert it to a percent of type 'Single,' and send it to the subroutine
    DrawBrightness_API Picture1, Picture1, CSng(Val(TxtBrightness.Text)) / 100!
    
End Sub

'A simple subroutine that will change the brightness of a picturebox using nothing but PSet and Point.
Public Sub DrawBrightness_API(ByRef dstPicture As PictureBox, ByRef srcPicture As PictureBox, ByVal newBrightness As Single)
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Build a look-up table for all possible brightness values
    Dim bTable(0 To 255) As Long
    Dim tmpColor As Long
    For x = 0 To 255
        
        'Calculate the brightness for pixel value x
        tmpColor = x * newBrightness
        
        'Make sure that the calculated value is between 0 and 255 (so we don't get an error)
        bTable(x) = ByteMe(tmpColor)
        
    Next x
    
    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim tmpWidth As Long, tmpHeight As Long
    tmpWidth = dstPicture.ScaleWidth - 1
    tmpHeight = dstPicture.ScaleHeight - 1
    
    'run a loop through the picture to change every pixel
    For x = 0 To tmpWidth
    For y = 0 To tmpHeight
        
        'Get the color (using GetPixel) and extract the red, green, and blue values
        tmpColor = GetPixel(srcPicture.hDC, x, y)
        r = ExtractR(tmpColor)
        g = ExtractG(tmpColor)
        b = ExtractB(tmpColor)
        
        'Use the values in the look-up table to quickly change the brightness values
        'of the selected colors.  The look-up table is much faster than doing the math
        'over and over for each individual pixel.
        r = bTable(r)
        g = bTable(g)
        b = bTable(b)
        
        'Now set that data using the SetPixelV command
        SetPixelV dstPicture.hDC, x, y, RGB(r, g, b)
        
    Next y
        
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        If ((x Mod 25) = 0) Then
            If dstPicture.AutoRedraw Then dstPicture.Refresh
        End If
        
    Next x
    
    'final picture refresh
    If dstPicture.AutoRedraw Then
        Set dstPicture.Picture = dstPicture.Image
        dstPicture.Refresh
    End If
    
End Sub

'Standardized routines for color extraction
Public Function ExtractR(ByVal srcColor As Long) As Byte
    ExtractR = srcColor And &HFF&
End Function

Public Function ExtractG(ByVal srcColor As Long) As Byte
    ExtractG = (srcColor \ 256) And &HFF&
End Function

Public Function ExtractB(ByVal srcColor As Long) As Byte
    ExtractB = (srcColor \ 65536) And &HFF&
End Function

'Standardized routine for converting to absolute byte values
Public Function ByteMe(ByVal srcValue As Long) As Long
    If (srcValue > 255) Then srcValue = 255
    If (srcValue < 0) Then srcValue = 0
    ByteMe = srcValue
End Function

