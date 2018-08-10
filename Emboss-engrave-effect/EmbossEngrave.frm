VERSION 5.00
Begin VB.Form frmEmbossEngrave 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emboss / Engrave Image Filters - www.tannerhelland.com"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9615
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
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optRelief 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Relief"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5400
      MousePointer    =   99  'Custom
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "to Color..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton optEngrave 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Engrave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton optEmboss 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Emboss"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Reset the Image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   2655
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
      Top             =   720
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
      Height          =   5280
      Left            =   120
      ScaleHeight     =   350
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   620
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   9330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmEmbossEngrave"
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

'Use the currently selected color to generate an "embossed" or "engraved" version of the images
Private Sub chkColor_Click()
    UpdateEffect
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'When the program is first loaded, copy the back buffer onto the front buffer
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
    Dim newWidth As Long, minWidth As Long
    minWidth = (cmdReset.Left + cmdReset.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 2.5) * Screen.TwipsPerPixelY
    
End Sub

Private Sub UpdateEffect()
    If optEmboss.Value Then
        If (chkColor.Value = vbChecked) Then
            DrawEmboss picBack, picMain, ExtractR(picColor.BackColor), ExtractG(picColor.BackColor), ExtractB(picColor.BackColor)
        Else
            DrawEmboss picBack, picMain
        End If
    ElseIf optEngrave.Value Then
        If (chkColor.Value = vbChecked) Then
            DrawEngrave picBack, picMain, ExtractR(picColor.BackColor), ExtractG(picColor.BackColor), ExtractB(picColor.BackColor)
        Else
            DrawEngrave picBack, picMain
        End If
    ElseIf optRelief.Value Then
        DrawRelief picBack, picMain
    End If
End Sub

'Apply a generic grayscale "Engrave" filter to an image
Public Sub DrawEngrave(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, Optional ByVal eR As Byte = 127, Optional ByVal eG As Byte = 127, Optional ByVal eB As Byte = 127)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'This array will hold the new image's pixel data
    Dim newImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    fDraw.GetImageData2D srcPic, newImageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long, quickX2 As Long
    For x = 0 To iWidth - 2
        quickX = x * 3
        quickX2 = (x + 1) * 3
    For y = 0 To iHeight - 2
        b = Abs(CLng(imageData(quickX2, y + 1)) - CLng(imageData(quickX, y)) + eB)
        g = Abs(CLng(imageData(quickX2 + 1, y + 1)) - CLng(imageData(quickX + 1, y)) + eG)
        r = Abs(CLng(imageData(quickX2 + 2, y + 1)) - CLng(imageData(quickX + 2, y)) + eR)
        ByteMe r
        ByteMe g
        ByteMe b
        newImageData(quickX, y) = CByte(b)
        newImageData(quickX + 1, y) = CByte(g)
        newImageData(quickX + 2, y) = CByte(r)
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, newImageData()

End Sub

'Apply a generic grayscale "Emboss" filter to an image
Public Sub DrawEmboss(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, Optional ByVal eR As Byte = 127, Optional ByVal eG As Byte = 127, Optional ByVal eB As Byte = 127)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'This array will hold the new image's pixel data
    Dim newImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    fDraw.GetImageData2D srcPic, newImageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long, quickX2 As Long
    For x = 0 To iWidth - 2
        quickX = x * 3
        quickX2 = (x + 1) * 3
    For y = 0 To iHeight - 1
        b = Abs(CInt(imageData(quickX, y)) - CInt(imageData(quickX2, y)) + eB)
        g = Abs(CInt(imageData(quickX + 1, y)) - CInt(imageData(quickX2 + 1, y)) + eG)
        r = Abs(CInt(imageData(quickX + 2, y)) - CInt(imageData(quickX2 + 2, y)) + eR)
        ByteMe r
        ByteMe g
        ByteMe b
        newImageData(quickX, y) = b
        newImageData(quickX + 1, y) = g
        newImageData(quickX + 2, y) = r
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, newImageData()

End Sub

'Apply a "Relief" filter to an image
Public Sub DrawRelief(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'This array will hold the new image's pixel data
    Dim newImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, imageData()
    fDraw.GetImageData2D srcPic, newImageData()
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long

    'Run a quick loop through the image, adjusting pixel values as we go
    Dim quickX As Long, quickX2 As Long, quickX3 As Long
    
    For x = 1 To iWidth - 2
        quickX = x * 3
        quickX2 = (x + 1) * 3
        quickX3 = (x - 1) * 3
    For y = 1 To iHeight - 2
    
        b = 2 * (imageData(quickX3, y - 1)) + (imageData(quickX3, y)) + (imageData(quickX, y - 1))
        b = b - (imageData(x * 3, y + 1)) - (imageData(quickX3, y)) - 2 * (imageData(quickX2, y + 1))
        
        g = 2 * (imageData(quickX3 + 1, y - 1)) + (imageData(quickX3 + 1, y)) + (imageData(quickX + 1, y - 1))
        g = g - (imageData(x * 3 + 1, y + 1)) - (imageData(quickX3 + 1, y)) - 2 * (imageData(quickX2 + 1, y + 1))
        
        r = 2 * (imageData(quickX3 + 2, y - 1)) + (imageData(quickX3 + 2, y)) + (imageData(quickX + 2, y - 1))
        r = r - (imageData(x * 3 + 2, y + 1)) - (imageData(quickX3 + 2, y)) - 2 * (imageData(quickX2 + 2, y + 1))
        
        b = ((imageData(quickX, y)) + b) \ 2 + 50
        g = ((imageData(quickX + 1, y)) + g) \ 2 + 50
        r = ((imageData(quickX + 2, y)) + r) \ 2 + 50
        
        b = Abs(CInt(imageData(quickX, y)) - CInt(imageData(quickX2, y)) + imageData(quickX, y))
        g = Abs(CInt(imageData(quickX + 1, y)) - CInt(imageData(quickX2 + 1, y)) + imageData(quickX + 1, y))
        r = Abs(CInt(imageData(quickX + 2, y)) - CInt(imageData(quickX2 + 2, y)) + imageData(quickX + 2, y))
        
        ByteMe r
        ByteMe g
        ByteMe b
        
        newImageData(quickX, y) = b
        newImageData(quickX + 1, y) = g
        newImageData(quickX + 2, y) = r
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, newImageData()

End Sub

'Convert to absolute byte values (Long-type)
Public Sub ByteMe(ByRef tempVar As Long)
    If (tempVar > 255) Then tempVar = 255
    If (tempVar < 0) Then tempVar = 0
End Sub

'If no color is specified, the Emboss function will automatically use gray
Private Sub optEmboss_Click()
    chkColor.Enabled = True
    UpdateEffect
End Sub

'If no color is specified, the Engrave function will automatically use gray
Private Sub optEngrave_Click()
    chkColor.Enabled = True
    UpdateEffect
End Sub

'Generate a new "relief" version of the image
Private Sub optRelief_Click()
    chkColor.Enabled = False
    UpdateEffect
End Sub

'Clicking on the picture box allows the user to select a new color
Private Sub picColor_Click()
    
    Dim clrDialog As cSystemColorDialog, newColor As Long
    Set clrDialog = New cSystemColorDialog
    If clrDialog.ShowColorDialog(newColor, Me.hWnd, True, picColor.BackColor) Then
        picColor.BackColor = newColor
        UpdateEffect
    End If
    
End Sub

'These functions strip the red, green, or blue value from a Long-type color
Public Function ExtractR(ByVal currentColor As Long) As Long
    ExtractR = currentColor Mod 256
End Function

Public Function ExtractG(ByVal currentColor As Long) As Long
    ExtractG = (currentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal currentColor As Long) As Long
    ExtractB = (currentColor \ 65536) And 255
End Function
