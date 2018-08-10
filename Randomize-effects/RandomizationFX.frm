VERSION 5.00
Begin VB.Form frmLineEffect 
   BackColor       =   &H80000005&
   Caption         =   "Line Randomization Effect - www.tannerhelland.com"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9900
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
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDrawTriangles 
      Caption         =   "Draw Triangles!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLength 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3600
      TabIndex        =   6
      Text            =   "40"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtIterations 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Text            =   "100000"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDrawLines 
      Appearance      =   0  'Flat
      Caption         =   "Draw Lines!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
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
      Height          =   5025
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   0
      Top             =   720
      Width           =   9570
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
      Height          =   5025
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   9570
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Line Length:"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Iterations:"
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
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   885
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmLineEffect"
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


'Drawing polygons
Private Type POINTAPI
    x As Long
    y As Long
End Type

Dim fxTri(0 To 2) As POINTAPI

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Option Explicit

'Draw a cool line effect
Private Sub cmdDrawLines_Click()
    DrawLineEffect picBack, picMain, Val(CLng(txtIterations.Text)), Val(CLng(txtLength.Text))
End Sub

'Draw a cool triangle effect
Private Sub cmdDrawTriangles_Click()
    DrawTriangleEffect picBack, picMain, Val(CLng(txtIterations.Text)), Val(CLng(txtLength.Text))
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()

    Dim fDraw As New FastDrawing
    Dim ImageData() As Byte
    Dim iWidth As Long, iHeight As Long
    iWidth = fDraw.GetImageWidth(Me.picBack)
    iHeight = fDraw.GetImageHeight(Me.picBack)
    fDraw.GetImageData2D Me.picBack, ImageData()
    fDraw.SetImageData2D Me.picMain, iWidth, iHeight, ImageData()
    
End Sub

Private Sub Form_Load()

    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
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
        LoadImageAutosized sFile
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
    minWidth = (cmdReset.Left + cmdReset.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 3) * Screen.TwipsPerPixelY
    
End Sub

'Apply a cool triangle-draw filter to an image
Public Sub DrawTriangleEffect(srcPic As PictureBox, dstPic As PictureBox, Optional numLoops As Long = 10000, Optional lenLine As Long = 50)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'This array will hold the new image's pixel data
    Dim newImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long
    
    'Loop Counter
    Dim n As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, ImageData(), True
    fDraw.GetImageData2D srcPic, newImageData(), True
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long

    'Reset the randomizer
    Randomize Timer
    
    'These make it faster
    Dim quickX As Long, quickX2 As Long, quickX3 As Long, lenHalf As Long
    
    lenHalf = lenLine \ 2
    
    'Run a quick loop through the image, drawing lines as we go
    For n = 0 To numLoops
        
        'Generate three random coordinates
        x = Int(Rnd * iWidth)
        x2 = Int(Rnd * lenLine) - lenHalf + x
        If x2 < 0 Then x2 = 0
        If x2 > iWidth - 1 Then x2 = iWidth - 1
        x3 = Int(Rnd * lenLine) - lenHalf + x
        If x3 < 0 Then x3 = 0
        If x3 > iWidth - 1 Then x3 = iWidth - 1
        
        y = Int(Rnd * iHeight)
        y2 = Int(Rnd * lenLine) - lenHalf + y
        If y2 < 0 Then y2 = 0
        If y2 > iHeight - 1 Then y2 = iHeight - 1
        y3 = Int(Rnd * lenLine) - lenHalf + y
        If y3 < 0 Then y3 = 0
        If y3 > iHeight - 1 Then y3 = iHeight - 1
        
        quickX = x * 3
        quickX2 = x2 * 3
        quickX3 = x3 * 3
        
        r = ImageData(quickX + 2, y)
        r = r + ImageData(quickX2 + 2, y2)
        r = r + ImageData(quickX3 + 2, y3)
        g = ImageData(quickX + 1, y)
        g = g + ImageData(quickX2 + 1, y2)
        g = g + ImageData(quickX3 + 1, y3)
        b = ImageData(quickX, y)
        b = b + ImageData(quickX2, y2)
        b = b + ImageData(quickX3, y3)
    
        r = r \ 3
        g = g \ 3
        b = b \ 3
        
        fxTri(0).x = x
        fxTri(0).y = y
        fxTri(1).x = x2
        fxTri(1).y = y2
        fxTri(2).x = x3
        fxTri(2).y = y3
        
        dstPic.FillColor = RGB(r, g, b)
        dstPic.ForeColor = RGB(r, g, b)
        Polygon dstPic.hdc, fxTri(0), 3
        
        'dstPic.Line (x, y)-(x2, y2), RGB(R, G, B)
    
        If (n And 95) = 0 Then
            dstPic.Picture = dstPic.Image
            dstPic.Refresh
            DoEvents
        End If
    
    Next n
    
    dstPic.Picture = dstPic.Image
    dstPic.Refresh
    
    'Draw the new image data to the screen
    'fDraw.SetImageData2D dstPic, iWidth, iHeight, NewImageData()

End Sub



'Apply a cool line-draw filter to an image
Public Sub DrawLineEffect(srcPic As PictureBox, dstPic As PictureBox, Optional numLoops As Long = 10000, Optional lenLine As Long = 50)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'This array will hold the new image's pixel data
    Dim newImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long, x2 As Long, y2 As Long
    
    'Loop Counter
    Dim n As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, ImageData(), True
    fDraw.GetImageData2D srcPic, newImageData(), True
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long

    'Reset the randomizer
    Randomize Timer
    
    'These make it faster
    Dim quickX As Long, quickX2 As Long, lenHalf As Long
    
    lenHalf = lenLine \ 2
    
    'Run a quick loop through the image, drawing lines as we go
    For n = 0 To numLoops
        
        'Generate two random coordinates
        x = Int(Rnd * iWidth)
        x2 = Int(Rnd * lenLine) - lenHalf + x
        If x2 < 0 Then x2 = 0
        If x2 > iWidth - 1 Then x2 = iWidth - 1
        
        y = Int(Rnd * iHeight)
        y2 = Int(Rnd * lenLine) - lenHalf + y
        If y2 < 0 Then y2 = 0
        If y2 > iHeight - 1 Then y2 = iHeight - 1
        
        quickX = x * 3
        quickX2 = x2 * 3
        
        r = ImageData(quickX + 2, y)
        r = r + ImageData(quickX2 + 2, y2)
        g = ImageData(quickX + 1, y)
        g = g + ImageData(quickX2 + 1, y2)
        b = ImageData(quickX, y)
        b = b + ImageData(quickX2, y2)
    
        r = r \ 2
        g = g \ 2
        b = b \ 2
        
        dstPic.Line (x, y)-(x2, y2), RGB(r, g, b)
    
        If (n And 95) = 0 Then
            dstPic.Picture = dstPic.Image
            dstPic.Refresh
            DoEvents
        End If
    
    Next n
    
    dstPic.Picture = dstPic.Image
    dstPic.Refresh
    
    'Draw the new image data to the screen
    'fDraw.SetImageData2D dstPic, iWidth, iHeight, NewImageData()

End Sub

'Convert to absolute byte values (Long-type)
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub

