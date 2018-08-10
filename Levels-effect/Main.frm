VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Real-time Image Levels - www.tannerhelland.com"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12375
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
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDispHistogram 
      Appearance      =   0  'Flat
      Caption         =   "Display Histogram"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Frame frmLevels 
      BackColor       =   &H80000005&
      Caption         =   "Levels:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton cmdReset 
         Appearance      =   0  'Flat
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   2640
         Width           =   1335
      End
      Begin VB.HScrollBar hsInM 
         Height          =   220
         Left            =   1560
         Max             =   254
         Min             =   1
         TabIndex        =   18
         Top             =   990
         Value           =   127
         Width           =   3855
      End
      Begin VB.HScrollBar hsInL 
         Height          =   220
         Left            =   1560
         Max             =   253
         TabIndex        =   13
         Top             =   735
         Width           =   3855
      End
      Begin VB.HScrollBar hsInR 
         Height          =   220
         Left            =   1560
         Max             =   255
         Min             =   2
         TabIndex        =   12
         Top             =   1250
         Value           =   255
         Width           =   3855
      End
      Begin VB.HScrollBar hsOutR 
         Height          =   220
         Left            =   1560
         Max             =   255
         TabIndex        =   9
         Top             =   2205
         Value           =   255
         Width           =   3855
      End
      Begin VB.HScrollBar hsOutL 
         Height          =   220
         Left            =   1560
         Max             =   255
         TabIndex        =   5
         Top             =   1950
         Width           =   3855
      End
      Begin VB.Label lblRightL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Segoe UI Light"
            Size            =   9.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1185
         TabIndex        =   22
         Top             =   1200
         Width           =   105
      End
      Begin VB.Label lblMiddleL 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   1185
         TabIndex        =   21
         Top             =   960
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midtones:"
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
         TabIndex        =   20
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblMiddleR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "254"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left limit:     0"
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
         TabIndex        =   17
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lblLeftR 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "253"
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
         Left            =   5520
         TabIndex        =   16
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right limit:"
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
         TabIndex        =   15
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5520
         TabIndex        =   14
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Input levels:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5520
         TabIndex        =   10
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right limit:  0"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "255"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label lblOutputL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Left limit:    0"
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
         TabIndex        =   4
         Top             =   1920
         Width           =   1125
      End
      Begin VB.Label lblOutputLevels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Output levels:"
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
         TabIndex        =   3
         Top             =   1560
         Width           =   1185
      End
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
      Left            =   6240
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   120
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
      Left            =   6240
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open Image"
      End
   End
End
Attribute VB_Name = "frmMain"
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

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'Used to track the ratio of the midtones scrollbar, so that when the left and
'right values get changed, we automatically set the midtone to the same ratio
'(i.e. as Photoshop does it)
Dim midRatio As Double

'Whether or not changing the midtone scrollbar is user-generated or program-generated
'(so we only refresh if the user moved it - otherwise we get bad looping)
Dim iRefresh As Boolean

'When the program starts, automatically initialize several things...
Private Sub Form_Load()
    
    'Upon loading the form, automatically set two histogram variables:
     'Luminance is the default histogram source
     lastHistSource = DRAWMETHOD_LUMINANCE
     'Line graph is the default drawing option
     lastHistMethod = DRAWMETHOD_BARS

    'Also, set the default midtone scrollbar ratio to 1/2
    midRatio = 0.5
    
    '...and allow refreshing
    iRefresh = True
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'The histogram information will be displayed on a separate form
Private Sub cmdDispHistogram_Click()
    frmHistogram.Show
End Sub


'This will reset the scrollbars to default levels
Private Sub cmdReset_Click()

    'Allow refreshing
    iRefresh = True
    
    'Set the output levels to (0-255)
    hsOutL.Value = 0
    hsOutR.Value = 255
    
    'Set the input levels to (0-255)
    hsInL.Value = 0
    hsInR.Value = 255
    FixScrollBars
    
    'Set the midtone level to default (127)
    midRatio = 0.5
    hsInM.Value = 127
    FixScrollBars
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmHistogram
End Sub

'*********************************************************************************
'The following 10 subroutines are for changing/scrolling any of the scrollbars
'on the main form
'*********************************************************************************
Private Sub hsInL_Change()
    FixScrollBars
    UpdateEffect
End Sub

Private Sub hsInL_Scroll()
    FixScrollBars
    UpdateEffect
End Sub

Private Sub hsInM_Change()
    If iRefresh Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        UpdateEffect
    End If
End Sub

Private Sub hsInM_Scroll()
    If iRefresh Then
        midRatio = (CDbl(hsInM.Value) - CDbl(hsInL.Value)) / (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        FixScrollBars True
        UpdateEffect
    End If
End Sub

Private Sub hsInR_Change()
    FixScrollBars
    UpdateEffect
End Sub

Private Sub hsInR_Scroll()
    FixScrollBars
    UpdateEffect
End Sub

Private Sub hsOutL_Change()
    UpdateEffect
End Sub

Private Sub hsOutL_Scroll()
    UpdateEffect
End Sub

Private Sub hsOutR_Change()
    UpdateEffect
End Sub

Private Sub hsOutR_Scroll()
    UpdateEffect
End Sub

'To load a new image...
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
        
        'Reset all scrollbars
        hsOutL.Value = 0
        hsOutR.Value = 255
        hsInL.Value = 0
        hsInR.Value = 255
        FixScrollBars
        hsInM.Value = 127
        midRatio = 0.5
        FixScrollBars
        
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
    Dim newHeight As Long, minHeight As Long
    minHeight = (frmLevels.Top + frmLevels.Height + 60) * Screen.TwipsPerPixelY
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then
        newHeight = (imgHeight + picMain.Top * 10) * Screen.TwipsPerPixelY
        If (newHeight < minHeight) Then newHeight = minHeight
        Me.Height = newHeight
    End If
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then Me.Width = (imgWidth + picMain.Left * 1.1) * Screen.TwipsPerPixelX
    
End Sub

Private Sub UpdateEffect()
    MapImageLevels picBack, picMain, hsInL.Value, hsInM.Value, hsInR.Value, hsOutL.Value, hsOutR.Value
End Sub

'Draw an image based on user-adjusted input and output levels
Private Sub MapImageLevels(srcPic As PictureBox, dstPic As PictureBox, ByVal inLLimit As Long, ByVal inMLimit As Long, ByVal inRLimit As Long, ByVal outLLimit As Long, ByVal outRLimit As Long)

    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.picBack)
    iHeight = fDraw.GetImageHeight(frmMain.picBack)
    fDraw.GetImageData2D frmMain.picBack, imageData()
    
    'These variables will hold temporary pixel color values
    Dim R As Long, G As Long, B As Long, l As Long
    
    'Look-up table for the midtone (gamma) leveled values
    Dim gValues(0 To 255) As Double
    
    'WARNING: This next chunk of code is a lot of messy math.  Don't worry too much
    'if you can't make sense of it ;)
    
    'Fill the gamma table with appropriate gamma values (from 10 to .1, ranged quadratically)
    'NOTE: This table is constant, and could be loaded from file instead of generated mathematically every time we run this function
    Dim gStep As Double
    gStep = (MAXGAMMA + MIDGAMMA) / 127
    For x = 0 To 127
        gValues(x) = (CDbl(x) / 127) * MIDGAMMA
    Next x
    For x = 128 To 255
        gValues(x) = MIDGAMMA + (CDbl(x - 127) * gStep)
    Next x
    For x = 0 To 255
        gValues(x) = 1 / ((gValues(x) + 1 / ROOT10) ^ 2)
    Next x
    
    'Because we've built our look-up tables on a 0-255 scale, correct the inMLimit
    'value (from the midtones scroll bar) to simply represent a ratio on that scale
    Dim tRatio As Double
    tRatio = (inMLimit - inLLimit) / (inRLimit - inLLimit)
    tRatio = tRatio * 255
    'Then convert that ratio into a byte (so we can access a look-up table with it)
    Dim bRatio As Byte
    bRatio = CByte(tRatio)
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim gLevels(0 To 255) As Byte
    Dim tmpGamma As Double
    For x = 0 To 255
        tmpGamma = CDbl(x) / 255
        tmpGamma = tmpGamma ^ (1 / gValues(bRatio))
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        gLevels(x) = tmpGamma
    Next x
    
    'Look-up table for the input leveled values
    Dim newLevels(0 To 255) As Byte
    
    'Fill the look-up table with appropriately mapped input limits
    Dim pStep As Single
    pStep = 255 / (CSng(inRLimit) - CSng(inLLimit))
    For x = 0 To 255
        If x < inLLimit Then
            newLevels(x) = 0
        ElseIf x > inRLimit Then
            newLevels(x) = 255
        Else
            newLevels(x) = ByteMe(((CSng(x) - CSng(inLLimit)) * pStep))
        End If
    Next x
    
    'Now run all input-mapped values through our midtone-correction look-up
    For x = 0 To 255
        newLevels(x) = gLevels(newLevels(x))
    Next x
    
    'Last of all, remap all image values to match the user-specified output limits
    Dim oStep As Double
    oStep = (CSng(outRLimit) - CSng(outLLimit)) / 255
    For x = 0 To 255
        newLevels(x) = ByteMe(CSng(outLLimit) + (CSng(newLevels(x)) * oStep))
    Next x
    
    'Now run a quick loop through the image, adjusting pixel values with the look-up tables
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
    
        'Grab red, green, and blue
        B = imageData(quickX, y)
        G = imageData(quickX + 1, y)
        R = imageData(quickX + 2, y)
        
        'Correct them all
        imageData(quickX, y) = newLevels(B)
        imageData(quickX + 1, y) = newLevels(G)
        imageData(quickX + 2, y) = newLevels(R)
        
    Next y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D picMain, iWidth, iHeight, imageData()

End Sub

'Used to make sure the scroll bars have appropriate limits
Private Sub FixScrollBars(Optional midMoving As Boolean = False)

    'Make sure that the input scrollbar values don't overlap, and update the labels
    'to display such
    hsInM.Min = hsInL.Value + 1
    lblMiddleL.Caption = hsInL.Value + 1
    hsInR.Min = hsInL.Value + 2
    lblRightL.Caption = hsInL.Value + 2
    hsInL.Max = hsInR.Value - 2
    lblLeftR.Caption = hsInR.Value - 2
    hsInM.Max = hsInR.Value - 1
    lblMiddleR.Caption = hsInR.Value - 1
    
    'If the user hasn't moved the midtones scrollbar, attempt to preserve its ratio
    If (Not midMoving) Then
        iRefresh = False
        Dim newValue As Long
        newValue = hsInL.Value + midRatio * (CDbl(hsInR.Value) - CDbl(hsInL.Value))
        If newValue > hsInM.Max Then
            newValue = hsInM.Max
        ElseIf newValue < hsInM.Min Then
            newValue = hsInM.Min
        End If
        hsInM.Value = newValue
        DoEvents
        iRefresh = True
    End If
    
End Sub

'Used to restrict values to the (0-255) range
Private Function ByteMe(ByVal val As Long) As Byte
    If val > 255 Then
        ByteMe = 255
    ElseIf val < 0 Then
        ByteMe = 0
    Else
        ByteMe = val
    End If
End Function

