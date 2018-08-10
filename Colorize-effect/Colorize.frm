VERSION 5.00
Begin VB.Form frmColorize 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colorize Image Filter - tannerhelland.com"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11190
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
   ScaleHeight     =   613
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   746
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hScrollSaturation 
      Height          =   255
      Left            =   6000
      Max             =   100
      TabIndex        =   9
      Top             =   1200
      Value           =   50
      Width           =   2175
   End
   Begin VB.OptionButton optKeepSaturation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Keep original pixel saturation"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.HScrollBar hScrollHue 
      Height          =   255
      Left            =   720
      Max             =   359
      TabIndex        =   5
      Top             =   240
      Value           =   180
      Width           =   7455
   End
   Begin VB.PictureBox picHueDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   3
      Top             =   600
      Width           =   6990
   End
   Begin VB.CommandButton cmdReset 
      Appearance      =   0  'Flat
      Caption         =   "Show the Original Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8400
      TabIndex        =   2
      Top             =   240
      Width           =   2535
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
      Top             =   1680
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
      Height          =   6780
      Left            =   120
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   10830
   End
   Begin VB.OptionButton optForceSaturation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Force all saturation to"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblSaturation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "50%"
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
      Left            =   5340
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hue:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmColorize"
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

'When the program is first loaded, several things needs to happen...
Private Sub Form_Load()

    'Draw a "rainbow" into the hue box to help the user colorize the image properly
    Dim hVal As Single
    Dim r As Long, g As Long, b As Long
    Dim x As Long, y As Long
    
    For x = 0 To picHueDemo.ScaleWidth
        'Based on our x-position, interpolate an x value between -1 and 5
        hVal = x / picHueDemo.ScaleWidth
        hVal = hVal * 360
        hVal = (hVal - 60) / 60
        
        'Use our x-coordinate to generate a hue for this position
        tHSLToRGB hVal, 1, 0.5, r, g, b
        
        'Draw the color to the box
        picHueDemo.Line (x, 0)-(x, picHueDemo.ScaleHeight), RGB(r, g, b)
    Next x
    
    picHueDemo.Picture = picHueDemo.Image
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'Apply a Colorize filter to an image using a user-specified hue (defined between -1 and 5) and, optionally, a user-specified saturation (defined between 0.0 and 1.0)
Public Sub DrawColorize(srcPic As PictureBox, dstPic As PictureBox, ByVal newHue As Single, Optional ByVal forceSaturation As Boolean = False, Optional ByVal newSaturation As Single = 0.5)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic)
    iHeight = fDraw.GetImageHeight(srcPic)
    fDraw.GetImageData2D srcPic, ImageData()
    
    'These variables will hold temporary pixel color values (red, green, blue / hue, saturation, luminance)
    Dim r As Long, g As Long, b As Long
    Dim HH As Single, SS As Single, LL As Single
    
    'Let's do simple checking to make sure the hue and saturation values we were passed are valid
    If newHue < -1 Then newHue = -1
    If newHue > 5 Then newHue = 5
    If newSaturation < 0 Then newSaturation = 0
    If newSaturation > 1 Then newSaturation = 1
    
    'Loop through the image, adjusting pixel values as we go
    Dim xStride As Long
    
    For y = 0 To iHeight - 1
    For x = 0 To iWidth - 1
        
        xStride = x * 3
        
        b = ImageData(xStride, y)
        g = ImageData(xStride + 1, y)
        r = ImageData(xStride + 2, y)
        
        'Get the hue and saturation of the current pixel
        tRGBToHSL r, g, b, HH, SS, LL
        
        'Convert back to RGB using the hue specified by the user (our artificial saturation value?)
        If forceSaturation Then
            tHSLToRGB newHue, newSaturation, LL, r, g, b  'Use this line to force pixel saturation values to 50%
        Else
            tHSLToRGB newHue, SS, LL, r, g, b   'Use this line to keep original pixel saturation values
        End If
        
        'Assign our new colors back into the image array
        ImageData(xStride, y) = CByte(b)
        ImageData(xStride + 1, y) = CByte(g)
        ImageData(xStride + 2, y) = CByte(r)
        
    Next x
    Next y
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth, iHeight, ImageData()

End Sub

'Whenever the scroll bar is moved, redraw the colorized image
Private Sub hScrollHue_Change()
    UpdateEffect
End Sub

Private Sub hScrollHue_Scroll()
    UpdateEffect
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'When this scroll bar is changed, make the associated label match its value
Private Sub hScrollSaturation_Change()
    optForceSaturation.Value = True
    lblSaturation.Caption = hScrollSaturation.Value & "%"
    UpdateEffect
End Sub

Private Sub hScrollSaturation_Scroll()
    optForceSaturation.Value = True
    lblSaturation.Caption = hScrollSaturation.Value & "%"
    UpdateEffect
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
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 1.5) * Screen.TwipsPerPixelY
    
End Sub

Private Sub UpdateEffect()
    If optForceSaturation.Value Then
        DrawColorize picBack, picMain, CSng((CSng(hScrollHue.Value) - 60) / 60), True, CSng(CSng(hScrollSaturation.Value) / 100)
    Else
        DrawColorize picBack, picMain, CSng((CSng(hScrollHue.Value) - 60) / 60), False
    End If
End Sub

'The next four functions are required to convert between the HSL and RGB colorspaces
Public Sub tRGBToHSL(ByVal r As Long, ByVal g As Long, ByVal b As Long, h As Single, s As Single, l As Single)
    
    Dim Max As Single, Min As Single, delta As Single
    Dim rR As Single, rG As Single, rB As Single
    
    rR = r / 255
    rG = g / 255
    rB = b / 255

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2
    
    If Max = Min Then
        s = 0
        h = 0
    Else
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If

        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta
        End If
    End If

End Sub

Public Sub tHSLToRGB(ByVal h As Single, ByVal s As Single, ByVal l As Single, r As Long, g As Long, b As Long)

    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single

    If s = 0 Then
        rR = l
        rG = l
        rB = l
    Else
        If l <= 0.5 Then Min = l * (1 - s) Else Min = l - s * (1 - l)
        Max = 2 * l - Min
      
        If (h < 1) Then
            rR = Max
            If (h < 0) Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If (h < 2) Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (h < 4) Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
        End If
    End If
    
    r = rR * 255
    g = rG * 255
    b = rB * 255
   
End Sub

'Return the maximum of three variables
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
    If (rR > rG) Then
        If (rR > rB) Then Maximum = rR Else Maximum = rB
    Else
        If (rB > rG) Then Maximum = rB Else Maximum = rG
    End If
End Function

'Return the minimum of three variables
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
    If (rR < rG) Then
        If (rR < rB) Then Minimum = rR Else Minimum = rB
    Else
        If (rB < rG) Then Minimum = rB Else Minimum = rG
    End If
End Function

'If the user changes the saturation option, redraw the main image to match
Private Sub optForceSaturation_Click()
    UpdateEffect
End Sub

Private Sub optKeepSaturation_Click()
    UpdateEffect
End Sub
