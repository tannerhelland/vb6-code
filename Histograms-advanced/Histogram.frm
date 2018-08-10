VERSION 5.00
Begin VB.Form frmHistogram 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Histogram"
   ClientHeight    =   6810
   ClientLeft      =   13200
   ClientTop       =   2805
   ClientWidth     =   4125
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
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Statistics"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   3855
      Begin VB.Label lblMaxCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum count:"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblCount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Count:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   3135
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3720
         Y1              =   530
         Y2              =   530
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblTotalPixels 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total pixels:"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox picGradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   5
      Top             =   4950
      Width           =   3870
   End
   Begin VB.ComboBox cmbHistSource 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox cmbHistMethod 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox picH 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   960
      Width           =   3870
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drawing Method:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   540
      Width           =   1230
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Histogram Source:"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "frmHistogram"
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

'Throughout this form, the following array locations refer to a type of histogram:
'0 - Red
'1 - Green
'2 - Blue
'3 - Luminance
'This applies especially to the hData() and hMax() arrays

'Histogram data for each particular type (r/g/b/luminance)
Dim hData(0 To 3, 0 To 255) As Single
'Maximum histogram values (r/g/b/luminance)
Dim hMax(0 To 3) As Single
'Loop and position variables
Dim x As Long, y As Long

'When either combo box is used, remember what was selected and then draw the new
'histogram
Private Sub cmbHistMethod_Click()
    lastHistMethod = cmbHistMethod.ListIndex
    DrawHistogram cmbHistSource.ListIndex, cmbHistMethod.ListIndex
End Sub

Private Sub cmbHistSource_Click()
    lastHistSource = cmbHistSource.ListIndex
    DrawHistogram cmbHistSource.ListIndex, cmbHistMethod.ListIndex
End Sub


'Generate the histogram data upon form load (we only need to do it once per image)
Private Sub Form_Load()
        
    GenerateHistogram
    
    'Clear out the combo box that displays histogram sources and fill it with
    'appropriate choices
    cmbHistSource.Clear
    cmbHistSource.AddItem "Red"
    cmbHistSource.AddItem "Green"
    cmbHistSource.AddItem "Blue"
    cmbHistSource.AddItem "Luminance"
    'Set the current combo box option to be whatever we used last
    cmbHistSource.ListIndex = lastHistSource
    
    'Clear out the combo box that displays histogram drawing methods and fill it
    'with appropriate choices
    cmbHistMethod.Clear
    cmbHistMethod.AddItem "Connecting lines"
    cmbHistMethod.AddItem "Solid bars"
    'Set the current combo box option to be whatever we used last
    cmbHistMethod.ListIndex = lastHistMethod
    
End Sub

'Subroutine to draw a histogram.  hType tells us what histogram to draw:
'0 - Red
'1 - Green
'2 - Blue
'3 - Luminance
'drawMethod tells us what kind of histogram to draw:
'0 - Connected lines (like a line graph)
'1 - Solid bars (like a bar graph)
Private Sub DrawHistogram(ByVal hType As Long, ByVal drawMethod As Long)
    
    'Clear out whatever was there before
    picH.Cls
    
    'tHeight is used to determine the height of the maximum value in the
    'histogram.  We want it to be slightly shorter than the height of the
    'picture box; this way the tallest histogram value fills the entire picture
    Dim tHeight As Long
    tHeight = picH.ScaleHeight - 2
    
    'LastX and LastY are used to draw a connecting line between histogram points
    Dim LastX As Long, LastY As Long
    
    'The type of histogram we're drawing will determine the color of the histogram
    'line - we'll make it match what we're drawing (red/green/blue/black)
    Select Case hType
        'Red
        Case 0
            picH.ForeColor = RGB(255, 0, 0)
        'Green
        Case 1
            picH.ForeColor = RGB(0, 255, 0)
        'Blue
        Case 2
            picH.ForeColor = RGB(0, 0, 255)
        'Luminance
        Case 3
            picH.ForeColor = RGB(0, 0, 0)
    End Select
    
    'Now draw a little gradient below the histogram window (just 'cause Photoshop
    'does it and it adds a nice touch, hehe ;)
    If hType < 3 Then
        'Draw a gradient from black to the color of the current histogram model
        DrawGradient picGradient, RGB(0, 0, 0), picH.ForeColor
    Else
        'For luminance, draw a gradient from black to white
        DrawGradient picGradient, RGB(0, 0, 0), RGB(255, 255, 255)
    End If
    
    'Now we'll draw the histogram.  Pay careful attention to this section of code
    
    'For the first point there is no last 'x' or 'y', so we'll just make it the
    'same as the first value in the histogram. (We care about this only if we're
    'drawing a "connected lines" type of histogram.)
    LastX = 0
    LastY = tHeight - (hData(hType, 0) / hMax(hType)) * tHeight
    
    'Run a loop through every histogram value...
    For x = 0 To 255
        'This is the most complicated line in the project.  The y-value of the
        'histogram is drawn as a percentage (RData(x) / MaxVal) * tHeight) with
        'tHeight being the tallest possible value (when RData(x) = MaxVal).  We
        'then subtract that value from tHeight because y values INCREASE as we
        'move DOWN a picture box - remember that (0,0) is in the top left.
        y = tHeight - (hData(hType, x) / hMax(hType)) * tHeight
        'For connecting lines...
        If drawMethod = 0 Then
            'Then draw a line from the last (x,y) to the current (x,y)
            picH.Line (LastX, LastY)-(x, y)
            LastX = x
            LastY = y
        'For a bar graph...
        ElseIf drawMethod = 1 Then
            'Draw a line from the bottom of the picture box to the calculated y-value
            picH.Line (x, tHeight)-(x, y)
        End If
    Next x
    
    'Last but not least, generate the statistics at the bottom of the form
    
    'Total number of pixels
    Dim fDraw As New FastDrawing
    Dim iWidth As Long, iHeight As Long
    iWidth = fDraw.GetImageWidth(frmMain.PicMain)
    iHeight = fDraw.GetImageHeight(frmMain.PicMain)
    lblTotalPixels.Caption = "Total pixels: " & (iWidth * iHeight)
    
    'Maximum value
    lblMaxCount.Caption = hMax(hType)

End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

'When the mouse moves over the histogram, display the level and count for the histogram
'entry at the x-value over which the mouse passes
Private Sub picH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblLevel.Caption = x
    lblCount.Caption = hData(lastHistSource, x)
End Sub

Public Sub GenerateHistogram()
    
    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.PicMain)
    iHeight = fDraw.GetImageHeight(frmMain.PicMain)
    fDraw.GetImageData2D frmMain.PicMain, imageData()
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, L As Long
    
    'If the histogram has already been used, we need to clear out all the
    'maximum values and histogram values
    For x = 0 To 3
        hMax(x) = 0
        For y = 0 To 255
            hData(x, y) = 0
        Next y
    Next x
    
    'Run a quick loop through the image, gathering what we need to
    'calculate our histogram
    Dim quickX As Long
    For x = 0 To iWidth - 1
        quickX = x * 3
    For y = 0 To iHeight - 1
        'We have to gather the red, green, and blue in order to calculate luminance
        r = imageData(quickX + 2, y)
        g = imageData(quickX + 1, y)
        b = imageData(quickX, y)
        'Rather than generate authentic luminance (which requires an HSL
        'conversion routine), we'll use the shortcut formula - averaging
        'R, G, and B.  This is plenty accurate for a project like this.
        L = (r + g + b) \ 3
        'Increment each value in the array, depending on its present value;
        'this will let us see how many of each color value (and luminance
        'value) there is in the image
        'Red
        hData(0, r) = hData(0, r) + 1
        'Green
        hData(1, g) = hData(1, g) + 1
        'Blue
        hData(2, b) = hData(2, b) + 1
        'Luminance
        hData(3, L) = hData(3, L) + 1
    Next y
    Next x
    
    'Run a quick loop through the completed array to find maximum values
    For x = 0 To 255
        'Red
        If hData(0, x) > hMax(0) Then hMax(0) = hData(0, x)
        'Green
        If hData(1, x) > hMax(1) Then hMax(1) = hData(1, x)
        'Blue
        If hData(2, x) > hMax(2) Then hMax(2) = hData(2, x)
        'Luminance
        If hData(3, x) > hMax(3) Then hMax(3) = hData(3, x)
    Next x
    
    'When all has been completed, off we go to draw the histogram.  Yay.
    'We'll always draw the luminance histogram first; if the user has already looked
    'at a particular histogram, remember which one and display it instead.
    DrawHistogram lastHistSource, lastHistMethod

End Sub


'StretchHistogram - if an image histogram doesn't reach from 0 to 255, make it.
Public Sub StretchHistogram()
    
    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.PicMain)
    iHeight = fDraw.GetImageHeight(frmMain.PicMain)
    fDraw.GetImageData2D frmMain.PicMain, imageData()
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long
    
    'These variables will store the maximum red, green, and blue values in this image
    Dim rMax As Byte, gMax As Byte, bMax As Byte
    Dim rMin As Byte, gMin As Byte, bMin As Byte

    'Reset the minimum values to a non-minimum value
    rMin = 255: gMin = 255: bMin = 255
    
    'This is used to access array locations quickly
    Dim quickVal As Long
    
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        'Grab values for this pixel
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        
        'Track maximum values of each color
        If r < rMin Then rMin = r
        If r > rMax Then rMax = r
        If g < gMin Then gMin = g
        If g > gMax Then gMax = g
        If b < bMin Then bMin = b
        If b > bMax Then bMax = b
        
    Next y
    Next x
    
    'Based on maximum and minimum values, calculate the current range of
    ' red, green, and blue values
    Dim rDif As Integer, gDif As Integer, bDif As Integer
    rDif = rMax - rMin
    gDif = gMax - gMin
    bDif = bMax - bMin
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        'Use this formula to force pixels to a new "stretched" value
        r = 255 * CSng(CSng(r - rMin) / CSng(rDif))
        g = 255 * CSng(CSng(g - gMin) / CSng(gDif))
        b = 255 * CSng(CSng(b - bMin) / CSng(bDif))
        imageData(quickVal + 2, y) = r
        imageData(quickVal + 1, y) = g
        imageData(quickVal, y) = b
    Next y
    Next x
    
    'Draw the new image
    fDraw.SetImageData2D frmMain.PicMain, iWidth, iHeight, imageData()
    
    'Draw the new histogram
    frmHistogram.GenerateHistogram
    
End Sub

'EqualizeHistogram - attempt to redistribute values across the brightness spectrum,
' with roughly the same amount of pixels at each brightness level
Public Sub EqualizeHistogram(ByVal HandleR As Boolean, ByVal HandleG As Boolean, ByVal HandleB As Boolean)
    
    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.PicMain)
    iHeight = fDraw.GetImageHeight(frmMain.PicMain)
    fDraw.GetImageData2D frmMain.PicMain, imageData()
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, L As Long
    
    'These arrays will hold brightness counts (used for the equalization formula)
    Dim rData(0 To 255) As Long, gData(0 To 255) As Long, bData(0 To 255) As Long
    
    'This is used to access array locations quickly
    Dim quickVal As Long
    
    'First, tally the amount of each color (i.e. build the histogram)
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        rData(r) = rData(r) + 1
        gData(g) = gData(g) + 1
        bData(b) = bData(b) + 1
    Next y
    Next x
    
    'Compute our scaling factor
    Dim scaleFactor As Single
    scaleFactor = 255 / (iWidth * iHeight)
    
    'Handle red as necessary
    If HandleR Then
        rData(0) = rData(0) * scaleFactor
        For x = 1 To 255
            rData(x) = rData(x - 1) + (scaleFactor * rData(x))
        Next x
    End If
    
    'Handle green as necessary
    If HandleG Then
        gData(0) = gData(0) * scaleFactor
        For x = 1 To 255
            gData(x) = gData(x - 1) + (scaleFactor * gData(x))
        Next x
    End If
    
    'Handle blue as necessary
    If HandleB Then
        bData(0) = bData(0) * scaleFactor
        For x = 1 To 255
            bData(x) = bData(x - 1) + (scaleFactor * bData(x))
        Next x
    End If
    
    'Integerize all the look-up values
    For x = 0 To 255
        rData(x) = Int(rData(x))
        If rData(x) > 255 Then rData(x) = 255
        gData(x) = Int(gData(x))
        If gData(x) > 255 Then gData(x) = 255
        bData(x) = Int(bData(x))
        If bData(x) > 255 Then bData(x) = 255
    Next x
    
    'Apply the equalized values
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        If HandleR Then imageData(quickVal + 2, y) = rData(imageData(quickVal + 2, y))
        If HandleG Then imageData(quickVal + 1, y) = gData(imageData(quickVal + 1, y))
        If HandleB Then imageData(quickVal, y) = bData(imageData(quickVal, y))
    Next y
    Next x
    
    'Draw the new image
    fDraw.SetImageData2D frmMain.PicMain, iWidth, iHeight, imageData()
    
    'Draw the new histogram
    frmHistogram.GenerateHistogram

End Sub

'EqualizeLuminance - attempt to redistribute values across the luminance spectrum,
' with roughly the same amount of pixels at each luminance level
Public Sub EqualizeLuminance()
    
    'This array will hold the image's pixel data
    Dim imageData() As Byte
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmMain.PicMain)
    iHeight = fDraw.GetImageHeight(frmMain.PicMain)
    fDraw.GetImageData2D frmMain.PicMain, imageData()
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, L As Long
    
    'This is used to access array locations quickly
    Dim quickVal As Long
    
    'This array will hold our luminance values
    Dim lum(0 To 255) As Single
    
    'These values are used to convert between RGB and HSL
    Dim hh As Single, ss As Single, ll As Single
    
    'First, tally the luminance amounts (i.e. build the histogram)
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        tRGBToHSL r, g, b, hh, ss, ll
        lum(ll) = lum(ll) + 1
    Next y
    Next x
    
    'Compute our scaling factor
    Dim scaleFactor As Double
    scaleFactor = 255 / (iWidth * iHeight)
    'Equalize the luminance
    lum(0) = lum(0) * scaleFactor
    For x = 1 To 255
        lum(x) = lum(x - 1) + (scaleFactor * lum(x))
    Next x
   'Integerize all the look-up values
    For x = 0 To 255
        lum(x) = Int(lum(x))
        If lum(x) > 255 Then lum(x) = 255
    Next x
    
    'Apply the equalized values
    For x = 0 To iWidth - 1
        quickVal = x * 3
    For y = 0 To iHeight - 1
        'Get the temporary values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        'Get the hue and saturation
        tRGBToHSL r, g, b, hh, ss, ll
        'Convert back to RGB using our artificial luminance values
        tHSLToRGB hh, ss, lum(ll) / 255, r, g, b
        'Assign those values into the array
        imageData(quickVal + 2, y) = r
        imageData(quickVal + 1, y) = g
        imageData(quickVal, y) = b
    Next y
    Next x
    
    'Draw the new image
    fDraw.SetImageData2D frmMain.PicMain, iWidth, iHeight, imageData()
    
    'Draw the new histogram
    frmHistogram.GenerateHistogram

End Sub

'All routines below this point are used to convert pixels between the RGB color space
' and an HSL color space.  An explanation of these functions is beyond the scope of
' this project - so I'm afraid you'll just have to take them on faith for now.  :)
Private Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Single, s As Single, L As Single)
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single
   rR = r / 255: rG = g / 255: rB = b / 255

        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        L = (Max + Min) / 2
        If Max = Min Then
            s = 0
            h = 0
        Else
           If L <= 0.5 Then
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
    
    L = Int(L * 255)
    
End Sub

Public Sub tHSLToRGB(h As Single, s As Single, L As Single, r As Long, g As Long, b As Long)
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

   If s = 0 Then
      rR = L: rG = L: rB = L
   Else
      If L <= 0.5 Then
         Min = L * (1 - s)
      Else
         Min = L - s * (1 - L)
      End If
      Max = 2 * L - Min
      
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
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Private Function Maximum(ByVal rR As Single, ByVal rG As Single, ByVal rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

Private Function Minimum(ByVal rR As Single, ByVal rG As Single, ByVal rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

