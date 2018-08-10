VERSION 5.00
Begin VB.Form frmCurves 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Curves - tannerhelland.com"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   8145
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
   ScaleHeight     =   589
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExplanation 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   3735
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
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   1
      Top             =   4200
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
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.PictureBox picCurve 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   3855
      Left            =   4080
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image..."
      End
   End
End
Attribute VB_Name = "frmCurves"
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
'Description:
'
'This project demonstrates how to use a cubic spline to adjust image levels
' (nearly identical to Photoshop's method).  The code is well-commented, but there are
' some fairly involved math sections.  Don't feel bad if you don't understand all the details!
'
'Despite the complexity, however, the main routine is a (fairly simple) complete sub that
' could be dropped into any VB project after a few minor adjustments.
'
'Because a large portion of this project relies on DIB sections, I would recommend
' that you first read "From PSet to DIB Sections - your comprehensive guide to VB
' Graphics Programming."  This article can be downloaded from several places, most
' notably http://www.tannerhelland.com/42/vb-graphics-programming-3/
'
'Also, I owe GREAT thanks to the original author of the cubic spline routine I've used
' (Jason Bullen).
' His original cubic spline code can be downloaded from:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=11488&lngWId=1
'
'**************************************************************************************
' JASON'S ORIGINAL CUBIC SPLINE CODE COMMENTS:
' Here is an absolute minimum Cubic Spline routine.
' It's a VB rewrite from a Java applet I found by by Anthony Alto 4/25/99
' Computes coefficients based on equations mathematically derived from the curve
' constraints.   i.e. :
'    curves meet at knots (predefined points)  - These must be sorted by X
'    first derivatives must be equal at knots
'    second derivatives must be equal at knots
'**************************************************************************************


Option Explicit

'Modified cubic spline variables:
Private Const maxNPoints As Byte = 32
Dim nPoints As Byte
Private iX() As Single
Private iY() As Single
Private p() As Single
Private u() As Single

'Tanner's new variables:
Private isMouseDown As Boolean  'Track mouse status between MouseDown and MouseMove events
Private selPoint As Long        'Currently selected knot in the spline
Private results(-1 To 256) As Integer   'Stores the y-values for each x-value in the final spline
Private minX As Integer, maxX As Integer    'Used for calculating leading and trailing values
Private Const MOUSE_ACCURACY As Byte = 6 'How close to a knot the user must click to select that knot

'This routine draws gridlines, knots, and the spline on the picture box
Private Function DrawCubicSpline()
    
    'Tanner's inserted code: draw the background grid
    picCurve.Cls
    Dim i As Long
    picCurve.ForeColor = RGB(128, 128, 128)
    For i = 0 To 255 Step 64
        picCurve.Line (i, 0)-(i, 255)
        picCurve.Line (0, i)-(255, i)
    Next i
    'Now draw the knots
    picCurve.ForeColor = RGB(255, 0, 0)
    For i = 1 To nPoints
        'If this is the currently selected knot, fill it in with red
        If i = selPoint Then
            picCurve.FillStyle = 0
            picCurve.FillColor = RGB(255, 0, 0)
        End If
        picCurve.Circle (iX(i), iY(i)), 4, RGB(255, 0, 0)
        If i = selPoint Then
            picCurve.FillStyle = 1
            picCurve.FillColor = RGB(0, 0, 0)
        End If
    Next i
    picCurve.ForeColor = RGB(0, 0, 0)
    'Clear the results array and reset the max/min variables
    For i = -1 To 256
        results(i) = -1
    Next i
    minX = 256
    maxX = -1
    
    'Now run a loop through the knots, calculating spline values as we go
    Call SetPandU
    Dim xPos As Long, yPos As Single
    For i = 1 To nPoints - 1
        For xPos = iX(i) To iX(i + 1)
            yPos = GetCurvePoint(i, xPos)
            If xPos < minX Then minX = xPos
            If xPos > maxX Then maxX = xPos
            If yPos > 255 Then yPos = 254       'Force values to be in the 1-254 range (0-255 also
            If yPos < 0 Then yPos = 1           ' works, but is harder to see on the picture box)
            results(xPos) = yPos
        Next xPos
    Next i
    
    'Based on the maximum and minimum, calculate preceding and trailing y-values
    For i = -1 To minX - 1
        results(i) = results(minX)
    Next i
    For i = 256 To maxX + 1 Step -1
        results(i) = results(maxX)
    Next i
    
    'Draw the finished spline
    For i = 0 To 255
        picCurve.Line (i, results(i))-(i - 1, results(i - 1))
    Next i
    picCurve.Refresh
    
    'Last, but certainly not least, draw the curves-adjusted image
    DrawCurves picBack, picMain
    
End Function

'Original required spline function:
Private Function GetCurvePoint(ByVal i As Long, ByVal v As Single) As Single
    Dim t As Single
    'derived curve equation (which uses p's and u's for coefficients)
    t = (v - iX(i)) / u(i)
    GetCurvePoint = t * iY(i + 1) + (1 - t) * iY(i) + u(i) * u(i) * (F(t) * p(i + 1) + F(1 - t) * p(i)) / 6#
End Function

'Original required spline function:
Private Function F(x As Single) As Single
    F = x * x * x - x
End Function

'Original required spline function:
Private Sub SetPandU()
    Dim i As Integer
    Dim d() As Single
    Dim w() As Single
    ReDim d(nPoints) As Single
    ReDim w(nPoints) As Single
'Routine to compute the parameters of our cubic spline.  Based on equations derived from some basic facts...
'Each segment must be a cubic polynomial.  Curve segments must have equal first and second derivatives
'at knots they share.  General algorithm taken from a book which has long since been lost.

'The math that derived this stuff is pretty messy...  expressions are isolated and put into
'arrays.  we're essentially trying to find the values of the second derivative of each polynomial
'at each knot within the curve.  That's why theres only N-2 p's (where N is # points).
'later, we use the p's and u's to calculate curve points...

    For i = 2 To nPoints - 1
        d(i) = 2 * (iX(i + 1) - iX(i - 1))
    Next
    For i = 1 To nPoints - 1
        u(i) = iX(i + 1) - iX(i)
    Next
    For i = 2 To nPoints - 1
        w(i) = 6# * ((iY(i + 1) - iY(i)) / u(i) - (iY(i) - iY(i - 1)) / u(i - 1))
    Next
    For i = 2 To nPoints - 2
        w(i + 1) = w(i + 1) - w(i) * u(i) / d(i)
        d(i + 1) = d(i + 1) - u(i) * u(i) / d(i)
    Next
    p(1) = 0#
    For i = nPoints - 1 To 2 Step -1
        p(i) = (w(i) - u(i) * p(i + 1)) / d(i)
    Next
    p(nPoints) = 0#
End Sub

'********************FORM LOADING********************
Private Sub Form_Load()
    
    txtExplanation.Text = "If you've never used Photoshop's Curves dialog, here's a quick primer: select the red dots (or click to create your own) and drag them around to adjust the cubic spline." & vbCrLf & vbCrLf
    txtExplanation.Text = txtExplanation.Text & "The resulting curve is used to adjust the gamma of the image." & vbCrLf & vbCrLf
    txtExplanation.Text = txtExplanation.Text & "(X coords = input level, Y coords = output level)"
    
    'Set form-wide variables to their default values
    isMouseDown = False
    selPoint = -1
    minX = 256
    maxX = -1
    
    'Create 3 default points in a straight line (a good starting point for working with curves)
    nPoints = 3
    ReDim iX(nPoints) As Single
    ReDim iY(nPoints) As Single
    ReDim p(nPoints) As Single
    ReDim u(nPoints) As Single
    Dim i As Long
    For i = 1 To nPoints
        iX(i) = (i - 1) * (256 / (nPoints - 1))
        iY(i) = 255 - iX(i)
    Next i
    
    'Draw the initial effect, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    DrawCubicSpline
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
        DrawCurves picBack, picMain
        
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
    minWidth = (picCurve.Left + picCurve.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.TOp * 1.2) * Screen.TwipsPerPixelX
    
End Sub

'When the user clicks on the picture box, see if they've selected a control point or not
Private Sub picCurve_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'No point selected yet
    selPoint = -1
    
    'Search to see if the user has clicked on (or very near) an existing point
    Dim found As Long
    found = CheckClick(x, Y)
    
    'If the user has selected an existing point, mark it
    If found > -1 Then
        selPoint = found
    Else
        'No match was found, so create a new point here if:
        '  1) This x-coordinate isn't already occupied
        Dim i As Long
        For i = 1 To nPoints
            'The user has clicked on an already occupied x-coordinate. Our spline formula doesn't
            'allow two knots to have the same x-value, so instead of creating a new knot just
            'select the knot already occupying this coordinate
            If x = iX(i) Then
                selPoint = i
                picCurve.MousePointer = 5
                isMouseDown = True
                Exit Sub
            End If
        Next i
        
        '  2) We haven't reached the maximum allowed limit yet
        If nPoints < maxNPoints Then
            'Increase the total number of points and fix all our arrays
            nPoints = nPoints + 1
            ReDim Preserve iX(nPoints) As Single
            ReDim Preserve iY(nPoints) As Single
            ReDim Preserve p(nPoints) As Single
            ReDim Preserve u(nPoints) As Single
            'Figure out which points surround this location
            Dim nextX As Long
            nextX = nPoints
            For i = 1 To nPoints
                If iX(i) > x Then
                    nextX = i
                    Exit For
                End If
            Next i
                        
            'Shift all points after this to the right
            For i = nPoints - 1 To nextX Step -1
                iX(i + 1) = iX(i)
                iY(i + 1) = iY(i)
            Next i
            iX(nextX) = x
            iY(nextX) = Y
            
            'Draw the new spline, change the mousepointer to the move pointer, select this point
            DrawCubicSpline
            picCurve.MousePointer = 5
            selPoint = nextX
            
        End If
    End If
    
    'We mark the mouse state here for use in the MouseMove sub
    isMouseDown = True
End Sub

'Simple distance formula here - we use this to calculate if the user has clicked on (or near) a knot
Private Function pDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Single
    pDistance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'MouseMove allows the user to interactively adjust existing knots and add new knots
Private Sub picCurve_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    'Button down AND a point is current selected
    If isMouseDown = True And selPoint > -1 Then
        'The first knot has to be checked specially (no point before it)
        If selPoint = 0 Then
            If (x >= 0) And (x < iX(selPoint + 1)) Then iX(selPoint) = x
            If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
            DrawCubicSpline
            Exit Sub
        End If
        'The last knot also has to be checked specially (no point after it)
        If selPoint = nPoints Then
            If (x > iX(selPoint - 1)) And (x <= 255) Then iX(selPoint) = x
            If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
            DrawCubicSpline
            Exit Sub
        End If
        'For all middle knots, don't allow them to be moved past neighboring knots
        If (x > iX(selPoint - 1)) And (x < iX(selPoint + 1)) Then iX(selPoint) = x
        If (Y >= 0) And (Y <= 255) Then iY(selPoint) = Y
    End If
    DrawCubicSpline
    
    'Button up
    If isMouseDown = False Then
        'If the user is close to a knot, change the mousepointer to 'move'
        Dim found As Long
        found = CheckClick(x, Y)
        If found > -1 Then
            picCurve.MousePointer = 5
        Else
            picCurve.MousePointer = 0
        End If
    End If
    
End Sub

'When the mouse is lifted, reset the mousestate boolean and the mouse pointer
Private Sub picCurve_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    isMouseDown = False
    picCurve.MousePointer = 0
End Sub

'Simple distance routine to see if a location on the picture box is near an existing knot
Private Function CheckClick(ByVal x As Long, ByVal Y As Long) As Long
    Dim dist As Single
    Dim i As Long
    For i = 1 To nPoints
        dist = pDistance(x, Y, iX(i), iY(i))
        'If we're close to an existing point, return the index of that point
        If dist < MOUSE_ACCURACY Then
            CheckClick = i
            Exit Function
        End If
    Next i
    'Returning -1 says we're not close to an existing point (so try to create a new one!)
    CheckClick = -1
End Function

Public Sub DrawCurves(srcPic As PictureBox, dstPic As PictureBox)

    'This array will hold the image's pixel data
    Dim ImageData() As Byte
    
    'Coordinate variables
    Dim x As Long, Y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(frmCurves.picBack)
    iHeight = fDraw.GetImageHeight(frmCurves.picBack)
    fDraw.GetImageData2D frmCurves.picBack, ImageData()
    
    'These variables will hold temporary pixel color values
    Dim R As Long, G As Long, B As Long, L As Long

    'Look-up table calculation for new gamma values
    Dim newGamma(0 To 255) As Byte
    Dim tmpGamma As Double
    For x = 0 To 255
        tmpGamma = CDbl(x) / 255
        'This 'if' statement is necessary to match a weird trend with Photoshop's Curves dialog -
        ' for darker gamma values, Photoshop increases the force of the gamma conversion.  This is
        ' good for emphasizing subtle dark shades that the human eye doesn't normally pick up...
        ' I think.  If this 'if' statement is removed, the top statement will yield more mathematically
        ' consistent results, but this version ends up yielding results closer to what Photoshop's
        ' Curves dialog does.  Go figure!
        If results(x) <= (256 - x) Then
            tmpGamma = tmpGamma ^ (1 / ((256 - x) / (results(x) + 1)))
        Else
            tmpGamma = tmpGamma ^ ((1 / ((256 - x) / (results(x) + 1))) ^ 1.5)
        End If
        tmpGamma = tmpGamma * 255
        If tmpGamma > 255 Then
            tmpGamma = 255
        ElseIf tmpGamma < 0 Then
            tmpGamma = 0
        End If
        newGamma(x) = tmpGamma
    Next x
    
    'Now run a quick loop through the image, adjusting pixel values with the look-up tables
    Dim QuickX As Long
    For x = 0 To iWidth - 1
        QuickX = x * 3
    For Y = 0 To iHeight - 1
        'Grab red, green, and blue
        R = ImageData(QuickX + 2, Y)
        G = ImageData(QuickX + 1, Y)
        B = ImageData(QuickX, Y)
        'Correct them all
        ImageData(QuickX + 2, Y) = newGamma(R)
        ImageData(QuickX + 1, Y) = newGamma(G)
        ImageData(QuickX, Y) = newGamma(B)
    Next Y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D picMain, iWidth, iHeight, ImageData()


End Sub
