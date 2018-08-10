VERSION 5.00
Begin VB.Form frmEdgeDetection 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edge Detection Algorithms - www.tannerhelland.com"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6270
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
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Frame frmDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Description:"
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
      Height          =   2415
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.Label lblDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.ListBox lstEdgeOptions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2655
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
      TabIndex        =   0
      Top             =   3240
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
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmEdgeDetection"
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

'FM() is our Filter Matrix - look at individual edge filters to see how this works
Private FM() As Long

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'Populate the possible edge detection routines
Private Sub Form_Load()

    lstEdgeOptions.Clear
    lstEdgeOptions.AddItem "Prewitt Horizontal"
    lstEdgeOptions.AddItem "Prewitt Vertical"
    lstEdgeOptions.AddItem "Sobel Horizontal"
    lstEdgeOptions.AddItem "Sobel Vertical"
    lstEdgeOptions.AddItem "Laplacian"
    lstEdgeOptions.AddItem "Hilite"
    lstEdgeOptions.AddItem "Helland Linear"
    lstEdgeOptions.AddItem "Helland Cubic"
    
    lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1  0  1" & vbCrLf & "-1  0  1" & vbCrLf & "-1  0  1"
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    lstEdgeOptions.ListIndex = 0
    Me.Show
    
End Sub

'When an edge detection routine is selected, run it and display its corresponding matrix in the label on the main form
Private Sub lstEdgeOptions_Click()
    
    Dim l As String
    l = lstEdgeOptions.List(lstEdgeOptions.ListIndex)
    
    'Prewitt Horizontal
    If l = "Prewitt Horizontal" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1  0  1" & vbCrLf & "-1  0  1" & vbCrLf & "-1  0  1"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = -1
            FM(-1, 0) = -1
            FM(-1, 1) = -1
            FM(1, -1) = 1
            FM(1, 0) = 1
            FM(1, 1) = 1
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Prewitt Vertical" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1 -1 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  1  1"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = 1
            FM(0, -1) = 1
            FM(1, -1) = 1
            FM(-1, 1) = -1
            FM(0, 1) = -1
            FM(1, 1) = -1
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Sobel Horizontal" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1  0  1" & vbCrLf & "-2  0  2" & vbCrLf & "-1  0  1"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = -1
            FM(-1, 0) = -2
            FM(-1, 1) = -1
            FM(1, -1) = 1
            FM(1, 0) = 2
            FM(1, 1) = 1
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Sobel Vertical" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1 -2 -1" & vbCrLf & " 0  0  0" & vbCrLf & " 1  2  1"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = 1
            FM(0, -1) = 2
            FM(1, -1) = 1
            FM(-1, 1) = -1
            FM(0, 1) = -2
            FM(1, 1) = -1
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Laplacian" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & " 0 -1  0" & vbCrLf & "-1  4 -1" & vbCrLf & " 0 -1  0"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, 0) = -1
            FM(0, -1) = -1
            FM(0, 1) = -1
            FM(1, 0) = -1
            FM(0, 0) = 4
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Hilite" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-4 -2 -1" & vbCrLf & "-2 10  0" & vbCrLf & "-1  0  0"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = -4
            FM(-1, 0) = -2
            FM(0, -1) = -2
            FM(1, -1) = -1
            FM(-1, 1) = -1
            FM(0, 0) = 10
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    ElseIf l = "Helland Linear" Then
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & "-1  0 -1" & vbCrLf & " 0  4  0" & vbCrLf & "-1  0 -1"
            ReDim FM(-1 To 1, -1 To 1) As Long
            FM(-1, -1) = -1
            FM(-1, 1) = -1
            FM(1, -1) = -1
            FM(1, 1) = -1
            FM(0, 0) = 4
            DoFilter Me.picBack, Me.picMain, 3, 1, 0, True
    Else
        lblDesc = "Matrix:" & vbCrLf & vbCrLf & " 0  0  0  1  0" & vbCrLf & " 1  0  0  0  0" & vbCrLf & " 0  0 -4  0  0" & vbCrLf & " 0  0  0  0  1" & vbCrLf & " 0  1  0  0  0"
            ReDim FM(-2 To 2, -2 To 2) As Long
            FM(-1, -2) = 1
            FM(-2, 1) = 1
            FM(1, 2) = 1
            FM(2, -1) = 1
            FM(0, 0) = -4
            DoFilter Me.picBack, Me.picMain, 5, 1, 0, True
    End If
    
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
        lstEdgeOptions_Click
        
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
    minWidth = (frmDescription.Left + frmDescription.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.TOp * 1.3) * Screen.TwipsPerPixelY
    
End Sub

'The omnipotent DoFilter routine - it takes whatever is in FM() and applies it to the image
Public Sub DoFilter(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, ByVal filterSize As Long, ByVal filterWeight As Long, ByVal filterBias As Long, Optional ByVal invertResult As Boolean = False)
    
    'This array will hold the image's original pixel data
    Dim imageData() As Byte
    
    'tData holds the calculated values of the filter (as opposed to the original values)
    Dim tData() As Byte
        
    'Coordinate variables
    Dim x As Long, Y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, imageData()
    fDraw.GetImageData2D srcPic, tData()  'We only do this so that tData() is sized appropriately
    
    'These variables will hold temporary pixel color values
    Dim r As Long, g As Long, b As Long
    
    'C and D are like X and Y - they are additional loop variables used for sub-loops
    Dim c As Long, d As Long
    
    'CalcX and CalcY are temporary width and height values for sub-loops
    Dim calcX As Long, calcY As Long
    
    'CalcVar determines the size of each sub-loop
    Dim calcVar As Long
    calcVar = (filterSize \ 2)
    
    'TempRef is like QuickX below, but for sub-loops
    Dim tempRef As Long
    
    'FilterWeightTemp will be reset for every pixel, and decremented appropriately if
    ' attempting to calculate the value for a pixel outside the perimeter of the image
    Dim filterWeightTemp As Long
    
    'Now that we're ready, run a quick loop through the image, calculating pixel values as we go
    Dim quickX As Long
    For x = 0 To iWidth
        quickX = x * 3
    For Y = 0 To iHeight
        
        'Reset red, green, and blue
        r = 0
        g = 0
        b = 0
        filterWeightTemp = filterWeight
        
        'Run a sub-loop around this pixel
        For c = x - calcVar To x + calcVar
            tempRef = c * 3
        For d = Y - calcVar To Y + calcVar
            
            calcX = c - x
            calcY = d - Y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming, but they ARE convenient :)
            If FM(calcX, calcY) = 0 Then GoTo 7
            
            'If this pixel lies outside the image perimeter, ignore it and adjust FilterWeight accordingly
            If c < 0 Or d < 0 Or c > iWidth Or d > iHeight Then
                filterWeightTemp = filterWeight - FM(calcX, calcY)
                GoTo 7
            End If
            
            'Adjust red, green, and blue according to the values in the filter matrix (FM)
            r = r + (imageData(tempRef + 2, d) * FM(calcX, calcY))
            g = g + (imageData(tempRef + 1, d) * FM(calcX, calcY))
            b = b + (imageData(tempRef, d) * FM(calcX, calcY))

7       Next d
        Next c
        
        'If a weight has been set, apply it now
        If filterWeight <> 1 Then
            r = r \ filterWeightTemp
            g = g \ filterWeightTemp
            b = b \ filterWeightTemp
        End If
        
        'If a bias has been specified, apply it now
        If filterBias <> 0 Then
            r = r + filterBias
            g = g + filterBias
            b = b + filterBias
        End If
        
        'Make sure all values are between 0 and 255
        ByteMe r
        ByteMe g
        ByteMe b
        
        'If inversion is specified, apply it now
        If invertResult Then
            r = 255 - r
            g = 255 - g
            b = 255 - b
        End If
        
        'Finally, remember the new value in our tData array
        tData(quickX, Y) = CByte(b)
        tData(quickX + 1, Y) = CByte(g)
        tData(quickX + 2, Y) = CByte(r)
        
    Next Y
    Next x
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth + 1, iHeight + 1, tData()
    
End Sub

'Convert to absolute byte values (Long-type)
Public Sub ByteMe(ByRef tempVar As Long)
    If (tempVar > 255) Then tempVar = 255
    If (tempVar < 0) Then tempVar = 0
End Sub
