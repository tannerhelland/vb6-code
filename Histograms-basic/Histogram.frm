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
    Dim r As Long, g As Long, b As Long, l As Long
    
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
        l = (r + g + b) \ 3
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
        hData(3, l) = hData(3, l) + 1
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
