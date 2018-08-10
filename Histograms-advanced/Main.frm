VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Image Histograms Demo 2 - www.tannerhelland.com"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6255
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
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset image"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame frmHistogramFunctions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Histogram Functions"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6015
      Begin VB.CommandButton cmdEqualizeLuminance 
         Appearance      =   0  'Flat
         Caption         =   "Equalize luminance"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkBlue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Equalize blue"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkGreen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Equalize green"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkRed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Equalize red"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CommandButton cmdEqualize 
         Appearance      =   0  'Flat
         Caption         =   "Equalize individual colors (as selected on the right)"
         Height          =   975
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdStretchHistogram 
         Appearance      =   0  'Flat
         Caption         =   "Stretch histogram"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   2040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   240
         Y2              =   1560
      End
   End
   Begin VB.CommandButton cmdDispHistogram 
      Caption         =   "Display histogram"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox PicMain 
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
      Top             =   2400
      Width           =   6030
   End
   Begin VB.PictureBox PicBack 
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
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpenImage 
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


'The histogram information is displayed on a separate form (launch from button)
Private Sub cmdDispHistogram_Click()
    frmHistogram.Show
End Sub

'Equalize the histogram using individual color components
Private Sub cmdEqualize_Click()
    frmHistogram.EqualizeHistogram chkRed.Value, chkGreen.Value, chkBlue.Value
End Sub

'Equalize the histogram using luminance only
Private Sub cmdEqualizeLuminance_Click()
    frmHistogram.EqualizeLuminance
End Sub

Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

'Stretch the histogram
Private Sub cmdStretchHistogram_Click()
    frmHistogram.StretchHistogram
End Sub

'Automatically initialized
Private Sub Form_Load()
    
    'Upon loading the form, reset two variables:
    
    'Luminance is the default histogram source
    lastHistSource = DRAWMETHOD_LUMINANCE
    
    'Line graph is the default drawing option
    lastHistMethod = DRAWMETHOD_BARS
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    Me.Show
    
    frmHistogram.Show 0, Me
    
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

    If (LenB(srcFilePath) <> 0) Then PicBack.Picture = LoadPicture(srcFilePath)
    
    'Copy the image, automatically resized, from the background picture box to the foreground one
    Dim fDraw As FastDrawing
    Set fDraw = New FastDrawing
    
    Dim imageData() As Byte, imgWidth As Long, imgHeight As Long
    imgWidth = fDraw.GetImageWidth(Me.PicBack)
    imgHeight = fDraw.GetImageHeight(Me.PicBack)
    fDraw.GetImageData2D Me.PicBack, imageData()
    
    Me.PicMain.Width = imgWidth + 2
    Me.PicMain.Height = imgHeight + 2
    fDraw.SetImageData2D Me.PicMain, imgWidth, imgHeight, imageData()
    
    'Resize the form to automatically contain the new picture box dimensions.
    ' (This performs some sloppy checks to keep the form from resizing larger than the primary display,
    '  but it is *not* good code - a proper solution would involve the AdjustWindowRect() API or similar!)
    Dim newWidth As Long, minWidth As Long
    minWidth = (frmHistogramFunctions.Left + frmHistogramFunctions.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + PicMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + PicMain.Top * 1.5) * Screen.TwipsPerPixelY
    
End Sub
