VERSION 5.00
Begin VB.Form frmCustomFilters 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Filters - www.tannerhelland.com"
   ClientHeight    =   7095
   ClientLeft      =   45
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   473
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFilter 
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
      Height          =   345
      Left            =   3240
      TabIndex        =   34
      Text            =   "Select a filter..."
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   29
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   28
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   27
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   26
      Text            =   "0"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   720
      TabIndex        =   24
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   23
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   22
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   21
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   720
      TabIndex        =   19
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   1320
      TabIndex        =   18
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   1920
      TabIndex        =   17
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   14
      Left            =   2520
      TabIndex        =   16
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   16
      Left            =   720
      TabIndex        =   14
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   1320
      TabIndex        =   13
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   18
      Left            =   1920
      TabIndex        =   12
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   19
      Left            =   2520
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   20
      Left            =   120
      TabIndex        =   10
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   21
      Left            =   720
      TabIndex        =   9
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   22
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   23
      Left            =   1920
      TabIndex        =   7
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtF 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   24
      Left            =   2520
      TabIndex        =   6
      Text            =   "0"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox TxtWeight 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdApply 
      Appearance      =   0  'Flat
      Caption         =   "Apply Filter"
      Default         =   -1  'True
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
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   120
      Width           =   2805
   End
   Begin VB.TextBox TxtBias 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "0"
      Top             =   2040
      Width           =   495
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
      Left            =   3240
      TabIndex        =   2
      Top             =   720
      Width           =   2775
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
      Top             =   2400
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
      Top             =   2400
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label lblCommon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Filters:"
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
      Height          =   225
      Left            =   3240
      TabIndex        =   33
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label lblScale 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scale:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   32
      Top             =   2085
      Width           =   480
   End
   Begin VB.Label lblBias 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bias:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1680
      TabIndex        =   31
      Top             =   2085
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenImage 
         Caption         =   "&Open image"
      End
   End
End
Attribute VB_Name = "frmCustomFilters"
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

'The combo box can be used to pre-populate the text boxes with some common filters
Private Sub cmbFilter_Click()
    
    'Loop variables
    Dim x As Long
    
    'Reset the text boxes to default values
    For x = 0 To 24
        TxtF(x) = 0
    Next x
    TxtWeight = 1
    TxtBias = 0
    
    Select Case cmbFilter.ListIndex
        'Blur
        Case 0
            TxtF(6) = 1
            TxtF(7) = 1
            TxtF(8) = 1
            TxtF(11) = 1
            TxtF(12) = 1
            TxtF(13) = 1
            TxtF(16) = 1
            TxtF(17) = 1
            TxtF(18) = 1
            TxtWeight = 9
        'Sharpen
        Case 1
            TxtF(7) = -1
            TxtF(11) = -1
            TxtF(12) = 5
            TxtF(13) = -1
            TxtF(17) = -1
        'Emboss
        Case 2
            TxtF(11) = -1
            TxtF(13) = 1
            TxtBias = 127
        'Engrave
        Case 3
            TxtF(11) = 1
            TxtF(13) = -1
            TxtBias = 127
        'Grease
        Case 4
            TxtF(6) = 1
            TxtF(7) = 1
            TxtF(8) = 1
            TxtF(11) = 1
            TxtF(12) = -7
            TxtF(13) = 1
            TxtF(16) = 1
            TxtF(17) = 1
            TxtF(18) = 1
        'Unfocus
        Case 5
            TxtF(6) = 1
            TxtF(7) = 1
            TxtF(8) = 1
            TxtF(11) = 1
            TxtF(12) = -4
            TxtF(13) = 1
            TxtF(16) = 1
            TxtF(17) = 1
            TxtF(18) = 1
            TxtWeight = 4
        'Vibrate
        Case 6
            TxtF(0) = 1
            TxtF(4) = -1
            TxtF(6) = -1
            TxtF(8) = 1
            TxtF(12) = 1
            TxtF(16) = 1
            TxtF(18) = -1
            TxtF(20) = -1
            TxtF(24) = 1
    End Select
    
    UpdateEffect
    
End Sub

'Apply the currently generated filter to the image
Private Sub cmdApply_Click()
    UpdateEffect
End Sub

'Reset the foreground image to look like the original image
Private Sub cmdReset_Click()
    LoadImageAutosized vbNullString
End Sub

Private Sub Form_Load()
    
    cmbFilter.Clear
    cmbFilter.AddItem "Blur"
    cmbFilter.AddItem "Sharpen"
    cmbFilter.AddItem "Emboss"
    cmbFilter.AddItem "Engrave"
    cmbFilter.AddItem "Grease"
    cmbFilter.AddItem "Unfocus"
    cmbFilter.AddItem "Vibrate"
    cmbFilter.ListIndex = 2
    
    'Draw the initial effect with a default value, then update the image
    If (App.LogMode = 0) Then Me.Caption = Me.Caption & " - compile for significantly better performance!"
    LoadImageAutosized App.Path & "\sample.jpg"
    UpdateEffect
    Me.Show
    
End Sub

'To make editing easier, automatically select all text when one of the text boxes is selected
Private Sub TxtBias_GotFocus()
    AutoSelectText TxtBias
End Sub

Private Sub TxtF_GotFocus(Index As Integer)
    AutoSelectText TxtF(Index)
End Sub

Private Sub TxtWeight_GotFocus()
    AutoSelectText TxtWeight
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
    minWidth = (cmdApply.Left + cmdApply.Width + 20) * Screen.TwipsPerPixelX
    If (imgWidth < (Screen.Width \ Screen.TwipsPerPixelX) - 50) Then
        newWidth = (imgWidth + picMain.Left * 4) * Screen.TwipsPerPixelX
        If (newWidth < minWidth) Then newWidth = minWidth
        Me.Width = newWidth
    End If
    If (imgHeight < (Screen.Height \ Screen.TwipsPerPixelY) - 50) Then Me.Height = (imgHeight + picMain.Top * 1.4) * Screen.TwipsPerPixelY
    
End Sub

'The omnipotent DoFilter routine - it takes whatever is in FM() and applies it to the image
Public Sub DoFilter(ByRef srcPic As PictureBox, ByRef dstPic As PictureBox, ByVal filterSize As Long, ByVal filterWeight As Long, ByVal filterBias As Long, Optional ByVal invertResult As Boolean = False)
    
    'This array will hold the image's original pixel data
    Dim ImageData() As Byte
    
    'tData holds the calculated values of the filter (as opposed to the original values)
    Dim tData() As Byte
        
    'Coordinate variables
    Dim x As Long, y As Long
    
    'Image dimensions
    Dim iWidth As Long, iHeight As Long
    
    'Instantiate a FastDrawing class and gather the image's data (into ImageData())
    Dim fDraw As New FastDrawing
    iWidth = fDraw.GetImageWidth(srcPic) - 1
    iHeight = fDraw.GetImageHeight(srcPic) - 1
    fDraw.GetImageData2D srcPic, ImageData()
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
    
    For y = 0 To iHeight
    For x = 0 To iWidth
    
        quickX = x * 3
        
        'Reset red, green, and blue
        r = 0
        g = 0
        b = 0
        filterWeightTemp = filterWeight
        
        'Run a sub-loop around this pixel
        For c = x - calcVar To x + calcVar
            tempRef = c * 3
        For d = y - calcVar To y + calcVar
            
            calcX = c - x
            calcY = d - y
            
            'If no filter value is being applied to this pixel, ignore it (GoTo's aren't generally a part of good programming, but they ARE convenient :)
            If FM(calcX, calcY) = 0 Then GoTo 7
            
            'If this pixel lies outside the image perimeter, ignore it and adjust FilterWeight accordingly
            If (c < 0) Or (d < 0) Or (c > iWidth) Or (d > iHeight) Then
                filterWeightTemp = filterWeight - FM(calcX, calcY)
                GoTo 7
            End If
            
            'Adjust red, green, and blue according to the values in the filter matrix (FM)
            r = r + (ImageData(tempRef + 2, d) * FM(calcX, calcY))
            g = g + (ImageData(tempRef + 1, d) * FM(calcX, calcY))
            b = b + (ImageData(tempRef, d) * FM(calcX, calcY))

7       Next d
        Next c
        
        'If a weight has been set, apply it now
        If (filterWeight <> 1) Then
            r = r \ filterWeightTemp
            g = g \ filterWeightTemp
            b = b \ filterWeightTemp
        End If
        
        'If a bias has been specified, apply it now
        If (filterBias <> 0) Then
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
        tData(quickX, y) = CByte(b)
        tData(quickX + 1, y) = CByte(g)
        tData(quickX + 2, y) = CByte(r)
        
    Next x
    Next y
    
    'Draw the new image data to the screen
    fDraw.SetImageData2D dstPic, iWidth + 1, iHeight + 1, tData()
    
End Sub

'Convert to absolute byte values (Long-type)
Private Sub ByteMe(ByRef TempVar As Long)
    If (TempVar > 255) Then TempVar = 255
    If (TempVar < 0) Then TempVar = 0
End Sub

'Pass this a text box and it will select all text currently in the text box
Private Function AutoSelectText(ByRef tBox As TextBox)
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

'Confirm that the variable passed in is a valid number
Public Function NumberValid(ByVal check As Variant) As Boolean
    If Not IsNumeric(check) Then
        MsgBox check & " is not a valid entry.  Please enter a numeric value.", vbCritical + vbOKOnly + vbApplicationModal, App.Title
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

Private Sub UpdateEffect()
    
    'Loop variables
    Dim x As Long, y As Long
    
    'Before we do anything else, check to make sure every text box contains a valid number
    For x = 0 To 24
        If Not NumberValid(TxtF(x)) Then
            AutoSelectText TxtF(x)
            Exit Sub
        End If
    Next x
    If Not NumberValid(TxtWeight) Then
        AutoSelectText TxtWeight
        Exit Sub
    End If
    If Not NumberValid(TxtBias) Then
        AutoSelectText TxtBias
        Exit Sub
    End If
    
    'Copy the values from the text boxes into our FM() array
    ReDim FM(-2 To 2, -2 To 2) As Long
    For x = -2 To 2
    For y = -2 To 2
        FM(x, y) = Val(TxtF((x + 2) + (y + 2) * 5))
    Next y
    Next x
    
    DoFilter Me.picBack, Me.picMain, 5, Val(TxtWeight.Text), Val(TxtBias.Text), False
    
End Sub
