VERSION 5.00
Begin VB.Form frmScanner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "VB6 Scanner Interface by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameScanOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Scan Options:"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   2895
      Begin VB.OptionButton optLoadImage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Load image at original size"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2655
      End
      Begin VB.OptionButton optLoadImage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Resize image to fit the picture box on the right"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picScan 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   3480
      ScaleHeight     =   463
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   463
      TabIndex        =   2
      Top             =   360
      Width           =   6975
   End
   Begin VB.CommandButton cmdScanImage 
      Caption         =   "Step 2: Scan Image"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdSelectScanner 
      Caption         =   "Step 1: Select Scanner (optional)"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(If a scanner is not manually selected, the program will use the default system scanner)"
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblSuccess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Image acquired successfully!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Copyright 2018 by Tanner Helland
' www.tannerhelland.com
'
'This project includes code for basic scanner interaction.  It's primarily an interface to the
' "EZTW32" DLL file, which is required because VB does not have native scanner support.
'
'The EZTW32 library is a free, public domain TWAIN32-compliant library.  You can learn more
' about it at http://eztwain.com/
'
'This project was designed against v1.19 of the EZTW32 library (2009.02.22).  It may not work with
' other versions of the library.  Additional documentation regarding the use of EZTW32 is
' available from the EZTW32 developers at http://eztwain.com/ezt1_download.htm
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

'EZTW32 functions for interfacing with a scanner
Private Declare Function TWAIN_AcquireToFilename Lib "EZTW32.dll" (ByVal hwndApp As Long, ByVal sFile As String) As Long
Private Declare Function TWAIN_SelectImageSource Lib "EZTW32.dll" (ByVal hwndApp As Long) As Long
Private Declare Function TWAIN_IsAvailable Lib "EZTW32.dll" () As Long

'API calls for explicitly allocating DLLs.  These allow you to load a DLL from an arbitrary location
' (i.e. without manually copying it into your system directory)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'API calls for resizing an image (faster and more accurate than VB's PaintPicture function)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal ClipX As Long, ByVal ClipY As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As Long) As Long
Private Const STRETCHBLT_HALFTONE As Long = 4

'Used to store the memory address of the EZTW32 library once it's been loaded
Dim hLib As Long

'Location of the project's folder (used to find the required DLL)
Dim ProgramPath As String

'Whether or not a valid scanner is available for use
Dim ScanEnabled As Boolean

Private Sub cmdScanImage_Click()

    'If this is not the first scan the user has performed, hide the "successful scan" label
    lblSuccess.Visible = False

    'Make sure that the scanner was successfully found before attempting to scan an image
    If (Not ScanEnabled) Then
        MsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into this application directory and reload the program.", vbCritical + vbOKOnly + vbApplicationModal, "Scanner Interface Error"
        Exit Sub
    End If

    'Basic scanner capture code
    On Error GoTo ScanError
    Dim ScannerCaptureFile As String, ScanCheck As Long
    'ScanCheck is used to store the return values of the EZTW32.dll scanner functions.  We start by setting it
    ' to an arbitrary value that only we know; if an error occurs and this value is still present, it means an
    ' error occurred outside of the EZTW32 library.
    ScanCheck = -5
    
    'A temporary file is required by the scanner; we will place it in the project folder, then delete it when finished
    ScannerCaptureFile = ProgramPath & "VBScanInterface.bmp"
    
    'This line uses the EZTW32.dll file to scan the image and send it to a temporary file
    ScanCheck = TWAIN_AcquireToFilename(frmScanner.hWnd, ScannerCaptureFile)
    
    'If the image was successfully scanned, load it into the available picture box
    If ScanCheck = 0 Then
        picBuffer.Picture = LoadPicture(ScannerCaptureFile)
        
        'Be polite and remove the temporary file acquired from the scanner
        Kill ScannerCaptureFile
        
        'Resize the image if the user has requested it
        If optLoadImage(0).Value Then
            '...no resize
            picScan.AutoSize = True
            picScan.Picture = picBuffer.Picture
            picScan.Refresh
        Else
            '...resize
            picScan.AutoSize = False
            
            'Setting the HALFTONE mode results in a slower, but higher-quality image resize
            SetStretchBltMode picScan.hDC, STRETCHBLT_HALFTONE
            StretchBlt picScan.hDC, 0, 0, picScan.ScaleWidth, picScan.ScaleHeight, picBuffer.hDC, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, vbSrcCopy
            picScan.Picture = picScan.Image
            picScan.Refresh
        End If
        
        'Let the user know that everything worked
        lblSuccess.Visible = True
        
        'Return focus to this project
        frmScanner.SetFocus
    Else
        'If the scan was unsuccessful, let the user know what happened
        GoTo ScanError
    End If
    
    Exit Sub

'Something went wrong
ScanError:
    
    Dim scanErrMessage As String
    
    Select Case ScanCheck
        Case -5
            scanErrMessage = "Unknown error occurred."
        Case -4
            scanErrMessage = "Scan successful, but temporary file save failed.  Is it possible that your hard drive is full (or almost full)?"
        Case -3
            scanErrMessage = "Unable to acquire DIB lock.  Please make sure no other programs are accessing the scanner.  If the problem persists, reboot and try again."
        Case -2
            scanErrMessage = "Temporary file access error.  This can be caused when running on a system with limited access rights.  Please enable admin rights and try again."
        Case -1
            scanErrMessage = "Scan failed.  (This may happen if the user canceled the scan dialog.)"
        Case Else
            scanErrMessage = "The scanner returned an error code that wasn't specified in the EZTW32.dll documentation.  Please visit http://www.eztwain.com for more information."
    End Select
    
    MsgBox scanErrMessage, vbCritical + vbOKOnly + vbApplicationModal, "Scanner Interface Error"

End Sub

'This button allows the user to select which scanner they want to use.  If the user does not select a scanner, the default system scanner will be used.
Private Sub cmdSelectScanner_Click()
    
    'Only launch the scanner select form if the EZTW32.dll file was successfully found at program load
    If ScanEnabled Then
        TWAIN_SelectImageSource (frmScanner.hWnd)
    Else
    'If the EZTW32.dll file doesn't exist...
        MsgBox "The scanner/digital camera interface plug-in (EZTW32.dll) was marked as missing upon program initialization." & vbCrLf & vbCrLf & "To enable scanner support, please copy the EZTW32.dll file (available for download from http://eztwain.com/ezt1_download.htm) into this application directory and reload the program.", vbCritical + vbOKOnly + vbApplicationModal, "Scanner Interface Error"
        Exit Sub
    End If
    
End Sub

'When the project is first loaded, several parameters need to be determined.
Private Sub Form_Load()
    
    'Remember the folder from which the program was launched
    ProgramPath = App.Path
    If Right(ProgramPath, 1) <> "\" Then ProgramPath = ProgramPath & "\"

    'This code allows us to load the scanner DLL from any location; this is easier than forcing the user to
    ' copy the file into their system directory.  However, it requires a bit of extra code.
    hLib = LoadLibrary(ProgramPath & "EZTW32.dll")

    If TWAIN_IsAvailable() = 0 Then
        ScanEnabled = False
        MsgBox "Unfortunately, this program was unable to locate the DLL file necessary for scanner support.  Please ensure that a file called EZTW32.dll is located in the same directory as this project file, then restart the project.  Thanks!", vbOKOnly + vbCritical, "Scanner DLL missing"
    Else
        ScanEnabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Unload the scanner DLL from memory.
    FreeLibrary hLib
    
End Sub
