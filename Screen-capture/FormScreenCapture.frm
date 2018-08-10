VERSION 5.00
Begin VB.Form FormScreenCapture 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Screen Capture Demo - www.tannerhelland.com"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Capture Demo -- www.tannerhelland.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Capture Demo -- www.tannerhelland.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   0
      Top             =   255
      Width           =   3615
   End
End
Attribute VB_Name = "FormScreenCapture"
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

'This gives us the hWnd (window handle) of the screen
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'This assigns an hDC (device context handle) from an hWnd
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

'BitBlt lets us draw an image from a hDC to another hDC (in our case, from an hDC of the screen capture to the hDC of a VB picture box)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal opCode As Long) As Long

'ReleaseDC is used to free the hDC we generate for the screen capture.
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'This sample project copies the screen when the form loads; you could also place this code in a command button (or any other input)
Private Sub Form_Load()

    'First, minimize this window
    Me.WindowState = vbMinimized
    
    'Get the hWnd of the screen
    Dim scrHwnd As Long
    scrHwnd = GetDesktopWindow
    
    'Now, assign an hDC to the hWnd we generated
    Dim shDC As Long
    shDC = GetDC(scrHwnd)
    
    'Determine the size of the screen
    Dim screenWidth As Long, screenHeight As Long
    screenWidth = Screen.Width \ Screen.TwipsPerPixelX
    screenHeight = Screen.Height \ Screen.TwipsPerPixelY
        
    'Copy the pixel data from the screen into our form
    BitBlt Me.hDC, 0, 0, screenWidth, screenHeight, shDC, 0, 0, vbSrcCopy
        
    'Release our hold on the screen DC
    ReleaseDC scrHwnd, shDC
    
    Me.Picture = Me.Image
    
    'Restore the window
    Me.WindowState = vbNormal
    
End Sub

