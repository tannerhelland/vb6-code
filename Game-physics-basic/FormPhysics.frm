VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Simple Game Physics (in VB6) - www.tannerhelland.com"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFramerate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lock Framerate at 60 fps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      TabIndex        =   16
      Top             =   9360
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   15
      Top             =   9360
      Width           =   855
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8940
      Left            =   120
      ScaleHeight     =   594
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   802
      TabIndex        =   10
      Top             =   120
      Width           =   12060
   End
   Begin VB.PictureBox PicBulletM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   9840
      Picture         =   "FormPhysics.frx":0000
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox PicBullet 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   9840
      Picture         =   "FormPhysics.frx":0074
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox PicTM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   6600
      Picture         =   "FormPhysics.frx":00E8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   5520
      Picture         =   "FormPhysics.frx":1E2D
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicRM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   4440
      Picture         =   "FormPhysics.frx":3D10
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   3360
      Picture         =   "FormPhysics.frx":573E
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicLM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   2280
      Picture         =   "FormPhysics.frx":7373
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      Picture         =   "FormPhysics.frx":8E05
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picShipMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   8760
      Picture         =   "FormPhysics.frx":AA12
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox picShip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   7680
      Picture         =   "FormPhysics.frx":C757
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicScreenBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8940
      Left            =   120
      ScaleHeight     =   594
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   802
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   12060
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Use the arrow keys to move the ship Use the space button to fire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   9330
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VELOCITY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   9360
      Width           =   840
   End
   Begin VB.Label LblVel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   9600
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9960
      TabIndex        =   11
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Code Copyright 2018 by Tanner Helland
' www.tannerhelland.com
'
'Documentation for this project can be found at https://tannerhelland.com/code/
'
'Art assets for this demo were borrowed from the old DOS game Raptor: Call of the Shadows
' https://en.wikipedia.org/wiki/Raptor:_Call_of_the_Shadows
' They are used here under Fair Use for educational purposes only.
'
'The source code in this project is licensed under a Simplified BSD license.
' For more information, please review LICENSE.md at https://github.com/tannerhelland/thdc-code/
'
'If you find this code useful, please consider a small donation to https://www.paypal.me/TannerHelland
'
'***************************************************************************

Option Explicit

Private Sub chkFramerate_Click()
    PicMain.SetFocus
End Sub

'START GAME AT FORM LOAD...
Private Sub Form_Load()

    'Upon form load, it's necessary to initialize some key variables
    InitializeGameEngine
    
    'When everything is set, start the game!
    GameActive = True
    MainLoop   'MainLoop appears at the bottom of the form
    
End Sub

'END GAME
Private Sub cmdExit_Click()
    GameActive = False
End Sub

'WHEN A KEY IS PRESSED...
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Depending on the arrow key being pressed, set the correct direction variable
    ' and the correct ship picture
    
    'Left
    If KeyCode = vbKeyLeft Then
        sLeft = 1
        picShip.Picture = PicL.Picture
        picShipMask.Picture = PicLM.Picture
    End If
    
    'Right
    If KeyCode = vbKeyRight Then
        sRight = 1
        picShip.Picture = PicR.Picture
        picShipMask.Picture = PicRM.Picture
    End If
    
    'Up
    If KeyCode = vbKeyUp Then sUp = 1
    
    'Down
    If KeyCode = vbKeyDown Then sDown = 1
    
    'Space (fire the gun)
    If KeyCode = vbKeySpace Then Firing = True
    
    'Escape (end the game)
    If KeyCode = vbKeyEscape Then GameActive = False
    
End Sub

'WHEN A KEY IS RELEASED...
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    'Whenever an arrow key is released, reset motion in that direction
    If KeyCode = vbKeyLeft Then sLeft = 0
    If KeyCode = vbKeyRight Then sRight = 0
    If KeyCode = vbKeyUp Then sUp = 0
    If KeyCode = vbKeyDown Then sDown = 0
    
    'When the space bar is released, stop firing
    If KeyCode = vbKeySpace Then Firing = False
    
End Sub

'MAIN LOOP
Public Sub MainLoop()

    'This variable is used to track time and reduce framerate (if necessary)
    Dim timeDelay As Single
    
    Do While GameActive
        
        If (Not GameActive) Then Exit Do
        
        'Calculate time at the start of the function
        timeDelay = Timer
        
        'Draw the background stars
        DrawStars
        
        'Calculate the ship's motion and location
        VelocityCode
        
        'If necessary, fire some bullets
        FireBullets
        
        'Change the velocity caption
        LblVel.Caption = CStr(50 - sVelVert) & " kps"
        If GameActive Then DoEvents Else Exit Do
        
        'If framerate limitations are active, loop until enough time has passed
        If (chkFramerate.Value = vbChecked) Then
            Do While (GameActive And ((Timer - timeDelay) < 1.66666666666667E-02))
                If GameActive Then DoEvents Else Exit Do
            Loop
        End If
        
    'Then do it all again...
    Loop
    
    'When the loop exits, release the main form
    Unload Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    GameActive = False
End Sub

