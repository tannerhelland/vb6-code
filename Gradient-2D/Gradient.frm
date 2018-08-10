VERSION 5.00
Begin VB.Form frmGradient 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Gradient Demo - www.tannerhelland.com"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbStyle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Text            =   "Horizontal"
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox picColor2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
      Height          =   735
      Left            =   480
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.PictureBox picColor1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
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
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   705
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   $"Gradient.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmGradient"
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


'Any time the combo box is changed, redraw the gradient
Private Sub cmbStyle_Change()
    DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
End Sub

Private Sub cmbStyle_Click()
    DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
End Sub

'Generate the combo box options upon form load
Private Sub Form_Load()

    cmbStyle.AddItem "Horizontal"
    cmbStyle.AddItem "Vertical"
    cmbStyle.ListIndex = 1
    
End Sub

'When the user clicks the form, flip the gradient 180 degrees
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Use a temporary variable to store the value of the first color
    Dim tempColor As Long
    tempColor = picColor1.BackColor
    'Change the first color to the second color
    picColor1.BackColor = picColor2.BackColor
    'Change the second color to the first color
    picColor2.BackColor = tempColor
    'Redraw the gradient
    DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
End Sub

'Redraw the gradient if the form is resized
Private Sub Form_Resize()
    DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
End Sub

'When the user clicks on one of the colored picture boxes, use the CommonDialog
'control to let them select a new color.  After the color has been selected,
'redraw the new gradient
Private Sub picColor1_Click()
    
    Dim showColor As cSystemColorDialog
    Set showColor = New cSystemColorDialog
    
    Dim newColor As Long
    If showColor.ShowColorDialog(newColor, Me.hWnd, True, picColor1.BackColor) Then
        picColor1.BackColor = newColor
        DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
    End If
    
End Sub

Private Sub picColor2_Click()
    
    Dim showColor As cSystemColorDialog
    Set showColor = New cSystemColorDialog
    
    Dim newColor As Long
    If showColor.ShowColorDialog(newColor, Me.hWnd, True, picColor2.BackColor) Then
        picColor2.BackColor = newColor
        DrawGradient frmGradient, picColor1.BackColor, picColor2.BackColor, cmbStyle.ListIndex
    End If
    
End Sub

'This gradient subroutine could be easily changed for drawing on picture boxes
'(and with several API calls, on other objects as well!)
Public Sub DrawGradient(ByRef dstObject As Form, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Direction As Long)
    
    'R, G, B variables for each end of the gradient
    Dim R As Long, G As Long, B As Long
    Dim R2 As Long, G2 As Long, B2 As Long
    
    'Fill our RGB values from the longs supplied by the calling routine
    R = Color1 Mod 256
    G = (Color1 \ 256) And 255
    B = (Color1 \ 65536) And 255
    R2 = Color2 Mod 256
    G2 = (Color2 \ 256) And 255
    B2 = (Color2 \ 65536) And 255
    
    'Always use variables for storing object values - it's loads faster
    Dim TempWidth As Long, TempHeight As Long
    TempWidth = dstObject.ScaleWidth
    TempHeight = dstObject.ScaleHeight
    
    'Several calculation variables for generating the gradient
    Dim VR As Single, VG As Single, VB As Single
    
    'Vertical gradient
    If Direction = 1 Then
        
        'First, create a calculation variable for determining the step
        'between each level of the gradient (large if the destination form
        'is small, small if the destination form is large); for example, this
        'value will be exactly 1 for each variable if the form is 255 pixels
        'tall and the gradient is going from pure black to pure white
        VR = Abs(R - R2) / TempHeight
        VG = Abs(G - G2) / TempHeight
        VB = Abs(B - B2) / TempHeight
        
        'If the second value is lower then the first value, make the step
        'negative (so that we subtract as we go along, not add)
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < B Then VB = -VB
        
        'Lastly, run a loop through the height of the form, incrementing (or if
        'negative, decrementing) the gradient color according to the y-coordinate
        'of the current line of the form
        For Y = 0 To TempHeight
            R2 = R + VR * Y
            G2 = G + VG * Y
            B2 = B + VB * Y
            dstObject.Line (0, Y)-(TempWidth, Y), RGB(R2, G2, B2)
        Next Y

    'Horizontal gradients work exactly the same, except that they (obviously)
    'run from left-to-right instead of up-and-down
    Else

        VR = Abs(R - R2) / TempWidth
        VG = Abs(G - G2) / TempWidth
        VB = Abs(B - B2) / TempWidth
        
        If R2 < R Then VR = -VR
        If G2 < G Then VG = -VG
        If B2 < B Then VB = -VB
        
        For X = 0 To TempWidth
            R2 = R + VR * X
            G2 = G + VG * X
            B2 = B + VB * X
            dstObject.Line (X, 0)-(X, TempHeight), RGB(R2, G2, B2)
        Next X
    
    End If

End Sub
