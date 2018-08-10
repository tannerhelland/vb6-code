VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Basic Tile Map Editor - www.tannerhelland.com"
   ClientHeight    =   8775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbDrawMethod 
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.HScrollBar HSMain 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   8400
      Width           =   12135
   End
   Begin VB.VScrollBar VSMain 
      Height          =   6735
      Left            =   12360
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.HScrollBar HSTileBar 
      Height          =   255
      Left            =   120
      Max             =   5
      TabIndex        =   4
      Top             =   1200
      Width           =   12495
   End
   Begin VB.PictureBox PicTiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   120
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   834
      TabIndex        =   0
      Top             =   120
      Width           =   12540
   End
   Begin VB.PictureBox PicTilesBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      Picture         =   "Main Editor.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   15390
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   4560
      Width           =   990
   End
   Begin VB.PictureBox MainPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   447
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   807
      TabIndex        =   2
      Top             =   1560
      Width           =   12135
      Begin VB.PictureBox Pic2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   360
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   7
         Top             =   3360
         Width           =   990
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuNew 
         Caption         =   "&New Map"
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open Map"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save Current Map"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnuResize 
         Caption         =   "Resize the Map"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuRefresh 
         Caption         =   "&Refresh Screen"
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Change &Zoom"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "Main"
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


'When the combo box is changed, update our drawing method
Private Sub cmbDrawMethod_Change()
    
    Select Case cmbDrawMethod.ListIndex
        Case 0
            UseStretchBlt = False
        Case 1
            UseStretchBlt = True
    End Select
    
End Sub

Private Sub cmbDrawMethod_Click()

    Select Case cmbDrawMethod.ListIndex
        Case 0
            UseStretchBlt = False
        Case 1
            UseStretchBlt = True
    End Select
    
End Sub

'The InitializeEditor routine can be found in Sub_Module
Private Sub Form_Load()
    
    InitializeEditor

    'Also, populate the DrawMethod combo box
    cmbDrawMethod.AddItem "Use PaintPicture"
    cmbDrawMethod.AddItem "Use StretchBlt"
    cmbDrawMethod.ListIndex = 0
    UseStretchBlt = False
    
    'Load a default map
    On Error GoTo MapMissing
    
    Open App.Path & "\Demo.map" For Binary As #1
        'At present, map files are pretty simple - the width, height, and data is all we currently store
        Get #1, 1, SizeX
        Get #1, , SizeY
        ReDim MapArray(0 To SizeX, 0 To SizeY) As Byte
        Get #1, , MapArray
    Close #1
    
    'Once MapArray has been successfully created, Draw the map!
    DrawMap
    
MapMissing:

End Sub

'Whenever a scrollbar is used, update the map
Private Sub HSMain_Change()
    ScrollMap
End Sub

Private Sub HSMain_Scroll()
    ScrollMap
End Sub

Private Sub VSMain_Change()
    ScrollMap
End Sub

Private Sub VSMain_Scroll()
    ScrollMap
End Sub

Private Sub HSTileBar_Change()
    DrawTileBar
End Sub

'Whenever the top scrollbar is used, update the tiles
Private Sub HSTileBar_Scroll()
    DrawTileBar
End Sub

'When a mouse button is pressed, remember the mouse state (1 = pressed) and draw a tile at the current mouse position
Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseState = True
    'Left button
    If Button = 1 Then DrawTile x, Y, CurPicIndex
    'Right button
    If Button = 2 Then DrawTile x, Y, CurPicIndex2
End Sub

'When the mouse is moved, only draw if a mouse button has been pressed
Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Only draw if a button is down
    If MouseState Then
        If Button = 1 Then DrawTile x, Y, CurPicIndex
        If Button = 2 Then DrawTile x, Y, CurPicIndex2
    End If
End Sub

'Return the mouse button state variable to unpressed
Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseState = False
End Sub

'Menu -> Exit
Private Sub MnuExit_Click()
    Unload Me
End Sub

'Menu -> New
Private Sub MnuNew_Click()
    
    'Keep the current size settings and reset the map to zero (default tile)
    For x = 0 To SizeX
    For Y = 0 To SizeY
        MapArray(x, Y) = 0
    Next Y
    Next x
    
    'Last, draw the new map
    DrawMap
    
End Sub

'Menu -> Open
Private Sub MnuOpen_Click()
    
    'Initialize the common dialog interface
    Dim newDialog As pdOpenSaveDialog
    Set newDialog = New pdOpenSaveDialog
    
    'String returned from the common dialog wrapper
    Dim sFile As String
    
    'This string contains the filters for loading different kinds of images.  Using
    'this feature correctly makes our common dialog box a LOT more pleasant to use.
    Dim cdfStr As String
    cdfStr = "VB Map Files (.map)|*.map|All Files|*.*"
    
    'If cancel isn't selected, load a map from the user-specified file
    If newDialog.GetOpenFileName(sFile, , True, , cdfStr, 1, , "Open a map file", , Me.hWnd) Then
        
        'Open the file
        Open sFile For Binary As #1
            'At present, map files are pretty simple - the width, height, and data is all we currently store
            Get #1, 1, SizeX
            Get #1, , SizeY
            ReDim MapArray(0 To SizeX, 0 To SizeY) As Byte
            Get #1, , MapArray
        Close #1
        
        'Once MapArray has been successfully created, Draw the map!
        DrawMap
        
    End If
    

End Sub

'Menu -> Refresh
Private Sub MnuRefresh_Click()
    RefreshAll
End Sub

'Menu -> Resize
Private Sub MnuResize_Click()
    FrmResize.Visible = True
End Sub

'Menu -> Save
Private Sub MnuSave_Click()

    'Initialize the common dialog interface
    Dim newDialog As pdOpenSaveDialog
    Set newDialog = New pdOpenSaveDialog
    
    'String returned from the common dialog wrapper
    Dim sFile As String
    
    'This string contains the filters for loading different kinds of images.  Using
    'this feature correctly makes our common dialog box a LOT more pleasant to use.
    Dim cdfStr As String
    cdfStr = "VB Map Files (.map)|*.map|All Files|*.*"
    
    'If cancel isn't selected, save a map to the user-specified file
    If newDialog.GetSaveFileName(sFile, , , cdfStr, 1, , "Save a map file", ".map", Me.hWnd) Then
        
        'Open the file
        Open sFile For Binary As #1
            'At present, map files are pretty simple - the width, height, and data is all we currently store
            Put #1, 1, SizeX
            Put #1, , SizeY
            Put #1, , MapArray
        Close #1
        
    End If

End Sub

'Menu -> Zoom
Private Sub mnuZoom_Click()
    ChangeZoom
End Sub

'When the top tile box is clicked, remember which tile was clicked (and draw that tile in its preview box)
Private Sub PicTiles_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    'Left button...
    If Button = 1 Then
        'Current picture is the x value integer divided by 64 (tile width) + whatever the scrollbar is at
        CurPicIndex = (x \ 64) + HSTileBar.Value
        'Draw this tile in its picture box on the left
        If UseStretchBlt Then
            StretchBlt Pic1.hdc, 0, 0, 64, 64, PicTilesBuffer.hdc, CurPicIndex * 64, 0, 64, 64, vbSrcCopy
        Else
            Pic1.PaintPicture PicTilesBuffer.Picture, 0, 0, 64, 64, CurPicIndex * 64, 0, 64, 64, vbSrcCopy
        End If
    'Right button...
    ElseIf Button = 2 Then
        CurPicIndex2 = (x \ 64) + HSTileBar.Value
        If UseStretchBlt Then
            StretchBlt Pic2.hdc, 0, 0, 64, 64, PicTilesBuffer.hdc, CurPicIndex2 * 64, 0, 64, 64, vbSrcCopy
        Else
            Pic2.PaintPicture PicTilesBuffer.Picture, 0, 0, 64, 64, CurPicIndex2 * 64, 0, 64, 64, vbSrcCopy
        End If
    End If
    
    'Because AutoRedraw is set to true for these pictures boxes, we must refresh them and update the .Picture property
    Pic1.Picture = Pic1.Image
    Pic2.Picture = Pic2.Image
    Pic1.Refresh
    Pic2.Refresh

End Sub

