VERSION 5.00
Begin VB.Form FrmResize 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resize Map"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3840
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
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   256
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox chkPreserve 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preserve Map Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Select this if you don't want your current map data erased"
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox TxtYSize 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "20"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox TxtXSize 
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
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "20"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: these measurements represent TILES, not PIXELS"
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Map Height:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Map Width:"
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmResize"
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

'OK Button
Private Sub cmdOK_Click()

    'If the map data will NOT be preserved...
    If chkPreserve.Value = 0 Then
        
        'Resize the array (and consequently erase all map data)
        SizeX = Val(TxtXSize) - 1
        SizeY = Val(TxtYSize) - 1
        ReDim MapArray(0 To SizeX, 0 To SizeY) As Byte
    
    'If the map data WILL be preserved...
    Else
        
        'This array will temporarily hold all of our existing map data
        Dim TempArray() As Byte
        
        'If the new map size will be smaller than the old map size, shrink the temporary array accordingly
        If TxtXSize.Text < SizeX Then SizeX = Val(TxtXSize) - 1
        If TxtYSize.Text < SizeY Then SizeY = Val(TxtYSize) - 1
        
        ReDim TempArray(0 To SizeX, 0 To SizeY) As Byte
        
        'Run a quick loop through the map data, copying points into our temporary array
        For x = 0 To SizeX
        For Y = 0 To SizeY
            TempArray(x, Y) = MapArray(x, Y)
        Next Y
        Next x
        
        'Now resize the map itself to the new size...
        ReDim MapArray(0 To SizeX, 0 To SizeY) As Byte
        
        '...and copy all the map data back
        For x = 0 To SizeX
        For Y = 0 To SizeY
            MapArray(x, Y) = TempArray(x, Y)
        Next Y
        Next x
        
    End If
    
    'Set the new size values and refresh everything
    Me.Visible = False
    RefreshAll
    Unload Me
    
End Sub

'Cancel Button
Private Sub cmdCancel_Click()
    Me.Visible = False
    RefreshAll
    Unload Me
End Sub

