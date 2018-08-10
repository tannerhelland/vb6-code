VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Artificial Life Simulator - www.tannerhelland.com"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
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
   ScaleHeight     =   580
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   818
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFront 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7530
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   27
      Top             =   360
      Width           =   7530
   End
   Begin VB.CommandButton cmdSaveData 
      Caption         =   "Save simulation data to file..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   26
      Top             =   8040
      Width           =   2775
   End
   Begin VB.CheckBox chkDisplayDead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Display dead creatures?"
      BeginProperty Font 
         Name            =   "Segoe UI"
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
      TabIndex        =   21
      Top             =   8160
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Frame frmStartSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Initial Simulation Settings:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   7800
      TabIndex        =   7
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox txtMutations 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   24
         Text            =   "15"
         Top             =   3690
         Width           =   1695
      End
      Begin VB.TextBox txtMutateTurns 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         TabIndex        =   22
         Text            =   "750"
         Top             =   3330
         Width           =   2055
      End
      Begin VB.CheckBox chkMultiply 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Allow creatures to multiply and mutate"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin VB.TextBox txtFoodWorth 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   18
         Text            =   "35"
         Top             =   2490
         Width           =   1455
      End
      Begin VB.TextBox txtFoodGen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Text            =   "5"
         Top             =   2130
         Width           =   1695
      End
      Begin VB.TextBox txtFoodRegen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "20"
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox txtInitialEnergy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Text            =   "500"
         Top             =   690
         Width           =   2655
      End
      Begin VB.TextBox txtInitialFood 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "1500"
         Top             =   1410
         Width           =   2295
      End
      Begin VB.TextBox txtOrganisms 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "15"
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New DNA mutates (x) bases:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   2130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multiply every (x) cycles:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   3360
         Width           =   1800
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Food grants this much energy: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regenerate this much food: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   2190
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regenerate food every (n) cycles: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   2550
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial energy: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial food amount: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1590
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# of creatures: "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame frmOrganisms 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select a creature for detailed information:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   7800
      TabIndex        =   4
      Top             =   5400
      Width           =   4215
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbOrganisms 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Text            =   "Please start the simulator"
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "STOP current simulation"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START new simulation"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      DrawWidth       =   2
      FillColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7530
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.PictureBox picFood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      DrawWidth       =   2
      FillColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   7530
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
      Width           =   7530
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press the start button to begin the simulator -------------------------------------------------->"
      BeginProperty Font 
         Name            =   "Segoe UI"
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
      Top             =   120
      Width           =   7455
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
'Description:
'
'This program was the final project for one of my university bioinformatics courses.
' The basic premise is simple - generate an initial group of random critters,
' then allow them to move, eat, reproduce, and die.
'
'Many evolutionary principles are demonstrated by this little project, including
' genetic drift, population equilibrium, the "bottleneck" effect, and other aspects
' of small, closed populations.  Many settings can be manually specified, which makes
' for a kind of "game" if you try to create a population that can last longer than
' several generations.  (The default settings are actually quite close to achieving
' equilibrium, depending on the iteration.)
'
'***************************************************************************

Option Explicit

'Array of artificial "organisms"
Dim creatures() As Organism

'Whether or not the main program loop is actually running...
Dim runSim As Boolean

'Variables used to show the user information about the simulation
Dim totalCreatures As Long
Dim numAlive As Long
Dim numOfCycles As Long

'Option to display deceased creatures
Private Sub chkDisplayDead_Click()
    drawDeadCreatures = chkDisplayDead
    picMap.Cls
    If (Not runSim) Then DrawMap
End Sub

'Option to allow the creatures to reproduce (for mature audiences only)
Private Sub chkMultiply_Click()
    toMultiply = chkMultiply
End Sub

'The bottom-right combo box that displays information about a particular creature
Private Sub cmbOrganisms_Click()

    'Only update the box if a simulation is running (or has been run before)
    If (UBound(creatures) > 0) Then
    
        'Mark all creatures as unselected
        For i = 0 To UBound(creatures)
        
            creatures(i).Selected = False
            
            'Erase any currently selected creature
            If drawDeadCreatures Then
                creatures(i).DrawOrganism picMap
            Else
                If creatures(i).isAlive Then creatures(i).DrawOrganism picMap
            End If
            
        Next i
        
        '"Select" the current creature
        creatures(cmbOrganisms.ListIndex).Selected = True
        
        'If this creature is dead, we have to go out of our way to draw it
        If drawDeadCreatures Then creatures(cmbOrganisms.ListIndex).DrawOrganism picMap
        
        'Now we write a bunch of text to the creature text box
        With txtInfo
            .Text = "Creature #" & cmbOrganisms.ListIndex & vbCrLf
            .Text = .Text & "------------------------" & vbCrLf
            .Text = .Text & "Lifetime: " & creatures(cmbOrganisms.ListIndex).Lifetime & " cycles"
            If creatures(cmbOrganisms.ListIndex).isAlive Then
                .Text = .Text & " (still alive)"
            Else
                .Text = .Text & " (now dead)"
            End If
            .Text = .Text & vbCrLf
            .Text = .Text & "Size: " & creatures(cmbOrganisms.ListIndex).gSize & vbCrLf
            .Text = .Text & "Range: " & creatures(cmbOrganisms.ListIndex).gRange & vbCrLf
            .Text = .Text & "Speed: " & creatures(cmbOrganisms.ListIndex).gSpeed & vbCrLf
            .Text = .Text & "Current Energy: " & creatures(cmbOrganisms.ListIndex).Energy & vbCrLf
            .Text = .Text & "Parent: "
            'The first generation of creatures are marked with -1 as their ParentID and AncestralID
            If (creatures(cmbOrganisms.ListIndex).ParentID = -1) Then
                .Text = .Text & "none (original)" & vbCrLf
            Else
                .Text = .Text & creatures(cmbOrganisms.ListIndex).ParentID & vbCrLf
            End If
            .Text = .Text & "Original ancestor: "
            If (creatures(cmbOrganisms.ListIndex).AncestralID = -1) Then
                .Text = .Text & "none (original)" & vbCrLf
            Else
                .Text = .Text & creatures(cmbOrganisms.ListIndex).AncestralID & vbCrLf
            End If
        End With
    End If
    
End Sub

'Save dialog allows for export to a tab-delimited text file.  This allows for surprisingly
' powerful analyses using any modern spreadsheet or statistical software.
Private Sub cmdSaveData_Click()

    'New CommonDialog object (prevents us from having to bundle the OCX)
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'Use the dialog to get a save file location from the user
    Dim sFile As String
    If cDialog.GetSaveFileName(sFile, , True, "Simulation data (.txt)|*.txt|All files|*.*", , , "Save simulation data", ".txt", frmMain.hWnd) Then
        
        If FileExist(sFile) Then Kill sFile
        
        Open sFile For Output As #1
        
        'Create a header row
        Dim hInfo As String
        hInfo = "Creature#" & vbTab & "Lifetime" & vbTab & "Size" & vbTab & "Range" & vbTab & "Speed" & vbTab & "Parent" & vbTab & "AncestralID" & vbTab & "DNA"
        Print #1, hInfo
        
        'Don't write anything further unless a simulation has been ran
        If (UBound(creatures) > 0) Then
        
            'Loop through every creature, dumping info into the file as we go
            For i = 0 To UBound(creatures)
                
                'A temporary string for storing the information on this creature
                Dim tInfo As String
                
                tInfo = creatures(i).ID & vbTab
                tInfo = tInfo & creatures(i).Lifetime & vbTab
                tInfo = tInfo & creatures(i).gSize & vbTab
                tInfo = tInfo & creatures(i).gRange & vbTab
                tInfo = tInfo & creatures(i).gSpeed & vbTab
                If (creatures(i).ParentID = -1) Then
                    tInfo = tInfo & "n/a" & vbTab
                Else
                    tInfo = tInfo & creatures(i).ParentID & vbTab
                End If
                If (creatures(i).AncestralID = -1) Then
                    tInfo = tInfo & "n/a" & vbTab
                Else
                    tInfo = tInfo & creatures(i).AncestralID & vbTab
                End If
      
                'As a bonus, include the creature's actual "DNA" strand
                For x = 0 To creatures(i).GetMaxDNA
                    tInfo = tInfo & creatures(i).GetDNABase(x)
                Next x
                
                'Write the line to file
                Print #1, tInfo
        
            Next i
            
        End If
        
        Close #1
        
    End If
    
End Sub


'START button: begin a new simulation
Private Sub cmdStart_Click()
    
    'Clear out the buffer picture boxes
    picMap.Picture = LoadPicture(vbNullString)
    picFood.Picture = LoadPicture(vbNullString)
    
    'Collect important values from text boxes and check boxes
    startEnergy = CLng(txtInitialEnergy)
    InitialCreatures = CLng(txtOrganisms) - 1
    foodRegen = CLng(txtFoodRegen)
    foodGen = CLng(txtFoodGen)
    foodWorth = CLng(txtFoodWorth)
    toMultiply = chkMultiply
    drawDeadCreatures = chkDisplayDead
    mutateTurns = CLng(txtMutateTurns)
    numOfMutations = CLng(txtMutations)
    
    'Randomly fill the food array with the specified amount of food
    ReDim Food(0 To WORLDWIDTH, 0 To WORLDHEIGHT) As Long
    
    For i = 0 To CLng(txtInitialFood)
        x = Int(Rnd * WORLDWIDTH)
        y = Int(Rnd * WORLDHEIGHT)
        Food(x, y) = 255
    Next i
    
    'Draw the food to its own special picture box
    For x = 0 To WORLDWIDTH
    For y = 0 To WORLDHEIGHT
        If (Food(x, y) > 0) Then SetPixelV picFood.hDC, x, y, RGB(0, 64, 0)
    Next y
    Next x
    
    'Empty the "select a creature" combo box
    cmbOrganisms.Clear
    
    'Create the original batch of creatures
    ReDim creatures(0 To InitialCreatures) As Organism
    
    For i = 0 To InitialCreatures
        Set creatures(i) = New Organism
        creatures(i).CreateRandom
        creatures(i).ID = i
        creatures(i).oX = Int(Rnd * WORLDWIDTH)
        creatures(i).osX = creatures(i).oX
        creatures(i).oY = Int(Rnd * WORLDHEIGHT)
        creatures(i).osY = creatures(i).oY
        cmbOrganisms.AddItem "Creature #" & i
    Next i
    
    'Reset other tracking data
    numOfCycles = 0
    numAlive = InitialCreatures
    totalCreatures = InitialCreatures
    
    'Start the simulation
    runSim = True
    MainLoop
    
End Sub

'STOP button: only stops the simulation; it does NOT erase any data
' (in case the user wants to export to a file for further analyses)
Private Sub cmdStop_Click()
    runSim = False
End Sub

'When the program is first started...
Private Sub Form_Load()
    
    'Set the main loop as "not running"
    runSim = False
    
    'Seed the random number generator
    Randomize Timer
    
    'Prepare the creature array
    ReDim creatures(0) As Organism
    
    'Display some explanatory text
    txtInfo.Text = "Please select a creature from the drop-down box."
    
End Sub

'DRAW EVERYTHING: copy the food buffer to a new buffer, add the creatures, then flip the
' composited buffer to the screen
Private Sub DrawMap()
    
    'Copy the food buffer to the creature buffer
    BitBlt picMap.hDC, 0, 0, picMap.ScaleWidth, picMap.ScaleHeight, picFood.hDC, 0, 0, vbSrcCopy
    
    'Next, draw all the creatures over the food
    For i = 0 To UBound(creatures)
        If drawDeadCreatures Then
            creatures(i).DrawOrganism picMap
        Else
            If creatures(i).isAlive Then creatures(i).DrawOrganism picMap
        End If
    Next i
    
    'Last, copy the composited image to the screen (doing it this way prevents flickering)
    BitBlt picFront.hDC, 0, 0, picFront.ScaleWidth, picFront.ScaleHeight, picMap.hDC, 0, 0, vbSrcCopy
    
End Sub

'Food is randomly regenerated at a rate specified by the user
Private Sub GenerateFood()
    
    Dim setNewFood As Long
    
    'Assign setNewFood an arbitrary value between zero and a million
    setNewFood = Int(Rnd * 1000000)
    
    'This formula makes it so that roughly every foodRegen turns, new food is drawn
    If (setNewFood Mod foodRegen = 0) Then
        'Draw new food (the amount specified by the user) at random locations
        For i = 0 To foodGen
            x = Int(Rnd * WORLDWIDTH)
            y = Int(Rnd * WORLDHEIGHT)
            'Mark this spot in the food array as being "filled"
            Food(x, y) = 255
            'Draw food as a pleasant dark green color
            SetPixelV picFood.hDC, x, y, RGB(0, 64, 0)
        Next i
    End If
    
End Sub


'THE MAIN PROGRAM LOOP. It runs when start is pushed and stops when stop is pushed.  Advanced, I know.
Private Sub MainLoop()
    
    'Count keeps track of how many times the loop has been run
    Dim count As Long
    count = 1
    
    'If "stop" has not been pressed, process another iteration of the loop
    Do While runSim = True
    
        'First, attempt to generate food
        GenerateFood
        
        'The "think" routine allows all the creatures one turn worth of brain activity
        Think
        
        'If the creatures have been allowed to multiply, do that now
        If toMultiply = True Then
            
            'Only reproduce once every <mutateTurns>
            If count Mod mutateTurns = 0 Then
            
                'Cycle through each creature one-at-a-time
                For i = 0 To UBound(creatures)
                    
                    'If this creature is a) alive and b) has existed for at least one lifecycle
                    ' (i.e. it wasn't just created), let it reproduce
                    If (creatures(i).isAlive = True And creatures(i).Lifetime >= mutateTurns) Then
                        'Track some basic statistics
                        totalCreatures = totalCreatures + 1
                        numAlive = numAlive + 1
                        'Make room in the array for a new creature
                        ReDim Preserve creatures(0 To totalCreatures) As Organism
                        'Create the creature using its parent's data
                        Set creatures(totalCreatures) = New Organism
                        creatures(totalCreatures).CreateFromCreature creatures(i)
                        creatures(totalCreatures).ID = totalCreatures
                        creatures(totalCreatures).oX = creatures(i).oX
                        creatures(totalCreatures).osX = creatures(i).osX
                        creatures(totalCreatures).oY = creatures(i).oY
                        creatures(totalCreatures).osY = creatures(i).osY
                        'Add this creature's entry to the combo box
                        cmbOrganisms.AddItem "Creature #" & totalCreatures
                    End If
                Next i
            End If
        End If
        
        'Update the statistics label
        lblTitle.Caption = "Running simulator: " & numAlive & " alive, " & (totalCreatures - numAlive + 1) & " dead @ " & count & " cycles"
        
        'Draw everything
        DrawMap
        
        'Pause for user events
        DoEvents
        
        'If all our creatures have died, end the loop automatically
        If numAlive = 0 Then runSim = False
        
        count = count + 1
    Loop
    
End Sub

'THINK: the routine that allows each creature one round of brain activity
Private Sub Think()

    'Reset the number of living creatures
    numAlive = 0

    'Loop through every creature, allowing each a moment of simple thought processing
    For i = 0 To UBound(creatures)
        If creatures(i).isAlive = True Then
            creatures(i).Brain
            'Track the number of living creatures while we're at it
            numAlive = numAlive + 1
        End If
    Next i

End Sub

'If the user shuts down the program, make sure our creatures don't remain alive in the background
Private Sub Form_Unload(Cancel As Integer)
    runSim = False
End Sub

'If the user clicks on a particular creature, display that creature's information
Private Sub picFront_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Allow left mouseclicks to select a creature...
    If Button = 1 Then
    
        Dim minDistance As Single
        Dim curDist As Single
        Dim minIndex As Long
        
        minDistance = 10000
        minIndex = 0
        
        'Ignore mouse clicks if a simulation hasn't been run
        If UBound(creatures) > 0 Then
        
            'Loop through each creature and find the one whose center is closest to the
            ' click location
            For i = 0 To UBound(creatures)
                'If dead creatures aren't being drawn, and this creature is dead, ignore it
                If drawDeadCreatures = False And creatures(i).isAlive = False Then GoTo NEXTI
                'Calculate distance
                curDist = Distance(x, y, creatures(i).oX, creatures(i).oY)
                'If this one is closer than what we currently have, promote it
                If curDist < minDistance Then
                    minDistance = curDist
                    minIndex = i
                End If
NEXTI:      Next i
            
            'Only allow a creature to be selected if the click was physically inside the
            ' creature's radius
            If minDistance < (creatures(minIndex).gRange * 2) Then
                If drawDeadCreatures = True Then
                    cmbOrganisms.ListIndex = minIndex
                Else
                    If creatures(minIndex).isAlive = True Then
                        cmbOrganisms.ListIndex = minIndex
                    End If
                End If
                
            '...otherwise, ignore this click
            Else
                For i = 0 To UBound(creatures)
                    creatures(i).Selected = False
                Next i
                cmbOrganisms.Text = "Please select a creature"
                txtInfo.Text = "Please select a creature from the drop-down box.  Once selected, it will turn BLACK for easy identification."
            End If
        
        End If
    
    '...while right and middle-clicks remove any previous selections
    Else
    
        If UBound(creatures) > 0 Then
            For i = 0 To UBound(creatures)
                creatures(i).Selected = False
            Next i
        End If
        
        cmbOrganisms.Text = "Please select a creature"
        txtInfo.Text = "Please select a creature from the drop-down box.  Once selected, it will turn BLACK for easy identification."
    
    End If

End Sub

'Simple Euclidean distance function
Private Function Distance(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    Distance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
End Function

'Check to see if a file with this name already exists
Private Function FileExist(fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(fName)
    FileExist = Not CBool(Err)
End Function
