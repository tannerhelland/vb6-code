VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Hidden Markov Models and CpG Islands in VB6 - www.tannerhelland.com"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13830
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
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   922
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Step 4: Run a sliding window analysis to determine possible CG island locations"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   39
      Top             =   4920
      Width           =   13575
      Begin VB.PictureBox picOE 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   887
         TabIndex        =   45
         Top             =   1080
         Width           =   13335
      End
      Begin VB.PictureBox picResults 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   887
         TabIndex        =   44
         Top             =   2280
         Width           =   13335
      End
      Begin VB.TextBox txtWindowSize 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   42
         Text            =   "200"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdWindow 
         Caption         =   "Scan with sliding window"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblRecommendation 
         BackStyle       =   0  'Transparent
         Caption         =   " After running your analyses, any estimated CG islands will appear here."
         Height          =   375
         Left            =   6360
         TabIndex        =   49
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label lblPercent 
         BackStyle       =   0  'Transparent
         Caption         =   " I/B Ratio:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   13335
      End
      Begin VB.Label lblCGHeader 
         BackStyle       =   0  'Transparent
         Caption         =   " C/G Observed vs. Expected:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   13335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sliding window size:"
         Height          =   195
         Left            =   3480
         TabIndex        =   41
         Top             =   480
         Width           =   1410
      End
   End
   Begin VB.Frame frm3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Step 3: Run the HMM and Viterbi algorithm on your data "
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   37
      Top             =   3840
      Width           =   7455
      Begin VB.CommandButton cmdHMM 
         Caption         =   "Run the algorithm!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblStep3Results 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Results: *none processed*"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3480
         TabIndex        =   43
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame frm2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Step 2: Input HMM Parameters "
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   7455
      Begin VB.CommandButton cmdEstimate 
         Caption         =   "Estimate state probabilities"
         Height          =   495
         Left            =   5160
         TabIndex        =   48
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Frame frmInitial 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Initial Probabilities:"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5160
         TabIndex        =   32
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txtInitProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   600
            TabIndex        =   36
            Text            =   ".5"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtInitProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   35
            Text            =   ".5"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame frmState 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "State Probabilities:"
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   2040
            TabIndex        =   31
            Text            =   ".25"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   2040
            TabIndex        =   29
            Text            =   ".25"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   2040
            TabIndex        =   27
            Text            =   ".25"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   2040
            TabIndex        =   25
            Text            =   ".25"
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   23
            Text            =   ".4"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   21
            Text            =   ".25"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   19
            Text            =   ".1"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtStateProb 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   17
            Text            =   ".25"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(t|I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   30
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(g|I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   28
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(c|I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(a|I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(t|B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(g|B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(c|B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(a|B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame frmIslandOrNot 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "I/B Probabilities"
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txtIBprob 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   14
            Text            =   ".5"
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtIBprob 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   12
            Text            =   ".5"
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtIBprob 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   10
            Text            =   ".3"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtIBprob 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   8
            Text            =   ".7"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(I->I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(I->B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(B->I):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "p(B->B):"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.TextBox txtTest 
      Appearance      =   0  'Flat
      Height          =   4575
      Left            =   7680
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmHMM.frx":0000
      Top             =   240
      Width           =   6015
   End
   Begin VB.Frame frm1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Step 1: Load information from a FASTA file "
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdOpenFASTA 
         Caption         =   "Open a FASTA file..."
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Current file: *none selected*"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Label lblMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Useful messages related to program operations will appear in this box."
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   8400
      Width           =   13575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'HMM, Viterbi, and Sliding Window Algorithms (for purposes of locating genes in human DNA)
'Copyright 2018 by Tanner Helland
' www.tannerhelland.com
'
'This project demonstrates use of a Hidden Markov Model (HMM) to define the relationship
' between normal states (B) and island states (I) within a region of the human chromosome.
' This algorithm proves particularly useful for gene finding, as CpG islands (regions of DNA
' with a larger than expected number of adjacent cytosine and guanine nucleotides) tend to
' appear near the promoters of some 40% of mammalian genes.  The automation of this process
' is critical in gene research because CpG islands are impossible to locate by simply looking
' at a strand of DNA (as they may cover several hundred bases and their C/G content may be
' only slightly higher than the expected value).
'
'The algorithm is fast, robust, and it includes a wealth of good routines (such as loading FASTA
' files into VB).  I've tried to include comments where applicable.
'
'The source code in this project is licensed under a Simplified BSD license.
' For more information, please review LICENSE.md at https://github.com/tannerhelland/thdc-code/
'
'If you find this code useful, please consider a small donation to https://www.paypal.me/TannerHelland
'
'***************************************************************************

Option Explicit

'Stores all lines from the FASTA file
Dim fText() As String
'Maximum number of lines in the FASTA file
Dim eFile As Long

'Byte array of all our DNA data
'a = 0, c = 1, g = 2, t = 3
Dim DNA() As Byte
'Length of the DNA strand
Dim lenDNA As Long
'Total amounts of A, C, G, and T nucleotides
Dim aTot As Long, cTot As Long, gTot As Long, tTot As Long

'These look-up tables will hold data for various probabilities
Dim pIorB(0 To 3) As Double
Dim pState(0 To 7) As Double
Dim pInitial(0 To 1) As Double

'The results of the Viterbi algorithm (0 = B, 1 = I)
Dim vHMM() As Double

'Multi-use variables
Dim x As Long, y As Long

'Estimate state probabilities based on total amounts of A, C, G, and T
Private Sub cmdEstimate_Click()
    
    Dim total As Long
    total = aTot + cTot + gTot + tTot
    
    txtStateProb(0).Text = aTot / total
    txtStateProb(1).Text = cTot / total
    txtStateProb(2).Text = gTot / total
    txtStateProb(3).Text = tTot / total
    txtStateProb(4).Text = txtStateProb(0) / 2
    txtStateProb(5).Text = txtStateProb(1) * 2
    txtStateProb(6).Text = txtStateProb(2) * 2
    txtStateProb(7).Text = txtStateProb(3) / 2
    
End Sub

'This routine performs the actual HMM and Viterbi algorithms
Private Sub cmdHMM_Click()

    Message "Building look-up tables..."

    'First, collect all of probabilities into look-up tables
    For x = 0 To 3
        pIorB(x) = Log(CDbl(txtIBprob(x)))
    Next x
    For x = 0 To 7
        pState(x) = Log(CDbl(txtStateProb(x)))
    Next x
    For x = 0 To 1
        pInitial(x) = Log(CDbl(txtInitProb(x)))
    Next x

    Message "Calculating initial states..."

    'Arrays for holding all HMM entries
    Dim bHMM() As Double, iHMM() As Double
    ReDim bHMM(0 To lenDNA) As Double
    ReDim iHMM(0 To lenDNA) As Double
    
    'We use these strings to show the B and I states output by the algorithms
    Dim sHMM As String
    sHMM = ""
    Dim sViterbi As String
    sViterbi = ""
    
    'Manually perform the first calculation for the first entry
    bHMM(0) = pState(DNA(0)) + pInitial(0)
    iHMM(0) = pState(DNA(0) + 4) + pInitial(1)
    If bHMM(0) > iHMM(0) Then sHMM = sHMM & "B" Else sHMM = sHMM & "I"
    
    'Next, loop through every entry and calulate HMM data as we go
    For x = 1 To lenDNA - 1

        bHMM(x) = pState(DNA(x)) + dMax(bHMM(x - 1) + pIorB(0), iHMM(x - 1) + pIorB(2))
        iHMM(x) = pState(DNA(x) + 4) + dMax(iHMM(x - 1) + pIorB(3), bHMM(x - 1) + pIorB(1))
        
        'Remember states
        If bHMM(x) > iHMM(x) Then sHMM = sHMM & "B" Else sHMM = sHMM & "I"
        
        'Display a friendly message to let the user know how things are coming
        If (x Mod 100) = 0 Then Message "Generating HMM (Base " & x & " of " & lenDNA - 1 & ")..."
        
    Next x
    
    'Now proceed back through the data and generate Viterbi info
    Dim bHMM1() As Double
    Dim bHMM2() As Double
    Dim iHMM1() As Double
    Dim iHMM2() As Double
     
    ReDim bHMM1(0 To lenDNA) As Double
    ReDim bHMM2(0 To lenDNA) As Double
    ReDim iHMM1(0 To lenDNA) As Double
    ReDim iHMM2(0 To lenDNA) As Double
    
    'Assemble initial Viterbi data
    For x = 1 To lenDNA - 1

        'Another friendly message
        If (x Mod 100) = 0 Then Message "Generating Viterbi data (Base " & x & " of " & lenDNA - 1 & ")..."
        
        bHMM1(x) = bHMM(x - 1) + pIorB(0)
        bHMM2(x) = iHMM(x - 1) + pIorB(2)
        iHMM1(x) = bHMM(x - 1) + pIorB(1)
        iHMM2(x) = iHMM(x - 1) + pIorB(3)
        
    Next x
    
    'Next comes the Viterbi algorithm itself
    ReDim vHMM(0 To lenDNA) As Double
    
    Dim totalI As Long, totalB As Long
    
    If bHMM(lenDNA - 1) > iHMM(lenDNA - 1) Then
        vHMM(lenDNA - 1) = 0
        sViterbi = "B"
        totalB = totalB + 1
    Else
        vHMM(lenDNA - 1) = 1
        sViterbi = "I"
        totalI = totalI + 1
    End If
    
    For x = lenDNA - 2 To 0 Step -1

        'Who doesn't like friendly messages?
        If (x Mod 100) = 0 Then Message "Performing Viterbi algorithm (Base " & x & " of " & lenDNA - 1 & ")..."
        
        If vHMM(x + 1) = 0 Then
            If bHMM1(x + 1) > bHMM2(x + 1) Then
                sViterbi = "B" & sViterbi
                vHMM(x) = 0
                totalB = totalB + 1
            Else
                sViterbi = "I" & sViterbi
                vHMM(x) = 1
                totalI = totalI + 1
            End If
        Else
            If iHMM1(x + 1) > iHMM2(x + 1) Then
                sViterbi = "B" & sViterbi
                vHMM(x) = 0
                totalB = totalB + 1
            Else
                sViterbi = "I" & sViterbi
                vHMM(x) = 1
                totalI = totalI + 1
            End If
        End If
        
    Next x
    
    'Display the results of our analysis and encourage the user to continue with step 4
    txtTest = "Raw HMM analysis: " & vbCrLf & sHMM & vbCrLf & vbCrLf & "Viterbi analysis: " & vbCrLf & sViterbi

    lblStep3Results.Caption = "Results: " & CSng(totalB) / (totalB + totalI) * 100 & "% B, " & CSng(totalI) / (totalB + totalI) * 100 & "% I"

    Message "Step 3 completed successfully.  Please proceed with Step 4."
    
    cmdWindow.Enabled = True

End Sub

'Handy routine for processing a FASTA file
Private Sub cmdOpenFASTA_Click()
    
    'Simple open dialog
    Dim newDialog As pdOpenSaveDialog
    Set newDialog = New pdOpenSaveDialog
    
    Dim sFile As String
    If newDialog.GetOpenFileName(sFile, , True, , "FASTA files (*.fas, *.fasta)|*.fas;*.fasta|All files|*.*", , , "Open a FASTA file", , frmMain.hWnd) Then

        Message "Opening FASTA file..."
        
        lblFile.Caption = "Current file: " & sFile
        
        Dim fNum As Long
        fNum = FreeFile
        
        Open sFile For Input As #fNum
        
        'This is an arbitrary "default" max # of lines in the file
        eFile = 100
        ReDim fText(0 To eFile) As String
        
        'Temporary line read from the file
        Dim hInfo As String
        
        'Total length of all lines
        Dim numChars As Long
        
        'Total number of lines
        Dim numOfLines As Long
        numOfLines = -1
        
        Message "Reading in FASTA data..."
        
        'Read through to the end of the file
        Do While Not EOF(fNum)
            Line Input #fNum, hInfo
            hInfo = Trim(hInfo)
            
            'Line width must be greater than 0 and not a FASTA comment line
            If (Len(hInfo) > 0) And Not (Left(hInfo, 1) = ">") Then
                numOfLines = numOfLines + 1
    
                'If the current line exceeds the maximum number of lines allowed, double the size of the
                ' array and continue onward.
                If numOfLines > eFile Then
                    eFile = eFile * 2
                    ReDim Preserve fText(0 To eFile) As String
                End If
                
                fText(numOfLines) = LCase(hInfo) 'Make EVERYTHING lowercase (to simplify text comparisons)
                numChars = numChars + Len(hInfo)
            
            End If
        
        Loop
        
        'Close the file
        Close #fNum
        
        'Resize our array to the minimum size necessary
        ReDim Preserve fText(0 To numOfLines) As String
        ReDim DNA(0 To numChars) As Byte
        
        'Prepare to convert the FASTA from text to numbers
        Dim curChar As Long
        curChar = 0
        aTot = 0
        cTot = 0
        gTot = 0
        tTot = 0
        
        Dim entry As String * 1
        
        Message "Converting FASTA data to byte data for HMM analysis..."
        
        'Travel through each line, converting the FASTA data as we go
        For x = 0 To numOfLines
            
            'On each line, travel through each character individually
            For y = 1 To Len(fText(x))
                
                'Grab the current character
                entry = Mid(fText(x), y, 1)
                Select Case entry
                    'Convert A/C/G/T to numbers, and count how many total A's, C's, G's, T's we have
                    Case "a"
                        DNA(curChar) = 0
                        aTot = aTot + 1
                    Case "c"
                        DNA(curChar) = 1
                        cTot = cTot + 1
                    Case "g"
                        DNA(curChar) = 2
                        gTot = gTot + 1
                    Case "t"
                        DNA(curChar) = 3
                        tTot = tTot + 1
                    'Ignore anything that isn't A/C/G/T
                    Case Else
                        Message "FASTA ERROR: a character other than a, c, g, t was detected (" & entry & ")."
                        curChar = curChar - 1
                End Select
                curChar = curChar + 1
            Next y
        Next x
        
        lenDNA = curChar
        
        Message "Step 1 completed successfully.  Please begin Step 2, and when complete, please click the Step 3 button."

        cmdHMM.Enabled = True

    End If
    
End Sub

'Displays relevant messages at the bottom of the screen
Private Sub Message(ByRef iMessage As String)
    lblMessage.Caption = " " & iMessage
    lblMessage.Refresh
    DoEvents
End Sub

'Determine the larger of two Double-type variables
Private Function dMax(ByVal d1 As Double, ByVal d2 As Double) As Double
    If (d1 > d2) Then dMax = d1 Else dMax = d2
End Function

'This routine performs a sliding window analysis of the HMM/Viterbi data
Private Sub cmdWindow_Click()
    
    Dim winSize As Long
    winSize = CLng(txtWindowSize)
    
    Dim wResults() As Long
    ReDim wResults(0 To lenDNA - winSize) As Long
    
    'Observed vs expected ratio of C/G content
    Dim oeResults() As Single
    ReDim oeResults(0 To lenDNA - winSize) As Single
    
    'Calculate an expected value
    Dim CGexpected As Single
    CGexpected = (cTot + gTot) / (cTot + gTot + aTot + tTot)
    Dim cgCount As Single
    Dim cgMax As Single
    cgMax = 0
    
    Dim tCount As Long
    Dim iMax As Long, iMin As Long
    iMax = 0
    iMin = winSize
    
    'Travel through the data, calculating sliding window values as we go
    For x = 0 To UBound(wResults)
        'Reset the I counter
        tCount = 0
        'Reset the c/g observed count
        cgCount = 0
        If (x Mod 100) = 0 Then Message "Analyzing window " & x & " of " & UBound(wResults)
        
        'Analyze this sliding window
        For y = x To x + winSize
            If vHMM(y) = 1 Then tCount = tCount + 1
            If DNA(y) = 1 Or DNA(y) = 2 Then cgCount = cgCount + 1
        Next y
        
        'Store the value of this sliding window
        wResults(x) = tCount
        cgCount = cgCount / (winSize * CGexpected)
        oeResults(x) = cgCount
        If cgCount > cgMax Then cgMax = cgCount
        If tCount > iMax Then iMax = tCount
        If tCount < iMin Then iMin = tCount
    Next x
    
    Dim tIndex As Single
    
    lblPercent.Caption = " I/B Ratio (Maximum = " & iMax / winSize & "): "
    lblCGHeader.Caption = " C/G Observed vs. Expected (Maximum = " & cgMax & "):"
    
    'Draw the results
    For x = 0 To picResults.ScaleWidth
    
        tIndex = (x / picResults.ScaleWidth) * UBound(wResults)
        
        picResults.Line (x + 1, 0)-(x + 1, picResults.ScaleHeight), RGB(255, 0, 0)
        picOE.Line (x + 1, 0)-(x + 1, picOE.ScaleHeight), RGB(255, 0, 0)
        
        'Percentage plot
        picResults.Line (x, 0)-(x, picResults.ScaleHeight), RGB(255, 255, 255)
        picResults.Line (x, picResults.ScaleHeight)-(x, picResults.ScaleHeight - (CSng(wResults(tIndex)) / CSng(iMax)) * picResults.ScaleHeight)
    
        'O/E plot
        picOE.Line (x, 0)-(x, picOE.ScaleHeight), RGB(255, 255, 255)
        picOE.Line (x, picOE.ScaleHeight)-(x, picOE.ScaleHeight - (CSng(oeResults(tIndex)) / CSng(cgMax)) * picOE.ScaleHeight)
    
        picResults.Refresh
        picOE.Refresh
        DoEvents
    
    Next x
    
    'Lastly, let's attempt to locate an island
    For x = 0 To UBound(wResults)
        
        If wResults(x) > (iMax * 0.9) Then
            y = x + 1
            
            Do While wResults(y) > (iMax * 0.8)
                y = y + 1
            Loop
            
            'See how far our search went
            If (y + winSize - x) > 200 Then
                lblRecommendation.Caption = "There is a likely CpG island from base " & x & " to base " & y + winSize & "."
                tIndex = (x / UBound(wResults)) * picResults.ScaleWidth
                picResults.Line (tIndex, 0)-(tIndex, picResults.ScaleHeight), RGB(255, 0, 0)
                picOE.Line (tIndex, 0)-(tIndex, picOE.ScaleHeight), RGB(255, 0, 0)
                tIndex = (y / UBound(wResults)) * picResults.ScaleWidth
                picResults.Line (tIndex, 0)-(tIndex, picResults.ScaleHeight), RGB(255, 0, 0)
                picOE.Line (tIndex, 0)-(tIndex, picOE.ScaleHeight), RGB(255, 0, 0)
                GoTo NoMoreLook
                
            End If
            
        End If
        
    Next x
    
    lblRecommendation.Caption = "No potential CpG islands were found (minimum length = " & winSize & " bases)"
    
NoMoreLook:
    
    Message " Sliding window analyses are complete."
    
End Sub
