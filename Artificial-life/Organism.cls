VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Organism"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'
'This Organism class gives life to the creatures in the simulation.
'
'***************************************************************************

Option Explicit

'As you can imagine, this is the length (in "nucelotides") of each creature's "DNA"
Private Const DNALENGTH As Long = 299

'A unique identifer for this creature
Public ID As Long
'The ID number of this creature's parent
Public ParentID As Long
'The ID number of the original ancestor of this creature
Public AncestralID As Long
'Is the creature currently alive or dead?
Public isAlive As Boolean
'An array for the creature's DNA
Private DNA() As Byte
'The creature's current coordinates
Public oX As Long, oY As Long
Public osX As Single, osY As Single

'Currently selected or not
Public Selected As Boolean

'Next, the values for specific genes
'(These are determined ONLY at birth)

'Radius of the creature
Public gSize As Long
'Speed of the creature
Public gSpeed As Single
'Eyesight range of the creature (also treated as a radius)
Public gRange As Long

'Next, creature properties that exist independent of genes.
'These are changed as the program progresses.

'Energy of the creature
Public Energy As Long
'How many cycles the creature has been alive
Public Lifetime As Long
'The type of our target coordinates (food or a random destination)
Public targetType As Long
'The creature's target coordinates
Public tX As Long, tY As Long

'Loop variables
Dim i As Long, j As Long
Dim x As Long, y As Long

'Return the length of this creature's DNA
Public Function GetMaxDNA() As Long
    GetMaxDNA = DNALENGTH
End Function

'Return a DNA base at position (pos)
Public Function GetDNABase(ByVal pos As Long) As Byte
    GetDNABase = DNA(pos)
End Function

'Mutate a base randomly (typically performed just after a creature has been created)
Public Sub MutateBase()
    
    'Generate a random number between -1 and 1
    Dim Mutate As Long
    Mutate = Int(Rnd * 3) - 1
    
    'Select a random base, then apply the mutation to it.
    '(Note that Mutate can have a value of 0, in which case it is an "inconsequential" mutation)
    Dim pos As Long
    pos = Rnd * DNALENGTH
    If DNA(pos) = 3 And Mutate = 1 Then Mutate = 0
    If DNA(pos) = 0 And Mutate = -1 Then Mutate = 0
    DNA(pos) = DNA(pos) + Mutate

End Sub

'Two possible methods for creating a new creature:
'1: Create a creature with random DNA
Public Sub CreateRandom()

    'Creatures are born with full energy
    Energy = startEnergy
    
    'Give this one random DNA
    InitializeRandomDNA
    
    'Generate initial values for all genetically-set attributes
    GetInitialSize
    GetInitialSpeed
    GetInitialRange
    
    'The only time this routine is run is for the initial generation, so
    ' unilaterally set ParentID to -1 (indicative of first generation creatures)
    ParentID = -1
    AncestralID = -1
    
End Sub

'2: Create a creature from an existing creature ("reproduction", of the asexual variety)
Public Sub CreateFromCreature(ByRef srcCreature As Organism)
    
    'Split the parent creature's energy with the child
    srcCreature.Energy = srcCreature.Energy / 2
    Energy = srcCreature.Energy
    ParentID = srcCreature.ID
    
    'If the parent was a member of the founding set of creatures, set
    ' the AncestralID to match the ParentID
    If srcCreature.AncestralID = -1 Then
        AncestralID = srcCreature.ID
    Else
        AncestralID = srcCreature.AncestralID
    End If
    
    'Copy the parent's DNA over to the child
    InitializeDNAFromCreature srcCreature
    
    'If mutations are enabled, mutate away!
    For i = 0 To numOfMutations
        MutateBase
    Next i
    
    'Generate initial values for all genetically-set attributes
    GetInitialSize
    GetInitialSpeed
    GetInitialRange
    
End Sub

'This routine simply copies one creature's DNA (typically a parent)
' into another creature (typically a child)
Public Sub InitializeDNAFromCreature(ByRef srcDNA As Organism)
    ReDim DNA(0 To DNALENGTH) As Byte
    For x = 0 To DNALENGTH
        DNA(x) = srcDNA.GetDNABase(x)
    Next x
End Sub

'This routine generates a random strand of DNA
Public Sub InitializeRandomDNA()
    ReDim DNA(0 To DNALENGTH) As Byte
    For x = 0 To DNALENGTH
        DNA(x) = Int(Rnd * 4)
    Next x
End Sub

'When a creature is "born"...
Private Sub Class_Initialize()
    
    'Mark the creature as not selected
    Selected = False
    
    'Mark the creature as alive
    isAlive = True
    
    'Start counting the number of cycles this creature has lived (0)
    Lifetime = 0
        
    'Set the creature's target type as "undefined"
    targetType = -1
    tX = 0
    tY = 0
    
End Sub

'****************************************************
'This set of routines calculates values for the creature's attributes (phenotype)
' based off scanning segments of DNA (genotype)

'Creature size
Public Sub GetInitialSize()

    For i = 0 To 49
        If DNA(i) = 3 Then gSize = gSize + 1
    Next i
    
End Sub

'Creature speed
Public Sub GetInitialSpeed()
    
    For i = 50 To 149
        If DNA(i) = 3 Then gSpeed = gSpeed + 1
    Next i
    
    'We don't want speed to be too fast, so divide it by 20
    gSpeed = gSpeed / 20
       
End Sub

'Creature range (how far it can see)
Public Sub GetInitialRange()

    For i = 150 To 199
        If DNA(i) = 3 Then gRange = gRange + 1
    Next i
    
    'Add size to this value (so range is never smaller than the size of the creature)
    gRange = gRange + gSize

End Sub

'END gene creators
'************************************************************************


'This routine will draw the creature onto (PictureBox) at its present coordinates
Public Sub DrawOrganism(ByVal pBox As PictureBox)

    'If the creature is alive, draw it in blue and red.
    If isAlive Then
        
        oX = Int(osX)
        oY = Int(osY)
    
        'If the creature is selected, draw it in black.
        If Me.Selected Then
            pBox.Circle (oX, oY), gRange, RGB(0, 0, 0)
            pBox.Circle (oX, oY), gSize, RGB(0, 0, 0)
        Else
            pBox.Circle (oX, oY), gRange, RGB(255, 0, 0)
            pBox.Circle (oX, oY), gSize, RGB(0, 0, 255)
        End If
    
    'If it's dead, draw it in light gray
    Else
    
        'If the creature is selected, draw it in black.
        If Me.Selected Then
            pBox.Circle (oX, oY), gRange, RGB(0, 0, 0)
            pBox.Circle (oX, oY), gSize, RGB(0, 0, 0)
        Else
            pBox.Circle (oX, oY), gRange, RGB(224, 224, 224)
            pBox.Circle (oX, oY), gSize, RGB(208, 208, 208)
        End If
    End If
        
End Sub

'This routine handles the main thought process for this creature
Public Sub Brain()
    
    'Thinking takes energy...
    Energy = Energy - 1
    
    'If it has no energy, it dies
    If (Energy <= 0) Then
        Energy = 0
        isAlive = False
        
    'If the creature is still alive, count this cycle
    Else
        Lifetime = Lifetime + 1
    End If
    
    'Determine what kind of a target square the creature has (random or food)
    If (targetType = -1) Then
    
        'If it doesn't have a target, scan for food
        ScanForFood
        
    Else
    
        'Check to see if the target is close enough to evaluate
        If (Abs(osX - tX) < gSize) And (Abs(osY - tY) < gSize) Then
        
            'Depending on the target type, the creature needs to re-evaluate all options
            
            'If the target is food and we've arrived, eat it!
            If (targetType = 1) Then
            
                Food(tX, tY) = 0  'Erase this food from the array
                SetPixelV frmMain.picFood.hDC, tX, tY, frmMain.picFood.BackColor  'Remove it from the map, too
                Energy = Energy + foodWorth  'Add the food value to this creature's energy
                targetType = -1  'Reset the target type (since it no longer has a target)
                
            'If the target isn't food, it's just a random target - so reset it!
            Else
                targetType = -1
            End If
            
        'If the target isn't close enough, the creature needs to move closer
        Else
    
            If (tY < osY) Then
                osY = osY - gSpeed
            Else
                osY = osY + gSpeed
            End If
            If (tX < osX) Then
                osX = osX - gSpeed
            Else
                osX = osX + gSpeed
            End If
            
        End If
    
    End If
    
End Sub

'This subroutine allows the creature to scan for food within its maximum scan radius
' NOTE: The creatures maybe represented as circles, but they actually search in a
'       square pattern. It's noticeably easier to program than a circular pattern.
Public Sub ScanForFood()

    'The initial x and y coordinates to search from
    Dim iX As Long, iY As Long
    
    'The initial x and y coordinates to search to
    Dim fX As Long, fY As Long
    
    'Trim search coordinates (don't want them outside the limit of the food array)
    iX = gThan(0, oX - gRange)
    iY = gThan(0, oY - gRange)
    fX = lThan(WORLDWIDTH, oX + gRange)
    fY = lThan(WORLDHEIGHT, oY + gRange)
    
    'Search for food
    For x = iX To fX
    For y = iY To fY
    
        'If the creature finds food...
        If Food(x, y) = 255 Then
            'Mark the target as being of type "Food"
            targetType = 1
            'Remember these x and y coordinates as the new target
            tX = x
            tY = y
            'Quit searching for food
            GoTo noMoreSearch
        End If
        
    Next y
    Next x
    
'Come here when searching is over(one way or another)
noMoreSearch:

    'If the creature found food, there's nothing more to do
    If targetType = 1 Then
        Exit Sub
        
    'If the creature didn't find food, set a random target location and start moving there
    Else
        targetType = 0  'targetType = 0 denotes a random target
        tX = Rnd * WORLDWIDTH
        tY = Rnd * WORLDHEIGHT
    End If

End Sub

'Returns the smaller of two numbers
Private Function lThan(ByVal n1 As Long, ByVal n2 As Long) As Long
    If (n1 < n2) Then lThan = n1 Else lThan = n2
End Function

'Returns the larger of two numbers
Private Function gThan(ByVal n1 As Long, ByVal n2 As Long) As Long
    If (n1 > n2) Then gThan = n1 Else gThan = n2
End Function
