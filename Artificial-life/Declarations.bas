Attribute VB_Name = "Declarations"

'API calls
'Standard image blitting; faster than VB's internal PaintPicture method
Public Declare Function BitBlt Lib "gdi32" (ByVal hDstDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Standard pixel setting; much faster than VB's internal PSet and Point methods
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte


'Size of the map; these dimensions can theoretically be changed, but you'll need
' to manually resize the various picture boxes to match
Public Const WORLDWIDTH As Long = 499
Public Const WORLDHEIGHT As Long = 499

'Amount of food to generate every round
Public foodGen As Long
'How much energy the food grants
Public foodWorth As Long
'Default starting energy for each creature
Public startEnergy As Long
'Regenerate food every (foodRegen) number of turns
Public foodRegen As Long

'Mutation values
'Whether or not to allow reproduction (a.k.a. the "celibacy" variable)
Public toMultiply As Boolean
'How many turns must pass before creatures are allowed to reproduce
Public mutateTurns As Long
'How many potential mutations to cause at reproduction time
Public numOfMutations As Long

'Amount of creatures to start with
Public InitialCreatures As Long

'The food array
Public Food() As Long

'Whether or not to draw deceased creatures
Public drawDeadCreatures As Boolean

'Loop variables
Public i As Long, j As Long
Public x As Long, y As Long
