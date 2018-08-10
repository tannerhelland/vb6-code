Attribute VB_Name = "Sub_Module"
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

'This routine is run when the program starts up
Public Sub InitializeEditor()
    'Default tiles are nasty brown school carpet on the left, water (?) on the right
    CurPicIndex = 0
    CurPicIndex2 = 1
    
    'Our default map size is 20 tiles by 20 tiles
    SizeX = 20
    SizeY = 20
    ReDim MapArray(0 To SizeX, 0 To SizeY) As Byte
    
    'Set the default zoom to 100% (64, or the size of our tiles)
    Zoom = 64
    
    'Set appropriate values for the main scroll bars
    Main.HSTileBar.Max = Abs(Main.PicTiles.ScaleWidth - Main.PicTilesBuffer.ScaleWidth) / 64
    Main.HSMain.Max = (SizeX * Zoom) - Main.MainPic.ScaleWidth
    Main.VSMain.Max = (SizeY * Zoom) - Main.MainPic.ScaleHeight
    
    'Initialize the map array
    For x = 0 To SizeX
    For Y = 0 To SizeY
        MapArray(x, Y) = 0
    Next Y
    Next x
    
    'Show the form and draw everything to it
    Main.Show
    RefreshAll
    
End Sub

'Render the current map data to the main picture box
Public Sub DrawMap()
    
    'Clear out the current map
    Main.MainPic.Picture = LoadPicture(vbNullString)
    
    For x = 0 To SizeX
    For Y = 0 To SizeY
        'For speed purposes, it would be better to remove this "If/Then" statement to be outside the loop.
        'However, in an editor environment this is somewhat unimportant.
        If UseStretchBlt Then
            StretchBlt Main.MainPic.hdc, (x * Zoom) - ViewX, (Y * Zoom) - ViewY, Zoom, Zoom, Main.PicTilesBuffer.hdc, MapArray(x, Y) * 64, 0, 64, 64, vbSrcCopy
        Else
            Main.MainPic.PaintPicture Main.PicTilesBuffer.Picture, (x * Zoom) - ViewX, (Y * Zoom) - ViewY, Zoom, Zoom, MapArray(x, Y) * 64, 0, 64, 64, vbSrcCopy
        End If
    Next Y
    Next x
    
    'Because AutoRedraw is set to true for our main map, update the .Picture property and force a refresh
    Main.MainPic.Picture = Main.MainPic.Image
    Main.MainPic.Refresh
    
End Sub

'When scrollbars are used, set the new viewport and redraw the map
Public Sub ScrollMap()
    ViewX = Main.HSMain.Value
    ViewY = Main.VSMain.Value
    DrawMap
End Sub

'This routine copies the correct tileset from the invisible buffer to the visible bar
Public Sub DrawTileBar()
    
    If UseStretchBlt Then
        StretchBlt Main.PicTiles.hdc, 0, 0, Main.PicTiles.ScaleWidth, 64, Main.PicTilesBuffer.hdc, Main.HSTileBar.Value * 64, 0, Main.PicTiles.ScaleWidth, 64, vbSrcCopy
    Else
        Main.PicTiles.PaintPicture Main.PicTilesBuffer.Picture, 0, 0, Main.PicTiles.ScaleWidth, 64, Main.HSTileBar.Value * 64, 0, Main.PicTiles.ScaleWidth, 64, vbSrcCopy
    End If
    
    Main.MainPic.Picture = Main.MainPic.Image
    Main.MainPic.Refresh

End Sub

'Draw a tile of type "ArrayIndex" at location (x, y)
Public Sub DrawTile(ByVal x As Integer, ByVal Y As Integer, ByVal ArrayIndex As Byte)
    
    'Calculate the actual x and y values (accounting for zoom and scrolling)
    CurX = (x \ Zoom) + (ViewX \ Zoom)
    CurY = (Y \ Zoom) + (ViewY \ Zoom)
    
    'If the user moved the mouse outside of the map dimensions, cancel this routine
    If CurX > SizeX Or CurY > SizeY Or CurX < 0 Or CurY < 0 Then Exit Sub
    
    'Store this tile value in MapArray
    MapArray(CurX, CurY) = ArrayIndex
    
    'Draw the new tile
    If UseStretchBlt Then
        StretchBlt Main.MainPic.hdc, (CurX * Zoom) - ViewX, (CurY * Zoom) - ViewY, Zoom, Zoom, Main.PicTilesBuffer.hdc, ArrayIndex * 64, 0, 64, 64, vbSrcCopy
    Else
        Main.MainPic.PaintPicture Main.PicTilesBuffer.Picture, (CurX * Zoom) - ViewX, (CurY * Zoom) - ViewY, Zoom, Zoom, ArrayIndex * 64, 0, 64, 64, vbSrcCopy
    End If
    
    Main.MainPic.Picture = Main.MainPic.Image
    Main.MainPic.Refresh
    
End Sub

'If the zoom is changed, update all scrollbars and draw the new, properly zoomed map
Public Sub ChangeZoom()

    'Get a new zoom value from the user
    Zoom = InputBox("Enter the percent zoom (100 is default)")
    Zoom = Int((Zoom / 100) * 64)
    
    'Change the scroll bars depending on the new zoom value
    If (SizeX * Zoom) < Main.MainPic.ScaleWidth Then
        Main.HSMain.Value = 0
        Main.HSMain.Enabled = False
    Else
        Main.HSMain.Enabled = True
        Main.HSMain.Max = (SizeX * Zoom) - Main.MainPic.ScaleWidth
    End If
    
    If (SizeY * Zoom) < Main.MainPic.ScaleHeight Then
        Main.VSMain.Value = 0
        Main.VSMain.Enabled = False
    Else
        Main.VSMain.Enabled = True
        Main.VSMain.Max = (SizeY * Zoom) - Main.MainPic.ScaleHeight
    End If
    
    'Draw the new map image
    DrawMap
    
End Sub

'As you may have guessed, this routine refreshes everything on the screen
Public Sub RefreshAll()
    
    If UseStretchBlt Then
        StretchBlt Main.Pic1.hdc, 0, 0, 64, 64, Main.PicTilesBuffer.hdc, CurPicIndex * 64, 0, 64, 64, vbSrcCopy
        StretchBlt Main.Pic2.hdc, 0, 0, 64, 64, Main.PicTilesBuffer.hdc, CurPicIndex2 * 64, 0, 64, 64, vbSrcCopy
    Else
        Main.Pic1.PaintPicture Main.PicTilesBuffer.Picture, 0, 0, 64, 64, CurPicIndex * 64, 0, 64, 64, vbSrcCopy
        Main.Pic2.PaintPicture Main.PicTilesBuffer.Picture, 0, 0, 64, 64, CurPicIndex2 * 64, 0, 64, 64, vbSrcCopy
    End If
    DrawTileBar
    DrawMap
    
End Sub
