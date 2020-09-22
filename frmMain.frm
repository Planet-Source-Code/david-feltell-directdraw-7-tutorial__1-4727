VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   720
      Top             =   2760
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare DirectX and DirectDraw objects
Dim dx As New DirectX7
Dim dd As DirectDraw7

'Declare the surface description of the screen - Surface
'descriptions are the 'options' that are passed to the
'CreateSurface method. The CreateSurface method is what creates
'the directdraw surface. Surfaces either
'(1) store a bitmap for use in the application or
'(2) are used to draw sprites to, and then display on the
'    screen once ready (a backbuffer in other words).
Dim sdScreen As DDSURFACEDESC2
'Declare a rectangular region that represents the screen, not
'used right now but would/could be if any background picture was
'to be used.
Dim rScreen As RECT
'Finaly declare the surface itself.
Dim ScreenSurf As DirectDrawSurface7

'Similarly with the backbuffer ( (2) above)
Dim sdBackBuffer As DDSURFACEDESC2
Dim BackBufferSurf As DirectDrawSurface7
Dim rBack

'Similarly with the sprite ( (1) above)
Dim ShipSurf As DirectDrawSurface7
Dim sdShip As DDSURFACEDESC2
Dim rShip As RECT

'The X and Y values that will be the position of the sprite
'later on
Dim x
Dim y

'''''''''''''''''''
Dim fx As DDBLTFX



Private Sub Form_KeyPress(KeyAscii As Integer)
'Bring us back to our desktop settings  when we press a key
dd.RestoreDisplayMode
dd.SetCooperativeLevel frmMain.hWnd, DDSCL_NORMAL
End
End Sub

Private Sub Form_Load()
'This brings up the sub below
InitDDraw

'The x and y values of where to place the sprite
x = 10
y = 10

'enable the timer that displays the stuff
tmrMain.Enabled = True
End Sub

Sub InitDDraw()

'Create the DirectDraw object that will be used in turn to create
'the ssurfaces and other objects.
Set dd = dx.DirectDrawCreate("")

'Set the application to fullscreen mode and exclusive mode
'fullscreen is self explanitory, exclusive means that no other
'applications can use the display
Call dd.SetCooperativeLevel(frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE)
'this sets the display mode to 800x600 resolution, in 8 bit
'(256) colors, with the default refresh rate (set as 0 indicates
'default) and uses the default display mode (not Mode13).
Call dd.SetDisplayMode(800, 600, 8, 0, DDSDM_DEFAULT)

' 'lFlags' is in charge of telling the DirectDraw surface what
'members are valid.  This means that the CAPS member (which tells
'the DirectDraw surface what it is and what it is allowed to do)
'is allowed to be set up.  Similarly the BACKBUFFERCOUNT member
'(which sets the amount of backbuffers the surface has) is
'allowed.
sdScreen.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
'This sets the lcaps member of the ddscaps member of the surface
'to tell the surface description (the set of options
'passed when you create the surface) that,
'1) This surface is to be used as the primary surface, i.e.
'   is visible to the user.
'2) tells the surface that it can 'flip' (this is the term given
'   to the ability to swap between surfaces to create fast
'   animatiion)
'3) Tells the surface that it is complex.  That means that it is
'   a root surface of more surfaces (more surfcaces can be created
'   from it - see later when the backbuffer is created)
sdScreen.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX

'Tells the surface that it has one backbuffer attached to it,
'created later.
sdScreen.lBackBufferCount = 1

'Set up the right and bottom properties of the rectangle that
'represents the screen to be the same as the width ad height of
'the 'Screen' surface
rScreen.Right = sdScreen.lWidth
rScreen.Bottom = sdScreen.lHeight

'Finally after all the options are setup in the surface
'description (sdScreen), create the surface and set its name
'as 'ScreenSurf'.
Set ScreenSurf = dd.CreateSurface(sdScreen)

'Declare a Temprary DDSDCAPS variable to use in setting up a
'backbuffer.
Dim caps As DDSCAPS2
'Set the lcaps member of the temprary  DDSCAPS to tell the
'backbuffer that it is a backbuffer
caps.lCaps = DDSCAPS_BACKBUFFER
'Attaches the backbuffersurf surface to the ScreenSurf surface
'and like the lflags member (above) of the ScreenSurf surface, it
'tells the backbuffer to set itself up as what it says in the caps
'object (that it is a backbuffer)
Set BackBufferSurf = ScreenSurf.GetAttachedSurface(caps)

'Set up the BackBufferSurf surface's description - not too sure
'about this, or rather how to explain this, you just need it when
'creating a backbuffer linked to a primary surface OK?
BackBufferSurf.GetSurfaceDesc sdBackBuffer

'This is pointless really just good practice to do.  It just sets
'the ShipSurf surface to nothing to ensure it is blank before we
'use it
Set ShipSurf = Nothing

'As above this just tells the surface that the caps member is
'allowed to be used.
sdShip.lFlags = DDSD_CAPS
'This uses the now 'allowed' ddsdCaps member to tell the surface
'that it is a an off screen (memory only) surface. These are
'primarily used to store bitmaps which are drawn to the screen
'later.
sdShip.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN

'This creates a surface from the bitmap file 'ship.bmp' using the
'options specified in the sdShip surface description
Set ShipSurf = dd.CreateSurfaceFromFile(App.Path & "\" & "ship.bmp", sdShip)

'Set the rectangle to be used as a reference to the size of the
'sprite/bitmap's bottom and right properties to be the same as the
'respective height and width properties of the surface
'description. Although, as you saw, there are no height and
'width properties specified in the surface description, that is
'because if none are specified when loading a bitmap, the
'height and width of the bitmap are automatically specified.
rShip.Bottom = sdShip.lHeight
rShip.Right = sdShip.lWidth

'Ok this is the main reason I got into directX at first -
'transparent blitting.  This line declares a temprary variable as
'a 'key' type.  This type is used to set a range of colors to NOT
'be blitted, which means that a certain color can be set as a
'transparent color
Dim key As DDCOLORKEY
'The next two lines set up the low and high colors in the range to
'be made transparent, currently only one color is set up
'(range 0 to 0) which is black.
key.low = 0
key.high = 0
'Bind the colorkey to the ShipSurf DirectDraw surface
ShipSurf.SetColorKey DDCKEY_SRCBLT, key

End Sub

Private Sub tmrMain_Timer()

'Now to declare a temporary rectangle as a destination rectangle
Dim rTemp As RECT
'Y coord of destination rectangle = 'Y' variable defined earlier
rTemp.Top = y
'X coord of destination rectangle = 'X' variable defined earlier
rTemp.Left = x
rTemp.Bottom = y + sdShip.lHeight
rTemp.Right = x + sdShip.lWidth

'This line sets the forecolor of the DirectDraw object which I use
'as the color of the text
BackBufferSurf.SetForeColor vbGreen
'Now this fills the screen with black, a bug in the old tutorial was
'that this line wasnt in there.  This prevents flickering... try and
'see what happens if you comment out this line to see what I mean.
BackBufferSurf.BltColorFill rScreen, vbBlack

'Next line blt's the ship.  The KEYSRC means 'use the transparency
'color key', set up before, and the WAIT means 'whilst drawing,
'dont let anything else do anything until finished'
BackBufferSurf.Blt rTemp, ShipSurf, rShip, DDBLT_KEYSRC Or DDBLT_WAIT

'Next line bltfast's the ship (look up bltfast and blt in help to
'see the difference).
'BackBufferSurf.BltFast 10, 10, ShipSurf, rScreen, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT


'This line draws the text to the screen
Call BackBufferSurf.DrawText(10, 10, "Press key to Exit", False)

'This is the magic line that flips the backbuffer (which is what
'we just blitted to) to the front
ScreenSurf.Flip BackBufferSurf, DDFLIP_WAIT

'Uncomment the line below to see the ship move left slowly, wooo!
'x = x + 1
End Sub
