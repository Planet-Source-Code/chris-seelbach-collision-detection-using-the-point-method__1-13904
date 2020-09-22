VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "DD Animation"
   ClientHeight    =   5625
   ClientLeft      =   2355
   ClientTop       =   1620
   ClientWidth     =   7065
   ForeColor       =   &H00000000&
   Icon            =   "DDtut4.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "DDtut4.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this example uses the Point method for
'collision detection between objects that
'interact with each other within the screen
'boundries. In this case, we want to know if
'a point on the form is not a certain color.
'You can also use the opposite; if a point is
'a certain color...
Option Explicit
'DX stuff
Dim binit As Boolean
Dim dx As New DirectX7
Dim dd As DirectDraw7
Dim flagsurf As DirectDrawSurface7
Dim spritesurf As DirectDrawSurface7
Dim spritesurf1 As DirectDrawSurface7 'blue donut
Dim spritesurf2 As DirectDrawSurface7 'gold donut
Dim spritesurf3 As DirectDrawSurface7 'red donut that wobbles oposite
Dim spritesurf4 As DirectDrawSurface7 'gold donut that wobbles oposite
Dim primary As DirectDrawSurface7
Dim backbuffer As DirectDrawSurface7
Dim ddsd1 As DDSURFACEDESC2
Dim ddsd2 As DDSURFACEDESC2
Dim ddsd3 As DDSURFACEDESC2
Dim ddsd4 As DDSURFACEDESC2
Dim ddsd5 As DDSURFACEDESC2
Dim ddsd6 As DDSURFACEDESC2
Dim ddsd7 As DDSURFACEDESC2
Dim ddsd8 As DDSURFACEDESC2
Dim ddsd9 As DDSURFACEDESC2
Dim ddsd10 As DDSURFACEDESC2
Dim spriteWidth As Integer
Dim spriteHeight As Integer
Dim cols As Integer
Dim rows As Integer
Dim row As Integer
Dim col As Integer
Dim fowardFrame As Integer
Dim brunning As Boolean
Dim CurModeActiveStatus As Boolean
Dim bRestore As Boolean
Dim sMedia As String
Dim Bmotion As Integer 'blue donut select case
Dim Gmotion As Integer 'gold donut select case
Dim Begin As Integer





Sub Init()
    On Local Error GoTo errOut
    
    Dim file As String
    
    Set dd = dx.DirectDrawCreate("")
    Me.Show
    
    'indicate that we dont need to change display depth
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    Call dd.SetDisplayMode(640, 480, 32, 0, DDSDM_DEFAULT)
    
    'get the screen surface and create a back buffer too
    ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsd1.lBackBufferCount = 1
    Set primary = dd.CreateSurface(ddsd1)
    
    'Get the backbuffer
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set backbuffer = primary.GetAttachedSurface(caps)
    backbuffer.GetSurfaceDesc ddsd4
         
    'Create DrawableSurface class form backbuffer
    'if you want to draw text and have the
    'background transparent;
    backbuffer.SetFontTransparency True
    'set the text forecolor to white
    backbuffer.SetForeColor vbWhite
         
    ' init the surfaces
    InitSurfaces
Bmotion = 3
Gmotion = 1
    
    binit = True
    brunning = True
    Do While brunning
        blt
        DoEvents
    Loop

errOut:
    EndIt
End Sub

Sub InitSurfaces()


    Set flagsurf = Nothing
    Set spritesurf = Nothing
    
    'sMedia = "blank.bmp"
    'If sMedia = vbNullString Then sMedia = AddDirSep(CurDir)
    
    'load the bitmap into a surface -  the blank
    ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd2.lWidth = ddsd4.lWidth
    ddsd2.lHeight = ddsd4.lHeight
    Set flagsurf = dd.CreateSurfaceFromFile("blank.bmp", ddsd2)
                        
    'load the bitmap into a surface
    'this bitmap has many frames of animation
    'each is 64 by 64 in layed out in cols x rows
    ddsd3.lFlags = DDSD_CAPS
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set spritesurf = dd.CreateSurfaceFromFile("rdonut.bmp", ddsd3)
    Set spritesurf = dd.CreateSurfaceFromFile("rdonut.bmp", ddsd5)
    Set spritesurf = dd.CreateSurfaceFromFile("rdonut.bmp", ddsd6)
    Set spritesurf1 = dd.CreateSurfaceFromFile("bdonut.bmp", ddsd7)
    Set spritesurf2 = dd.CreateSurfaceFromFile("gdonut.bmp", ddsd8)
    Set spritesurf3 = dd.CreateSurfaceFromFile("rdonutrev.bmp", ddsd9)
    Set spritesurf4 = dd.CreateSurfaceFromFile("gdonutrev.bmp", ddsd10)
    spriteWidth = 64
    spriteHeight = 64
    cols = ddsd3.lWidth / spriteWidth
    rows = ddsd3.lHeight / spriteHeight
    
    'use black for transparent color key which is on
    'the source bitmap -> use src keying
    Dim key As DDCOLORKEY
    key.low = 0 'black
    key.high = 0 'black
    spritesurf.SetColorKey DDCKEY_SRCBLT, key
    spritesurf1.SetColorKey DDCKEY_SRCBLT, key
    spritesurf2.SetColorKey DDCKEY_SRCBLT, key
    spritesurf3.SetColorKey DDCKEY_SRCBLT, key
    spritesurf4.SetColorKey DDCKEY_SRCBLT, key
End Sub


Sub blt()
    On Local Error GoTo errOut
    If binit = False Then Exit Sub
    
    Dim ddrval As Long
    Static i As Integer
    
    Dim rBack As RECT
    Dim rFlag As RECT
    Dim rSprite As RECT
    'the donuts
    Dim rSprite2 As RECT
    Dim rSprite3 As RECT
    Dim rSprite4 As RECT
    Dim rSprite5 As RECT
    Dim rSprite6 As RECT
    '
    Dim rPrim As RECT
    
    Static a As Single
    Static a1 As Single
    Static x As Single
    Static y As Single
    Static x1 As Single
    Static y1 As Single
    Static x2 As Single
    Static y2 As Single
    Static x3 As Single
    Static y3 As Single
    Static t As Single
    Static t2 As Single
    Static tLast As Single
    Static fps As Single
    
    
    ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
    bRestore = False
    Do Until ExModeActive
        DoEvents
        bRestore = True
    Loop
    
    ' if we lost and got back the surfaces, then restore them
    DoEvents
    If bRestore Then
        bRestore = False
        dd.RestoreAllSurfaces
        InitSurfaces ' must init the surfaces again if they we're lost
    End If
    
    'get the area of the screen where our window is
    
    rBack.Bottom = ddsd4.lHeight
    rBack.Right = ddsd4.lWidth
    
    'get the area of the bitmap we want ot blt
    rFlag.Bottom = ddsd2.lHeight
    rFlag.Right = ddsd2.lWidth



    
    
    'blt to the backbuffer from our  surface to
    'the screen surface such that our bitmap
    'appears over the window
    ddrval = backbuffer.BltFast(0, 0, flagsurf, rFlag, DDBLTFAST_WAIT)

    
    'Calculate the frame rate
    If i = 30 Then
        If tLast <> 0 Then fps = 30 / (Timer - tLast)
        tLast = Timer
        i = 0
    End If
    i = i + 1
    'if we want to draw text on the screen;
    'Call backbuffer.DrawText(10, 10, "640 X 480 X 16 Frames per Second: " + Format$(fps, "#.0"), False)
    'Call backbuffer.DrawText(10, 30, "Click on the screen to Quit...", False)
    
    
    'calculate the angle from the center
    'at witch to place the sprite
    'calcultate wich frame# we are on in the sprite bitmap
    t2 = Timer
    If t <> 0 Then
        a = a + (t2 - t) * 40
        a1 = a1 + (t2 - t) * 40
        If a > 360 Then a = a - 360
        If a1 > 360 Then a1 = a1 - 360
        'keep tract of the picture frames
        fowardFrame = fowardFrame + (t2 - t) * 40
        If fowardFrame > rows * cols - 1 Then fowardFrame = 0
    End If
    t = t2
    'position the Gold donut initally
    If Begin = False Then
    x3 = 400
    y3 = 100
    Begin = True
    Else
    End If
    'calculat the x and y position of the sprite,
    'this red donut moves clockwise (+a)
    x = Cos((a / 360) * 2 * 3.141) * 100
    y = Sin((a / 360) * 2 * 3.141) * 100
    rSprite2.Top = y + Me.ScaleHeight / 2 - spriteWidth / 2
    rSprite2.Left = x + Me.ScaleWidth / 2.2
    rSprite2.Right = rSprite2.Left + spriteWidth
    rSprite2.Bottom = rSprite2.Top + spriteHeight
    
    'this red donut sits in the center of the screen
    rSprite3.Top = Me.ScaleHeight / 2 - spriteWidth / 2
    rSprite3.Left = Me.ScaleWidth / 2 - spriteWidth / 2
    
    Dim HereX, HereY As Integer 'where the sprite is
    'move the blue donut
    Select Case Bmotion
    Case 1
    HereX = x1 + 10 'since the donut is round, we'll
    HereY = y1 + 10 'set the X,Y inside the rect.
        'Move the donut left and up by 2 pixels.
        x1 = x1 - 2: y1 = y1 - 2
        rSprite5.Left = x1
        rSprite5.Top = y1
        'If the donut reaches the left edge of the form, move it to the right and up.
        If rSprite5.Left <= 0 Then
            Bmotion = 2
        'If the donut hits another donut
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 3
        ElseIf Point(HereX + 44, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 4
        'If the donut reaches the top edge of the form, move it to the left and down.
        ElseIf rSprite5.Top <= 0 Then
        Bmotion = 4
        End If
    Case 2
    HereX = x1 + 54
    HereY = y1 + 10
        'Move the donut right and up by 2 pixels.
        x1 = x1 + 2: y1 = y1 - 2
        rSprite5.Left = x1
        rSprite5.Top = y1
        
        If rSprite5.Left >= (640 - spriteWidth) Then
            Bmotion = 1
            'If the donut hits another donut
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 4
        ElseIf Point(HereX - 44, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 3
        'If the donut reaches the top edge of the form, move it to the right and down.
        ElseIf rSprite5.Top <= 0 Then
            Bmotion = 3
        End If
    Case 3
    HereX = x1 + 54
    HereY = y1 + 54
        'Move the donut right and down by 2 pixels.
        x1 = x1 + 2: y1 = y1 + 2
        rSprite5.Left = x1
        rSprite5.Top = y1
        'If the donut reaches the right edge of the form, move it to the left and down.
        If rSprite5.Left >= (640 - spriteWidth) Then
            Bmotion = 4
        'If the donut hits another donut
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 1
        ElseIf Point(HereX - 44, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 2
        ElseIf rSprite5.Top >= (480 - spriteWidth) Then
            Bmotion = 2
        End If
    Case 4
    HereX = x1 + 10
    HereY = y1 + 54
        'Move the donut left and down by 2 pixels.
        x1 = x1 - 2: y1 = y1 + 2
        rSprite5.Left = x1
        rSprite5.Top = y1
        'If the donut reaches the left edge of the form, move it to the right and down.
        If rSprite5.Left <= 0 Then
            Bmotion = 3
        'If the donut hits another donut
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 2
        ElseIf Point(HereX + 44, HereY) <> RGB(0, 0, 0) Then
            Bmotion = 1
        'If the donut reaches the bottom edge of the form, move it to the left and up.
        ElseIf rSprite5.Top >= (480 - spriteWidth) Then
            Bmotion = 1
        End If
    End Select
    'move the Gold donut, same as above
    Select Case Gmotion
    Case 1
    HereX = x3 + 10
    HereY = y3 + 10
        
        x3 = x3 - 2: y3 = y3 - 2
        rSprite6.Left = x3
        rSprite6.Top = y3
        
        If rSprite6.Left <= 0 Then
            Gmotion = 2
         
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 3
        ElseIf Point(HereX + 44, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 4
        
        ElseIf rSprite6.Top <= 0 Then
            Gmotion = 4
        End If
    Case 2
    HereX = x3 + 54
    HereY = y3 + 10
        
        x3 = x3 + 2: y3 = y3 - 2
        rSprite6.Left = x3
        rSprite6.Top = y3
        
        If rSprite6.Left >= (640 - spriteWidth) Then
            Gmotion = 1
             
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 4
        ElseIf Point(HereX - 44, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 3
        
        ElseIf rSprite6.Top <= 0 Then
            Gmotion = 3
        End If
    Case 3
    HereX = x3 + 54
    HereY = y3 + 54
        
        x3 = x3 + 2: y3 = y3 + 2
        rSprite6.Left = x3
        rSprite6.Top = y3
        
        If rSprite6.Left >= (640 - spriteWidth) Then
            Gmotion = 4
        
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 1
        ElseIf Point(HereX - 44, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 2
        ElseIf rSprite6.Top >= (480 - spriteWidth) Then
            Gmotion = 2
        End If
    Case 4
    HereX = x3 + 10
    HereY = y3 + 54
        
        x3 = x3 - 2: y3 = y3 + 2
        rSprite6.Left = x3
        rSprite6.Top = y3
        
        If rSprite6.Left <= 0 Then
            Gmotion = 3
        
        ElseIf Point(HereX, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 2
        ElseIf Point(HereX + 44, HereY) <> RGB(0, 0, 0) Then
            Gmotion = 1
        
        ElseIf rSprite6.Top >= (480 - spriteWidth) Then
            Gmotion = 1
        End If
    End Select
    'this red donut moves counterclockwise (-a1)
    'and wobbles in the oposite direction
    x2 = Cos((-a1 / 360) * 2 * 3.141) * 200
    y2 = Sin((-a1 / 360) * 2 * 3.141) * 200
    rSprite4.Top = y2 + Me.ScaleHeight / 2 - spriteWidth / 2
    rSprite4.Left = x2 + Me.ScaleWidth / 2.2
    rSprite4.Right = rSprite4.Left + spriteWidth
    rSprite4.Bottom = rSprite4.Top + spriteHeight
    
    'from the current frame select the bitmap we want
    col = fowardFrame Mod cols
    row = Int(fowardFrame / cols)
    rSprite.Left = col * spriteWidth
    rSprite.Top = row * spriteHeight
    rSprite.Right = rSprite.Left + spriteWidth
    rSprite.Bottom = rSprite.Top + spriteHeight
   
   
    'blt to the backbuffer our animated sprite
    ddrval = backbuffer.BltFast(rSprite2.Left, rSprite2.Top, spritesurf3, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    ddrval = backbuffer.BltFast(rSprite3.Left, rSprite3.Top, spritesurf, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    ddrval = backbuffer.BltFast(rSprite4.Left, rSprite4.Top, spritesurf, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    ddrval = backbuffer.BltFast(rSprite5.Left, rSprite5.Top, spritesurf1, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'blue donut
    ddrval = backbuffer.BltFast(rSprite6.Left, rSprite6.Top, spritesurf4, rSprite, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT) 'gold donut
    'flip the back buffer to the screen
    primary.Flip Nothing, DDFLIP_WAIT

errOut:

End Sub

Sub EndIt()
    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    End
End Sub

Private Sub Form_Click()
'quit
EndIt
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'quit
EndIt
End Sub

Private Sub Form_Load()
Init
End Sub

Private Sub Form_Paint()
blt
End Sub

Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = dd.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
    
End Function
