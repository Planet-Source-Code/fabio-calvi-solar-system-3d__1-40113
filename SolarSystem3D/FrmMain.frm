VERSION 5.00
Begin VB.Form Frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   240
   ClientTop       =   1155
   ClientWidth     =   8130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi all.
'This was an test application about my Space Defender 3D game program I am writing on.
'It's only an experiment about the reference through the object in a scene, but I think
'you can find some interesting things (like as the reference between objects and the
'rotating background in static height view applications,material override and so on).
'
'NOTES:
'
'The planets don't applying Kepler's Laws: the orbits of the planets are circles, not ellipses.
'All the planets spin on an axis perpendicular to the plane of the ecliptic, Uranus too.
'
'Rotation and orbital motion velocity of each planet are referenced to Earth's rotation and
'Earth's orbital motion velocity.
'
'There is not scale between dimensions and distance but the scales of the distances
'from a planet to another and the scale of the dimensions throught them are correct.
'
'Only Earth has the satellite and the Moon don't show the same face to Earth, sorry.
'There are not the asteroids between Mars and Jupiter.
'Any planet has axial inclination in this example and Jupiter,Uranus and Neptune has
'not rings.
'
'The textures of the planets are taken from http://maxfreek.tripod.com/Material.htm
'For Mercury texture I used a modified Charon's map.
'
'A good site about Solar System info is http://seds.lpl.arizona.edu/nineplanets/nineplanets/nineplanets.html
'
'Everybody can modify the code or employ parts of it within own projects, I don't care (just
'a little credit, please) but NOBODY MUST USE ANY PART OF THIS PROGRAM IN COMMERCIAL
'PURPOSES.
'
'Every comments, suggestions, ideas and e-mails are always welcomed to:
'fabiocalvi@ yahoo.com
'
'Happy coding and have fun!
'Goodbye, Fabio.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Arrow Keys moving you around scene
'Home key: increase orbital motion velocity
'End key: decrease orbital motion velocity
'Pageup key: increase rotation velocity
'Pagedown key: decrease rotation velocity
'S key: rotate robot to left
'L key: rotate robot to right
'F1 key: closed view to Mercury
'F2 key: closed view to Venus
'F3 key: closed view to Earth
'F4 key: closed view to Mars
'F5 key: closed view to Jupiter
'F6 key: closed view to Saturn
'F7 key: closed view to Uranus
'F8 key: closed view to Neptune
'F9 key: closed view to Pluto
'F0 key: reset view
'Esc key: Quit

Option Explicit
Const pi = 3.1415927

' direct x objects
Dim dx As New DirectX7
Dim dd As DirectDraw4

Dim SurfPrimary As DirectDrawSurface4
Dim SurfBack As DirectDrawSurface4
Dim DDSDPrimary As DDSURFACEDESC2
Dim DDCapsBack As DDSCAPS2

Dim dev As Direct3DRMDevice3

Dim clip As DirectDrawClipper
Dim d3drm As Direct3DRM3

Dim scene As Direct3DRMFrame3
Dim cam As Direct3DRMFrame3
Dim view As Direct3DRMViewport2

Dim DxMeshSun As Direct3DRMMeshBuilder3
Dim DxMeshflare As Direct3DRMMeshBuilder3
Dim DxMeshMercury As Direct3DRMMeshBuilder3
Dim DxMeshVenus As Direct3DRMMeshBuilder3
Dim DxMeshEarth As Direct3DRMMeshBuilder3
Dim DxMeshClouds As Direct3DRMMeshBuilder3
Dim DxMeshMoon As Direct3DRMMeshBuilder3
Dim DxMeshMars As Direct3DRMMeshBuilder3
Dim DxMeshJupiter As Direct3DRMMeshBuilder3
Dim DxMeshSaturn As Direct3DRMMeshBuilder3
Dim DxMeshRing As Direct3DRMMeshBuilder3
Dim DxMeshUranus As Direct3DRMMeshBuilder3
Dim DxMeshNeptune As Direct3DRMMeshBuilder3
Dim DxMeshPluto As Direct3DRMMeshBuilder3

Dim mesh As Direct3DRMMeshBuilder3

Dim mFrSSun As Direct3DRMFrame3
Dim mFrSun As Direct3DRMFrame3
Dim mFrFlare As Direct3DRMFrame3

Dim mFrMercury As Direct3DRMFrame3
Dim mFrSMercury As Direct3DRMFrame3

Dim mFrVenus As Direct3DRMFrame3
Dim mFrSVenus As Direct3DRMFrame3

Dim mFrEarth As Direct3DRMFrame3
Dim mFrClouds As Direct3DRMFrame3
Dim mFrSearth As Direct3DRMFrame3
Dim mFrMoon As Direct3DRMFrame3

Dim mFrSMars As Direct3DRMFrame3
Dim mFrMars As Direct3DRMFrame3

Dim mFrSJupiter As Direct3DRMFrame3
Dim mFrJupiter As Direct3DRMFrame3

Dim mFrSSaturn As Direct3DRMFrame3
Dim mFrSaturn As Direct3DRMFrame3
Dim mFrRing As Direct3DRMFrame3

Dim mFrSUranus As Direct3DRMFrame3
Dim mFrUranus As Direct3DRMFrame3

Dim mFrSNeptune As Direct3DRMFrame3
Dim mFrNeptune As Direct3DRMFrame3

Dim mFrSPluto As Direct3DRMFrame3
Dim mFrPluto As Direct3DRMFrame3

Dim tex As Direct3DRMTexture3
Dim mo As D3DRMMATERIALOVERRIDE

Dim EarthRt As Single
Dim EarthRv As Single

Dim Sfondo As Direct3DRMTexture3

Dim campos As D3DVECTOR

Dim I As Integer, j As Integer

Dim LastTime As Long
Dim NumFramesDone As Integer
Dim FrameText As String

Dim StartGameTime As Long, nowTime As Long
Dim MaxSpeed As Integer

Dim pressed As Boolean
Dim Keyright As Boolean
Dim Keyleft As Boolean
Dim Keydown As Boolean
Dim Keyup As Boolean
Dim KeyS As Boolean
Dim KeyL As Boolean
Dim KeyF0 As Boolean
Dim KeyF1 As Boolean
Dim KeyF2 As Boolean
Dim KeyF3 As Boolean
Dim KeyF4 As Boolean
Dim KeyF5 As Boolean
Dim KeyF6 As Boolean
Dim KeyF7 As Boolean
Dim KeyF8 As Boolean
Dim KeyF9 As Boolean
Dim Keyescape As Boolean
Dim KeyPagedown As Boolean
Dim KeyPageup As Boolean
Dim KeyHome As Boolean
Dim KeyEnd As Boolean

Dim Planetinfo, index As Integer, textstring As String

Dim seenspeed As Single

Dim Background As DirectDrawSurface4
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyRight Then Keyright = True
   If KeyCode = vbKeyLeft Then Keyleft = True
   If KeyCode = vbKeyDown Then Keydown = True
   If KeyCode = vbKeyEscape Then Keyescape = True
   If KeyCode = vbKeyUp Then Keyup = True
   If KeyCode = vbKeyS Then KeyS = True
   If KeyCode = vbKeyL Then KeyL = True
   If KeyCode = vbKeyPageUp Then KeyPageup = True
   If KeyCode = vbKeyPageDown Then KeyPagedown = True
   If KeyCode = vbKeyHome Then KeyHome = True
   If KeyCode = vbKeyEnd Then KeyEnd = True
   If KeyCode = vbKeyF1 Then KeyF1 = True
   If KeyCode = vbKeyF2 Then KeyF2 = True
   If KeyCode = vbKeyF3 Then KeyF3 = True
   If KeyCode = vbKeyF4 Then KeyF4 = True
   If KeyCode = vbKeyF5 Then KeyF5 = True
   If KeyCode = vbKeyF6 Then KeyF6 = True
   If KeyCode = vbKeyF7 Then KeyF7 = True
   If KeyCode = vbKeyF8 Then KeyF8 = True
   If KeyCode = vbKeyF9 Then KeyF9 = True
   If KeyCode = vbKeyF12 Then KeyF0 = True
End Sub
Private Sub form_Keyup(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyRight Then Keyright = False
   If KeyCode = vbKeyLeft Then Keyleft = False
   If KeyCode = vbKeyDown Then Keydown = False
   If KeyCode = vbKeyEscape Then Keyescape = False
   If KeyCode = vbKeyUp Then Keyup = False
   If KeyCode = vbKeyS Then KeyS = False
   If KeyCode = vbKeyL Then KeyL = False
   If KeyCode = vbKeyPageUp Then KeyPageup = False
   If KeyCode = vbKeyPageDown Then KeyPagedown = False
   If KeyCode = vbKeyHome Then KeyHome = False
   If KeyCode = vbKeyEnd Then KeyEnd = False
   If KeyCode = vbKeyF1 Then KeyF1 = False
   If KeyCode = vbKeyF2 Then KeyF2 = False
   If KeyCode = vbKeyF3 Then KeyF3 = False
   If KeyCode = vbKeyF4 Then KeyF4 = False
   If KeyCode = vbKeyF5 Then KeyF5 = False
   If KeyCode = vbKeyF6 Then KeyF6 = False
   If KeyCode = vbKeyF7 Then KeyF7 = False
   If KeyCode = vbKeyF8 Then KeyF8 = False
   If KeyCode = vbKeyF9 Then KeyF9 = False
   If KeyCode = vbKeyF12 Then KeyF0 = False
End Sub
' main sub
Public Sub init_dx(nWidth As Integer, nHeight As Integer, nDepth As Integer, nGUID As String, nDetail As Integer)
    Dim t1 As Long
    Dim Starttick As Long, LastTick As Long
    Dim ClearRec(1) As D3DRECT
    Dim ang, bakm
    Dim DDS As DDSURFACEDESC2
    Dim RECT As RECT
    Dim Systemframe, Oldtime
    Dim sx, sy, sz
    Dim countup As Boolean, countdown As Boolean
    Dim viewmode As Byte, pressed As Boolean
    Dim dist As Byte, sun As Boolean, shadows As Boolean
    Dim MyFont As New StdFont

    
    MyFont.Name = "Impact"
    MyFont.Size = 20
    MyFont.Bold = True

    StartGameTime = dx.TickCount
    MaxSpeed = 30
    Unload frmSplash
    
    t1 = dx.TickCount()
    
    Set dd = dx.DirectDraw4Create("")
    dd.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    ' this will be full-screen, so set the display mode
    dd.SetDisplayMode CLng(nWidth), CLng(nHeight), CLng(nDepth), 0, DDSDM_DEFAULT
    
    ' create the primary surface
    DDSDPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    DDSDPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    DDSDPrimary.lBackBufferCount = 1
    Set SurfPrimary = dd.CreateSurface(DDSDPrimary)
           
    ' get the back buffer
    DDCapsBack.lCaps = DDSCAPS_BACKBUFFER
    Set SurfBack = SurfPrimary.GetAttachedSurface(DDCapsBack)
    
    ' Create the Retained Mode object
    Set d3drm = dx.Direct3DRMCreate()
    
    SurfBack.SetForeColor RGB(0, 255, 0)

    Set dev = d3drm.CreateDeviceFromSurface(nGUID, dd, SurfBack, D3DRMDEVICE_DEFAULT)
    
    dev.SetBufferCount 2
    
    Select Case nDetail
       Case 0
          dev.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID
       Case 1
          dev.SetQuality D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD
          'Linear texturing looks better
          dev.SetTextureQuality D3DRMTEXTURE_LINEAR
          dev.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY
    End Select
    
    Set scene = d3drm.CreateFrame(Nothing)
    Set cam = d3drm.CreateFrame(scene)
    
    dev.SetDither D_TRUE
    Set view = d3drm.CreateViewport(dev, cam, 0, 0, Me.ScaleWidth, Me.ScaleHeight)
    view.SetBack 10000!
   
    Set mesh = d3drm.CreateMeshBuilder()
    mesh.SetPerspective D_TRUE
    scene.AddVisual mesh
  
    scene.SetZbufferMode D3DRMZBUFFER_ENABLE
    scene.SetSortMode D3DRMSORT_FRONTTOBACK
    
    viewmode = 0
    pressed = False
    
    EarthRt = 0.05
    EarthRv = 0.001
    sun = True
    
    Planets
    
    ' add light to camera
    Dim light(3) As Direct3DRMLight
    
    Set light(1) = d3drm.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
    light(1).SetColorRGB 255, 255, 255
    mFrSSun.AddLight light(1)
     
    ' add a bit of ambient light to the scene
    Set light(2) = d3drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 1, 1, 1)
    mFrSSun.AddLight light(2)
    shadows = False
    'The problem is that sun is the origin of light(see light(1)-point light) but is
    'also an object of the scene with the light source inside it.
    'Or, most probably, I don't understand how to use lights well and so I don't know how
    'to illuminate the sun without 'touch' the other planets.
    Set light(3) = d3drm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0, 0, 0)
    'Try to change these light's color values above (ex. 0.5,0.5,0.5) to see the difference

    countup = True
    countdown = False
     
    cam.SetPosition scene, -1000, 0, 0
    cam.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 90 * (pi / 180)
   
    DDS.lFlags = DDSD_WIDTH Or DDSD_HEIGHT
    DDS.lHeight = view.GetHeight / 2 + 12
    DDS.lWidth = view.GetWidth * 5 / 2

    Set Background = dd.CreateSurfaceFromFile("stars.bmp", DDS)
     
    Oldtime = dx.TickCount
    textstring = ""
    
    ShowCursor 0
    ' start main loop
    Do While DoEvents()
       Starttick = dx.TickCount
       Do Until nowTime - LastTick > MaxSpeed
          DoEvents
          seenspeed = (nowTime - Starttick) / 100
          nowTime = dx.TickCount
        Loop
        LastTick = nowTime
        
       'Camera
       'Move forward
        If Keyup = True Then
            cam.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, 5
        End If
        
        'Move back
        If Keydown = True Then
           cam.AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, -5
        End If
        
        'Rotate left
        If Keyleft = True Then
            cam.AddRotation D3DRMCOMBINE_BEFORE, 0, -1, 0, 0.05
           ang = ang + 0.05
        End If
        
        'Rotate right
        If Keyright = True Then
           cam.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, 0.05
           ang = ang - 0.05
        End If
        
        'Rotation and orbital motion velocity of each planet are referenced to Earth's
        'rotation and orbital motion velocity;
        
        'Increase velocity rotation
        If KeyPageup = True Then
           If EarthRt < 0.09 Then EarthRt = EarthRt + 0.01
           mFrMercury.SetRotation scene, 0, 1, 0, -EarthRt / 59
           mFrVenus.SetRotation scene, 0, 1, 0, EarthRt / 243
           mFrEarth.SetRotation scene, 0, 1, 0, -EarthRt
           mFrMars.SetRotation scene, 0, 1, 0, -EarthRt
           mFrJupiter.SetRotation scene, 0, 1, 0, -EarthRt * 1.98
           mFrSaturn.SetRotation scene, 0, 1, 0, -EarthRt * 1.95
           mFrUranus.SetRotation scene, 0, 1, 0, EarthRt * 1.92
           mFrNeptune.SetRotation scene, 0, 1, 0, -EarthRt * 1.75
           mFrPluto.SetRotation scene, 0, 1, 0, -EarthRt / 6
           KeyPageup = False
        End If
        'Decrease velocity rotation
        If KeyPagedown = True Then
           If EarthRt > 0.01 Then EarthRt = EarthRt - 0.01
           mFrMercury.SetRotation scene, 0, 1, 0, -EarthRt / 59
           mFrVenus.SetRotation scene, 0, 1, 0, EarthRt / 243
           mFrEarth.SetRotation scene, 0, 1, 0, -EarthRt
           mFrMars.SetRotation scene, 0, 1, 0, -EarthRt
           mFrJupiter.SetRotation scene, 0, 1, 0, -EarthRt * 1.98
           mFrSaturn.SetRotation scene, 0, 1, 0, -EarthRt * 1.95
           mFrUranus.SetRotation scene, 0, 1, 0, EarthRt * 1.92
           mFrNeptune.SetRotation scene, 0, 1, 0, -EarthRt * 1.75
           mFrPluto.SetRotation scene, 0, 1, 0, -EarthRt / 6
           KeyPagedown = False
        End If
          
        'Increase orbital motion velocity
        If KeyHome = True Then
           If EarthRv < 0.009 Then EarthRv = EarthRv + 0.001
           mFrSMercury.SetRotation mFrSun, 0, 1, 0, EarthRv / 0.24
           mFrSVenus.SetRotation mFrSun, 0, 1, 0, EarthRv / 0.615
           mFrSearth.SetRotation mFrSun, 0, 1, 0, EarthRv
           mFrSMars.SetRotation mFrSun, 0, 1, 0, EarthRv / 1.88
           mFrSJupiter.SetRotation mFrSun, 0, 1, 0, EarthRv / 11.86
           mFrSSaturn.SetRotation mFrSun, 0, 1, 0, EarthRv / 29.457
           mFrSUranus.SetRotation mFrSun, 0, 1, 0, EarthRv / 84
           mFrSNeptune.SetRotation mFrSun, 0, 1, 0, EarthRv / 164
           mFrSPluto.SetRotation mFrSun, 0, 1, 0, EarthRv / 248
           KeyHome = False
        End If
        'Decrease orbital motion velocity
        If KeyEnd = True Then
           If EarthRv > 0.001 Then EarthRv = EarthRv - 0.001
           mFrSMercury.SetRotation mFrSun, 0, 1, 0, EarthRv / 0.24
           mFrSVenus.SetRotation mFrSun, 0, 1, 0, EarthRv / 0.615
           mFrSearth.SetRotation mFrSun, 0, 1, 0, EarthRv
           mFrSMars.SetRotation mFrSun, 0, 1, 0, EarthRv / 1.88
           mFrSJupiter.SetRotation mFrSun, 0, 1, 0, EarthRv / 11.86
           mFrSSaturn.SetRotation mFrSun, 0, 1, 0, EarthRv / 29.457
           mFrSUranus.SetRotation mFrSun, 0, 1, 0, EarthRv / 84
           mFrSNeptune.SetRotation mFrSun, 0, 1, 0, EarthRv / 164
           mFrSPluto.SetRotation mFrSun, 0, 1, 0, EarthRv / 248
           KeyEnd = False
        End If
          
        If KeyF0 = True Then
           viewmode = 0
           pressed = True
           textstring = ""
           index = -1
           Timer1.Enabled = False
        End If
        If KeyF1 = True Then
           viewmode = 1
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,1, ,-, ,N,A,M,E,:, ,M,E,R,C,U,R,Y", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF2 = True Then
           viewmode = 2
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,2, ,-, ,N,A,M,E,:, ,V,E,N,U,S", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF3 = True Then
           viewmode = 3
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,3, ,-, ,N,A,M,E,:, ,E,A,R,T,H", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF4 = True Then
           viewmode = 4
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,4, ,-, ,N,A,M,E,:, ,M,A,R,S", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF5 = True Then
           viewmode = 5
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,5, ,-, ,N,A,M,E,:, ,J,U,P,I,T,E,R", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF6 = True Then
           viewmode = 6
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,6, ,-, ,N,A,M,E,:, ,S,A,T,U,R,N", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF7 = True Then
           viewmode = 7
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,7, ,-, ,N,A,M,E,:, ,U,R,A,N,U,S", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF8 = True Then
           viewmode = 8
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,8, ,-, ,N,A,M,E,:, ,N,E,P,T,U,N,E", ",", -1, 1)
           Timer1.Enabled = True
        End If
        If KeyF9 = True Then
           viewmode = 9
           index = -1
           textstring = ""
           Planetinfo = Split(",P,L,A,N,E,T, ,N,°,9, ,-, ,N,A,M,E,:, ,P,L,U,T,O", ",", -1, 1)
           Timer1.Enabled = True
        End If

        If viewmode = 0 And pressed = True Then
           cam.SetPosition mFrSun, campos.x + dist, campos.y, campos.z
           pressed = False
        End If
        If viewmode = 1 Then
           mFrMercury.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 20, campos.y, campos.z
           cam.LookAt mFrMercury, Nothing, D3DRMCONSTRAIN_Z
           dist = 20
        End If
        If viewmode = 2 Then
           mFrVenus.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 20, campos.y, campos.z
           cam.LookAt mFrVenus, Nothing, D3DRMCONSTRAIN_Z
           dist = 20
        End If
        If viewmode = 3 Then
           mFrEarth.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 20, campos.y, campos.z
           cam.LookAt mFrEarth, Nothing, D3DRMCONSTRAIN_Z
           dist = 20
        End If
        If viewmode = 4 Then
           mFrMars.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 10, campos.y, campos.z
           cam.LookAt mFrMars, Nothing, D3DRMCONSTRAIN_Z
           dist = 10
        End If
        If viewmode = 5 Then
           mFrJupiter.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 200, campos.y, campos.z
           cam.LookAt mFrJupiter, Nothing, D3DRMCONSTRAIN_Z
           dist = 200
        End If
        If viewmode = 6 Then
           mFrSaturn.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 180, campos.y, campos.z
           cam.LookAt mFrSaturn, Nothing, D3DRMCONSTRAIN_Z
           dist = 180
        End If
        If viewmode = 7 Then
           mFrUranus.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 80, campos.y, campos.z
           cam.LookAt mFrUranus, Nothing, D3DRMCONSTRAIN_Z
           dist = 80
        End If
        If viewmode = 8 Then
           mFrNeptune.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 80, campos.y, campos.z
           cam.LookAt mFrNeptune, Nothing, D3DRMCONSTRAIN_Z
           dist = 80
        End If
        If viewmode = 9 Then
           mFrPluto.GetPosition mFrSun, campos
           cam.SetPosition mFrSun, campos.x + 10, campos.y, campos.z
           cam.LookAt mFrPluto, Nothing, D3DRMCONSTRAIN_Z
           dist = 10
        End If

        'Sun on/off
        If KeyS = True Then
            If sun = True Then
              mFrSun.DeleteVisual DxMeshSun
              mFrFlare.DeleteVisual DxMeshflare
              sun = False
           Else
              mFrSun.AddVisual DxMeshSun
              mFrFlare.AddVisual DxMeshflare
              sun = True
           End If
           KeyS = False
        End If
        
        'Shadow on/off
        If KeyL = True Then
           If shadows = False Then
               mFrSSun.DeleteLight light(2)
               mFrSSun.AddLight light(3)
               shadows = True
           Else
               mFrSSun.DeleteLight light(3)
               mFrSSun.AddLight light(2)
               shadows = False
           End If
            KeyL = False
        End If
       
        'fix the angle of rotation
        If ang > 3.141592 * 2 Then ang = ang - 3.141592 * 2
        If ang < 0 Then ang = ang + 3.141592 * 2
        
        'Uncomment these line below if you want have a sort of pulsing Sun
'        If countup = True Then
'        sx = sx + 0.5: sy = sy + 0.5: sz = sz + 0.5
'        mFrFlare.AddScale D3DRMCOMBINE_REPLACE, sx, sy, sz
'        mFrFlare.SetPosition mFrSun, 0, 0, 0
'        mo.dcDiffuse.a = mo.dcDiffuse.a + 0.15
'        mFrFlare.SetMaterialOverride mo
'        If mo.dcDiffuse.a > 0.85 Then
'           countup = False
'           countdown = True
'        End If
'        End If
'
'        If countdown = True Then
'        sx = sx - 0.5: sy = sy - 0.5: sz = sz - 0.5
'        mFrFlare.AddScale D3DRMCOMBINE_REPLACE, sx, sy, sz
'        mFrFlare.SetPosition mFrSun, 0, 0, 0
'        mo.dcDiffuse.a = mo.dcDiffuse.a - 0.15
'        mFrFlare.SetMaterialOverride mo
'        If mo.dcDiffuse.a < 0.25 Then
'           countup = True
'           countdown = False
'        End If
'        End If
        
        'Let's do movements
        mFrSSun.Move 1
        mFrSMercury.Move 1
        mFrSVenus.Move 1
        mFrMoon.Move 1
        
        mFrSearth.Move 1
        mFrSMars.Move 1
        mFrSJupiter.Move 1
        
        mFrSSaturn.Move 1
        
        mFrSUranus.Move 1
        mFrSNeptune.Move 1
        mFrSPluto.Move 1
       
        'Check to exit
        If Keyescape = True Then Unload Me: EndIT
        
       'Clear view
        view.Clear D3DRMCLEAR_ALL

       'Clear Rect
       ClearRec(0).X1 = view.GetX
       ClearRec(0).Y1 = view.GetY
       ClearRec(0).X2 = view.GetWidth
       ClearRec(0).Y2 = view.GetHeight
       'Calculate the Background X to display
       bakm = (6.28 - ang) / 6.28 * view.GetWidth * 2
       If bakm > view.GetWidth * 2 Then bakm = view.GetWidth * 2
       If bakm < 0 Then bakm = 0
       'Display the part of the background
       SurfBack.Blt REC(0, 0, 0, 0), Background, REC(bakm, 0, bakm + view.GetWidth / 2, view.GetHeight / 2), DDBLT_WAIT
  
       'Render the scene
       view.Render scene
       dev.Update
       
       SurfBack.SetForeColor RGB(0, 255, 0)
       SurfBack.SetFont MyFont
       SurfBack.DrawText 10, Frmmain.ScaleHeight - 50, textstring, False
       SurfPrimary.Flip Nothing, DDFLIP_WAIT
    
    Loop
    
End Sub

Sub EndIT()
'Goodbye at all
    Set DxMeshSun = Nothing
    Set DxMeshflare = Nothing
    Set DxMeshMercury = Nothing
    
    Set DxMeshVenus = Nothing
    
    Set DxMeshEarth = Nothing
    Set DxMeshClouds = Nothing
    Set DxMeshMoon = Nothing
    
    Set DxMeshMars = Nothing
    
    Set DxMeshJupiter = Nothing
    
    Set DxMeshSaturn = Nothing
    Set DxMeshRing = Nothing
    
    Set DxMeshUranus = Nothing
    
    Set DxMeshNeptune = Nothing
    
    Set DxMeshPluto = Nothing
    Set mFrSSun = Nothing
    Set mFrSun = Nothing
    Set mFrFlare = Nothing
   
    Set mFrSMercury = Nothing
    Set mFrMercury = Nothing

    Set mFrSVenus = Nothing
    Set mFrVenus = Nothing
    
    Set mFrSearth = Nothing
    Set mFrEarth = Nothing
    Set mFrClouds = Nothing
    Set mFrMoon = Nothing
    
    Set mFrSMars = Nothing
    Set mFrMars = Nothing
    
    Set mFrSJupiter = Nothing
    Set mFrJupiter = Nothing
    
    Set mFrSSaturn = Nothing
    Set mFrRing = Nothing
    Set mFrSaturn = Nothing
       
    Set mFrSUranus = Nothing
    Set mFrUranus = Nothing
        
    Set mFrSNeptune = Nothing
    Set mFrNeptune = Nothing
        
    Set mFrSPluto = Nothing
    Set mFrPluto = Nothing
    
    Set scene = Nothing
    Set cam = Nothing
    Set tex = Nothing

    Call dd.RestoreDisplayMode
    Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    
    Set d3drm = Nothing
    Set dd = Nothing
    Set dx = Nothing
    
    ShowCursor 1
    
    End
End Sub
Function REC(x, y, xa, ya) As RECT
    REC.Bottom = ya
    REC.Left = x
    REC.Right = xa
    REC.Top = y
End Function
Private Sub Planets()
   
    'Meshes
    Set DxMeshSun = d3drm.CreateMeshBuilder()
    Set DxMeshflare = d3drm.CreateMeshBuilder()
    
    Set DxMeshMercury = d3drm.CreateMeshBuilder()
    
    Set DxMeshVenus = d3drm.CreateMeshBuilder()
    
    Set DxMeshEarth = d3drm.CreateMeshBuilder()
    Set DxMeshClouds = d3drm.CreateMeshBuilder()
    Set DxMeshMoon = d3drm.CreateMeshBuilder()
    
    Set DxMeshMars = d3drm.CreateMeshBuilder()
    
    Set DxMeshJupiter = d3drm.CreateMeshBuilder()
    
    Set DxMeshSaturn = d3drm.CreateMeshBuilder()
    Set DxMeshRing = d3drm.CreateMeshBuilder()
    
    Set DxMeshUranus = d3drm.CreateMeshBuilder()
    
    Set DxMeshNeptune = d3drm.CreateMeshBuilder()
    
    Set DxMeshPluto = d3drm.CreateMeshBuilder()
    
    'Frames
    Set mFrSSun = d3drm.CreateFrame(scene)
    Set mFrSun = d3drm.CreateFrame(mFrSSun)
    Set mFrFlare = d3drm.CreateFrame(mFrSSun)
   
    Set mFrSMercury = d3drm.CreateFrame(mFrSSun)
    Set mFrMercury = d3drm.CreateFrame(mFrSMercury)

    Set mFrSVenus = d3drm.CreateFrame(mFrSSun)
    Set mFrVenus = d3drm.CreateFrame(mFrSVenus)
    
    Set mFrSearth = d3drm.CreateFrame(mFrSSun)
    Set mFrEarth = d3drm.CreateFrame(mFrSearth)
    Set mFrClouds = d3drm.CreateFrame(mFrEarth)
    Set mFrMoon = d3drm.CreateFrame(mFrEarth)
    
    Set mFrSMars = d3drm.CreateFrame(mFrSSun)
    Set mFrMars = d3drm.CreateFrame(mFrSMars)
    
    Set mFrSJupiter = d3drm.CreateFrame(mFrSSun)
    Set mFrJupiter = d3drm.CreateFrame(mFrSJupiter)
    
    Set mFrSSaturn = d3drm.CreateFrame(mFrSSun)
    Set mFrRing = d3drm.CreateFrame(mFrSSaturn)
    Set mFrSaturn = d3drm.CreateFrame(mFrSSaturn)
       
    Set mFrSUranus = d3drm.CreateFrame(mFrSSun)
    Set mFrUranus = d3drm.CreateFrame(mFrSUranus)
        
    Set mFrSNeptune = d3drm.CreateFrame(mFrSSun)
    Set mFrNeptune = d3drm.CreateFrame(mFrSNeptune)
        
    Set mFrSPluto = d3drm.CreateFrame(mFrSSun)
    Set mFrPluto = d3drm.CreateFrame(mFrSPluto)
       
    'Sun system
    With DxMeshSun
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("sun.bmp")
        .ScaleMesh 100, 100, 100
    End With
      
    mFrSun.SetPosition mFrSun, 0, 0, 0
    mFrSun.SetRotation mFrSSun, 0, 1, 0, -EarthRt / 25
   
    Set tex = d3drm.LoadTexture("flare.bmp")
    tex.SetDecalTransparency D_TRUE
    tex.SetDecalTransparentColor 0
    
    With DxMeshflare
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture tex
        .ScaleMesh 110, 110, 110
    End With
        
    mFrFlare.SetRotation mFrSSun, 0, 1, 0, -EarthRt / 25
    mFrFlare.SetPosition mFrSun, 0, 0, 0
    
    With mo
        .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
        .dcDiffuse.a = 0.55
    End With
    mFrFlare.SetMaterialOverride mo
       
    mFrSSun.SetPosition mFrSun, -500, 0, 0
       
    'Mercury System
    With DxMeshMercury
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Mercury.bmp")
        .ScaleMesh 3.9, 3.9, 3.9
    End With
       
    mFrMercury.SetPosition mFrSun, 57.9, 0, 0
    mFrMercury.SetRotation scene, 0, 1, 0, -EarthRt / 59
    mFrSMercury.SetRotation mFrSun, 0, 1, 0, (EarthRv / 0.24)
    
    
    'Venus System
    With DxMeshVenus
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Venus.bmp")
        .ScaleMesh 4, 4, 4
    End With
       
    mFrVenus.SetPosition mFrSun, 108.2, 0, 0
    mFrVenus.SetRotation scene, 0, 1, 0, EarthRt / 243
    mFrSVenus.SetRotation mFrSun, 0, 1, 0, EarthRv / 0.615
    
    
    'Earth System
    With DxMeshEarth
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("earth.bmp")
        .ScaleMesh 4, 4, 4
    End With
       
    mFrEarth.SetPosition mFrSun, 149.5, 0, 0
    mFrEarth.SetRotation scene, 0, 1, 0, -EarthRt
    
    Set tex = d3drm.LoadTexture("clouds.bmp")
    tex.SetDecalTransparency D_TRUE
    tex.SetDecalTransparentColor 0
    
    With DxMeshClouds
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture tex
        .ScaleMesh 4.25, 4.25, 4.25
    End With
        
    mFrClouds.SetRotation mFrEarth, 0, 1, 0, -0.075
    mFrClouds.SetPosition mFrEarth, 0, 0, 0
    
    With mo
        .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
        .dcDiffuse.a = 0.55
    End With
    mFrClouds.SetMaterialOverride mo
    
    With DxMeshMoon
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("moon.bmp")
        .ScaleMesh 2, 2, 2
    End With
   
    mFrMoon.SetPosition mFrEarth, 5, 0, 5
    mFrMoon.SetRotation scene, 0, 1, 0, -EarthRt / 29
    mFrMoon.SetRotation mFrEarth, 0, 1, 0, -EarthRt / 29
    mFrSearth.SetRotation mFrSun, 0, 1, 0, EarthRv
    
    
    'Mars System
    With DxMeshMars
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Mars.bmp")
        .ScaleMesh 2, 2, 2
    End With
       
    mFrMars.SetPosition mFrSun, 228.1, 0, 0
    mFrMars.SetRotation scene, 0, 1, 0, -EarthRt
    mFrSMars.SetRotation mFrSun, 0, 1, 0, EarthRv / 1.88
    
    'Jupiter System
    With DxMeshJupiter
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Jupiter.bmp")
        .ScaleMesh 44.2, 44.2, 44.2
    End With
       
    mFrJupiter.SetPosition mFrSun, 779, 0, 0
    mFrJupiter.SetRotation scene, 0, 1, 0, -EarthRt * 1.98
    mFrSJupiter.SetRotation mFrSun, 0, 1, 0, EarthRv / 11.86
  
    'Saturn System
    With DxMeshSaturn
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Saturn.bmp")
        .ScaleMesh 40, 40, 40
    End With
    
    With DxMeshRing
        .LoadFromFile "ring.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("venus.bmp")
        .ScaleMesh 3, 3, 0.1
    End With

    With mo
        .lFlags = D3DRMMATERIALOVERRIDE_DIFFUSE_ALPHAONLY
        .dcDiffuse.a = 0.75
    End With
    mFrRing.SetMaterialOverride mo

    mFrRing.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 4.85
    mFrRing.SetPosition mFrSun, 1428, 0, 0
    
    mFrSaturn.SetPosition mFrSun, 1428, 0, -9
    mFrSaturn.SetRotation scene, 0, 1, 0, -EarthRt * 1.95
    
    mFrSSaturn.SetRotation mFrSun, 0, 1, 0, EarthRv / 29.457
    
     
    'Uranus System
    With DxMeshUranus
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Uranus.bmp")
        .ScaleMesh 17, 17, 17
    End With
       
    mFrUranus.SetPosition mFrSun, 2871, 0, 0
    mFrUranus.SetRotation scene, 0, 1, 0, EarthRt * 1.92
    mFrSUranus.SetRotation mFrSun, 0, 1, 0, EarthRv / 84
   
    'Neptune System
    With DxMeshNeptune
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Neptune.bmp")
        .ScaleMesh 16, 16, 16
    End With
       
    mFrNeptune.SetPosition mFrSun, 4499, 0, 0
    mFrNeptune.SetRotation mFrSun, 0, 1, 0, -EarthRt * 1.75
    mFrSNeptune.SetRotation mFrSun, 0, 1, 0, EarthRv / 164
    
    
    'Pluto System
    With DxMeshPluto
        .LoadFromFile "sphere.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
        .SetTexture d3drm.LoadTexture("Pluto.bmp")
        .ScaleMesh 2, 2, 2
    End With
       
    mFrPluto.SetPosition mFrSun, 5908, 0, 0
    mFrPluto.SetRotation mFrSun, 0, 1, 0, -EarthRt / 6
    mFrSPluto.SetRotation mFrSun, 0, 1, 0, EarthRv / 248
   
    ''''''''''''''''''''''''''''''''''''''''''''
    
    mFrSun.AddVisual DxMeshSun
    mFrFlare.AddVisual DxMeshflare
  
    mFrMercury.AddVisual DxMeshMercury
    
    mFrVenus.AddVisual DxMeshVenus
    
    mFrEarth.AddVisual DxMeshEarth
    mFrClouds.AddVisual DxMeshClouds
    mFrMoon.AddVisual DxMeshMoon
   
    mFrMars.AddVisual DxMeshMars
    mFrJupiter.AddVisual DxMeshJupiter
    mFrSaturn.AddVisual DxMeshSaturn
    mFrRing.AddVisual DxMeshRing
    
    mFrUranus.AddVisual DxMeshUranus
    mFrNeptune.AddVisual DxMeshNeptune
    mFrPluto.AddVisual DxMeshPluto
  
    scene.AddVisual mFrSSun
    
    scene.AddVisual mFrSMercury
    scene.AddVisual mFrSVenus
    scene.AddVisual mFrSearth
    scene.AddVisual mFrSMars
    scene.AddVisual mFrSJupiter
    scene.AddVisual mFrSSaturn
    scene.AddVisual mFrSUranus
    scene.AddVisual mFrNeptune
    scene.AddVisual mFrPluto

   
End Sub

Private Sub TypeWrite(text)
index = index + 1
If index <= UBound(text) Then textstring = textstring & Planetinfo(index)

End Sub
Private Sub Timer1_Timer()
TypeWrite (Planetinfo)
End Sub


