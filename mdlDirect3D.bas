Attribute VB_Name = "mdlDirect3D"
'Copyright (C) 2003 Richard W. Allen
'Program Name  DX Project
'Author        Richard W. Allen
'Version       V1.0B
'Date Started  March 01, 2002
'Date End      April 8, 2003
'DX Project Copyright (C) 2003 Richard W. Allen Dx Project Comes with ABSOLUTELY NO WARRANTY;
'DX Project is licensed under the GNU GENERAL PUBLIC LICENSE Version 2.
'for details see the license.txt include with this program.
Option Explicit

Public D3D As Direct3D8      'The Direct3D Interface

Public lngAdapters As Long 'How many adapters found
Public AdapterInfo As D3DADAPTER_IDENTIFIER8 'A Structure holding information on the adapter

Public lngModes As Long 'How many display modes found

Public D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Public bRunning As Boolean 'Controls whether the program is running or not...

Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

'New March 22, 2003
Public sngUpDown As Single
Public sngLeftRight As Single
Public sngBackFront As Single
'end of New

'New March 25, 2003
Public Visable As Boolean
'End New

Dim TriStrip(0 To 3) As TLVERTEX '//We're going to have two squares - one with colour, the other without
Dim TriStrip2(0 To 3) As TLVERTEX
Dim TrisStrip3(0 To 3) As TLVERTEX '//This is going to be our transparent part - it'll follow the mouse...

'New March 23, 2003
Dim intTextureNum As Integer
Dim intVertexNum As Integer
Dim sngVertical As Single
Dim sngHorizontal As Single

'New March 25, 2003
Dim sngTextureVertical As Single
Dim sngTextureHorizontal As Single

Dim TriStripX(0 To 3, 0 To 3) As TLVERTEX
'End New

'New March 22, 2003
Dim TriStripColour(0 To 2, 0 To 3) As Integer
'End New

Dim D3DX As D3DX8 '//A helper library
Dim Texture As Direct3DTexture8
Dim TransTexture As Direct3DTexture8 '//This texture will have transparency information encoded into it....

Dim ColourDown As Boolean

'Display Setting's
Public intVsyncOn As Integer
Public intDisplayFTP As Integer
Public intDisplayWidth As Integer
Public intDisplayHight As Integer
Public strDisplayColour As String
Public strDisplayDevice As String
'New March 22, 2003
Public intLowResTextures As Integer
Public intModescreen As Integer
'End New

'New March 27, 2003         For text
Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont
'End New Code




Public Declare Function GetTickCount Lib "kernel32" () As Long '//This is used to get the frame rate.
Public FPS_LastCheck As Long
Public FPS_Count As Long
Public FPS_Current As Integer




'// Initialise : This procedure kick starts the whole process.
'// It'll return true for success, false if there was an error.
Public Function Initialise() As Boolean
On Error GoTo ErrHandler:

    sngUpDown = 200
    sngLeftRight = 200
    
    Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode
    Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport
    Dim ColorKeyVal As Long '//What colour becomes transparent...
    
    Set Dx = New DirectX8  '//Create our Master Object
    Set D3D = Dx.Direct3DCreate() '//Make our Master Object create the Direct3D Interface
    Set D3DX = New D3DX8 '//Create our helper library...
    
    
    'New March 26, 2003
    Set DSEnum = Dx.GetDSEnum
    Set DS = Dx.DirectSoundCreate(DSEnum.GetGuid(1))
    
    DS.SetCooperativeLevel frmSinglePlayerGame.hWnd, DSSCL_NORMAL
    
    DSBDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    Set DSBuffer = DS.CreateSoundBufferFromFile(App.Path & "\Sample.wav", DSBDesc)
    
   ' lngFrequency = 1102
    
   ' DSBuffer.SetFrequency lngFrequency
   ' DSBuffer.SetPan 0
    
   ' DSBuffer.SetVolume 0
    
    
    'DSBDesc.fxFormat.lSamplesPerSec
    'scrlFrq.Value = CInt(DSBDesc.fxFormat.lSamplesPerSec / 10)
    'end new Code
    
    
    
    D3DWindow.Windowed = intModescreen 'New March 22, 2003 intModescreen Var add so user can pic between Full screen and Windows modes
    
    'Sets the Colour up
    Select Case (strDisplayColour)
        Case Is = "D3DFMT_X8R8G8B8"
            DispMode.Format = D3DFMT_X8R8G8B8
        Case Is = "D3DFMT_R5G6B5"
            DispMode.Format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...
    End Select
    'Sets the Display Hight and Width
    DispMode.Width = intDisplayWidth
    DispMode.Height = intDisplayHight
    
    If intVsyncOn = 1 Then
        D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    Else
        D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
    End If
    
    
    D3DWindow.BackBufferCount = 1 '//1 backbuffer only
    D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
    D3DWindow.BackBufferHeight = DispMode.Height
    D3DWindow.BackBufferWidth = DispMode.Width
    D3DWindow.hDeviceWindow = frmSinglePlayerGame.hWnd
    
    '//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
    '//See the lesson text for more information on this line...
    'Sets the the Display Redering device to D3DDEVTYPE_HAL for Hardware or D3DDEVTYPE_REF for software
    Select Case (strDisplayDevice)
        Case Is = "D3DDEVTYPE_HAL"
            Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmSinglePlayerGame.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
        Case Is = "D3DDEVTYPE_REF"
            Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_REF, frmSinglePlayerGame.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                            D3DWindow)
    End Select
        
    '//Set the vertex shader to use our vertex format
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    
    '//We now want to load our texture;
    If intLowResTextures = 1 Then
        Set Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\ExampleTexture.bmp", 8, 8, _
                                                                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                                D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                                D3DX_FILTER_POINT, ColorKeyVal, _
                                                                                ByVal 0, ByVal 0)
    Else
        Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\ExampleTexture.bmp")
    End If
    
    'New march 27, 2003
    'D3DDevice.SetRenderState D3DRS_ZENABLE, 1
    'End New code
    
    '//Choose one of the following depending on what you
    '   should need. Other colours can be made up, but these
    '   ones should be okay for most uses...
    
    'ColorKeyVal = &HFF000000 '//Black
    'ColorKeyVal = &HFFFF0000 '//Red
    ColorKeyVal = &HFF00FF00 '//Green
    'ColorKeyVal = &HFF0000FF '//Blue
    'ColorKeyVal = &HFFFF00FF '//Magenta
    'ColorKeyVal = &HFFFFFF00 '//Yellow
    'ColorKeyVal = &HFF00FFFF '//Cyan
    'ColorKeyVal = &HFFFFFFFF '//White
    
    Set TransTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\transtexture.bmp", 64, 64, _
                                                                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                                D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                                D3DX_FILTER_POINT, ColorKeyVal, _
                                                                                ByVal 0, ByVal 0)
    
    '//We can only continue if Initialise Geometry succeeds;
    '   If it doesn't we'll fail this call as well...
    If InitialiseGeometry() = True Then
        Initialise = True '//We succeeded
        Exit Function
    End If
    
 
    
    
ErrHandler:
    '//We failed; for now we wont worry about why.
    Debug.Print "Error Number Returned: " & Err.Number
    Initialise = False
End Function
Public Sub Render()
'//1. We need to clear the render device before we can draw anything
'       This must always happen before you start rendering stuff...
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0 '//Clear the screen black

'//2. Rendering the graphics...

D3DDevice.BeginScene
    'All rendering calls go between these two lines\
    
    Call InitialiseGeometry
    
    D3DDevice.SetTexture 0, Texture '//Tell the device which texture we want to use...
    
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    
        
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip2(0), Len(TriStrip2(0))
    
    D3DDevice.SetTexture 0, TransTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TrisStrip3(0), Len(TrisStrip3(0))
    
    'New March 25, 2003
    Dim intTextNum As Integer
    
    
    'If Visable = True Then
        D3DDevice.SetTexture 0, Texture
        Do While intTextNum < 4                     'Note TriStripX is an array of Squares TriStripX(Vertex Number, Square Number)
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStripX(0, intTextNum), Len(TriStripX(0, intTextNum))
            intTextNum = intTextNum + 1
        Loop
    'End If
    
    'end New
    
    'New code march 27, 2003
    '//Render boring 2D text
    If intDisplayFTP = 1 Then
        TextRect.Top = 1
        TextRect.Left = 1
        TextRect.bottom = 100
        TextRect.Right = 300
        D3DX.DrawText MainFont, &HFFCCCCFF, "Current Frame Rate: " & FPS_Current, TextRect, DT_TOP Or DT_LEFT
    End If
    'end new code
    
    
    
    'New March 25, 2003 Not need now See loop above
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStripX(0, 1), Len(TriStripX(0, 1))
    'End New
    
    
D3DDevice.EndScene

'//3. Update the frame to the screen...
'       This is the same as the Primary.Flip method as used in DirectX 7
'       These values below should work for almost all cases...
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Function InitialiseGeometry() As Boolean
    
    
On Error GoTo BailOut: '//Setup our Error handler

'//NOTE THAT WE ARE PASSING VALUES FOR THE tu AND tv ARGUMENTS


'## FIRST SQUARE ##
            
                 
            'vertex 0
            TriStrip(0) = CreateTLVertex(0, 0, 1, 1, RGB(TriStripColour(0, 0), TriStripColour(1, 0), TriStripColour(2, 0)), 0, 0, 0)
                    
            'vertex 1
            TriStrip(1) = CreateTLVertex(200, 0, 1, 1, RGB(TriStripColour(0, 1), TriStripColour(1, 1), TriStripColour(2, 1)), 0, 1, 0)
          
            'vertex 2
            TriStrip(2) = CreateTLVertex(0, 200, 1, 1, RGB(TriStripColour(0, 2), TriStripColour(1, 2), TriStripColour(2, 2)), 0, 0, 1)
         
            'vertex 3
            TriStrip(3) = CreateTLVertex(200, 200, 1, 1, RGB(TriStripColour(0, 3), TriStripColour(1, 3), TriStripColour(2, 3)), 0, 1, 1)
           

'## SECOND SQUARE ##

            'vertex 0
            TriStrip2(0) = CreateTLVertex(sngLeftRight, sngUpDown, sngBackFront, 1, RGB(255, 255, 255), 0, 0, 0)
            
            'vertex 1
            TriStrip2(1) = CreateTLVertex(sngLeftRight + 200, sngUpDown, sngBackFront, 1, RGB(255, 255, 255), 0, 1, 0)
            
            'vertex 2
            TriStrip2(2) = CreateTLVertex(sngLeftRight, sngUpDown + 200, sngBackFront, 1, RGB(255, 255, 255), 0, 0, 1)
            
            'vertex 3
            TriStrip2(3) = CreateTLVertex(sngLeftRight + 200, sngUpDown + 200, sngBackFront, 1, RGB(255, 255, 255), 0, 1, 1)
            
            'New March 23, 2003
            
          '  TriStripX(0, 0) = CreateTLVertex(0, 400, 1, 1, RGB(255, 255, 255), 0, 0, 0)
                    
         '   TriStripX(1, 0) = CreateTLVertex(200, 400, 1, 1, RGB(255, 255, 255), 0, 1, 0)
               
         '   TriStripX(2, 0) = CreateTLVertex(0, 600, 1, 1, RGB(255, 255, 255), 0, 0, 1)
               
         '   TriStripX(3, 0) = CreateTLVertex(200, 600, 1, 1, RGB(255, 255, 255), 0, 1, 1)
            
            'End New
            
            'New March 25, 2003
          '  TriStripX(0, 1) = CreateTLVertex(200, 400, 1, 1, RGB(255, 255, 255), 0, 0, 0)
                    
          '  TriStripX(1, 1) = CreateTLVertex(400, 400, 1, 1, RGB(255, 255, 255), 0, 1, 0)
               
         '   TriStripX(2, 1) = CreateTLVertex(200, 600, 1, 1, RGB(255, 255, 255), 0, 0, 1)
               
         '   TriStripX(3, 1) = CreateTLVertex(400, 600, 1, 1, RGB(255, 255, 255), 0, 1, 1)
            'End new
            
               
            'New March 23, 2003
            intVertexNum = 0
            intTextureNum = 0
            
            Do While intTextureNum < 4
                Do While intVertexNum < 4
                    Call PointChange
                    TriStripX(intVertexNum, intTextureNum) = CreateTLVertex(sngVertical, sngHorizontal, 0, 1, RGB(255, 255, 255), 0, sngTextureVertical, sngTextureHorizontal)
                    intVertexNum = intVertexNum + 1
                Loop
                intVertexNum = 0
                intTextureNum = intTextureNum + 1
            Loop
            'End New
                    
            
InitialiseGeometry = True

If GetTickCount >= 60000 Then
    Call ColourChange
End If
    
Exit Function
BailOut:
InitialiseGeometry = False
End Function

'//This is just a simple wrapper function that makes filling the structures much much easier...
Public Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX

    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
    CreateTLVertex.X = X
    CreateTLVertex.Y = Y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.specular = specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv
End Function

Public Sub ColourChange() 'New Public Sub March 22, 2003
    
    'This Sub incoments The colour light for the First texture up and down.
    Dim intColour As Integer
    Dim intPoint As Integer
        
    If TriStripColour(0, 0) >= 255 Then
        ColourDown = True
    End If
    If TriStripColour(0, 0) <= 0 Then
        ColourDown = False
    End If
    
    If ColourDown = False Then
        Do While intPoint < 4
            Do While intColour < 3
                TriStripColour(intColour, intPoint) = TriStripColour(intColour, intPoint) + 1
                intColour = intColour + 1
            Loop
            intPoint = intPoint + 1
            intColour = 0
        Loop
    End If
    
       
    If ColourDown = True Then
        intColour = 0
        intPoint = 0
        Do While intPoint < 4
            Do While intColour < 3
                TriStripColour(intColour, intPoint) = TriStripColour(intColour, intPoint) - 1
                intColour = intColour + 1
            Loop
            intPoint = intPoint + 1
            intColour = 0
        Loop
    End If

    
    
End Sub

Public Sub PointChange()
    Select Case intTextureNum
        Case (0)
            Select Case intVertexNum
                Case (0)
                    sngVertical = 0
                    sngHorizontal = 400
                    sngTextureVertical = 0
                    sngTextureHorizontal = 0
                Case (1)
                    sngVertical = 200
                    sngHorizontal = 400
                    sngTextureVertical = 1
                    sngTextureHorizontal = 0
                Case (2)
                    sngVertical = 0
                    sngHorizontal = 600
                    sngTextureVertical = 0
                    sngTextureHorizontal = 1
                Case (3)
                    sngVertical = 200
                    sngHorizontal = 600
                    sngTextureVertical = 1
                    sngTextureHorizontal = 1
            End Select
        Case (1)
             Select Case intVertexNum
                Case (0)
                    sngVertical = 0 + 200
                    sngHorizontal = 400
                    sngTextureVertical = 0
                    sngTextureHorizontal = 0
                Case (1)
                    sngVertical = 200 + 200
                    sngHorizontal = 400
                    sngTextureVertical = 1
                    sngTextureHorizontal = 0
                Case (2)
                    sngVertical = 0 + 200
                    sngHorizontal = 600
                    sngTextureVertical = 0
                    sngTextureHorizontal = 1
                Case (3)
                    sngVertical = 200 + 200
                    sngHorizontal = 600
                    sngTextureVertical = 1
                    sngTextureHorizontal = 1
            End Select
        Case (2)
             Select Case intVertexNum
                Case (0)
                    sngVertical = 0 + 400
                    sngHorizontal = 400
                    sngTextureVertical = 0
                    sngTextureHorizontal = 0
                Case (1)
                    sngVertical = 200 + 400
                    sngHorizontal = 400
                    sngTextureVertical = 1
                    sngTextureHorizontal = 0
                Case (2)
                    sngVertical = 0 + 400
                    sngHorizontal = 600
                    sngTextureVertical = 0
                    sngTextureHorizontal = 1
                Case (3)
                    sngVertical = 200 + 400
                    sngHorizontal = 600
                    sngTextureVertical = 1
                    sngTextureHorizontal = 1
            End Select
        Case (3)
             Select Case intVertexNum
                Case (0)
                    sngVertical = 0 + 600
                    sngHorizontal = 400
                    sngTextureVertical = 0
                    sngTextureHorizontal = 0
                Case (1)
                    sngVertical = 200 + 600
                    sngHorizontal = 400
                    sngTextureVertical = 1
                    sngTextureHorizontal = 0
                Case (2)
                    sngVertical = 0 + 600
                    sngHorizontal = 600
                    sngTextureVertical = 0
                    sngTextureHorizontal = 1
                Case (3)
                    sngVertical = 200 + 600
                    sngHorizontal = 600
                    sngTextureVertical = 1
                    sngTextureHorizontal = 1
            End Select
    End Select
End Sub

Public Sub text()
    '## 2D TEXT ##
    'New Code 27, 2003
    fnt.Name = "Verdana"
    fnt.Size = 12
    fnt.Bold = True
    Set MainFontDesc = fnt
    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    'End New Code
End Sub
