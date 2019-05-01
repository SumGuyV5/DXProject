Attribute VB_Name = "mdlGraphicOptions"
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
Public Sub DisplayAdapters()
    
    Dim intForLoop As Integer, strTemp As String, intForLoopGetChr As Integer
    
    '//This'll either be 1 or 2
    lngAdapters = D3D.GetAdapterCount
    
    For intForLoop = 0 To lngAdapters - 1
        'Get the relevent Details
        D3D.GetAdapterIdentifier intForLoop, 0, AdapterInfo
        
        'Get the name of the current adapter - it's stored as a long
        'list of character codes that we need to parse into a string
        ' - Dont ask me why they did it like this; seems silly really :)
        strTemp = "" 'Reset the string ready for our use
        
        For intForLoopGetChr = 0 To 511
            strTemp = strTemp & Chr$(AdapterInfo.Description(intForLoopGetChr)) 'Gets Each character
        Next intForLoopGetChr
        
        strTemp = Replace(strTemp, Chr$(0), " ")
        frmGraphicOptions.cmbAdapters.AddItem strTemp
    Next intForLoop
End Sub
Public Sub DisplayDevices()
    On Local Error Resume Next '//We want to handle the errors...
    Dim Caps As D3DCAPS8

    D3D.GetDeviceCaps frmGraphicOptions.cmbAdapters.ListIndex, D3DDEVTYPE_HAL, Caps
        If Err.Number = D3DERR_NOTAVAILABLE Then
            'There is no hardware acceleration
            frmGraphicOptions.cmbDevice.AddItem "Software Rendering" 'Reference device will always be available
        Else
            frmGraphicOptions.cmbDevice.AddItem "Software Rendering" 'Reference device will always be available
            frmGraphicOptions.cmbDevice.AddItem "Hardware Acceleration Rendering"
        End If
End Sub
Public Sub DisplyModes(Renderer As Long)
    frmGraphicOptions.cmbRes.Clear '//Remove any existing entries...

    Dim intForLoop As Integer, ModeTemp As D3DDISPLAYMODE
    Dim blnHi32 As Boolean, blnLow32 As Boolean
    Dim blnHi16 As Boolean, blnLow16 As Boolean

    lngModes = D3D.GetAdapterModeCount(frmGraphicOptions.cmbAdapters.ListIndex)

    For intForLoop = 0 To lngModes - 1 '//Cycle through them and collect the data...
        Call D3D.EnumAdapterModes(frmGraphicOptions.cmbAdapters.ListIndex, intForLoop, ModeTemp)
        
        Select Case (ModeTemp.Format)
        
            Case Is = D3DFMT_X8R8G8B8
                If ModeTemp.Width = 800 And ModeTemp.Height = 600 Then
                    If blnLow32 = False Then
                        frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 32 bit"
                        blnLow32 = True
                    End If
                End If
                If ModeTemp.Width = 1024 And ModeTemp.Height = 768 Then
                    If blnHi32 = False Then
                        frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 32 bit"
                        blnHi32 = True
                    End If
                End If
            Case Is = D3DFMT_R5G6B5
                If ModeTemp.Width = 800 And ModeTemp.Height = 600 Then
                    If blnLow16 = False Then
                        frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 16 bit"
                        blnLow16 = True
                    End If
                End If
                If ModeTemp.Width = 1024 And ModeTemp.Height = 768 Then
                    If blnHi16 = False Then
                        frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 16 bit"
                        blnHi16 = True
                    End If
                End If
        
        End Select
        
           
        'DirectX for VB code
        
        'First we parse the modes into two catergories - 16bit and 32bit
        'If ModeTemp.Format = D3DFMT_R8G8B8 Or ModeTemp.Format = D3DFMT_X8R8G8B8 Or ModeTemp.Format = D3DFMT_A8R8G8B8 Then
        '    'Check that the device is acceptable and valid...
        '    If D3D.CheckDeviceType(frmGraphicOptions.cmbAdapters.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
                'then add it to the displayed list
         '       frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 32 bit" & "    [FMT: " & ModeTemp.Format & "]"
         '   End If
        'Else
        '    If D3D.CheckDeviceType(frmGraphicOptions.cmbAdapters.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
        '        frmGraphicOptions.cmbRes.AddItem ModeTemp.Width & "x" & ModeTemp.Height & " 16 bit" & "    [FMT: " & ModeTemp.Format & "]"
        '    End If
        'End If
    
    Next intForLoop

    'frmGraphicOptions.cmbRes.ListIndex = frmGraphicOptions.cmbRes.ListCount - 1 what the Fuck is this?
End Sub
Public Sub GraphicOptionsSave()
    'New March 22, 2003
    On Error GoTo SaveErrHandler
    'End New
    
    If UCase(Left(frmGraphicOptions.cmbDevice.text, 3)) = "SOF" Then
        strDisplayDevice = "D3DDEVTYPE_REF"
    Else
        strDisplayDevice = "D3DDEVTYPE_HAL"
    End If
    If UCase(Mid(frmGraphicOptions.cmbRes.text, 5, 1)) = "X" Then
        intDisplayWidth = Left(frmGraphicOptions.cmbRes.text, 4)
    Else
        intDisplayWidth = Left(frmGraphicOptions.cmbRes.text, 3)
    End If
    If UCase(Mid(frmGraphicOptions.cmbRes.text, 5, 1)) = "X" Then
        intDisplayHight = Mid(frmGraphicOptions.cmbRes.text, 6, 3)
    Else
        intDisplayHight = Mid(frmGraphicOptions.cmbRes.text, 5, 3)
    End If
    If Right(frmGraphicOptions.cmbRes.text, 6) = "32 bit" Then
        strDisplayColour = "D3DFMT_X8R8G8B8"
    Else
        strDisplayColour = "D3DFMT_R5G6B5"
    End If
   
    
    intDisplayFTP = frmGraphicOptions.chkFrameRate
    
    intVsyncOn = frmGraphicOptions.chkVsyncOn
    
    'New March 22, 2003
    intModescreen = frmGraphicOptions.chkModescreen
    
    intLowResTextures = frmGraphicOptions.chkLowResTexture
    'End New
        
    Open App.Path & "/Settings.dat" For Output As #1
        Write #1, intVsyncOn        'Vsync 1 is ON 0 is OFF
        Write #1, intDisplayFTP     'FPS Conter 1 is ON 0 is OFF Note: I sould rename it intDisplayFPS
        Write #1, intDisplayWidth   'Displays Width     Note: I sould have combind Display Width and Hight in to on Var
        Write #1, intDisplayHight   'Displays Hight
        Write #1, strDisplayColour  'Colour type 32bit or 16bit
        Write #1, strDisplayDevice  'Name of the Device
        'New March 22, 2003
        Write #1, intModescreen
        Write #1, intLowResTextures
        'End New
    Close #1
      
   ' Binary Verson of the file writeing code above
   ' Open App.Path & "Settings.dat" For Binary Access Write As #1
   '     Put #1, , blnVsyncOn
   '     Put #1, , blnDisplayFTP
   '     Put #1, , intDisplayWidth
   '     Put #1, , intDisplayHight
   '     Put #1, , strDisplayColour
   '     Put #1, , strDisplayDevice
   ' Close #1
'New March 22, 2003
   Exit Sub
SaveErrHandler:
'End New
    
End Sub
Public Sub GraphicOptionsOpen()
    'New March 22, 2003
    On Error GoTo OpenErrHandler
    'End New
    Open App.Path & "/Settings.dat" For Input As #1
        Input #1, intVsyncOn
        Input #1, intDisplayFTP
        Input #1, intDisplayWidth
        Input #1, intDisplayHight
        Input #1, strDisplayColour
        Input #1, strDisplayDevice
        'New March 22, 2003
        Input #1, intModescreen
        Input #1, intLowResTextures
        'End New
    Close #1
    
    frmMenu.cmdSinglePlayer.Visible = True
    
    ' Binary Verson of the file Reading code above
    'Open App.Path & "Settings.dat" For Binary Access Read As #1
    '    Get #1, , blnVsyncOn
    '    Get #1, , blnDisplayFTP
    '    Get #1, , intDisplayWidth
    '    Get #1, , intDisplayHight
    '    Get #1, , strDisplayColour
    '    Get #1, , strDisplayDevice
    'Close #1
'New March 22, 2003
    Exit Sub
OpenErrHandler:
    frmMenu.cmdSinglePlayer.Visible = False
'End New
End Sub
Public Sub GraphicOptionsSettings()
    
    'New March 22, 2003
    frmGraphicOptions.chkModescreen = intModescreen
    
    frmGraphicOptions.chkLowResTexture = intLowResTextures
    'End New
    
    frmGraphicOptions.chkFrameRate = intDisplayFTP
    
    frmGraphicOptions.chkVsyncOn = intVsyncOn
    
    Select Case (strDisplayDevice)
        Case Is = "D3DDEVTYPE_REF"
            frmGraphicOptions.cmbDevice.ListIndex = 0
        Case Is = "D3DDEVTYPE_HAL"
            frmGraphicOptions.cmbDevice.ListIndex = 1
        Case Else
    End Select
    Select Case (intDisplayWidth)
        Case Is = 800
            Select Case (strDisplayColour)
                Case Is = "D3DFMT_X8R8G8B8"
                    frmGraphicOptions.cmbRes.ListIndex = 1
                Case Is = "D3DFMT_R5G6B5"
                    frmGraphicOptions.cmbRes.ListIndex = 0
                Case Else
            End Select
        Case Is = 1024
            Select Case (strDisplayColour)
                Case Is = "D3DFMT_X8R8G8B8"
                    frmGraphicOptions.cmbRes.ListIndex = 3
                Case Is = "D3DFMT_R5G6B5"
                    frmGraphicOptions.cmbRes.ListIndex = 2
                Case Else
            End Select
        Case Else
    End Select
    
    
    
    
    
End Sub
