VERSION 5.00
Begin VB.Form frmSinglePlayerGame 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSinglePlayerGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Click()
    'bRunning = False
    'Unload Me 'New line March 22, 2003
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer) ' All new Sub March 22, 2003
    Select Case (KeyAscii)
        Case (81)           '"Q" to Quit
            bRunning = False
            Unload Me
        Case (113)          '"q" to Quit
            bRunning = False
            Unload Me
        Case (56)           '"8" to move texture Number 2 Up
            sngUpDown = sngUpDown - 10
        Case (50)           '"2" to move texture Number 2 Down
            sngUpDown = sngUpDown + 10
        Case (52)           '"4" to move texture Number 2 Left
            sngLeftRight = sngLeftRight - 10
        Case (54)           '"6" to move texture Number 2 right
            sngLeftRight = sngLeftRight + 10
        Case (55)           '"7" to move texture Number 2 Left and Up
            sngLeftRight = sngLeftRight - 10
            sngUpDown = sngUpDown - 10
        Case (57)           '"9" to move texture Number 2 Right and Up
            sngLeftRight = sngLeftRight + 10
            sngUpDown = sngUpDown - 10
        Case (49)           '"1" to move texture Number 2 Left and Down
            sngLeftRight = sngLeftRight - 10
            sngUpDown = sngUpDown + 10
        Case (51)           '"3" to move texture Number 2 Right and Down
            sngLeftRight = sngLeftRight + 10
            sngUpDown = sngUpDown + 10
        Case (45)           '"-" to move texture Number 2 Back
            sngBackFront = sngBackFront - 0.1
        Case (43)           '"+" to move texture Number 2 Forword
            sngBackFront = sngBackFront + 0.1
        Case (53) 'new March 26, 2003
            DSBuffer.Play DSBPLAY_DEFAULT
        Case (42) 'stop mp3 new code march 27, 2003
            DSControl.Pause
        Case (47) 'Starts Mp3 new code march 27, 2003
            DSControl.Run
        
    End Select
    
End Sub

Private Sub Form_Load()
        
    Me.Show '//Make sure our window is visible
    
    bRunning = Initialise()
    Debug.Print "Device Creation Return Code : ", bRunning 'So you can see what happens...
    
    Call text
    
    'New march 27, 2003
     Call TerminateEngine
        
    '//2. Setup a filter graph for the file
        Set DSControl = New FilgraphManager
        Call DSControl.RenderFile(App.Path + "\sample.mp3")
    
    '//3. Setup the basic audio object
        Set DSAudio = DSControl
        DSAudio.Volume = 0
        DSAudio.Balance = 0
    
    '//4. Setup the media event and position objects
        Set DSEvent = DSControl
        Set DSPosition = DSControl
        If ObjPtr(DSPosition) Then DSPosition.Rate = 1#
        DSPosition.CurrentPosition = 0
       
    'end new March 27, 2003
    
    Do While bRunning = True
        Render '//Update the frame...
        DoEvents '//Allow windows time to think; otherwise you'll get into a really tight (and bad) loop...
        If intDisplayFTP = 1 Then
            If GetTickCount() - FPS_LastCheck >= 100 Then
                FPS_Current = FPS_Count * 10 '//We check every 1/10 of a second, so we scale it up....
                FPS_Count = 0 'reset the counter
                FPS_LastCheck = GetTickCount()
            End If
            FPS_Count = FPS_Count + 1
        End If
    Loop '//Begin the next frame...
    
    '//If we've gotten to this point the loop must have been terminated
    '   So we need to clean up after ourselves. This isn't essential, but it'
    '   good coding practise.
    
    On Error Resume Next 'If the objects were never created;
    '                               (the initialisation failed) we might get an
    '                               error when freeing them... which we need to
    '                               handle, but as we're closing anyway...
    Set D3DDevice = Nothing
    Set D3D = Nothing
    'Set Dx = Nothing
    
    'New code March 26, 2003
    Set DSBuffer = Nothing
    Set DSEnum = Nothing
    Set DS = Nothing
    Set Dx = Nothing
    'end new code
    Debug.Print "All Objects Destroyed"
    
    '//Final termination:
    Load frmMenu
    Unload Me
    'End
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'New March 26, 2003
    'Could have done this better but me soo tiread
    'If Button = 1 And X >= 0 And X <= 600 And Y >= 400 And Y <= 800 And Visable = True Then
    '    Visable = False
    '    Exit Sub
    'End If
    'If Button = 1 And X >= 0 And X <= 600 And Y >= 400 And Y <= 800 And Visable = False Then
    '    Visable = True
    'End If
    'End New
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bRunning = False
    DSControl.Pause
End Sub
