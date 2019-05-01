Attribute VB_Name = "mdlDirectShow"
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

'//DirectShow Objects
Public DSAudio As IBasicAudio         'Basic Audio Objectt
Public DSEvent As IMediaEvent        'MediaEvent Object
Public DSControl As IMediaControl    'MediaControl Object
Public DSPosition As IMediaPosition 'MediaPosition Object

Public Function TerminateEngine() As Boolean
On Error GoTo BailOut:

    If ObjPtr(DSControl) > 0 Then
        DSControl.Stop
    End If
                
    If ObjPtr(DSAudio) Then Set DSAudio = Nothing
    If ObjPtr(DSEvent) Then Set DSEvent = Nothing
    If ObjPtr(DSControl) Then Set DSControl = Nothing
    If ObjPtr(DSPosition) Then Set DSPosition = Nothing
                
    TerminateEngine = True
    Exit Function
BailOut:
    TerminateEngine = False
    Debug.Print "ERROR: modDirectShow.TerminateEngine()"
    Debug.Print "     ", Err.Number, Err.Description
End Function
