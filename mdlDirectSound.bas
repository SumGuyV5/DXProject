Attribute VB_Name = "mdlDirectSound"
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

'DirectSound8: Looks after all of the sound playback interfaces
Public DS As DirectSound8
'DirectSoundSecondaryBuffer8: Stores the actual audio data for playback
Public DSBuffer As DirectSoundSecondaryBuffer8
'DirectSoundEnum8: Allows us to get information on available hardware/software devices.
Public DSEnum As DirectSoundEnum8

Public DSBDesc As DSBUFFERDESC

Public lngFrequency As Long



