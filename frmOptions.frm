VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back to Main Menu"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton cmdMultiplayerOptions 
      Caption         =   "&Multiplayer Options"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton cmdSinglePlayerOptions 
      Caption         =   "&Single Player Options"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton cmdAudioOptions 
      Caption         =   "&Audio Options"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdGraphicOptions 
      Caption         =   "&Graphic Options"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   4395
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmOptions"
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

Private Sub cmdBack_Click()
    Unload frmOptions
    Load frmMenu
    frmMenu.Show vbModal
    
    
End Sub

Private Sub cmdGraphicOptions_Click()
    Unload frmOptions
    Load frmGraphicOptions
    frmGraphicOptions.Show vbModal
    
    
End Sub

