VERSION 5.00
Begin VB.Form frmMenu 
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
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About Space Troopers"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
   End
   Begin VB.CommandButton cmdMultiplayer 
      Caption         =   "&Multiplayer"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   4455
   End
   Begin VB.CommandButton cmdSinglePlayer 
      Caption         =   "&Single Player"
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
   Begin VB.Label lblNoGraphicOptions 
      Alignment       =   2  'Center
      Caption         =   "Please Setup Your Graphics Options"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
   End
End
Attribute VB_Name = "frmMenu"
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

Private Sub cmdAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal
        
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOptions_Click()
    Unload frmMenu
    Load frmOptions
    frmOptions.Show vbModal
    
    
End Sub

Private Sub cmdSinglePlayer_Click()
    Unload frmMenu
    Load frmSinglePlayerGame
    
    
    
End Sub

Private Sub Form_Load()
    Call GraphicOptionsOpen
    Me.Show vbModal
    
End Sub

Private Sub lblNoGraphicOptions_Click()
    
    'New Lable March 22, 2003
    
End Sub
