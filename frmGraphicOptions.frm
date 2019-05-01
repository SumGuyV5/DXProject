VERSION 5.00
Begin VB.Form frmGraphicOptions 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5340.839
   ScaleMode       =   0  'User
   ScaleWidth      =   4620.76
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLowResTexture 
      Caption         =   "Low Res Textures"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox chkModescreen 
      Caption         =   "Windows Screen"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CheckBox chkVsyncOn 
      Caption         =   "V-sync On"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "V-sync on will force the frame to be drawn when the monitor refreshes."
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox chkFrameRate 
      Caption         =   "Display Frame Rate"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Displays How Many Frames Per Second are being rendered by the Game.."
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Options"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   4455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back to Options Menu"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   4455
   End
   Begin VB.ComboBox cmbRes 
      Height          =   315
      ItemData        =   "frmGraphicOptions.frx":0000
      Left            =   120
      List            =   "frmGraphicOptions.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   4455
   End
   Begin VB.ComboBox cmbDevice 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   4455
   End
   Begin VB.ComboBox cmbAdapters 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resolutions available:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1875
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rendering Devices Available:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label lblmisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hardware Adapters Available:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2565
   End
End
Attribute VB_Name = "frmGraphicOptions"
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


Private Sub cmbAdapters_Click()
   ' If UCase(Left(cmbDevice.Text, 3)) = "Sof" Then
   '     Call DisplyModes(2)
   ' Else
   '     Call DisplyModes(1)
   ' End If
End Sub


Private Sub cmbDevice_Click()
    If UCase(Left(cmbDevice.text, 3)) = "SOF" Then
        Call DisplyModes(2)
    Else
        Call DisplyModes(1)
    End If

End Sub

Private Sub cmdBack_Click()
    Unload frmGraphicOptions
    Load frmOptions
    frmOptions.Show vbModal
    
    
End Sub

Private Sub cmdSave_Click()
    Call GraphicOptionsSave
End Sub

Private Sub Form_Load()
        
    Set Dx = New DirectX8
    Set D3D = Dx.Direct3DCreate
    
    Call DisplayAdapters 'Information on physical hardware cards available
        cmbAdapters.ListIndex = 0
    
    Call DisplayDevices 'what rendering devices they support software or hardware
    
    
    If UCase(Left(cmbDevice.text, 3)) = "Sof" Then
        Call DisplyModes(2)
    Else
        Call DisplyModes(1)
    End If
        'cmbRes.ListIndex = cmbRes.ListCount - 1
    Call GraphicOptionsSettings
       
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set Dx = Nothing
    Set D3D = Nothing
    
End Sub
