VERSION 5.00
Begin VB.Form Main_Form 
   Caption         =   "Transparent Forms on fly"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3630
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Index           =   0
      Left            =   180
      ScaleHeight     =   1020
      ScaleWidth      =   1125
      TabIndex        =   7
      Top             =   3915
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Index           =   1
      Left            =   1665
      ScaleHeight     =   1020
      ScaleWidth      =   1125
      TabIndex        =   6
      Top             =   3945
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Index           =   2
      Left            =   3150
      ScaleHeight     =   1020
      ScaleWidth      =   1125
      TabIndex        =   5
      Top             =   3975
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Index           =   3
      Left            =   4605
      ScaleHeight     =   1020
      ScaleWidth      =   1125
      TabIndex        =   4
      Top             =   3990
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   1350
      Index           =   0
      Left            =   270
      MaskColor       =   &H000000FF&
      Picture         =   "Main Form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   1290
      Index           =   1
      Left            =   1920
      Picture         =   "Main Form.frx":12D2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   1260
      Index           =   2
      Left            =   270
      Picture         =   "Main Form.frx":2A32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1650
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   1275
      Index           =   3
      Left            =   1905
      Picture         =   "Main Form.frx":3B8C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1410
      Width           =   1650
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)

Select Case Index
    Case 0
        Transparent_form.Picture = Picture1(0).Picture
        rgnMain = ExtCreateRegion(ByVal 0&, STY_PLAYER_nBytes, STY_PLAYER_bytRegion(0))
        SetWindowRgn Transparent_form.hwnd, rgnMain, True
        Transparent_form.Caption = "STY Player"
    Case 1
        Transparent_form.Picture = Picture1(1).Picture
        rgnMain = ExtCreateRegion(ByVal 0&, OCTOPUS_nBytes, OCTOPUS_bytRegion(0))
        SetWindowRgn Transparent_form.hwnd, rgnMain, True
        Transparent_form.Caption = "Octopus"
    Case 2
        Transparent_form.Picture = Picture1(2).Picture
        rgnMain = ExtCreateRegion(ByVal 0&, MP3_REGENTE_nBytes, MP3_REGENTE_bytRegion(0))
        SetWindowRgn Transparent_form.hwnd, rgnMain, True
        Transparent_form.Caption = "MP3 Regente"
    Case 3
        Transparent_form.Picture = Picture1(3).Picture
        rgnMain = ExtCreateRegion(ByVal 0&, VIRTUAL_GUITAR_nBytes, VIRTUAL_GUITAR_bytRegion(0))
        SetWindowRgn Transparent_form.hwnd, rgnMain, True
        Transparent_form.Caption = "Virtual Guitar"
End Select

End Sub

Private Sub Form_Load()
Open App.Path & "\Sty player.byt" For Binary As 1
        Get #1, 1, STY_PLAYER_nBytes
        ReDim STY_PLAYER_bytRegion(STY_PLAYER_nBytes)
        Get #1, 8, STY_PLAYER_bytRegion
    Close
    
    Open App.Path & "\Octopus.byt" For Binary As 1
        Get #1, 1, OCTOPUS_nBytes
        ReDim OCTOPUS_bytRegion(OCTOPUS_nBytes)
        Get #1, 8, OCTOPUS_bytRegion
    Close
    
    Open App.Path & "\Mp3 Regente.byt" For Binary As 1
        Get #1, 1, MP3_REGENTE_nBytes
        ReDim MP3_REGENTE_bytRegion(MP3_REGENTE_nBytes)
        Get #1, 8, MP3_REGENTE_bytRegion
    Close
    
    Open App.Path & "\Virtual Guitar.byt" For Binary As 1
        Get #1, 1, VIRTUAL_GUITAR_nBytes
        ReDim VIRTUAL_GUITAR_bytRegion(VIRTUAL_GUITAR_nBytes)
        Get #1, 8, VIRTUAL_GUITAR_bytRegion
    Close
    
    Picture1(0).Picture = LoadPicture(App.Path & "\STY Player.jpg")
    Picture1(1).Picture = LoadPicture(App.Path & "\Octopus.jpg")
    Picture1(2).Picture = LoadPicture(App.Path & "\Mp3 Regente.jpg")
    Picture1(3).Picture = LoadPicture(App.Path & "\Virtual Guitar.jpg")
    
    Command1_Click (0)
    Transparent_form.Show
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Transparent_form
End
End Sub
