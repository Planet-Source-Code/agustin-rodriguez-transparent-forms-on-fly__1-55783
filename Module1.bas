Attribute VB_Name = "Module1"

Public STY_PLAYER_bytRegion() As Byte
Public STY_PLAYER_nBytes As Long

Public VIRTUAL_GUITAR_bytRegion() As Byte
Public VIRTUAL_GUITAR_nBytes As Long

Public OCTOPUS_bytRegion() As Byte
Public OCTOPUS_nBytes As Long

Public MP3_REGENTE_bytRegion() As Byte
Public MP3_REGENTE_nBytes As Long

Public Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
