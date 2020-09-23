VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Just Like MDI (Really Cool Indeed !)"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStyle 
      BackColor       =   &H8000000C&
      Caption         =   "Style"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox picWnd 
      Height          =   3735
      Left            =   120
      Picture         =   "frmHookWndLong.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_THICKFRAME = &H40000
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Sub UpDateStyle(ctl As Control, ByVal lngNewStyles As Long, Optional blnExtended As Boolean, Optional blnToggle As Boolean)
Const GWL_STYLE = -16
Const GWL_EXSTYLE = -20
Dim lngGWLType As Long
Dim lngOrigStyle As Long
lngGWLType = IIf(blnExtended, GWL_EXSTYLE, GWL_STYLE)
lngOrigStyle = GetWindowLong(ctl.hWnd, lngGWLType)
If blnToggle Then
lngNewStyles = lngOrigStyle Xor lngNewStyles
Else
lngNewStyles = lngNewStyles Or lngOrigStyle
End If
SetWindowLong ctl.hWnd, lngGWLType, lngNewStyles
If Err.LastDllError Then
Err.Raise 1
End If
Dim sOrigWidth As Single
With ctl
sOrigWidth = .Width
.Width = .Width + .ScaleX(10, .Container.ScaleMode, vbTwips)
.Width = sOrigWidth
End With
End Sub


Private Sub chkStyle_Click()
'UpDateStyle picWnd, WS_THICKFRAME, False, True
FlashWindow picWnd.hWnd, True 'Just to make it look active
If chkStyle Then
UpDateStyle picWnd, WS_CAPTION Or WS_BORDER Or WS_THICKFRAME
chkStyle.Caption = "Restore Style"
SetWindowText picWnd.hWnd, "My Caption Text"
Else
UpDateStyle picWnd, WS_CAPTION Or WS_BORDER Or WS_THICKFRAME, , True
chkStyle.Caption = "Update Style"
End If
End Sub

