VERSION 5.00
Begin VB.UserControl ActiveButton 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
   ToolboxBitmap   =   "ActiveButton.ctx":0000
   Begin VB.Timer MouseOverCheckTimer 
      Interval        =   1
      Left            =   600
      Top             =   960
   End
   Begin VB.Image ActiveButtonImage 
      Height          =   360
      Left            =   90
      Top             =   210
      Width           =   1800
   End
   Begin VB.Image SourceImage 
      Height          =   360
      Index           =   3
      Left            =   1950
      Picture         =   "ActiveButton.ctx":0312
      Top             =   1335
      Width           =   1800
   End
   Begin VB.Image SourceImage 
      Height          =   360
      Index           =   2
      Left            =   1950
      Picture         =   "ActiveButton.ctx":0B02
      Top             =   960
      Width           =   1800
   End
   Begin VB.Image SourceImage 
      Height          =   360
      Index           =   1
      Left            =   1950
      Picture         =   "ActiveButton.ctx":135E
      Top             =   585
      Width           =   1800
   End
   Begin VB.Image SourceImage 
      Height          =   360
      Index           =   0
      Left            =   1950
      Picture         =   "ActiveButton.ctx":1B47
      Top             =   210
      Width           =   1800
   End
End
Attribute VB_Name = "ActiveButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Event OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Click()
Private ucEnable As Boolean
Private ucState As ucStates
Private isMouseDown As Boolean
Private MouseDownKeyNumber As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Enum ucStates
  UNTOUCHED = 0
  MOUSEOVER = 1
  Pressed = 2
  DISABLED = 3
End Enum

Private Function IsMouseOver() As Boolean
  Dim Point As POINTAPI
  GetCursorPos Point
  IsMouseOver = (WindowFromPoint(Point.X, Point.Y) = UserControl.hWnd)
End Function

Private Sub InitButton()
  UserControl.Width = ActiveButtonImage.Width * Screen.TwipsPerPixelX
  UserControl.Height = ActiveButtonImage.Height * Screen.TwipsPerPixelY
  UserControl.Cls
  ActiveButtonImage.Top = 0
  ActiveButtonImage.Left = 0
End Sub

Private Sub AdjustButton()
  If ucEnable Then
    If IsMouseOver Then
      If isMouseDown Then
        If MouseDownKeyNumber = 1 Then 'Left mouse key
          ucState = Pressed
        End If
      Else
        ucState = MOUSEOVER
      End If
    Else
      ucState = UNTOUCHED
    End If
  Else
    ucState = DISABLED
  End If
  ActiveButtonImage.Picture = SourceImage(ucState).Picture
  ActiveButtonImage.ToolTipText = "Close"
End Sub

Public Property Let Enable(Val As Boolean)
  ucEnable = Val
End Property

Public Property Get Enable() As Boolean
  Enable = ucEnable
End Property

Private Sub MouseOverCheckTimer_Timer()
  AdjustButton
End Sub

Private Sub ActiveButtonImage_Click()
  If MouseDownKeyNumber = 1 Then 'Left mouse key
    RaiseEvent Click
  End If
End Sub

Private Sub ActiveButtonImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  isMouseDown = True
  MouseDownKeyNumber = Button
  If MouseDownKeyNumber = 1 Then  'Left mouse key
    RaiseEvent OnMouseDown(Button, Shift, X, Y)
  End If
End Sub

Private Sub ActiveButtonImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  isMouseDown = False
  MouseDownKeyNumber = Button
  If MouseDownKeyNumber = 1 Then  'Left mouse key
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
  End If
End Sub

Private Sub UserControl_Initialize()
  ucEnable = True
  ucState = UNTOUCHED
  InitButton
End Sub

Private Sub UserControl_Resize()
  InitButton
End Sub
