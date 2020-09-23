VERSION 5.00
Begin VB.UserControl MinMaxCloseButtons 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   144
   ToolboxBitmap   =   "MinMaxCloseButtons.ctx":0000
   Begin VB.Timer MouseOverCheckTimer 
      Interval        =   1
      Left            =   255
      Top             =   750
   End
   Begin VB.Image MaximizeImage 
      Height          =   240
      Index           =   3
      Left            =   1515
      Picture         =   "MinMaxCloseButtons.ctx":0312
      Top             =   675
      Width           =   240
   End
   Begin VB.Image MaximizeImage 
      Height          =   240
      Index           =   2
      Left            =   1275
      Picture         =   "MinMaxCloseButtons.ctx":0654
      Top             =   675
      Width           =   240
   End
   Begin VB.Image MaximizeImage 
      Height          =   240
      Index           =   1
      Left            =   1035
      Picture         =   "MinMaxCloseButtons.ctx":0996
      Top             =   675
      Width           =   240
   End
   Begin VB.Image MaximizeImage 
      Height          =   240
      Index           =   0
      Left            =   795
      Picture         =   "MinMaxCloseButtons.ctx":0CD8
      Top             =   675
      Width           =   240
   End
   Begin VB.Image ActiveButtonImage 
      Height          =   240
      Left            =   315
      Top             =   450
      Width           =   240
   End
   Begin VB.Image CloseImage 
      Height          =   240
      Index           =   3
      Left            =   1515
      Picture         =   "MinMaxCloseButtons.ctx":101A
      Top             =   915
      Width           =   240
   End
   Begin VB.Image CloseImage 
      Height          =   240
      Index           =   2
      Left            =   1275
      Picture         =   "MinMaxCloseButtons.ctx":135C
      Top             =   915
      Width           =   240
   End
   Begin VB.Image CloseImage 
      Height          =   240
      Index           =   1
      Left            =   1035
      Picture         =   "MinMaxCloseButtons.ctx":169E
      Top             =   915
      Width           =   240
   End
   Begin VB.Image CloseImage 
      Height          =   240
      Index           =   0
      Left            =   795
      Picture         =   "MinMaxCloseButtons.ctx":19E0
      Top             =   915
      Width           =   240
   End
   Begin VB.Image MinimizeImage 
      Height          =   240
      Index           =   3
      Left            =   1515
      Picture         =   "MinMaxCloseButtons.ctx":1D22
      Top             =   435
      Width           =   240
   End
   Begin VB.Image MinimizeImage 
      Height          =   240
      Index           =   2
      Left            =   1275
      Picture         =   "MinMaxCloseButtons.ctx":2064
      Top             =   435
      Width           =   240
   End
   Begin VB.Image MinimizeImage 
      Height          =   240
      Index           =   1
      Left            =   1035
      Picture         =   "MinMaxCloseButtons.ctx":23A6
      Top             =   435
      Width           =   240
   End
   Begin VB.Image MinimizeImage 
      Height          =   240
      Index           =   0
      Left            =   795
      Picture         =   "MinMaxCloseButtons.ctx":26E8
      Top             =   435
      Width           =   240
   End
End
Attribute VB_Name = "MinMaxCloseButtons"
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
Private ucType As ucTypes
Private ucState As ucStates
Private isMouseDown As Boolean
Private MouseDownKeyNumber As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Public Enum ucTypes
  MINIMIZE_BUTTON = 1
  MAXIMIZE_BUTTON = 2
  CLOSE_BUTTON = 3
End Enum

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
  Select Case ucType
    Case MINIMIZE_BUTTON
      ActiveButtonImage.Picture = MinimizeImage(ucState).Picture
      ActiveButtonImage.ToolTipText = "Minimize"
    Case MAXIMIZE_BUTTON
      ActiveButtonImage.Picture = MaximizeImage(ucState).Picture
      ActiveButtonImage.ToolTipText = "Maximize"
    Case CLOSE_BUTTON
      ActiveButtonImage.Picture = CloseImage(ucState).Picture
      ActiveButtonImage.ToolTipText = "Close"
  End Select
End Sub

Public Property Let Enable(Val As Boolean)
  ucEnable = Val
End Property

Public Property Get Enable() As Boolean
  Enable = ucEnable
End Property

Public Property Let ButtonType(Val As ucTypes)
  ucType = Val
End Property

Public Property Get ButtonType() As ucTypes
  ButtonType = ucType
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
  ucType = CLOSE_BUTTON
  ucState = UNTOUCHED
  InitButton
End Sub

Private Sub UserControl_Resize()
  InitButton
End Sub
