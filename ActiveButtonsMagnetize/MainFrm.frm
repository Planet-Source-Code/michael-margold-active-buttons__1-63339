VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "Move and magnetize to edge"
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "MainFrm.frx":1D82
   ScaleHeight     =   355
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   Begin Magnetize.ActiveButton ActiveButton1 
      Height          =   360
      Left            =   1050
      TabIndex        =   3
      Top             =   4350
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   635
   End
   Begin Magnetize.MinMaxCloseButtons CloseButton 
      Height          =   240
      Left            =   3540
      TabIndex        =   2
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin Magnetize.MinMaxCloseButtons MaximizeButton 
      Height          =   240
      Left            =   3270
      TabIndex        =   1
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin Magnetize.MinMaxCloseButtons MinimizeButton 
      Height          =   240
      Left            =   3000
      TabIndex        =   0
      Top             =   60
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Image imgMoving 
      Height          =   345
      Left            =   15
      Top             =   15
      Width           =   3885
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LastX As Integer  'Auxiliary variable for moving picture
Dim LastY As Integer  'Auxiliary variable for moving picture
'***********************************MainFrm Events*************************************
Private Sub Form_Load()
  'Buttons (Min/Max/Close)
  MinimizeButton.ButtonType = MINIMIZE_BUTTON
  MinimizeButton.Enable = True
  MaximizeButton.ButtonType = MAXIMIZE_BUTTON
  MaximizeButton.Enable = False
  CloseButton.ButtonType = CLOSE_BUTTON
  CloseButton.Enable = True
  SetWindowTopMost Me, True
End Sub
'Unload this form
Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
'**************************Move Window Mechanism*******************************************
Private Sub imgMoving_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  LastX = X
  LastY = Y
End Sub
Private Sub imgMoving_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim HorisontalMove As Long
  Dim VerticalMove As Long
  Dim ScreenWidthForLeftTopCourner As Long
  Dim ScreenHeightForLeftTopCourner As Long
  Dim Magnet As Long
  Dim CurrentXCoordinate As Long
  Dim CurrentYCoordinate As Long
  If Button = 1 Then
    CurrentXCoordinate = Me.Left
    CurrentYCoordinate = Me.Top
    HorisontalMove = X - LastX
    VerticalMove = Y - LastY
    'Determine max coordinates for upper left corner of this form
    ScreenWidthForLeftTopCourner = Screen.Width - Me.Width
    ScreenHeightForLeftTopCourner = Screen.Height - Me.Height
    'Magnet option
    Magnet = Screen.Height / 100
    If CurrentXCoordinate + HorisontalMove > Magnet And _
       CurrentXCoordinate + HorisontalMove < ScreenWidthForLeftTopCourner - Magnet Then
      Me.Left = CurrentXCoordinate + HorisontalMove
    Else
      If CurrentXCoordinate + HorisontalMove <= Magnet Then Me.Left = 0
      If CurrentXCoordinate + HorisontalMove >= ScreenWidthForLeftTopCourner - Magnet Then Me.Left = ScreenWidthForLeftTopCourner
    End If
    If CurrentYCoordinate + VerticalMove > Magnet And _
       CurrentYCoordinate + VerticalMove < ScreenHeightForLeftTopCourner - Magnet Then
      Me.Top = CurrentYCoordinate + VerticalMove
    Else
      If CurrentYCoordinate + VerticalMove <= Magnet Then Me.Top = 0
      If CurrentYCoordinate + VerticalMove >= ScreenHeightForLeftTopCourner - Magnet Then Me.Top = ScreenHeightForLeftTopCourner
    End If
  End If
End Sub
'**************************************Max/Min/Close Events**************************************
Private Sub MinimizeButton_Click()
  WindowState = 1
End Sub
Private Sub CloseButton_Click()
  Unload Me
End Sub
'**************************************ActiveButton Events**************************************
Private Sub ActiveButton1_OnMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Link
End Sub

