VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "- Pong - Score: 0 - Speed: 10 - "
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrLeft 
      Interval        =   20
      Left            =   4080
      Top             =   3360
   End
   Begin VB.Timer tmrUp 
      Interval        =   10
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Shape moveBall 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4080
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   495
   End
   Begin VB.Shape bgGrid 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   6  'Cross
      Height          =   6495
      Left            =   360
      Top             =   360
      Width           =   8055
   End
   Begin VB.Shape padRight 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   8400
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape padLeft 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   120
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape padTop 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3240
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape padBottom 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3240
      Top             =   6840
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Made by damenta
'This is free source code feel free to edit it as you see fit.
'Please leave the first line of code in with my name as credit if you
'sell, or give out to anyone.. Thanks.

'This is really basic code I made for a friend who wanted to learn commands
'This program covers almost all the basic commands you need to learn getting
'into VB.

'Set vars
Dim intSpeed As Integer
Dim intScore As Integer
Dim intBall As Integer
Dim intSide As Integer
Dim intUp As Integer
Private Sub Form_Load()
intUp = 0
intSide = 0
intSpeed = 10
intScore = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moves the top Bar
padTop.Left = X - padTop.Width / 2
    'Checks to make sure the bars to move off the screen
    If padTop.Left <= 360 Then padTop.Left = 360
    If padTop.Left >= 6120 Then padTop.Left = 6120
    
'Moves the bottom Bar
padBottom.Left = X - padBottom.Width / 2
    'Checks to make sure the bars to move off the screen
    If padBottom.Left <= 360 Then padBottom.Left = 360
    If padBottom.Left >= 6120 Then padBottom.Left = 6120

'Moves the left Bar
padLeft.Top = Y - padLeft.Height / 2
    'Checks to make sure the bars to move off the screen
    If padLeft.Top <= 360 Then padLeft.Top = 360
    If padLeft.Top >= 4560 Then padLeft.Top = 4560

'Moves the right Bar
padRight.Top = Y - padRight.Height / 2
    'Checks to make sure the bars to move off the screen
    If padRight.Top <= 360 Then padRight.Top = 360
    If padRight.Top >= 4560 Then padRight.Top = 4560

End Sub

Private Sub tmrLeft_Timer()


If intSide = 0 Then
    moveBall.Left = moveBall.Left + intSpeed * 10
Else
    moveBall.Left = moveBall.Left - intSpeed * 10
End If


End Sub

Private Sub tmrUp_Timer()
'By setting the 2 timers at diffrent speeds
'this gives us random up and down movements

'Moves the ball up or down
If intUp = 0 Then
    moveBall.Top = moveBall.Top + intSpeed * 10
Else
    moveBall.Top = moveBall.Top - intSpeed * 10
End If

'Checks to see if it hit padle or not
'Checks top bar
If moveBall.Top <= 360 Then
intBall = moveBall.Left - moveBall.Width / 2
Select Case intBall
'If hit
Case padTop.Left To padTop.Left + padTop.Width
    intUp = 0
    intScore = intScore + 1
    intSpeed = intSpeed + 1
    frmMain.Caption = "- Pong - Score: " & intScore & " - Speed: " & intSpeed & " -"
'If not
Case Else
    tmrUp.Enabled = False
    tmrLeft.Enabled = False
    MsgBox "You have lost! Your score was " & intScore & ", Your speed was " & intSpeed
    End
End Select
End If

'Checks bottom bar
If moveBall.Top >= 6360 Then
intBall = moveBall.Left - moveBall.Width / 2
Select Case intBall
'If hit
Case padTop.Left To padTop.Left + padTop.Width
    intUp = 1
    intScore = intScore + 1
    intSpeed = intSpeed + 1
    frmMain.Caption = "- Pong - Score: " & intScore & " - Speed: " & intSpeed & " -"
'If not
Case Else
    tmrUp.Enabled = False
    tmrLeft.Enabled = False
    MsgBox "You have lost! Your score was " & intScore & ", Your speed was " & intSpeed
    End
End Select
End If

'Checks to see if it hit padle or not
'Check left bar
If moveBall.Left <= 360 Then
intBall = moveBall.Top
Select Case intBall
'If hit
Case padLeft.Top To padLeft.Top + padLeft.Height
    intSide = 0
    intScore = intScore + 1
    intSpeed = intSpeed + 1
    frmMain.Caption = "- Pong - Score: " & intScore & " - Speed: " & intSpeed & " -"
'If not
Case Else
    tmrUp.Enabled = False
    tmrLeft.Enabled = False
    MsgBox "You have lost! Your score was " & intScore & ", Your speed was " & intSpeed
    End
End Select
End If

'Checks to see if it hit padle or not
'Check left bar
If moveBall.Left >= 7920 Then
intBall = moveBall.Top
Select Case intBall
'If hit
Case padLeft.Top To padLeft.Top + padLeft.Height
    intSide = 1
    intScore = intScore + 1
    intSpeed = intSpeed + 1
    frmMain.Caption = "- Pong - Score: " & intScore & " - Speed: " & intSpeed & " -"
'If not
Case Else
    tmrUp.Enabled = False
    tmrLeft.Enabled = False
    MsgBox "You have lost! Your score was " & intScore & ", Your speed was " & intSpeed
    End
End Select
End If

End Sub
