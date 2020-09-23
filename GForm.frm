VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Si's Super Snakes - Visit www.VBgames.co.uk for more games!!!"
   ClientHeight    =   5664
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   7464
   FillColor       =   &H0000FF00&
   FillStyle       =   6  'Cross
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PB 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   5400
      Left            =   120
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin VB.Timer FrameCountT 
         Interval        =   1000
         Left            =   240
         Top             =   240
      End
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Snake(1 To 4) As New Snake
Dim FrameCount As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyEscape
    EscapeNow = True
    GameOver (0)
  Case vbKeyUp
    Snake(1).Turn (NORTH)
  Case vbKeyRight
    Snake(1).Turn (EAST)
  Case vbKeyDown
    Snake(1).Turn (SOUTH)
  Case vbKeyLeft
    Snake(1).Turn (WEST)
  Case vbKeyW
    Snake(2).Turn (NORTH)
  Case vbKeyS
    Snake(2).Turn (EAST)
  Case vbKeyZ
    Snake(2).Turn (SOUTH)
  Case vbKeyA
    Snake(2).Turn (WEST)
End Select
End Sub

Private Sub Form_Load()
Show
CreateSnakes
MainLoop
End Sub

Public Sub MainLoop()
PlayerNo = PlayerNo + 1

Do
DoEvents
FrameCount = FrameCount + 1

For i3 = PlayerNo To 4
  Snake(i3).Think
Next

For i3 = 1 To 4
    Select Case Snake(i3).Move
      Case GREEN
        Snake(i3).Kill
        ReplaceSnake i3
        If i3 <> GREEN Then Snake(GREEN).Grow (GrowthRate)
      Case RED
        Snake(i3).Kill
        ReplaceSnake i3
        If i3 <> RED Then Snake(RED).Grow (GrowthRate)
      Case CYAN
        Snake(i3).Kill
        ReplaceSnake i3
        If i3 <> CYAN Then Snake(CYAN).Grow (GrowthRate)
      Case YELLOW
        Snake(i3).Kill
        ReplaceSnake i3
        If i3 <> YELLOW Then Snake(YELLOW).Grow (GrowthRate)
      Case WALL
        Snake(i3).Kill
        ReplaceSnake i3
    End Select

Next
Loop Until EscapeNow = True
Set GForm = Nothing
Hide
Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
EscapeNow = True
End Sub

Private Sub FrameCountT_Timer()
If FrameCount = 0 Then
  Visible = False
  Set GForm = Nothing
  End
End If
Caption = "Si's Super Snakes - Visit www.VBgames.co.uk for more games!!! - FPS = " & FrameCount
FrameCount = 0
i2 = 0
For i = 1 To 4
  If Snake(i).IsDead Then i2 = i2 + 1
Next
If i2 = 3 Then
  For i = 1 To 4
  i3 = i
    If Snake(i).IsDead = False Then
      GameOver i3
    End If
  Next
End If
End Sub

Public Sub GameOver(WhoWon As Byte)
Select Case WhoWon
  Case 0
    MsgBox "Game Quitted", vbExclamation, "Game Over!"
  Case GREEN
    MsgBox "The GREEN snake has won!!!", vbInformation, "Game Over!"
  Case RED
    MsgBox "The RED snake has won!!!", vbInformation, "Game Over!"
  Case CYAN
    MsgBox "The CYAN snake has won!!!", vbInformation, "Game Over!"
  Case YELLOW
    MsgBox "The YELLOW snake has won!!!", vbInformation, "Game Over!"
End Select

   MsgBox "Thankyou for playing, to download loads more games visit www.VBgames.co.uk"
EscapeNow = True
Set GForm = Nothing
End
End Sub

Private Sub PB_Paint()
PB.FillColor = vbBlack
PB.Line (0, 0)-(PBWIDTH - 1, PBHEIGHT - 1), vbWhite, B
PB.FillColor = vbGreen
End Sub

Private Sub CreateSnakes()
For i3 = 1 To 4
  ReplaceSnake i3
Next
End Sub

Private Sub ReplaceSnake(Color As Byte)

Select Case Color
  Case GREEN
    Snake(GREEN).Create vbGreen, NORTH, PBWIDTH \ 2, PBHEIGHT - 5
  Case RED
    Snake(RED).Create vbRed, SOUTH, PBWIDTH \ 2, 5
  Case CYAN
    Snake(CYAN).Create vbCyan, EAST, 5, PBHEIGHT \ 2
  Case YELLOW
    Snake(YELLOW).Create vbYellow, WEST, PBWIDTH - 5, PBHEIGHT \ 2
End Select

End Sub
