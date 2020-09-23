VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type tSnakePart
  x As Integer
  y As Integer
End Type

Dim Part() As tSnakePart
Dim Dir As Byte
Dim Score As Long
Dim Color As Long
Dim Length As Integer
Dim Dead As Boolean

Function Move() As Byte
On Error Resume Next

If Dead Then Exit Function

'shunt along the snake a bit
For i = UBound(Part) To 1 Step -1
  Part(i) = Part(i - 1)
Next

'now move the head
Select Case Dir
  Case NORTH
    x = 0: y = -1
  Case EAST
    x = 1: y = 0
  Case SOUTH
    x = 0: y = 1
  Case WEST
    x = -1: y = 0
End Select
Part(0).x = Part(0).x + x
Part(0).y = Part(0).y + y

'check for crashes
Select Case GetPixel(GForm.PB.hdc, Part(0).x, Part(0).y)
   Case vbGreen
     Move = GREEN
   Case vbRed
     Move = RED
   Case vbCyan
     Move = CYAN
   Case vbYellow
     Move = YELLOW
   Case vbWhite
     Move = WALL
   Case vbBlack
     Move = OK
End Select

'clear away the tail
SetPixel GForm.PB.hdc, Part(UBound(Part)).x, Part(UBound(Part)).y, vbBlack

'draw in the head
SetPixel GForm.PB.hdc, Part(0).x, Part(0).y, Color

End Function

Sub Create(WhatColor As Long, Facing As Byte, x As Integer, y As Integer)
On Error Resume Next
'create a new snake
ReDim Part(0 To Length)
'assign color
Color = WhatColor
'position it
Dir = Facing
For i = 0 To Length
  Part(i).x = x
  Part(i).y = y
Next
Alive = True
End Sub

Sub Grow(HowMuch As Byte)
On Error Resume Next
If Length - HowMuch <= 0 Then Die
Length = Length + HowMuch
ReDim Preserve Part(0 To UBound(Part) + HowMuch)
End Sub

Sub Kill()
For i = 1 To UBound(Part)
  SetPixel GForm.PB.hdc, Part(i).x, Part(i).y, vbBlack
Next
Length = Length - GrowthRate
If Length <= 0 Then Die
End Sub

Function Turn(WhichWay As Byte) As Byte
'turns the snake, checking that the move is legal
StartAgain:
Select Case WhichWay
  Case NORTH
    If Dir = SOUTH Then
      Turn = BADTURN
    Else
      Dir = NORTH
    End If
  Case EAST
    If Dir = WEST Then
      Turn = BADTURN
    Else
      Dir = EAST
    End If
  Case SOUTH
    If Dir = NORTH Then
      Turn = BADTURN
    Else
      Dir = SOUTH
    End If
  Case WEST
    If Dir = EAST Then
      Turn = BADTURN
    Else
      Dir = WEST
    End If
  Case RANDOM
    WhichWay = Int(Rnd * 4) + 1
    GoTo StartAgain
End Select

End Function

Private Sub Class_Initialize()
Length = InitLength
End Sub

Sub Die()
Dead = True
End Sub

Sub Think()
On Error Resume Next

If Int(Rnd * TurnTime) = 0 Then Turn (RANDOM)

Select Case Dir

  Case NORTH
    For i = 1 To LookDist
      If GetPixel(GForm.PB.hdc, Part(0).x, Part(0).y - i) > vbBlack Then
        If Int(Rnd * 2) = 0 Then
          Turn (EAST)
        Else
          Turn (WEST)
        End If
      End If
    Next

  Case EAST
    For i = 1 To LookDist
      If GetPixel(GForm.PB.hdc, Part(0).x + i, Part(0).y) > vbBlack Then
        If Int(Rnd * 2) = 0 Then
          Turn (NORTH)
        Else
          Turn (SOUTH)
        End If
      End If
    Next

  Case SOUTH
    For i = 1 To LookDist
      If GetPixel(GForm.PB.hdc, Part(0).x, Part(0).y + i) > vbBlack Then
        If Int(Rnd * 2) = 0 Then
          Turn (EAST)
        Else
          Turn (WEST)
        End If
      End If
    Next

  Case WEST
    For i = 1 To LookDist
      If GetPixel(GForm.PB.hdc, Part(0).x - i, Part(0).y) > vbBlack Then
        If Int(Rnd * 2) = 0 Then
          Turn (NORTH)
        Else
          Turn (SOUTH)
        End If
      End If
    Next

End Select

End Sub

Function IsDead() As Boolean
If Dead Then IsDead = True
End Function
