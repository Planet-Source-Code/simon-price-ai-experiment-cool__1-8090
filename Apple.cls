VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Apple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim xm As Integer
Dim ym As Integer
Const SIZE = 5
Const SIZETAKE1 = SIZE - 1

Function Move() As Byte

'clear away last apple pos
GForm.PB.Circle (x, y), SIZETAKE1, vbBlack
'ExtFloodFill GForm.PB.hdc, x, y, vbBlack, 0

'check for collisions
For i = -SIZE To SIZE

Select Case GetPixel(GForm.PB.hdc, x + i, y - SIZE)
   Case vbRed
   ym = -ym
   Move = WALLBOUNCE
   Case vbWhite
   ym = -ym
   Move = SNAKEBOUNCE
End Select

Select Case GetPixel(GForm.PB.hdc, x + i, y + SIZE)
   Case Is = vbRed
   ym = -ym
   Move = WALLBOUNCE
   Case vbWhite
   ym = -ym
   Move = SNAKEBOUNCE
End Select

Select Case GetPixel(GForm.PB.hdc, x - SIZE, y + i)
   Case vbRed
   xm = -xm
   Move = WALLBOUNCE
   Case vbWhite
   xm = -xm
   Move = SNAKEBOUNCE
End Select

Select Case GetPixel(GForm.PB.hdc, x + SIZE, y + i)
   Case vbRed
   xm = -xm
   Move = WALLBOUNCE
   Case vbWhite
   xm = -xm
   Move = SNAKEBOUNCE
End Select

Next

'move apple
x = x + xm
y = y + ym

'draw in new apple pos
GForm.PB.Circle (x, y), SIZETAKE1, vbGreen
    
End Function

Sub Create()
On Error Resume Next
'creates a new apple

'clear away last apple pos
GForm.PB.Line (x - 5, y - 5)-(x + 5, y + 5), vbBlack, BF

x = Int(Rnd * 400) + 50
y = Int(Rnd * 300) + 25 'position

xm = Int(Rnd * 2) - 1
ym = Int(Rnd * 2) - 1 'velocity
End Sub

