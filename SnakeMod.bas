Attribute VB_Name = "sNAKEmOD"
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'direction constants
Public Const NORTH = 1
Public Const EAST = 2
Public Const SOUTH = 3
Public Const WEST = 4
Public Const RANDOM = 5

'used for looping + stuff
Public i As Integer
Public i2 As Integer
Public i3 As Byte
Public x As Integer
Public y As Integer

'used for move function returns of snake
Public Const OK = 0
Public Const GREEN = 1
Public Const RED = 2
Public Const CYAN = 3
Public Const YELLOW = 4
Public Const WALL = 5
Public Const PICKUP = 6

Public Const PBWIDTH = 600
Public Const PBHEIGHT = 450

Public InitLength As Integer
Public GrowthRate As Integer

Public LookDist As Byte
Public TurnTime As Byte

Public PlayerNo As Byte

Public EscapeNow As Boolean

