VERSION 5.00
Begin VB.Form OForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Snakes - by Simon Price - visit www.VBgames.co.uk"
   ClientHeight    =   4416
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5364
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4416
   ScaleWidth      =   5364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO!!!"
      Height          =   492
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Click to play Super Snakes"
      Top             =   3840
      Width           =   1092
   End
   Begin VB.Frame Frame3 
      Caption         =   "AI Settings"
      Height          =   1452
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   2412
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   372
         Left            =   720
         TabIndex        =   12
         ToolTipText     =   "Click here to adjust the AI settings"
         Top             =   960
         Width           =   972
      End
      Begin VB.Label TwitchL 
         Caption         =   "Amount Of Twitching = 150"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label ThinkL 
         Caption         =   "Amount Of Thinking = 50"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2052
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Snake Length Settings"
      Height          =   1932
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2652
      Begin VB.HScrollBar GrowthS 
         Height          =   372
         Left            =   120
         Max             =   250
         Min             =   10
         TabIndex        =   7
         Top             =   1080
         Value           =   50
         Width           =   2412
      End
      Begin VB.HScrollBar InitLengthS 
         Height          =   372
         Left            =   120
         Max             =   500
         Min             =   10
         TabIndex        =   5
         Top             =   240
         Value           =   200
         Width           =   2412
      End
      Begin VB.Label GrowthL 
         Caption         =   "Growth Length = 50"
         Height          =   252
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "How fast the snakes grow/shrink"
         Top             =   1560
         Width           =   2412
      End
      Begin VB.Label InitLengthL 
         Caption         =   "Initial Length = 200"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "The length of each snake at the start"
         Top             =   720
         Width           =   2412
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select No. of Players"
      Height          =   2172
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2652
      Begin VB.OptionButton P 
         Caption         =   "green = human 1, red =human 2, other 2 are computer controlled"
         Height          =   612
         Index           =   2
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Play 2 player"
         Top             =   1320
         Width           =   2292
      End
      Begin VB.OptionButton P 
         Caption         =   "green = human, other 3 computer"
         Height          =   372
         Index           =   1
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Play 1 player"
         Top             =   840
         Width           =   2292
      End
      Begin VB.OptionButton P 
         Caption         =   "4 computer snakes"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Watch the computer play on it's own"
         Top             =   360
         Value           =   -1  'True
         Width           =   2292
      End
   End
   Begin VB.Label InfoL 
      BorderStyle     =   1  'Fixed Single
      Height          =   1932
      Left            =   2880
      TabIndex        =   13
      Top             =   1680
      Width           =   2412
   End
End
Attribute VB_Name = "OForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
Hide
AIForm.Visible = True
Unload Me
End Sub

Private Sub cmdGO_Click()
InitLength = InitLengthS
GrowthRate = GrowthS
Hide
Unload GForm
GForm.Visible = True
End Sub

Private Sub Form_Load()
WriteInfo
End Sub

Private Sub WriteInfo()
InfoL = "Welcome to Super Snakes, a small game I made in just one day. It is an experiment in AI, the challenge was to create snakes that would try to avoid each other. When a snake crashes, it gets shorter and the snake it crashed into gets longer. The last snake crawling wins."
End Sub

Private Sub GrowthS_Change()
GrowthL = "Growth = " & GrowthS
End Sub

Private Sub InitLengthS_Change()
InitLengthL = "Initial Length = " & InitLengthS
End Sub

Private Sub P_Click(Index As Integer)
PlayerNo = Index
End Sub
