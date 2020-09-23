VERSION 5.00
Begin VB.Form AIForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Snakes - AI settings - www.VBgames.co.uk"
   ClientHeight    =   5052
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4788
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   4788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   2880
      TabIndex        =   11
      Text            =   "Custom"
      ToolTipText     =   "The name of the AI setting"
      Top             =   3840
      Width           =   1692
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Current Settings"
      Height          =   492
      Left            =   3000
      TabIndex        =   10
      ToolTipText     =   "Saves the current AI settings under the name below"
      Top             =   3240
      Width           =   1452
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO!!!"
      Height          =   492
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Click when you are happy with the AI settings"
      Top             =   4440
      Width           =   1452
   End
   Begin VB.Frame AIFrame 
      Caption         =   "AI Settings"
      Height          =   4812
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2532
      Begin VB.FileListBox File1 
         Height          =   1032
         Left            =   120
         Pattern         =   "*.AI"
         TabIndex        =   8
         ToolTipText     =   "Choose a saved setting"
         Top             =   3720
         Width           =   2292
      End
      Begin VB.HScrollBar TwitchS 
         Height          =   372
         LargeChange     =   10
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   4
         Top             =   2520
         Value           =   150
         Width           =   2292
      End
      Begin VB.HScrollBar ThinkS 
         Height          =   372
         LargeChange     =   10
         Left            =   120
         Max             =   200
         Min             =   1
         TabIndex        =   2
         Top             =   1560
         Value           =   50
         Width           =   2292
      End
      Begin VB.Label Label2 
         Caption         =   "Choose preset AI settings..."
         Height          =   252
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Choose a saved setting"
         Top             =   3480
         Width           =   2292
      End
      Begin VB.Label TwitchL 
         Caption         =   "Amount Of Twitching = 150"
         Height          =   252
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "How often the snake turns"
         Top             =   3000
         Width           =   2292
      End
      Begin VB.Label ThinkL 
         Caption         =   "Amount Of Thinking = 50"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "How much time is spent thinking"
         Top             =   2040
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   $"AIForm.frx":0000
         Height          =   1212
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2292
      End
   End
   Begin VB.Label InfoL 
      BorderStyle     =   1  'Fixed Single
      Height          =   3012
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "AIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGO_Click()
LookDist = ThinkS.Value
TurnTime = 201 - TwitchS.Value
Hide
OForm.Visible = True
Unload Me
End Sub

Private Sub WriteInfo()
InfoL = "Here are the AI settings. It is not recommended that you change these before your first go, but afterwards you can muck around with them for a bit of fun. You can load and save settings if you find a good combination. To save, type in a filename in the textbox and click on the button 'Save AI Settings'."
End Sub

Private Sub cmdSave_Click()
LookDist = ThinkS.Value
TurnTime = 201 - TwitchS.Value
Open App.Path & "\" & Text1 & ".AI" For Random As #1 Len = 1
Put #1, 1, LookDist
Put #1, 3, TurnTime
Close #1
File1.Refresh
End Sub

Private Sub File1_DblClick()
For i = 1 To 5
Open App.Path & "\" & File1.FileName For Random As #1 Len = 1
Get #1, 1, LookDist
Get #1, 3, TurnTime
Close #1
ThinkS = LookDist
TwitchS = 201 - TurnTime
Next
End Sub

Private Sub Form_Load()
Randomize Timer
File1.Path = App.Path
WriteInfo
End Sub

Private Sub ThinkS_Change()
ThinkL = "Amount Of Thinking = " & ThinkS.Value
LookDist = ThinkS.Value
TurnTime = 201 - TwitchS.Value
End Sub

Private Sub TwitchS_Change()
TwitchL = "Amount Of Twitching = " & TwitchS.Value
LookDist = ThinkS.Value
TurnTime = 201 - TwitchS.Value
End Sub
