VERSION 5.00
Begin VB.Form SForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3024
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   Picture         =   "SForm.frx":0000
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGO 
      Caption         =   "Play Super Snakes"
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   2520
      Width           =   1692
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGO_Click()
Hide
OForm.Visible = True
Unload Me
End Sub

Private Sub Form_Load()
LookDist = 50
TurnTime = 50
End Sub


