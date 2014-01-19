VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game - Main Menu"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameMainMenu 
      BackColor       =   &H000080FF&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5655
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit The Game!"
         Height          =   855
         Left            =   1080
         TabIndex        =   4
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton cmdInstructions 
         Caption         =   "How to Play?"
         Height          =   855
         Left            =   1080
         TabIndex        =   3
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play The Game!"
         Height          =   855
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
   End
   Begin VB.Label lblTheGame 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "The Game"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
'Exit is Clicked, then the Program terminates!
    End

End Sub

Private Sub cmdInstructions_Click()
'Show Insturctions Form
    frmGameInfo.Show
    Unload Me

End Sub

Private Sub cmdPlay_Click()
'Show the Game Options Form
    frmGameStart.Show
    Unload Me

End Sub

