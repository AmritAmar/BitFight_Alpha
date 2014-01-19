VERSION 5.00
Begin VB.Form frmGameStart 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game - Game Options"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameGameOptions 
      BackColor       =   &H000080FF&
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
      Begin VB.OptionButton opoHard 
         BackColor       =   &H000080FF&
         Caption         =   "Hard!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton opoMedium 
         BackColor       =   &H000080FF&
         Caption         =   "Medium!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opoEasy 
         BackColor       =   &H000080FF&
         Caption         =   "Easy!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Return to Main Menu!"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start the Game!"
         Height          =   615
         Left            =   3480
         TabIndex        =   2
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Caption         =   "Select the Difficulty!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblDifficulty 
         BackColor       =   &H0000FFFF&
         Caption         =   "Select Your Difficulty Above!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   5175
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
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmGameStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
'if Return is Clicked!
    Unload Me
    frmMainMenu.Show

End Sub

Private Sub cmdStart_Click()
'Easy, Medium or Hard Form Opener after Start is Clicked!

    If Me.Caption = "The Game - Game Options - Easy!" Then
    
        MsgBox ("Get Ready!")
        Unload Me
        frmEasyGame.Show
        
            
    ElseIf Me.Caption = "The Game - Game Options - Medium!" Then
    
        MsgBox ("Get Ready!")
        Unload Me
        frmMediumGame.Show
        
        
    ElseIf Me.Caption = "The Game - Game Options - Hard!" Then
    
        MsgBox ("Get Ready!")
        Unload Me
        frmHardGame.Show
        
        
    End If

End Sub


Private Sub opoEasy_Click()
'If Option Easy is Selected!

    Me.Caption = "The Game - Game Options - Easy!"
    lblDifficulty.Caption = "Easy! - You have only 1 opponent. The Goal is to score 100 Points. Every Enemy - Wall Collison is 4 Points! Engine Burns out if Speed is over 150!"

End Sub

Private Sub opoHard_Click()
'If Option Hard is Selected!

    Me.Caption = "The Game - Game Options - Hard!"
    lblDifficulty.Caption = "Hard! - You have only 3 opponents. The Goal is to score 500 Points. Every Enemy to Wall Collison is 1 Point! Engine Burns out if Speed is over 150!"

End Sub

Private Sub opoMedium_Click()
'If Option Medium is Selected!

    Me.Caption = "The Game - Game Options - Medium!"
    lblDifficulty.Caption = "Medium! - You have only 2 opponents. The Goal is to score 200 Points. Every Enemy to Wall Collison is 2 Points! Engine Burns out if Speed is over 150!"

End Sub
