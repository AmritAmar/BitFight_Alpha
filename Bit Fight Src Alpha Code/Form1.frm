VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BitFight Test B"
   ClientHeight    =   2340
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   Palette         =   "Form1.frx":0000
   ScaleHeight     =   2340
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "START!"
      Height          =   1212
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":5F599
      Height          =   3132
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3252
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me
    frmHardGame.Show

End Sub
