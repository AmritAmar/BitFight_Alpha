VERSION 5.00
Begin VB.Form frmGameInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game - Instructions"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtDSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtLSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtRSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.Timer timEngine 
      Interval        =   5
      Left            =   5040
      Top             =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF00FF&
      Caption         =   "Your Speed Status. If you go over 150 then you Die!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   10
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Play!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label lblObjective 
      BackColor       =   &H0000FFFF&
      Caption         =   "The Objective of the Game is to Reach 100, 200 or 500 Points without Dying, whithin the Arena! The best way to learn is to Play!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   480
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label blbInstruction2 
      BackColor       =   &H00C00000&
      Caption         =   "Avoid the Red Squares! They are Enemies! If you Hit them, you will die!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Return to Main Menu!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Label lblInstruction1 
      BackColor       =   &H000000FF&
      Caption         =   "Use WASD or the Arrow Keys to Move your Blue Square."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
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
Attribute VB_Name = "frmGameInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Define Every Variable Option

Dim X As Integer 'User's Square X Co-ordinates
Dim Y As Integer 'User's Square Y Co-ordinates

Dim xspeed As Integer 'User's Square X Speed
Dim yspeed As Integer 'User's Square Y Speed

Dim a As Integer 'Enemy Square X Co-ordinates
Dim b As Integer 'Enemy Square Y Co-ordinates


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'This is the Code to Set the Speed of the User's Square...
'The Variables are then applied to the Coo-ordinates of the Square

    If KeyCode = KeyCodeConstants.vbKeyRight Or KeyCode = KeyCodeConstants.vbKeyD Then
    
        xspeed = xspeed + 10
        txtRSpeed.Text = txtRSpeed.Text + 10
        txtLSpeed.Text = txtLSpeed.Text - 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Or KeyCode = KeyCodeConstants.vbKeyA Then
    
        xspeed = xspeed - 10
        txtLSpeed.Text = txtLSpeed.Text + 10
        txtRSpeed.Text = txtRSpeed.Text - 10

    ElseIf KeyCode = KeyCodeConstants.vbKeyUp Or KeyCode = KeyCodeConstants.vbKeyW Then
    
        yspeed = yspeed - 10
        txtUSpeed.Text = txtUSpeed.Text + 10
        txtDSpeed.Text = txtDSpeed.Text - 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyDown Or KeyCode = KeyCodeConstants.vbKeyS Then
    
        yspeed = yspeed + 10
        txtDSpeed.Text = txtDSpeed.Text + 10
        txtUSpeed.Text = txtUSpeed.Text - 10
       
    End If

End Sub

Private Sub Form_Load()
'The Form Start Code
    
    'Define Position of User's Square
    
    X = 4680
    Y = 2160
    
    'Define Position of Enemy's Square
    
    a = 480
    b = 3120
    
    'Set Speed
    
    txtDSpeed.Text = 0
    txtLSpeed.Text = 0
    txtRSpeed.Text = 0
    txtUSpeed.Text = 0
    
End Sub

Private Sub Form_Paint()
'Draw the User's Square!

    Line (X, Y)-(X + 300, Y + 300), RGB(0, 0, 255), BF
    
'Draw the Enemy Square!

    Line (a, b)-(a + 300, b + 300), RGB(255, 0, 0), BF

End Sub

Private Sub lblExit_Click()
'When Exit is Clicked!
    timEngine.Enabled = False
    Unload Me
    frmMainMenu.Show

End Sub

Private Sub lblPlay_Click()
'When Play is Clicked!
    timEngine.Enabled = False
    Unload Me
    frmGameStart.Show
    
End Sub

Private Sub timEngine_Timer()
'Error Handler

    On Error GoTo erroroverflow
    
'This is where the Square actually moves!
'The timer refreshes the Form every 1 millisecond to Redraw the Position of the Squares!

    X = X + xspeed
    Y = Y + yspeed
    
    If txtUSpeed.Text > 150 Or txtDSpeed.Text > 150 Or txtLSpeed.Text > 150 Or txtRSpeed.Text > 150 Then
    
        MsgBox ("Your Engine Burnt out!")
        X = 4680
        Y = 2160
        
        xspeed = 0
        yspeed = 0
        
        txtDSpeed.Text = 0
        txtLSpeed.Text = 0
        txtRSpeed.Text = 0
        txtUSpeed.Text = 0
        
    End If
    
'Collisons... This code detects when the User's Square collides with the Enemy Square!

    If Abs(a - X) < 300 And Abs(b - Y) < 300 Then
    
        MsgBox ("You just died!")
        
        X = 4680
        Y = 2160
        
        xspeed = 0
        yspeed = 0
        
        txtDSpeed.Text = 0
        txtLSpeed.Text = 0
        txtRSpeed.Text = 0
        txtUSpeed.Text = 0
    End If
    
    Me.Refresh
    
Exit Sub

erroroverflow:

    MsgBox ("If you Escape the Box and Go too far, then you will be automatically reset!")
    X = 4680
    Y = 2160
        
    xspeed = 0
    yspeed = 0

End Sub
