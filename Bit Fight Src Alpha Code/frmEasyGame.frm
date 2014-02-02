VERSION 5.00
Begin VB.Form frmEasyGame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game - Easy!"
   ClientHeight    =   7920
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13848
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   13848
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12600
      TabIndex        =   5
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtLSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtDSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   11520
      TabIndex        =   3
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox txtUSpeed 
      Enabled         =   0   'False
      Height          =   495
      Left            =   11520
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.Timer timEngine 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   4560
      TabIndex        =   6
      Top             =   2640
      Width           =   492
   End
   Begin VB.Label lblScore 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label lblYourScore 
      BackColor       =   &H00000000&
      Caption         =   "Your Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6960
      Width           =   1455
   End
End
Attribute VB_Name = "frmEasyGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarations!


Option Explicit 'Define Every Variable

Dim X As Integer 'User's Square X co-ordinates
Dim Y As Integer 'User's Square Y co-ordinates

Dim a As Integer 'Enemy's Square X co-ordinates
Dim b As Integer 'Enemy's Square Y co-ordinates

Dim w(1 To 8) As Integer 'The Walls that enclose the Squares (Array!)

Dim TEST As Boolean
Dim currentSpeed As Double
Dim newspeed As Double

Dim xspeed As Integer 'User's Square X Speed
Dim yspeed As Integer 'User's Square Y Speed
Dim aspeed As Integer 'Enemy's Square X Speed
Dim bspeed As Integer 'Enemy's Square Y Speed

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
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyN Then
    
        If TEST = False Then
        
            TEST = True
            
        Else
        
            TEST = False
            
        End If
       
    End If
        
End Sub

Private Sub Form_Load()
'The Form Start Code

    'Define Enemy Square Speed
    
    aspeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    bspeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    
    'Define Position of User's Square
    
    X = 4500
    Y = 4900
    
    'Safety Precautions
    
    xspeed = 0
    yspeed = 0
  
    'Wall Co-ordinates
    
    w(1) = 250
    w(2) = 250
    w(3) = 250
    w(4) = 6240
    w(5) = 250
    w(6) = 250
    w(7) = 13550
    w(8) = 250
    
    'Enemy Square Position
    
    a = 11500
    b = 2900
    
    'Enable the Timer
    
    timEngine.Enabled = True
    
    'Set Speed
    
    txtDSpeed.Text = 0
    txtLSpeed.Text = 0
    txtRSpeed.Text = 0
    txtUSpeed.Text = 0
    
End Sub
Private Sub Form_Paint()

    'Draws the User's Square for use in the Engine
    
    Line (X, Y)-(X + 300, Y + 300), RGB(0, 0, 150), BF
    
    'Draws the Enemy Square for use in the Engine
    
    Line (a, b)-(a + 300, b + 300), RGB(255, 0, 0), BF
    
    'Draws the Walls for use in the Engine
    
    Line (w(1), w(2))-(w(1) + 13320, w(2) + 15), RGB(0, 255, 0), BF
    Line (w(3), w(4))-(w(3) + 13320, w(4) + 15), RGB(0, 255, 0), BF
    Line (w(5), w(6))-(w(5) + 15, w(6) + 5990), RGB(0, 255, 0), BF
    Line (w(7), w(8))-(w(7) + 15, w(8) + 5990), RGB(0, 255, 0), BF
    
End Sub

Private Sub timEngine_Timer()
'Error Handler!
    
    On Error GoTo errhandle
    Label1.Caption = TEST
    'Setting the User's Square Able to move!
    
    X = X + xspeed
    Y = Y + yspeed
    
    'Setting the Enemy's Square to move!
    
    a = a + aspeed
    b = b + bspeed
    
    'The Full Collison System!
    
     If Abs(w(1) - X) < 13335 And Abs(w(2) - Y) < 50 Then
    
        yspeed = yspeed * -1
        txtUSpeed.Text = txtUSpeed.Text * -1
        txtDSpeed.Text = txtDSpeed.Text * -1
    
    End If
    
    If Abs(w(1) - a) < 13335 And Abs(w(2) - b) < 50 Then
    
        bspeed = (bspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If
    
    If Abs(w(3) - X) < 13335 And Abs(w(4) - Y) < 270 Then
    
        yspeed = yspeed * -1
        txtUSpeed.Text = txtUSpeed.Text * -1
        txtDSpeed.Text = txtDSpeed.Text * -1
        
    End If
    
    If Abs(w(3) - a) < 13335 And Abs(w(4) - b) < 270 Then
    
        bspeed = (bspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If
    
    If Abs(w(5) - X) < 60 And Abs(w(6) - Y) < 5990 Then
    
        xspeed = xspeed * -1
        txtLSpeed.Text = txtLSpeed.Text * -1
        txtRSpeed.Text = txtRSpeed.Text * -1
    End If
    
    If Abs(w(5) - a) < 60 And Abs(w(6) - b) < 5990 Then
    
        aspeed = (aspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If
    
        If Abs(w(7) - X) < 270 And Abs(w(8) - Y) < 5990 Then
    
        xspeed = xspeed * -1
        txtLSpeed.Text = txtLSpeed.Text * -1
        txtRSpeed.Text = txtRSpeed.Text * -1
    
    End If
    
    If Abs(w(7) - a) < 270 And Abs(w(8) - b) < 5990 Then
    
        aspeed = (aspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
                
    End If
    
    If lblScore.Caption >= 100 Then
    
        MsgBox ("Congratulations! You have Won! The Program will now close!")
        End
                
    End If
    
        'Speed Limits!
    
    If txtUSpeed.Text > 150 Or txtDSpeed.Text > 150 Or txtLSpeed.Text > 150 Or txtRSpeed.Text > 150 Then
    
        MsgBox ("Your Engine Burnt out! Resetting Arena...")
        X = 2500
        Y = 2900

        xspeed = 0
        yspeed = 0
        
        a = 11500
        b = 2900
    
        txtDSpeed.Text = 0
        txtLSpeed.Text = 0
        txtRSpeed.Text = 0
        txtUSpeed.Text = 0
        
        lblScore.Caption = 0
        
        
    End If
    
    If Abs(a - X) < 300 And Abs(b - Y) < 300 Then
    
        MsgBox ("You Died! Resetting Arena...")
        X = 2500
        Y = 2900

        xspeed = 0
        yspeed = 0
        
        a = 11500
        b = 2900
    
        txtDSpeed.Text = 0
        txtLSpeed.Text = 0
        txtRSpeed.Text = 0
        txtUSpeed.Text = 0
        
        lblScore.Caption = 0
                
    End If
    
    'THE POWER OF MATHS
    If TEST = True Then
    
        If X < a And Y > b Then '1st QUADRANT
            'CSPEED, NSPEED = DOUBLE
            currentSpeed = Sqr((aspeed ^ 2) + (bspeed ^ 2))
            newspeed = Sqr((Abs(a - X) ^ 2) + (Abs(Y - b) ^ 2))
            
            aspeed = -currentSpeed * (Abs(a - X) / newspeed)
            bspeed = currentSpeed * (Abs(Y - b) / newspeed)
        
        End If
        
        If X < a And Y < b Then ' 2nd QUADRANT
            'CSPEED, NSPEED = DOUBLE
            currentSpeed = Sqr((aspeed ^ 2) + (bspeed ^ 2))
            newspeed = Sqr((Abs(a - X) ^ 2) + (Abs(b - Y) ^ 2))

            aspeed = -currentSpeed * (Abs(a - X) / newspeed)
            bspeed = -currentSpeed * (Abs(Y - b) / newspeed)

        End If

        If X > a And Y < b Then '3rd QUADRANT
            'CSPEED, NSPEED = DOUBLE
            currentSpeed = Sqr((aspeed ^ 2) + (bspeed ^ 2))
            newspeed = Sqr((Abs(X - a) ^ 2) + (Abs(b - Y) ^ 2))

            aspeed = currentSpeed * (Abs(a - X) / newspeed)
            bspeed = -currentSpeed * (Abs(Y - b) / newspeed)

        End If
        
        If X > a And Y > b Then '4th QUADRANT
            'CSPEED, NSPEED = DOUBLE
            currentSpeed = Sqr((aspeed ^ 2) + (bspeed ^ 2))
            newspeed = Sqr((Abs(X - a) ^ 2) + (Abs(Y - b) ^ 2))
            
            aspeed = currentSpeed * (Abs(a - X) / newspeed)
            bspeed = currentSpeed * (Abs(Y - b) / newspeed)
        
        End If
        

        
    End If
    
    Me.Refresh
    
Exit Sub

errhandle:

'Resets the Square and Square Speed (User)

    MsgBox ("Resetting Your Square!")
    
    X = 2500
    Y = 2900
    
    xspeed = 0
    yspeed = 0
    
    txtDSpeed.Text = 0
    txtLSpeed.Text = 0
    txtRSpeed.Text = 0
    txtUSpeed.Text = 0

End Sub
