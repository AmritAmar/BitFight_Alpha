VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6024
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6024
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1920
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   11400
      X2              =   11400
      Y1              =   0
      Y2              =   6000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Label3"
      Height          =   372
      Left            =   11400
      TabIndex        =   2
      Top             =   5640
      Width           =   372
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "Label2"
      Height          =   372
      Left            =   11400
      TabIndex        =   1
      Top             =   720
      Width           =   372
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Label1"
      Height          =   372
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   372
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UserBitX As Integer 'X Position of User Bit
Public UserBitY As Integer 'Y Position of User Bit

Public UserBitXSpeed As Integer 'Speed X of User Bit
Public UserBitYSpeed As Integer 'Speed Y of User Bit

Public Enemy1BitX As Integer 'X Position of User Bit
Public Enemy1BitY As Integer 'Y Position of User Bit

Public Enemy1BitXSpeed As Integer 'Speed X of User Bit
Public Enemy1BitYSpeed As Integer 'Speed Y of User Bit

Public Enemy2BitX As Integer 'X Position of User Bit
Public Enemy2BitY As Integer 'Y Position of User Bit

Public Enemy2BitXSpeed As Integer 'Speed X of User Bit
Public Enemy2BitYSpeed As Integer 'Speed Y of User Bit

Public SlowMo As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KeyCodeConstants.vbKeyRight Then
        
        UserBitXSpeed = UserBitXSpeed + 10
    
    ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Then
    
        UserBitXSpeed = UserBitXSpeed - 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyUp Then
    
        UserBitYSpeed = UserBitYSpeed - 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyDown Then
    
        UserBitYSpeed = UserBitYSpeed + 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyY Then
    
        Timer2.Enabled = True
        
    ElseIf KeyCode = KeyCodeConstants.vbKeySpace Then
    
        If SlowMo = True Then
        
            SlowMo = False
            
        Else
        
            SlowMo = True
            
        End If
    
    End If

End Sub

Private Sub Form_Load()

    UserBitX = 1000
    UserBitY = 3000
    
    Enemy1BitX = 11400
    Enemy2BitX = 11400
    
    Enemy1BitY = CInt(Int((5000 * Rnd()) + 1))
    Enemy2BitY = CInt(Int((5000 * Rnd()) + 1))
    'CInt(Math.Floor((upperbound - lowerbound + 1) * Rnd())) + lowerbound
    
    Enemy1BitYSpeed = 0
    Enemy2BitYSpeed = 0
    
    Enemy1BitXSpeed = CInt(Int((-75 * Rnd()) + -50))
    Enemy2BitXSpeed = CInt(Int((-75 * Rnd()) + -50))
    

End Sub

Private Sub Timer1_Timer()

    Label1.Top = UserBitY
    Label1.Left = UserBitX
    
    UserBitX = UserBitX + UserBitXSpeed
    UserBitY = UserBitY + UserBitYSpeed
    
    If SlowMo = True Then
    
        Timer1.Interval = 50
        Timer2.Interval = 0
        
    Else
    
        Timer1.Interval = 1
        Timer2.Interval = 2
        
    End If

End Sub

Private Sub Timer2_Timer()

    Label2.Top = Enemy1BitY
    Label2.Left = Enemy1BitX
    
    Label3.Top = Enemy2BitY
    Label3.Left = Enemy2BitX
    
    Enemy1BitX = Enemy1BitX + Enemy1BitXSpeed
    Enemy2BitX = Enemy2BitX + Enemy2BitXSpeed
    
    Enemy2BitY = Enemy2BitY + Enemy2BitYSpeed
End Sub

Private Sub Timer3_Timer()

    If (Form1.Width - Enemy1BitX) > Form1.Width Then
    
        Enemy1BitX = 11400
        Enemy1BitY = CInt(Int((5000 * Rnd()) + 1))
        
        Enemy1BitXSpeed = CInt(Int((-75 * Rnd()) + -50))
        
    End If
    
    
    If (Form1.Width - Enemy2BitX) > Form1.Width Then
    
        Enemy2BitX = 11400
        Enemy2BitY = CInt(Int((5000 * Rnd()) + 1))
        
        Enemy2BitXSpeed = CInt(Int((-75 * Rnd()) + -50))
        'Enemy2BitYSpeed = CInt(Int((75 * Rnd()) + 75))
        
    End If
    
    If (Form1.Height - Enemy2BitY) > Form1.Height Then
    
        Enemy2BitYSpeed = Enemy2BitYSpeed * -1
        
    End If
    
    If (Form1.Height - Enemy2BitY) < 0 Then
    
        Enemy2BitYSpeed = Enemy2BitYSpeed * -1
        
    End If
    
    
'    If (UserBitX - Enemy1BitX) > (-Label1.Width) Then ()CREATE A SPEEDING PIXEL AT THE BACK!
'
'        Enemy1BitX = Enemy1BitX - 500
'
'    End If

    If Abs(Enemy1BitX - UserBitX) < 372 And Abs(Enemy1BitY - UserBitY) < 372 Then 'SEND ENEMY BACK

        Enemy1BitX = Enemy1BitX - 500

    End If
    
End Sub
