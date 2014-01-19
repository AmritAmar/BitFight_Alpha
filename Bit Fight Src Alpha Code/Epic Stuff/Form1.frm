VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   6552
   ClientLeft      =   48
   ClientTop       =   696
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   6552
   ScaleWidth      =   9900
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   6120
      TabIndex        =   1
      Top             =   4080
      Width           =   300
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xspeed As Integer
Dim yspeed As Integer
Dim x As Integer
Dim y As Integer

Private Sub Form_Load()

    x = 3000
    y = 3000
    Timer1.Interval = 1

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'This is the Code to Set the Speed of the User's Square...
'The Variables are then applied to the Coo-ordinates of the Square

    If KeyCode = KeyCodeConstants.vbKeyRight Or KeyCode = KeyCodeConstants.vbKeyD Then
    
        xspeed = xspeed + 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Or KeyCode = KeyCodeConstants.vbKeyA Then
    
        xspeed = xspeed - 10

    ElseIf KeyCode = KeyCodeConstants.vbKeyUp Or KeyCode = KeyCodeConstants.vbKeyW Then
    
        yspeed = yspeed - 10
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyDown Or KeyCode = KeyCodeConstants.vbKeyS Then
    
        yspeed = yspeed + 10
    
    End If
        
End Sub

Private Sub Timer1_Timer()

    Label1.Left = x
    Label1.Top = y
    x = x + xspeed
    y = y + yspeed
    
    '-----------------------
    
    'If ((l2.Top - Label1.Top) > (-(l2.Height))) And ((l2.Top - Label1.Top) < Label1.Height) And _
        ((l2.Left - Label1.Left) < (Label1.Width)) And ((l2.Left - Label1.Left) > (-(l2.Width))) Then
    
        'Label2.Caption = "HO"
        'xspeed = xspeed * -1
        'yspeed = yspeed * -1
        
    'Else
        
        'Label2.Caption = "NO"
    
    'End If

    '--------------
    
    'For Low Side:((l2.Top - Label1.Top) > (-(l2.Height)))
    'For Up side:((l2.Top - Label1.Top) < Label1.Height)
    'For Left Side:((l2.Left - Label1.Left) < (Label1.Width))
    'For Right Side:((l2.Left - Label1.Left) > (-(l2.Width)))
    
    If ((Form1.Top - Label1.Top) < Label1.Height) Then
    
        y = y - 150
        yspeed = yspeed * -1
        
    End If
    
    If ((Form1.Top - Label1.Top) > (-(Form1.Height))) Then
    
        y = y + 150
        yspeed = yspeed * -1
        
    End If
    
    If ((Form1.Left - Label1.Left) > (-(Form1.Width))) Then
    
        x = x + 150
        xspeed = xspeed * -1
        
    End If
    
    If ((Form1.Left - Label1.Left) < (Label1.Width)) Then
    
        x = x - 150
        xspeed = xspeed * -1
        
    End If
    
    If Abs(Label2.Left - Label1.Left) < 300 And Abs(Label2.Top - Label2.Top) < 300 Then
    
        Label1.BackColor = vbRed
    
    End If
    
End Sub
