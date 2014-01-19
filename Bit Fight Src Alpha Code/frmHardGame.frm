VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHardGame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game - Hard!"
   ClientHeight    =   8292
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   13848
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmHardGame.frx":0000
   ScaleHeight     =   8292
   ScaleMode       =   0  'User
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   1200
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1692
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   13812
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "System Information"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   13.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1695
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   13572
         Begin MSComctlLib.ProgressBar prbShield 
            Height          =   372
            Left            =   5760
            TabIndex        =   15
            Top             =   840
            Width           =   3732
            _ExtentX        =   6583
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.TextBox txtRSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   12480
            TabIndex        =   8
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtLSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   10320
            TabIndex        =   7
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtDSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   11400
            TabIndex        =   6
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Text            =   "Your Name will Come here... Eventually."
            Top             =   1200
            Width           =   3015
         End
         Begin VB.TextBox txtUSpeed 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   11400
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   975
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   5760
            TabIndex        =   5
            Top             =   360
            Width           =   3735
            _ExtentX        =   6583
            _ExtentY        =   656
            _Version        =   393216
            Appearance      =   0
            Max             =   150
            Scrolling       =   1
         End
         Begin VB.Label Label8 
            BackColor       =   &H00000000&
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   9600
            TabIndex        =   19
            Top             =   840
            Width           =   612
         End
         Begin VB.Label Label7 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   4920
            TabIndex        =   18
            Top             =   840
            Width           =   732
         End
         Begin VB.Label Label6 
            BackColor       =   &H00000000&
            Caption         =   "  Shields:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   492
            Left            =   3600
            TabIndex        =   16
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            Caption         =   "OFF"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   9600
            TabIndex        =   11
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            Caption         =   "Slow Motion:"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3240
            TabIndex        =   10
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   13.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   5040
            TabIndex        =   9
            Top             =   360
            Width           =   630
         End
         Begin VB.Label lblScore 
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   2775
         End
      End
   End
   Begin VB.Timer timEngine 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   132
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13812
   End
   Begin VB.Label Bit 
      BackColor       =   &H00FF0000&
      Height          =   300
      Left            =   6960
      TabIndex        =   13
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   8040
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Shield 
      BackColor       =   &H000080FF&
      Height          =   400
      Left            =   6960
      TabIndex        =   17
      Top             =   3480
      Width           =   400
   End
End
Attribute VB_Name = "frmHardGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declarations!

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
 
Private Const PBM_SETBARCOLOR As Long = &H409
Private Const PBM_SETBARCOLOR1 As Long = &HFF00&
Private Const PBM_SETBKCOLOR As Long = &H2001
Private Const PROGBAR_DEF_COLOR = &HFF000000 '&H8000000D


Option Explicit 'Define Every Variable

Dim X As Integer 'User's Square X co-ordinates
Dim Y As Integer 'User's Square Y co-ordinates
'Randomize()
Dim P1 As Integer 'Power Up X and Y
Dim P2 As Integer

Dim PTimer As Integer
Dim npress As Boolean

Dim a As Integer 'Enemy's Square X co-ordinates
Dim b As Integer 'Enemy's Square Y co-ordinates
Dim c As Integer 'Enemy's Square X co-ordinates
Dim d As Integer 'Enemy's Square Y co-ordinates
Dim e As Integer 'Enemy's Square X co-ordinates
Dim f As Integer 'Enemy's Square Y co-ordinates

Dim w(1 To 8) As Integer 'The Walls that enclose the Squares (Arrays! Oh Hell Yeah!)

Dim xspeed As Integer 'User's Square X Speed
Dim yspeed As Integer 'User's Square Y Speed
Dim aspeed As Integer 'Enemy's Square X Speed
Dim bspeed As Integer 'Enemy's Square Y Speed
Dim cspeed As Integer 'Enemy's Square X Speed
Dim dspeed As Integer 'Enemy's Square Y Speed
Dim espeed As Integer 'Enemy's Square X Speed
Dim fspeed As Integer 'Enemy's Square Y Speed

Dim slowmospeed As Double 'Progress Bar Speed

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
        
    ElseIf KeyCode = KeyCodeConstants.vbKeySpace Then
    
        If ProgressBar1.Value > 0 Then
    
            If Label3.Caption = "OFF" Then
        
                Label3.Caption = "ON"
            
            Else
        
                Label3.Caption = "OFF"
            
            End If
            
        Else
        
            Label3.Caption = "OUT"
            
        End If
        
    ElseIf KeyCode = KeyCodeConstants.vbKeyN Then
    
        
        
        If prbShield.Value > 0 Then
            npress = True
            Label8.Caption = "ON"
            
        Else
            npress = False
            Label8.Caption = "OUT"
            
        End If
        
    End If
        
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = KeyCodeConstants.vbKeyN Then
    
        
        
        If prbShield.Value > 0 Then
        
            npress = False
            Label8.Caption = "OFF"
            
        Else
            
            npress = False
            Label8.Caption = "OUT"
            
        End If
        
        
    End If

End Sub

Private Sub Form_Load()
'The Form Start Code

    'Define Enemy Square Speed
    
    aspeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    bspeed = CInt(Int((-50 - -75 + 1) * Rnd() + -75))
    cspeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    dspeed = CInt(Int((75 - 50 + 1) * Rnd() + 50))
    espeed = CInt(Int((-75 - -80 + 1) * Rnd() + -80))
    fspeed = CInt(Int((-75 - -80 + 1) * Rnd() + -80))
    
    'Define Position of User's Square
    
    X = 2500
    Y = 2900
    
    'Progress Bar
    
    ProgressBar1.Value = 150
    prbShield.Value = 100
    Call SendMessageLong(ProgressBar1.hwnd, PBM_SETBARCOLOR1, 0&, ByVal 255) 'Red
    
    slowmospeed = ProgressBar1.Value * (1 / 6)
    Label1.Caption = ""
    
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
    b = 3000
    c = 11500
    d = 4000
    e = 11500
    f = 2000
    
    
    ' Generate random value between 1 and 6.
'Dim value As Integer = CInt(Int((6 * Rnd()) + 1))

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
    
    'Line (X, Y)-(X + 300, Y + 300), RGB(0, 0, 150), BF
    
    'power up
    'Line (P1, P2)-(P1 + 300, P2 + 300), RGB(0, 100, 100), BF
    
    'Draws the Enemy Square for use in the Engine
    
    Line (a, b)-(a + 300, b + 300), RGB(255, 0, 0), BF
    Line (c, d)-(c + 300, d + 300), RGB(255, 0, 0), BF
    Line (e, f)-(e + 300, f + 300), RGB(255, 0, 0), BF
    
    'Draws the Walls for use in the Engine
    
    Line (w(1), w(2))-(w(1) + 13320, w(2) + 15), RGB(0, 255, 0), BF
    Line (w(3), w(4))-(w(3) + 13320, w(4) + 15), RGB(0, 255, 0), BF
    Line (w(5), w(6))-(w(5) + 15, w(6) + 5990), RGB(0, 255, 0), BF
    Line (w(7), w(8))-(w(7) + 15, w(8) + 5990), RGB(0, 255, 0), BF
    
End Sub

Private Sub timEngine_Timer()
'Error Handler!

    On Error GoTo errhandle
    
    'SHIELD
    Label7.Caption = Round(CDbl(prbShield.Value), 1)
    
    If npress = True Then
    
        If Label8.Caption = "ON" Then
        
            If prbShield.Value > 0.3 Then
                
                    Shield.Top = Bit.Top - 50
                    Shield.Left = Bit.Left - 50
                    prbShield.Value = prbShield.Value - 0.2
                    
            Else
                
                npress = False
        
            End If
                
            
        Else
        
            Shield.Top = 10000
            Shield.Left = 10000
         
        End If
            
    Else

         Shield.Top = 10000
         Shield.Left = 10000
         
    End If
    
    If prbShield.Value < 0.3 Then
    
        npress = False
        prbShield.Value = 0
        Label8.Caption = "OUT"
        Shield.Top = 10000
        Shield.Left = 10000
        
    ElseIf prbShield.Value > 0.3 Then
    
        If npress = False Then
        
            Label8.Caption = "OFF"
            
        End If
        
    End If
    
    
    Bit.Left = X
    Bit.Top = Y
    
    'txt Speeds
    
    txtUSpeed = yspeed * -1
    txtDSpeed = yspeed
    txtLSpeed = xspeed * -1
    
    txtRSpeed = xspeed
    
    If Abs(P1 - X) < 300 And Abs(P2 - Y) < 300 And Label4.Visible = True Then
    
        If PTimer = 3 Then
        
            ProgressBar1.Value = 150
            Label4.Visible = False
            lblScore = lblScore + 50
        
        ElseIf PTimer = 6 Then
        
            prbShield.Value = 100
            Label4.Visible = False
            lblScore = lblScore + 50
            
       End If
                
    End If
   
    If Y - Label5.Top < 0 Then
        
        Y = Y + 300
        yspeed = yspeed * -1
        
    End If
    
    
    'Setting the User's Square Able to move!
    
    X = X + xspeed
    Y = Y + yspeed
    
    'Setting the Enemy's Square to move!
    
    a = a + aspeed
    b = b + bspeed
    c = c + cspeed
    d = d + dspeed
    e = e + espeed
    f = f + fspeed
    
    If ProgressBar1.Value = 0 Then
    
        Label3.Caption = "OUT"
        
        
    End If
    
    'ProgressBar Label Settings
    
    Label1.Caption = slowmospeed
    slowmospeed = Round(CDbl(ProgressBar1.Value * (1 / 6)), 1)
    
    'Progress Bar!
    
    If Label3.Caption = "ON" Then
    
        If timEngine.Interval = 1 Then
        
            timEngine.Interval = 50
            ProgressBar1.Value = ProgressBar1.Value - 0.5
            
        Else
        
            timEngine.Interval = 1
            
        End If
        
    ElseIf Label3.Caption = "OFF" Then
    
        timEngine.Interval = 1
            
                  
    ElseIf Label3.Caption = "OUT" Then
    
        timEngine.Interval = 1
    
    End If
    
    'ProgressBar Color!
    
'    If ProgressBar1.Value < 90 Then
'
'        Call SendMessageLong(ProgressBar1.hwnd, PBM_SETBARCOLOR, 0&, ByVal 255) 'Red
'
'    End If
'    If ProgressBar1.Value > 89 Then
'
'        Call SendMessageLong(ProgressBar1.hwnd, PBM_SETBARCOLOR1, 0&, ByVal 255) 'Red
'
'    End If
    
    'The Full Collison System!
    
     If Abs(w(1) - X) < 13335 And Abs(w(2) - Y) < 50 Then
     
        yspeed = yspeed * -1
        txtUSpeed.Text = txtUSpeed.Text * -1
        txtDSpeed.Text = txtDSpeed.Text * -1
    
    End If
    
    If Abs(w(1) - a) < 13335 And Abs(w(2) - b) < 50 Then
    
        bspeed = (bspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If
    
    If Abs(w(3) - X) < 13335 And Abs(w(4) - Y) < 270 Then
    
        yspeed = yspeed * -1
        txtUSpeed.Text = txtUSpeed.Text * -1
        txtDSpeed.Text = txtDSpeed.Text * -1
        
    End If
    
    If Abs(w(3) - a) < 13335 And Abs(w(4) - b) < 270 Then
    
        bspeed = (bspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If
    
    If Abs(w(5) - X) < 60 And Abs(w(6) - Y) < 5990 Then
    
        xspeed = xspeed * -1
        txtLSpeed.Text = txtLSpeed.Text * -1
        txtRSpeed.Text = txtRSpeed.Text * -1
    End If
    
    If Abs(w(5) - a) < 60 And Abs(w(6) - b) < 5990 Then
    
        aspeed = (aspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If
    
        If Abs(w(7) - X) < 270 And Abs(w(8) - Y) < 5990 Then
    
        xspeed = xspeed * -1
        txtLSpeed.Text = txtLSpeed.Text * -1
        txtRSpeed.Text = txtRSpeed.Text * -1
    
    End If
    
    If Abs(w(7) - a) < 270 And Abs(w(8) - b) < 5990 Then
    
        aspeed = (aspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
                
    End If
    
        
    If Abs(w(1) - c) < 13335 And Abs(w(2) - d) < 50 Then
    
        dspeed = (dspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If

    If Abs(w(3) - c) < 13335 And Abs(w(4) - d) < 270 Then
    
        dspeed = (dspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If
    
    If Abs(w(5) - c) < 60 And Abs(w(6) - d) < 5990 Then
    
        cspeed = (cspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
        
    End If
    
    If Abs(w(7) - c) < 270 And Abs(w(8) - d) < 5990 Then
    
        cspeed = (cspeed * -1)
        lblScore.Caption = lblScore.Caption + 2
                
    End If
    
        
    If Abs(w(1) - e) < 13335 And Abs(w(2) - f) < 50 Then
    
        fspeed = (fspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If

    If Abs(w(3) - e) < 13335 And Abs(w(4) - f) < 270 Then
    
        fspeed = (fspeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If
    
    If Abs(w(5) - e) < 60 And Abs(w(6) - f) < 5990 Then
    
        espeed = (espeed * -1)
        lblScore.Caption = lblScore.Caption + 4
        
    End If
    
    If Abs(w(7) - e) < 270 And Abs(w(8) - f) < 5990 Then
    
        espeed = (espeed * -1)
        lblScore.Caption = lblScore.Caption + 4
                
    End If
    
    'Score
    
    If lblScore.Caption >= 500 Then
    
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
        b = 3000
        c = 11500
        d = 4000
        e = 11500
        f = 2000
    
        txtDSpeed.Text = 0
        txtLSpeed.Text = 0
        txtRSpeed.Text = 0
        txtUSpeed.Text = 0
        
        lblScore.Caption = 0
        Unload Me
        frmHardGame.Show
                
    End If
    
    'SHIELD
'    KeyNumber = checkKey(KeyCodeA, 0)
'
'    If KeyNumber = 1 Then
'
'        Shield.Top = Bit.Top - 50
'        Shield.Left = Bit.Left - 50
'
'    Else
'
'        Shield.Top = 9000
'        Shield.Left = 9000
'
'    End If
'
'    'FORCE FIELD!
'    If Abs(a - X) < 5000 And Abs(b - Y) < 5000 Then
'
'        If X > a And Y < b Then
'
'
'
'            Label7.Caption = "Yes"
'            timEngine.Interval = 50
'
'        End If
'
'
'    Else
'
'        Label7.Caption = "No"
'        timEngine.Interval = 1
'
'    End If
    If Abs(a - Shield.Left) < 500 And Abs(b - Shield.Top) < 500 Then
        
        If aspeed > 0 Then
        
            a = a - 25
            aspeed = aspeed * -1
            
        Else
        
            a = a + 25
            aspeed = aspeed * -1
            
        End If
        
        If bspeed > 0 Then
        
            b = b - 25
            bspeed = bspeed * -1
            
        Else
        
            b = b + 25
            bspeed = bspeed * -1
            
        End If
    
    End If
    If Abs(c - Shield.Left) < 500 And Abs(d - Shield.Top) < 500 Then
        
        If cspeed > 0 Then
        
            c = c - 25
            cspeed = cspeed * -1
            
        Else
        
            c = c + 25
            cspeed = cspeed * -1
            
        End If
        
        If dspeed > 0 Then
        
            d = d - 25
            dspeed = dspeed * -1
            
        Else
        
            d = d + 25
            dspeed = dspeed * -1
            
        End If
    
    End If
    If Abs(e - Shield.Left) < 500 And Abs(f - Shield.Top) < 500 Then
        
        If espeed > 0 Then
        
            e = e - 25
            espeed = espeed * -1
            
        Else
        
            e = e + 25
            espeed = espeed * -1
            
        End If
        
        If fspeed > 0 Then
        
            f = f - 25
            fspeed = fspeed * -1
            
        Else
        
            f = f + 25
            fspeed = fspeed * -1
            
        End If
    
    End If
    If Abs(a - X) < 300 And Abs(b - Y) < 300 Then
    
        If npress = False Then
            MsgBox ("You Died! Resetting Arena...")
            X = 2500
            Y = 2900

            xspeed = 0
            yspeed = 0
        
            a = 11500
            b = 2900
            c = 11500
            d = 3100
            e = 11500
            f = 2000
            txtDSpeed.Text = 0
            txtLSpeed.Text = 0
            txtRSpeed.Text = 0
            txtUSpeed.Text = 0
        
            lblScore.Caption = 0
            Unload Me
            frmHardGame.Show
        End If
    End If

    If Abs(c - X) < 300 And Abs(d - Y) < 300 Then
        If npress = False Then
            MsgBox ("You Died! Resetting Arena...")
            X = 2500
            Y = 2900

            xspeed = 0
            yspeed = 0
        
            a = 11500
            b = 2900
            c = 11500
            d = 3100
            e = 11500
            f = 2000
            txtDSpeed.Text = 0
            txtLSpeed.Text = 0
            txtRSpeed.Text = 0
            txtUSpeed.Text = 0
        
            lblScore.Caption = 0
            Unload Me
            frmHardGame.Show
        End If
    End If
    If Abs(e - X) < 300 And Abs(f - Y) < 300 Then
        If npress = False Then
            MsgBox ("You Died! Resetting Arena...")
            X = 2500
            Y = 2900

            xspeed = 0
            yspeed = 0
        
            a = 11500
            b = 2900
            c = 11500
            d = 3100
            e = 11500
            f = 2000
            txtDSpeed.Text = 0
            txtLSpeed.Text = 0
            txtRSpeed.Text = 0
            txtUSpeed.Text = 0
        
            lblScore.Caption = 0
            Unload Me
            frmHardGame.Show
        End If
    End If
    
    Me.Refresh
    
Exit Sub

errhandle:

'Resets the Square and Square Speed (User)

    MsgBox ("An Error has appeared! Resetting Game! Error name:" + Error)

    X = 2500
    Y = 2900

    xspeed = 0
    yspeed = 0

    txtDSpeed.Text = 0
    txtLSpeed.Text = 0
    txtRSpeed.Text = 0
    txtUSpeed.Text = 0

    lblScore.Caption = 0
    Unload Me
    frmHardGame.Show

End Sub


Private Sub Timer1_Timer()

    PTimer = CInt(Int(6 * Rnd()) + 1)
    
    If PTimer = 3 Then
        
        Label4.Visible = True
        Timer1.Interval = 5000
        
    ElseIf PTimer = 6 Then
        
        Label4.Visible = True
        Timer1.Interval = 5000
        
    Else
    
        Label4.Visible = False
        Timer1.Interval = 1000
        
    End If
    
    P1 = CInt(Int((12000 * Rnd()) + 500))
    P2 = CInt(Int((5500 * Rnd()) + 500))
    Label4.Left = P1
    Label4.Top = P2

End Sub
