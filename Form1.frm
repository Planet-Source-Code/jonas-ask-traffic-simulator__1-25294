VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traffic"
   ClientHeight    =   12525
   ClientLeft      =   3795
   ClientTop       =   1770
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   835
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   918
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "Speed"
      Height          =   375
      Left            =   2100
      TabIndex        =   35
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Restart"
      Height          =   375
      Left            =   180
      TabIndex        =   34
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdCars 
      Caption         =   "Cars"
      Height          =   375
      Left            =   1200
      TabIndex        =   33
      Top             =   60
      Width           =   795
   End
   Begin VB.PictureBox pGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   6480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   6240
      Picture         =   "Form1.frx":01CE
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   6000
      Picture         =   "Form1.frx":039C
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   5760
      Picture         =   "Form1.frx":056A
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pGround 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   5520
      Picture         =   "Form1.frx":0738
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   11
      Left            =   10140
      Picture         =   "Form1.frx":0906
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   27
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   11
      Left            =   10140
      Picture         =   "Form1.frx":0AF8
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   26
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   10
      Left            =   9960
      Picture         =   "Form1.frx":0CEA
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   25
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   10
      Left            =   9960
      Picture         =   "Form1.frx":0EB8
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   24
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   9
      Left            =   9720
      Picture         =   "Form1.frx":1086
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   23
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   9
      Left            =   9720
      Picture         =   "Form1.frx":1278
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   22
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   8
      Left            =   9480
      Picture         =   "Form1.frx":146A
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   21
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   8
      Left            =   9480
      Picture         =   "Form1.frx":1638
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   20
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   7
      Left            =   9180
      Picture         =   "Form1.frx":1806
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   19
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   6
      Left            =   8940
      Picture         =   "Form1.frx":19D4
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   18
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   5
      Left            =   8700
      Picture         =   "Form1.frx":1BA2
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   17
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   8460
      Picture         =   "Form1.frx":1D70
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   16
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   7
      Left            =   9180
      Picture         =   "Form1.frx":1F3E
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   6
      Left            =   8940
      Picture         =   "Form1.frx":210C
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   5
      Left            =   8700
      Picture         =   "Form1.frx":22DA
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   8460
      Picture         =   "Form1.frx":24CC
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   8160
      Picture         =   "Form1.frx":269A
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   11
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   7920
      Picture         =   "Form1.frx":2868
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   10
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   7440
      Picture         =   "Form1.frx":2A36
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   9
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCarm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   7680
      Picture         =   "Form1.frx":2C04
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   8
      Top             =   300
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   8160
      Picture         =   "Form1.frx":2DD2
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   7920
      Picture         =   "Form1.frx":2FA0
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   7680
      Picture         =   "Form1.frx":316E
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox pCar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   7440
      Picture         =   "Form1.frx":333C
      ScaleHeight     =   11
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1920
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   2700
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3540
      Top             =   180
   End
   Begin VB.PictureBox picKilde 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   4800
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   11955
      Left            =   60
      ScaleHeight     =   793
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   909
      TabIndex        =   0
      Top             =   540
      Width           =   13695
   End
   Begin VB.Label lblMS 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS IS FAR FROM COMPLETE. IT'S A BETA OF A TRAFFIC SYSTEM
'FOR A SIM GAME, BUT YOU'RE WELCOME TO TRY IT OUT :)
'To make any adjustments to the map, just open map1.bmp and juse 0,0,0 for road
' - Jonas


Private Sub cmdCars_Click()
Dim Temp As Long
    Temp = InputBox("How many cars would you like?", "", NumCars)
    If Temp < 1 Or Temp > 99999 Then Exit Sub
    NumCars = Temp
    Command1_Click
End Sub

Private Sub cmdSpeed_Click()
Dim Temp As Currency
    Temp = InputBox("Please input the simulation speed (lower meens faster) ", "", GameSpeed)
    If Temp < 1 Or Temp > 99999999 Then Exit Sub
    GameSpeed = Temp
End Sub

Private Sub Command1_Click()
SetUpNet
End Sub

Private Sub Form_Load()
    ScaleMode = 3
    
    NumCars = 60
    GameSpeed = 500000
    
    ColAsp = RGB(135, 135, 135)
    ColGround = RGB(154, 233, 6)
    
    picKilde.Picture = LoadPicture(App.Path & "\map1.bmp")
    picKilde.AutoSize = True
    
    
    picMain.BackColor = ColGround
    picMain.Print vbNewLine & " PLEASE WAIT....."
    
    Bredde = picKilde.ScaleWidth
    Hoyde = picKilde.ScaleHeight
    
    picBuffer.Width = picMain.Width
    picBuffer.Height = picMain.Height
    
    SetUpNet
    Timer1.Enabled = True
End Sub
Sub MainLoop()
    Do
        t = GetTickCount
        
        For a = 1 To GameSpeed 'Frame Limiter
        
        Next a
        MoveCars
        Paintboard picMain
        DoEvents
        lblMS.Caption = "ms pr. tick: " & GetTickCount - t
    Loop

End Sub
    
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 MsgBox "This program was cooped togehter by Jonas Ask," & vbNewLine & "one really boring afternoon :)", vbInformation, "Note"
 End
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    HITX = Int(X / Size): HITY = Int(Y / Size)
    
    FirstTime = True 'Flag repaint
    If Nett(HITX, HITY) = 1 Then Nett(HITX, HITY) = 0 Else Nett(HITX, HITY) = 1
    
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MainLoop
End Sub
