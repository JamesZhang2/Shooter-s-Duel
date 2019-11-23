VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Times: 100000"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "C Wins:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "B Wins:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "A Wins:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Order: A-B-C"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "A:75%  B:50%  C:25%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AIsAlive, BIsAlive, CIsAlive As Integer
'是否存活
Dim AWin, BWin, CWin As Long
'每人获胜次数
Dim num As Single

Private Sub Command1_Click()
  AWin = 0
  BWin = 0
  CWin = 0
  For i = 1 To 100000
    StartGame
  Next i
  Label4.Caption = AWin
  Label6.Caption = BWin
  Label8.Caption = CWin
End Sub

Sub StartGame()
  Initialize
  '初始化
  Do
    If AIsAlive = 1 Then
    'A开枪
      If BIsAlive = 1 Then
      'B未死，射击B
        Randomize
        num = Rnd()
        If num < 0.75 Then
        '击中目标
          BIsAlive = 0
        End If
      Else
      'B已死
        If CIsAlive = 1 Then
        'C未死，射击C
          Randomize
          num = Rnd()
          If num < 0.75 Then
          '击中目标
            CIsAlive = 0
          End If
        Else
        'BC已死，A获胜
          AWin = AWin + 1
          Exit Do
        End If
      End If
    End If
  
    If BIsAlive = 1 Then
    'B开枪
      If AIsAlive = 1 Then
      'A未死，射击A
        Randomize
        num = Rnd()
        If num < 0.5 Then
        '击中目标
          AIsAlive = 0
        End If
      Else
      'A已死
        If CIsAlive = 1 Then
        'C未死，射击C
          Randomize
          num = Rnd()
          If num < 0.5 Then
          '击中目标
            CIsAlive = 0
          End If
        Else
        'AC已死，B获胜
          BWin = BWin + 1
          Exit Do
        End If
      End If
    End If
    
    If CIsAlive = 1 Then
    'C开枪
      If AIsAlive = 1 Then
      'A未死，射击A
        Randomize
        num = Rnd()
        If num < 0.25 Then
        '击中目标
          AIsAlive = 0
        End If
      Else
      'A已死
        If BIsAlive = 1 Then
        'B未死，射击B
          Randomize
          num = Rnd()
          If num < 0.25 Then
          '击中目标
            BIsAlive = 0
          End If
        Else
        'AB已死，C获胜
          CWin = CWin + 1
          Exit Do
        End If
      End If
    End If
  Loop
End Sub

Sub Initialize()
  AIsAlive = 1
  BIsAlive = 1
  CIsAlive = 1
End Sub
