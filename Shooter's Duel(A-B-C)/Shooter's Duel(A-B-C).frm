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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
'�Ƿ���
Dim AWin, BWin, CWin As Long
'ÿ�˻�ʤ����
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
  '��ʼ��
  Do
    If AIsAlive = 1 Then
    'A��ǹ
      If BIsAlive = 1 Then
      'Bδ�������B
        Randomize
        num = Rnd()
        If num < 0.75 Then
        '����Ŀ��
          BIsAlive = 0
        End If
      Else
      'B����
        If CIsAlive = 1 Then
        'Cδ�������C
          Randomize
          num = Rnd()
          If num < 0.75 Then
          '����Ŀ��
            CIsAlive = 0
          End If
        Else
        'BC������A��ʤ
          AWin = AWin + 1
          Exit Do
        End If
      End If
    End If
  
    If BIsAlive = 1 Then
    'B��ǹ
      If AIsAlive = 1 Then
      'Aδ�������A
        Randomize
        num = Rnd()
        If num < 0.5 Then
        '����Ŀ��
          AIsAlive = 0
        End If
      Else
      'A����
        If CIsAlive = 1 Then
        'Cδ�������C
          Randomize
          num = Rnd()
          If num < 0.5 Then
          '����Ŀ��
            CIsAlive = 0
          End If
        Else
        'AC������B��ʤ
          BWin = BWin + 1
          Exit Do
        End If
      End If
    End If
    
    If CIsAlive = 1 Then
    'C��ǹ
      If AIsAlive = 1 Then
      'Aδ�������A
        Randomize
        num = Rnd()
        If num < 0.25 Then
        '����Ŀ��
          AIsAlive = 0
        End If
      Else
      'A����
        If BIsAlive = 1 Then
        'Bδ�������B
          Randomize
          num = Rnd()
          If num < 0.25 Then
          '����Ŀ��
            BIsAlive = 0
          End If
        Else
        'AB������C��ʤ
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
