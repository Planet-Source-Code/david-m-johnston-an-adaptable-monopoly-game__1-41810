VERSION 5.00
Begin VB.Form FrmPlayers 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Players"
   ClientHeight    =   2760
   ClientLeft      =   2055
   ClientTop       =   3435
   ClientWidth     =   7410
   Icon            =   "FrmPlayers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   7410
   Begin VB.ListBox LstPlayerNo 
      Height          =   1035
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   495
   End
   Begin VB.ListBox LstPlayers 
      Height          =   1035
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox TxtPlayerName 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   900
      Width           =   2115
   End
   Begin VB.Label LbCFinished 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Finished"
      BeginProperty Font 
         Name            =   "Hobo"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LbCOptions 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Options"
      BeginProperty Font 
         Name            =   "Hobo"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LbCEnterPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enter Player"
      BeginProperty Font 
         Name            =   "Hobo"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label LblPlayerNumb 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   900
      Width           =   495
   End
   Begin VB.Image ImgChosenCounter 
      Height          =   480
      Left            =   3480
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   10
      Left            =   6840
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   9
      Left            =   6240
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   8
      Left            =   5640
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   1
      Left            =   4440
      Top             =   300
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   2
      Left            =   5040
      Top             =   300
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   3
      Left            =   5640
      Top             =   300
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   4
      Left            =   6240
      Top             =   300
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   7
      Left            =   5040
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   6
      Left            =   4440
      Top             =   840
      Width           =   480
   End
   Begin VB.Image ImgCounter 
      Height          =   480
      Index           =   5
      Left            =   6840
      Top             =   300
      Width           =   480
   End
   Begin VB.Label LblPlayerNameLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Name:"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   540
      Width           =   585
   End
   Begin VB.Label LblPlayerLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Player:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   765
   End
   Begin VB.Menu MnuGame 
      Caption         =   "&Game"
      Begin VB.Menu MnuEnterPlayer 
         Caption         =   "Enter &Player"
      End
      Begin VB.Menu MnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu MnuFinnished 
         Caption         =   "&Finnished"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Call ModPlayers.CreateForm      'Put available counters on Players form
End Sub

Private Sub ImgCounter_Click(Index As Integer)
FrmPlayers.ImgChosenCounter.Picture = ImgCounter(Index).Picture
'Put chosen counter on board
CounterNumb = ImgCounter(Index).Index
End Sub

Private Sub LbCEnterPlayer_Click()
Call EnterPlyr          'Add player to DataBase
End Sub

Private Sub LbCfinished_Click()
Call ModPlayers.finished
End Sub


Private Sub LbCOptions_Click()
Me.Hide
FrmOptions.Show     'Go to options
End Sub

Private Sub MnuEnterPlayer_Click()
Call EnterPlyr          'Add player to DataBase
End Sub

Private Sub MnuExit_Click()
Call ModFunctions.EndGame            'End Programme
End Sub

Private Sub Mnufinished_Click()
Call ModPlayers.finished
End Sub

Private Sub MnuOptions_Click()
Me.Hide
FrmOptions.Show     'Go to options
End Sub
