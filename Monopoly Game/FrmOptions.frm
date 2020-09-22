VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmOptions 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Options"
   ClientHeight    =   5040
   ClientLeft      =   2055
   ClientTop       =   3030
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6750
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtCounters 
      Height          =   285
      Left            =   128
      TabIndex        =   1
      Text            =   "Counters Folder"
      Top             =   2520
      Width           =   6495
   End
   Begin VB.Label LbCToGame 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back To &Game"
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
      Left            =   4695
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCPlayers 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Enter Players"
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
      Left            =   3570
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LbCBoardText 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Board &Text"
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
      Left            =   3495
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCBoardColour 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Board &Colour"
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
      Left            =   2295
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCEditDB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Edit DataBase"
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
      Left            =   1095
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LbCLoadDB 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Select DataBase"
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
      Left            =   2205
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label LblDBLocation 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   143
      TabIndex        =   3
      Top             =   480
      Width           =   6465
   End
   Begin VB.Label LblCountersLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Counters Folder Location"
      Height          =   255
      Left            =   2475
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label LblDBLocationLab 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Data Base Location"
      Height          =   255
      Left            =   2655
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
   Begin VB.Menu MnuGame 
      Caption         =   "&Game"
      Begin VB.Menu MnuBack 
         Caption         =   "&Back to Game"
      End
      Begin VB.Menu MnuNewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu MnuSelectDB 
         Caption         =   "&Select DataBase"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu MnuEditDB 
         Caption         =   "&Edit Data Base"
      End
      Begin VB.Menu MnuBoardCol 
         Caption         =   "&Board Colour"
      End
      Begin VB.Menu MnuTextProp 
         Caption         =   "&Text Properties"
      End
      Begin VB.Menu MnuEnterPlayers 
         Caption         =   "Enter &Players"
      End
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MnuBack_Click()
Call ModOptions.BackToGame  'Go back to Game
End Sub

Private Sub MnuBoardCol_Click()
Call ModOptions.BoardColour         'Change Board Colour
End Sub

Private Sub MnuEditDB_Click()
FrmEditDB.Show
Call EditDB         'Go to Edit DataBase Options
End Sub

Private Sub MnuEnterPlayers_Click()
CountersPath = FrmOptions.TxtCounters.Text
Me.Hide
FrmPlayers.Show             'Enter Players
End Sub

Private Sub MnuExit_Click()
Call ModFunctions.EndGame                'Exit Programme
End Sub

Private Sub Form_Load()
FrmOptions.TxtCounters.Text = CountersPath
End Sub

Private Sub LbCBoardColour_Click()
Call ModOptions.BoardColour      'Change Board Colour
End Sub

Private Sub LbCBoardText_Click()
Call ModOptions.BoardText       'Change board text appearance
End Sub

Private Sub LbCEditDB_Click()
FrmEditDB.Show
Call EditDB           'Go to Edit DataBase Options
End Sub

Private Sub LbCLoadDB_Click()
Call ModOptions.LoadDB  'Choose & Load a DataBase
End Sub

Private Sub LbCPlayers_Click()
CountersPath = FrmOptions.TxtCounters.Text
Me.Hide
FrmPlayers.Show             'Enter Players
End Sub

Private Sub LbCToGame_Click()
Call ModOptions.BackToGame  'Go back to Game
End Sub

Private Sub MnuNewGame_Click()
Call NewGame                'Start a New Game
End Sub

Private Sub MnuSelectDB_Click()
Call ModOptions.LoadDB  'Choose & Load a DataBase
End Sub

Private Sub MnuTextProp_Click()
Call ModOptions.BoardText       'Change board text appearance
End Sub
