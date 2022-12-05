VERSION 5.00
Begin VB.Form FrmTrade 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Trade"
   ClientHeight    =   5325
   ClientLeft      =   3660
   ClientTop       =   2025
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6750
   Begin VB.ListBox LstPlayers 
      Height          =   1620
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox LstPlayerProp 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label LbCSellProperty 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Sell Property"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label LbCSelectPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select &Buyer"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label LbCSellHouses 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Downgrade Property"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1260
      Width           =   3015
   End
   Begin VB.Label LbCBuyHouse 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Upgrade Property"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   810
      Width           =   3015
   End
   Begin VB.Label LbCMortgage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Mortgage/ Unmortgage"
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
      Left            =   3240
      TabIndex        =   4
      Top             =   1710
      Width           =   3015
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
      Left            =   4200
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label LbCDeed 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "&View Deed"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Menu MnuGame 
      Caption         =   "&Game"
      Begin VB.Menu MnuFinished 
         Caption         =   "&Finished"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Actions 
      Caption         =   "&Actions"
      Begin VB.Menu MnuViewDeed 
         Caption         =   "&View Deed"
      End
      Begin VB.Menu MnuUpgrade 
         Caption         =   "&Upgrade Property"
      End
      Begin VB.Menu MnuDowngrade 
         Caption         =   "&Downgrade Property"
      End
      Begin VB.Menu MnuMortgage 
         Caption         =   "&Mortgage/Unmortgage"
      End
      Begin VB.Menu MnuBuyer 
         Caption         =   "Select &Buyer"
      End
      Begin VB.Menu MnuSell 
         Caption         =   "&Sell Property"
      End
   End
End
Attribute VB_Name = "FrmTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Exit_Click()
Call ModFunctions.EndGame        'End programme
End Sub

Private Sub LbCBuyHouse_Click()
If SelectedProperty = "" Then Exit Sub
Call UpgradeProperty     'check if user can buy a house then allow them to
End Sub

Private Sub LbCDeed_Click()
If SelectedProperty = "" Then Exit Sub
Call Deed(LstPlayerProp.Text)   'Show Title Deed
End Sub

Private Sub LbCfinished_Click()
Call FinishedTrade             'Back to Game
End Sub

Private Sub LbCMortgage_Click()
If SelectedProperty = "" Then Exit Sub
Call Mortgage(LstPlayerProp.Text)   'Mortgage Property
End Sub

Private Sub LbCSelectPlayer_Click()
Call ModTrade.SelectPlayer  'Choose player to sell to
End Sub

Private Sub LbCSellHouses_Click()
If SelectedProperty = "" Then Exit Sub
Call SellHouses         'Sell Houses
End Sub

Private Sub LbCSellProperty_Click()
Call ModTrade.SellProp          'Sell Property
End Sub

Private Sub LstPlayerProp_Click()   'Change Mortgage/Unmortgage option
If FrmTrade.LstPlayerProp.Text = "" Then Exit Sub
Prop.Index = "Name"
Prop.Seek "=", FrmTrade.LstPlayerProp.Text
If Prop.Fields("Mortgaged") = True Then
    FrmTrade.LbCMortgage.Caption = "Unmortgage"
    FrmTrade.MnuMortgage.Caption = "Unmortgage"
Else
    FrmTrade.LbCMortgage.Caption = "Mortgage"
    FrmTrade.MnuMortgage.Caption = "Mortgage"
End If
Call EnableTrade(True)
End Sub

Private Sub MnuBuyer_Click()
Call ModTrade.SelectPlayer  'Choose player to sell to
End Sub

Private Sub MnuDowngrade_Click()
If SelectedProperty = "" Then Exit Sub
Call SellHouses         'Sell Houses
End Sub

Private Sub Mnufinished_Click()
Call FinishedTrade             'Back to Game
End Sub

Private Sub MnuMortgage_Click()
If SelectedProperty = "" Then Exit Sub
Call Mortgage(LstPlayerProp.Text)   'Mortgage Property
End Sub

Private Sub MnuSell_Click()
Call ModTrade.SellProp          'Sell Property
End Sub

Private Sub MnuUpgrade_Click()
If SelectedProperty = "" Then Exit Sub
Call UpgradeProperty     'check if user can buy a house then allow them to
End Sub

Private Sub MnuViewDeed_Click()
If SelectedProperty = "" Then Exit Sub
Call Deed(LstPlayerProp.Text)   'Show Title Deed
End Sub
