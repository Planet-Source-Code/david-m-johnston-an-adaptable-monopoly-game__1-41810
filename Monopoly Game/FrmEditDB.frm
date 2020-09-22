VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmEditDB 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Edit DataBase"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtDBPath 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   4365
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7605
      Left            =   68
      TabIndex        =   2
      Top             =   600
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   13414
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabHeight       =   520
      BackColor       =   12648384
      TabCaption(0)   =   "Property"
      TabPicture(0)   =   "FrmEditDB.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CboSet"
      Tab(0).Control(1)=   "TxtRent(6)"
      Tab(0).Control(2)=   "TxtPrice"
      Tab(0).Control(3)=   "LstProperty(11)"
      Tab(0).Control(4)=   "LstProperty(9)"
      Tab(0).Control(5)=   "LstProperty(7)"
      Tab(0).Control(6)=   "LstProperty(6)"
      Tab(0).Control(7)=   "LstProperty(3)"
      Tab(0).Control(8)=   "LstProperty(2)"
      Tab(0).Control(9)=   "LstProperty(1)"
      Tab(0).Control(10)=   "LstProperty(0)"
      Tab(0).Control(11)=   "TxtName"
      Tab(0).Control(12)=   "TxtRent(7)"
      Tab(0).Control(13)=   "TxtRent(8)"
      Tab(0).Control(14)=   "TxtRent(9)"
      Tab(0).Control(15)=   "TxtRent(10)"
      Tab(0).Control(16)=   "TxtRent(11)"
      Tab(0).Control(17)=   "LstProperty(8)"
      Tab(0).Control(18)=   "LstProperty(10)"
      Tab(0).Control(19)=   "LblPropNo"
      Tab(0).Control(20)=   "LbCAddRec"
      Tab(0).Control(21)=   "Label1"
      Tab(0).Control(22)=   "Label2"
      Tab(0).Control(23)=   "Label5"
      Tab(0).Control(24)=   "Label6"
      Tab(0).Control(25)=   "Label7"
      Tab(0).Control(26)=   "Label8"
      Tab(0).Control(27)=   "Label9"
      Tab(0).Control(28)=   "Label10"
      Tab(0).Control(29)=   "Label3"
      Tab(0).Control(30)=   "Label4"
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Cards"
      TabPicture(1)   =   "FrmEditDB.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FraCChest"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraChance"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Sets"
      TabPicture(2)   =   "FrmEditDB.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtHPrice(0)"
      Tab(2).Control(1)=   "TxtHPrice(1)"
      Tab(2).Control(2)=   "TxtHPrice(2)"
      Tab(2).Control(3)=   "TxtHPrice(3)"
      Tab(2).Control(4)=   "TxtHPrice(4)"
      Tab(2).Control(5)=   "TxtHPrice(5)"
      Tab(2).Control(6)=   "TxtHPrice(6)"
      Tab(2).Control(7)=   "TxtHPrice(7)"
      Tab(2).Control(8)=   "CD2"
      Tab(2).Control(9)=   "LblSet(0)"
      Tab(2).Control(10)=   "LblSet(1)"
      Tab(2).Control(11)=   "LblSet(2)"
      Tab(2).Control(12)=   "LblSet(3)"
      Tab(2).Control(13)=   "LblSet(4)"
      Tab(2).Control(14)=   "LblSet(5)"
      Tab(2).Control(15)=   "LblSet(6)"
      Tab(2).Control(16)=   "LblSet(7)"
      Tab(2).Control(17)=   "LblSet1Lab"
      Tab(2).Control(18)=   "LblSet2Lab"
      Tab(2).Control(19)=   "LblSet3Lab"
      Tab(2).Control(20)=   "LblSet4"
      Tab(2).Control(21)=   "LblSet5Lab"
      Tab(2).Control(22)=   "LblSet6Lab"
      Tab(2).Control(23)=   "LblSet7Lab"
      Tab(2).Control(24)=   "LblSet8Lab"
      Tab(2).Control(25)=   "LblHousePriceLab(0)"
      Tab(2).Control(26)=   "LblHousePriceLab(2)"
      Tab(2).Control(27)=   "LblHousePriceLab(3)"
      Tab(2).Control(28)=   "LblHousePriceLab(4)"
      Tab(2).Control(29)=   "LblHousePriceLab(5)"
      Tab(2).Control(30)=   "LblHousePriceLab(6)"
      Tab(2).Control(31)=   "LblHousePriceLab(7)"
      Tab(2).Control(32)=   "LblHousePriceLab(8)"
      Tab(2).Control(33)=   "LbCUpdateCols"
      Tab(2).Control(34)=   "LblSetInfoLab"
      Tab(2).ControlCount=   35
      TabCaption(3)   =   "Names"
      TabPicture(3)   =   "FrmEditDB.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtNamesName(8)"
      Tab(3).Control(1)=   "TxtNamesName(7)"
      Tab(3).Control(2)=   "TxtNamesName(6)"
      Tab(3).Control(3)=   "TxtNamesName(5)"
      Tab(3).Control(4)=   "TxtNamesName(4)"
      Tab(3).Control(5)=   "TxtNamesName(3)"
      Tab(3).Control(6)=   "TxtNamesName(2)"
      Tab(3).Control(7)=   "TxtNamesName(1)"
      Tab(3).Control(8)=   "TxtNamesName(0)"
      Tab(3).Control(9)=   "LblNamesLab"
      Tab(3).Control(10)=   "LbcUpdateNames"
      Tab(3).Control(11)=   "LblNameLab(0)"
      Tab(3).Control(12)=   "LblNameLab(1)"
      Tab(3).Control(13)=   "LblNameLab(2)"
      Tab(3).Control(14)=   "LblNameLab(3)"
      Tab(3).Control(15)=   "LblNameLab(4)"
      Tab(3).Control(16)=   "LblNameLab(5)"
      Tab(3).Control(17)=   "LblNameLab(6)"
      Tab(3).Control(18)=   "LblNameLab(7)"
      Tab(3).Control(19)=   "LblNameLab(8)"
      Tab(3).ControlCount=   20
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   8
         Left            =   -74520
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   7
         Left            =   -65880
         TabIndex        =   104
         Text            =   "Text1"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   6
         Left            =   -68760
         TabIndex        =   103
         Text            =   "Text1"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   5
         Left            =   -71640
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   4
         Left            =   -74520
         TabIndex        =   101
         Text            =   "Text1"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   3
         Left            =   -65880
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   2
         Left            =   -68760
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   1
         Left            =   -71640
         TabIndex        =   98
         Text            =   "Text1"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox TxtNamesName 
         Height          =   285
         Index           =   0
         Left            =   -74520
         TabIndex        =   97
         Text            =   "Text1"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox CboSet 
         Height          =   315
         ItemData        =   "FrmEditDB.frx":0070
         Left            =   -70440
         List            =   "FrmEditDB.frx":0095
         TabIndex        =   49
         Text            =   "Set No"
         Top             =   6540
         Width           =   600
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   6
         Left            =   -68850
         TabIndex        =   48
         Text            =   "Rent"
         Top             =   6540
         Width           =   765
      End
      Begin VB.TextBox TxtPrice 
         Height          =   315
         Left            =   -69720
         TabIndex        =   47
         Text            =   "Price"
         Top             =   6540
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   11
         Left            =   -64350
         TabIndex        =   46
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   9
         Left            =   -66150
         TabIndex        =   45
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   7
         Left            =   -67950
         TabIndex        =   44
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   6
         Left            =   -68850
         TabIndex        =   43
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   3
         Left            =   -69750
         TabIndex        =   42
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   2
         Left            =   -70440
         TabIndex        =   41
         Top             =   1230
         Width           =   600
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   1
         Left            =   -74040
         TabIndex        =   40
         Top             =   1230
         Width           =   3465
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   0
         Left            =   -74700
         TabIndex        =   39
         Top             =   1230
         Width           =   615
      End
      Begin VB.TextBox TxtName 
         Height          =   315
         Left            =   -74010
         TabIndex        =   38
         Text            =   "Property Name"
         Top             =   6540
         Width           =   3525
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   7
         Left            =   -67920
         TabIndex        =   37
         Text            =   "Rent 1 House"
         Top             =   6540
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   8
         Left            =   -66960
         TabIndex        =   36
         Text            =   "Rent 2 Houses"
         Top             =   6540
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   9
         Left            =   -66120
         TabIndex        =   35
         Text            =   "Rent 3 Houses"
         Top             =   6540
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   10
         Left            =   -65280
         TabIndex        =   34
         Text            =   "Rent 4 Houses"
         Top             =   6540
         Width           =   765
      End
      Begin VB.TextBox TxtRent 
         Height          =   315
         Index           =   11
         Left            =   -64320
         TabIndex        =   33
         Text            =   "Rent Hotel"
         Top             =   6540
         Width           =   765
      End
      Begin VB.Frame FraChance 
         Caption         =   "Chance"
         Height          =   3255
         Left            =   240
         TabIndex        =   23
         Top             =   780
         Width           =   10935
         Begin VB.ListBox LstChance 
            Height          =   2010
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox LstChance 
            Height          =   2010
            Index           =   1
            Left            =   840
            TabIndex        =   29
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox TxtText 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Text            =   "Card Text"
            Top             =   2400
            Width           =   7215
         End
         Begin VB.ComboBox CboAction 
            Height          =   315
            Index           =   0
            Left            =   8160
            TabIndex        =   27
            Text            =   "Action"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.ListBox LstChance 
            Height          =   2010
            Index           =   2
            Left            =   8160
            TabIndex        =   26
            Top             =   240
            Width           =   1935
         End
         Begin VB.ListBox LstChance 
            Height          =   2010
            Index           =   3
            Left            =   10200
            TabIndex        =   25
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TxtAmount 
            Height          =   285
            Index           =   0
            Left            =   10200
            TabIndex        =   24
            Text            =   "Amount"
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label LblCardNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Card No"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label LbCUpdateChance 
            Alignment       =   2  'Center
            Caption         =   "Change &Card"
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
            Left            =   5213
            TabIndex        =   31
            Top             =   2760
            Width           =   1455
         End
      End
      Begin VB.Frame FraCChest 
         Caption         =   "Community Chest"
         Height          =   3225
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   10935
         Begin VB.ListBox LstCChest 
            Height          =   2010
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox LstCChest 
            Height          =   2010
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox TxtText 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   18
            Text            =   "Card Text"
            Top             =   2400
            Width           =   7215
         End
         Begin VB.ComboBox CboAction 
            Height          =   315
            Index           =   1
            Left            =   8160
            TabIndex        =   17
            Text            =   "Action"
            Top             =   2400
            Width           =   1935
         End
         Begin VB.ListBox LstCChest 
            Height          =   2010
            Index           =   2
            Left            =   8160
            TabIndex        =   16
            Top             =   240
            Width           =   1935
         End
         Begin VB.ListBox LstCChest 
            Height          =   2010
            Index           =   3
            Left            =   10200
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox TxtAmount 
            Height          =   285
            Index           =   1
            Left            =   10200
            TabIndex        =   14
            Text            =   "Amount"
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label LblCardNo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Card No"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label LbCUpdateChest 
            Alignment       =   2  'Center
            Caption         =   "Change &Card"
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
            Left            =   5213
            TabIndex        =   21
            Top             =   2760
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -72907
         TabIndex        =   12
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -72907
         TabIndex        =   11
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5940
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -70387
         TabIndex        =   10
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -70387
         TabIndex        =   9
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5940
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   -67627
         TabIndex        =   8
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   -67627
         TabIndex        =   7
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5940
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   -64867
         TabIndex        =   6
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox TxtHPrice 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """£""#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   -64747
         TabIndex        =   5
         ToolTipText     =   "Type new cost of houses on this set here"
         Top             =   5940
         Width           =   735
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   8
         Left            =   -67050
         TabIndex        =   4
         Top             =   1230
         Width           =   765
      End
      Begin VB.ListBox LstProperty 
         Height          =   5130
         Index           =   10
         Left            =   -65250
         TabIndex        =   3
         Top             =   1230
         Width           =   765
      End
      Begin MSComDlg.CommonDialog CD2 
         Left            =   -74760
         Top             =   660
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LblNamesLab 
         Caption         =   "Enter new names for the following in the boxes provided"
         Height          =   255
         Left            =   -71107
         TabIndex        =   107
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label LbcUpdateNames 
         Alignment       =   2  'Center
         Caption         =   "&ApplyChanges"
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
         Left            =   -71280
         TabIndex        =   106
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label LblPropNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PropNo"
         Height          =   315
         Left            =   -74700
         TabIndex        =   96
         Top             =   6540
         Width           =   615
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   -74107
         TabIndex        =   95
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   3180
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   -74107
         TabIndex        =   94
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   5460
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   -71587
         TabIndex        =   93
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   3180
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   -71587
         TabIndex        =   92
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   5460
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   -68827
         TabIndex        =   91
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   3180
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   -68827
         TabIndex        =   90
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   5460
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   -66067
         TabIndex        =   89
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   3180
         Width           =   1935
      End
      Begin VB.Label LblSet 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   -65947
         TabIndex        =   88
         ToolTipText     =   "Click here to change this sets colour"
         Top             =   5460
         Width           =   1935
      End
      Begin VB.Label LblSet1Lab 
         Caption         =   "Set 1"
         Height          =   255
         Left            =   -73507
         TabIndex        =   87
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label LblSet2Lab 
         Caption         =   "Set 2"
         Height          =   255
         Left            =   -73507
         TabIndex        =   86
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label LblSet3Lab 
         Caption         =   "Set 3"
         Height          =   255
         Left            =   -70987
         TabIndex        =   85
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label LblSet4 
         Caption         =   "Set 4"
         Height          =   255
         Left            =   -71107
         TabIndex        =   84
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label LblSet5Lab 
         Caption         =   "Set 5"
         Height          =   255
         Left            =   -68227
         TabIndex        =   83
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label LblSet6Lab 
         Caption         =   "Set 6"
         Height          =   255
         Left            =   -68227
         TabIndex        =   82
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label LblSet7Lab 
         Caption         =   "Set 7"
         Height          =   255
         Left            =   -65227
         TabIndex        =   81
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label LblSet8Lab 
         Caption         =   "Set 8"
         Height          =   255
         Left            =   -65347
         TabIndex        =   80
         Top             =   5100
         Width           =   735
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   0
         Left            =   -74107
         TabIndex        =   79
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   2
         Left            =   -66067
         TabIndex        =   78
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   3
         Left            =   -68827
         TabIndex        =   77
         Top             =   5940
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   4
         Left            =   -68827
         TabIndex        =   76
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   5
         Left            =   -71587
         TabIndex        =   75
         Top             =   5940
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   6
         Left            =   -71587
         TabIndex        =   74
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   7
         Left            =   -74107
         TabIndex        =   73
         Top             =   5940
         Width           =   975
      End
      Begin VB.Label LblHousePriceLab 
         Caption         =   "House Price:"
         Height          =   255
         Index           =   8
         Left            =   -65947
         TabIndex        =   72
         Top             =   5940
         Width           =   975
      End
      Begin VB.Label LbCAddRec 
         Alignment       =   2  'Center
         Caption         =   "&Apply Changes"
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
         Left            =   -69900
         TabIndex        =   71
         Top             =   7140
         Width           =   1695
      End
      Begin VB.Label LbCUpdateCols 
         Alignment       =   2  'Center
         Caption         =   "&Apply Changes"
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
         Left            =   -70027
         TabIndex        =   70
         Top             =   6660
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Property Number"
         Height          =   375
         Left            =   -74640
         TabIndex        =   69
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Property Name"
         Height          =   375
         Left            =   -73680
         TabIndex        =   68
         Top             =   780
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Rent"
         Height          =   375
         Left            =   -68760
         TabIndex        =   67
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Rent 1 House"
         Height          =   375
         Left            =   -67920
         TabIndex        =   66
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "Rent 2 Houses"
         Height          =   375
         Left            =   -66960
         TabIndex        =   65
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Rent 3 Houses"
         Height          =   375
         Left            =   -66120
         TabIndex        =   64
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Rent 4 Houses"
         Height          =   375
         Left            =   -65160
         TabIndex        =   63
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label10 
         Caption         =   "Rent Hotel"
         Height          =   375
         Left            =   -64320
         TabIndex        =   62
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Set"
         Height          =   375
         Left            =   -70440
         TabIndex        =   61
         Top             =   780
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Price"
         Height          =   375
         Left            =   -69720
         TabIndex        =   60
         Top             =   780
         Width           =   735
      End
      Begin VB.Label LblSetInfoLab 
         Caption         =   $"FrmEditDB.frx":00BB
         Height          =   615
         Left            =   -72300
         TabIndex        =   59
         Top             =   1020
         Width           =   6375
      End
      Begin VB.Label LblNameLab 
         Caption         =   "House"
         Height          =   375
         Index           =   0
         Left            =   -74520
         TabIndex        =   58
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Station"
         Height          =   375
         Index           =   1
         Left            =   -74520
         TabIndex        =   57
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Hotel"
         Height          =   375
         Index           =   2
         Left            =   -71640
         TabIndex        =   56
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Go"
         Height          =   375
         Index           =   3
         Left            =   -68760
         TabIndex        =   55
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Jail"
         Height          =   375
         Index           =   4
         Left            =   -65880
         TabIndex        =   54
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Bank"
         Height          =   375
         Index           =   5
         Left            =   -74520
         TabIndex        =   53
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Deed"
         Height          =   375
         Index           =   6
         Left            =   -71640
         TabIndex        =   52
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Rent"
         Height          =   375
         Index           =   7
         Left            =   -68760
         TabIndex        =   51
         Top             =   3780
         Width           =   1455
      End
      Begin VB.Label LblNameLab 
         Caption         =   "Utility"
         Height          =   375
         Index           =   8
         Left            =   -65880
         TabIndex        =   50
         Top             =   3780
         Width           =   1455
      End
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
      Left            =   5393
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "FrmEditDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Exit_Click()        'End Programme
Call ModFunctions.EndGame
End Sub

Private Sub Form_Load()
FrmEditDB.TxtDBPath.Text = DBPath
End Sub


Private Sub LbCAddRec_Click()
Call UpdateDBProperty           'Update DataBase
Call CreateLists                'Update property list on Board
End Sub

Private Sub LbCfinished_Click()
Call DrawBoard          'Re-draw (Update) Board
Call UpdateHouses       'Re-draw House/Hotels
Me.Hide
End Sub

Private Sub LbCUpdateChance_Click()  'Update DataBase
Call UpdateDBChance
Call CreateLists
End Sub

Private Sub LbCUpdateChest_Click()  'Update DataBase
Call UpdateDBCChest
Call CreateLists
End Sub

Private Sub LbCUpdateCols_Click()
Call UpdateDBCols           'Update DataBase with new Colours
End Sub

Private Sub LbcUpdateNames_Click()
Dim n As Integer

With FrmEditDB
    Vers.MoveFirst
    Vers.Edit
    Vers.Fields("House") = .TxtNamesName(0).Text
    House = .TxtNamesName(0).Text
    Vers.Fields("Hotel") = .TxtNamesName(1).Text
    Hotel = .TxtNamesName(1).Text
    Vers.Fields("Go") = .TxtNamesName(2).Text
    Go = .TxtNamesName(2).Text
    Vers.Fields("Jail") = .TxtNamesName(3).Text
    Jail = .TxtNamesName(3).Text
    Vers.Fields("Bank") = .TxtNamesName(4).Text
    Bank = .TxtNamesName(4).Text
    Vers.Fields("Deed") = .TxtNamesName(5).Text
    PropInfo = .TxtNamesName(5).Text
    Vers.Fields("Rent") = .TxtNamesName(6).Text
    Rent = .TxtNamesName(6).Text
    Vers.Fields("Utility") = .TxtNamesName(7).Text
    Utility = .TxtNamesName(7).Text
    Vers.Fields("Station") = .TxtNamesName(8).Text
    Station = .TxtNamesName(8).Text
    Vers.Update
End With

n = MsgBox("Names Updated", vbInformation, "Updated")
End Sub

Private Sub LblSet_Click(Index As Integer)  'Allow user to Choose Colour
CD2.CancelError = True
On Error GoTo ErrHandler
CD2.Flags = cdlCCRGBInit
CD2.ShowColor
LblSet(Index).BackColor = CD2.Color
Exit Sub

ErrHandler:
MsgBox "Error"
Exit Sub

End Sub

Private Sub LstCChest_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstCChest(Index).Index
Call CChestSelected(Clicked)    'Select same card in all other fields
End Sub

Private Sub LstChance_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstChance(Index).Index
Call ChanceSelected(Clicked)    'Select same card in all other fields
End Sub

Private Sub LstProperty_Click(Index As Integer)
Dim Clicked As Integer
Clicked = LstProperty(Index).Index
Call PropertySelected(Clicked)  'Select same property in all other fields
End Sub

Private Sub TxtPrice_Click()
FrmEditDB.TxtPrice.Text = ""
End Sub

Private Sub TxtRent_Click(Index As Integer)
FrmEditDB.TxtRent(Index).Text = ""
End Sub
