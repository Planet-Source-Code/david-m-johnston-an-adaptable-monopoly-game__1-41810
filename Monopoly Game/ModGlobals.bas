Attribute VB_Name = "ModGlobals"
'***************************************************************************
'Programme:  Monopoly game
'
'Files:          Final Program.vbp,
'                FrmBoard.frm , FrmEditDB.frm, FrmOptions.frm
'                FrmPlayers.frm , FrmProperty.frm, FrmTrade.frm
'
'                ModBoard.bas , ModDataBase.bas, ModEditDB.bas, ModFunctions.bas
'                ModGame.bas , ModGameFunctions.bas, ModGlobals.bas
'                ModOptions.bas , ModPlayers.bas, ModTrade.bas
'
'                Clasic.mdb , IT.mdb, Software.mdb
'
'                cannon.ico , dog.ico, horse.ico, iron.ico, race_car.ico, ship.ico, shoe.ico
'                thimble.ico , top_hat.ico, wheel.ico
'
'Function:       To be a flexible game of monopoly that will allow the user to make
'                alterations to both the content of the game and the appearance of the
'                games' user interface.
'
'Description:    The game allows 2 to 6 players who take turns to click 'roll dice'.
'                Then the player moves a number of squares depending or the randomly
'                generated number on the dice.  The player will then be asked to
'                pay rent, buy property as appropriate.  In addition players may also
'                receive a chance or community chest card and, through the 'Trade' form
'                can buy & sell houses as well as sell property to other players or the bank.
'
'Author:         David Johnston
'
'Environments:   MS Visual Basic 6.0, Pentium III 500MHZ, 128mb RAM, Windows 98SE.
'                Pentium 100MHZ, 64mb RAM, Windows 98SE
'
'Notes:          In most instances the user is prevented from entering invalid data through
'                the use of error messages in the form of message boxes.
'                Message boxes are also used to provide the user with information.
'                This programme won't run at resolution of less than 800 x 600.
'
'Revisions:      1.00   12/3/2002 Version 1
'                2.00 21/4/2002 Version 2
'                20/12/2002 Final release - Greater error checking. Fewer bugs.
'                Improved User Interface
'***************************************************************************


Public DB, Prop, PropSet As Recordset
Public Plyr, Vers As Recordset
Public Counter, Chnce, CChest As Recordset
Public Path, FileName, DBPath, CountersPath As String
Public BrdColour, TextColour As String
Public LowRes, FHeight, FWidth, SqBShort, SqSShort, Corner As Integer
Public XPos, YPos As Integer
Public ViewPlayer, CurPlayer, TotPlayers As Integer
Public PlyrAdd, CounterNumb, Dice1, Dice2 As Integer
Public ResComp As Single
Public LowMon As Boolean
Public Hrs, Mins, Secs As Integer
Public AmountOwed As Currency
Public House, Hotel, Go, Jail, Bank, PropInfo, Rent, Utility, Station As String

Sub Main()  'First procedure to run
Dim i As Integer

CountersPath = (App.Path & "\Counters")
Call ResCheck

BrdColour = &HC0FFC0
TextColour = &H0&
FrmOptions.Show     'Show Options Screen

End Sub

Public Sub NewGame()    'Re-set the game
Call DBClearPlayers
FrmPlayers.LstPlayerNo.Clear
FrmPlayers.LstPlayers.Clear
FrmPlayers.LblPlayerNumb.Caption = ""
FrmOptions.LblDBLocation.Caption = ""
FrmBoard.CboViewPlayer.Clear
FrmBoard.LblDice1.Caption = ""
FrmBoard.LblDice2.Caption = ""
TotPlayers = 0
PlyrAdd = 1
FrmBoard.Hide
Call ModPlayers.CreateForm
Set DB = Nothing
Call Main
End Sub
