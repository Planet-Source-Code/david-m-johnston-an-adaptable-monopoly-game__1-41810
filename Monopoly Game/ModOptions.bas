Attribute VB_Name = "ModOptions"
Option Explicit

Public Sub LoadDB() 'Player selects DataBase to be loaded
Dim i As Integer

Set DB = Nothing
On Error GoTo ErrorCheck
FrmOptions.CD1.Filter = ("records|*.mdb")
If FrmOptions.CD1.CancelError = True Then Exit Sub
    FrmOptions.CD1.ShowOpen
    DBPath = FrmOptions.CD1.FileName
Call ModDataBase.LoadDatabase(DBPath)   'Load the DataBase

Call ModDataBase.SetRecordSets  'Set the record sets
Vers.MoveFirst
House = Vers.Fields("House")
Hotel = Vers.Fields("Hotel")
Go = Vers.Fields("Go")
Jail = Vers.Fields("Jail")
Bank = Vers.Fields("Bank")
PropInfo = Vers.Fields("Deed")
Rent = Vers.Fields("Rent")
Utility = Vers.Fields("Utility")
Station = Vers.Fields("Station")

FrmOptions.LbCBoardColour.ForeColor = &H80000012
FrmOptions.MnuBoardCol.Enabled = True
FrmOptions.LbCBoardColour.Enabled = True
FrmOptions.LbCBoardText.ForeColor = &H80000012
FrmOptions.LbCBoardText.Enabled = True
FrmOptions.MnuTextProp.Enabled = True
FrmOptions.LbCToGame.ForeColor = &H8000000F
FrmOptions.LbCToGame.Enabled = False
FrmOptions.MnuBack.Enabled = False
FrmOptions.LbCPlayers.ForeColor = &H80000012
FrmOptions.MnuEnterPlayers.Enabled = True
FrmOptions.LbCPlayers.Enabled = True
FrmOptions.LblDBLocation.Caption = DBPath

LowMon = False
CurPlayer = 1
Call DrawBoard

Plyr.Index = "Number"
Plyr.MoveFirst
TotPlayers = 0

Do Until Plyr.EOF
    If Plyr.Fields("Number") <> 0 And Plyr.Fields("Number") <> 99 Then
        TotPlayers = TotPlayers + 1
        CurPlayer = (Plyr.Fields("Number"))
        FrmBoard.ImgCounter(Plyr.Fields("Number")).Visible = True
        FrmBoard.ImgCounter(Plyr.Fields("Number")).Picture = LoadPicture(App.Path & (Plyr.Fields("CounterPath")))
        FrmBoard.CboViewPlayer.AddItem (Plyr.Fields("Name"))
        Call PositionPlayer(Plyr.Fields("Square"))
    End If
Plyr.MoveNext
Loop

Hrs = 0
Mins = 0
Secs = 0
Call BankProperty   'Create new list of property in Bank
Plyr.Index = "Number"
If TotPlayers > 1 Then
    CurPlayer = GetCurPlayer
    ViewPlayer = CurPlayer
    FrmOptions.Hide
    FrmBoard.Show
    Call UpdateBoard
    Call UpdateHouses
    Plyr.Seek "=", CurPlayer
    FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"
End If

ErrorCheck:
Exit Sub
End Sub

Public Sub BoardColour()    'Change board colour

FrmOptions.CD1.CancelError = True
On Error GoTo ErrHandler
FrmOptions.CD1.Flags = cdlCCRGBInit
FrmOptions.CD1.ShowColor
BrdColour = FrmOptions.CD1.Color
Call DrawBoard
Exit Sub

ErrHandler:
Exit Sub
End Sub

Public Sub BoardText()  'Change text settings
Dim Ctrl As Object: Dim i As Integer

With FrmOptions
    .CD1.CancelError = True
    On Error GoTo ErrHandler
    .CD1.Flags = cdlCFBoth Or cdlCFEffects
    .CD1.ShowFont

For Each Ctrl In FrmBoard.Controls
    If Ctrl.Name Like "LblName*" Or Ctrl.Name Like "LblPrice*" Then
        Ctrl.Font.Name = .CD1.FontName
        Ctrl.ForeColor = .CD1.Color
        Ctrl.Font.Size = .CD1.FontSize
        Ctrl.Font.Bold = .CD1.FontBold
        Ctrl.Font.Italic = .CD1.FontItalic
        Ctrl.Font.Underline = .CD1.FontUnderline
        Ctrl.Font.Strikethrough = .CD1.FontStrikethru
    End If
Next Ctrl
TextColour = .CD1.Color
Exit Sub
End With
ErrHandler:
Exit Sub
End Sub

Public Sub BackToGame() 'Go back to game
Dim n As Integer

Plyr.Index = "Number"
If Plyr.RecordCount < 2 Then    'Not enough players entered
    n = MsgBox("Please enter player details", vbCritical, "Options")
Exit Sub
End If
Call UpdateHouses   'Re-create houses/hotels
FrmOptions.Hide
FrmBoard.Show
End Sub
