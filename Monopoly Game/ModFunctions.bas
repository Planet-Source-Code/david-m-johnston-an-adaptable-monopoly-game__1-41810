Attribute VB_Name = "ModFunctions"
Option Explicit

Public Function PosX(ByVal Sqre) As Integer
    'Receives the Square (Sqre) and sets XPos to appropriate X co-ordinate
        'SqBShort = length of short side of squares on bottom
Select Case Sqre
Case 1: XPos = FWidth - Corner
Case 2 To 10: XPos = FWidth - Corner - (Sqre - 1) * SqBShort
Case 11 To 21: XPos = 0
Case 22 To 31: XPos = Corner + ((Sqre - 22) * SqBShort)
Case 31 To 40: XPos = FWidth - Corner - LowRes
End Select

End Function

Public Function PosY(ByVal Sqre) As Integer
    'Receives the Square (Sqre) and sets YPos to appropriate Y co-ordinate
        'SqSShort = length of short side of squares on side
Select Case Sqre
Case 1 To 11: YPos = FHeight - Corner
Case 12 To 20: YPos = FHeight - Corner - (Sqre - 11) * SqSShort
Case 21 To 31: YPos = 0
Case 32 To 40: YPos = Corner + ((Sqre - 32) * SqSShort)
End Select

End Function

Public Function SelectedProperty()
    'Returns name of property selected for trade action
Dim n As Integer
SelectedProperty = FrmTrade.LstPlayerProp.Text
End Function

Public Function SetOwned(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'Reterns true if yes, False if no
Dim i, SetNo, OwnerNumb As Integer

Prop.Index = "Number"
Prop.Seek "=", s
OwnerNumb = Prop.Fields("OwnerNo")
SetNo = Prop.Fields("Set")
SetOwned = True
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("OwnerNo") <> OwnerNumb Then
            SetOwned = False
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Function HousesOnSet(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'and if any property in the set has houses/hotels
        'Reterns true if houses exist, False if none
Dim i, OwnerNumb, SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s
SetNo = Prop.Fields("Set")
HousesOnSet = False
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("HousesOwned") > 0 Then
            HousesOnSet = True
            Exit Do
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Function MortgInSetSet(ByVal s) As Boolean
    'Determines if square (s) is part of a set owned by the same player
        'and if any property in that set is mortgaged
        'Reterns true if yes, False if none
Dim i, OwnerNumb, SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s
SetNo = Prop.Fields("Set")
MortgInSetSet = False
Prop.MoveFirst

Do Until Prop.EOF   'Check all properties
    If Prop.Fields("Set") = SetNo Then
        If Prop.Fields("Mortgaged") = True Then
            MortgInSetSet = True
            Exit Do
        End If
    End If
Prop.MoveNext
Loop
End Function

Public Sub MovePlayer(ByVal FinalSquare)
    'Moves players' counter forwards & updates database
Dim OldSqur, Move, s, i, n As Integer
Dim Start

ViewPlayer = CurPlayer

If CurPlayer = 0 Then
    n = MsgBox("Please Click Options to Enter Players", vbCritical, "Players")
    Exit Sub
End If

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
OldSqur = Plyr.Fields("Square")
Plyr.Fields("Square") = FinalSquare
Plyr.Update

    If OldSqur < FinalSquare Then
    For s = OldSqur + 1 To FinalSquare
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.05   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    Else: For s = OldSqur To 40
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.05   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    For s = 1 To FinalSquare
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.05   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    End If
End Sub

Public Sub MovePlayerBack(FinalSquare)
    'Moves players' counter backwards & updates database
Dim OldSqur, Move, s, i, n As Integer
Dim Start

ViewPlayer = CurPlayer

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
OldSqur = Plyr.Fields("Square")
Plyr.Fields("Square") = FinalSquare
Plyr.Update

    If OldSqur > FinalSquare Then
    For s = OldSqur To FinalSquare Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.1   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    Else: For s = OldSqur To 1 Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.1   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    For s = 40 To FinalSquare Step -1
    PosX (s)
    PosY (s)
    Start = Timer   ' Set start time.
        Do While Timer < Start + 0.1   'Pause between squares
        DoEvents   ' Yield to other processes.
        Loop
    Call PositionPlayer(s)  'Moves Counter to square (s)
    Next s
    
    End If
End Sub

Public Sub PositionPlayer(ByVal s)
    'Moves counter to new square (s)
PosX (s)
PosY (s)
If s >= 1 And s <= 11 Or s >= 21 And s <= 31 Then
    FrmBoard.ImgCounter(CurPlayer).Move XPos + SqBShort / 3, YPos + Corner / 2
Else
    FrmBoard.ImgCounter(CurPlayer).Move XPos + Corner / 6, YPos + SqSShort / 6
End If
End Sub

Public Sub BuyHouse(ByVal Numb)
    'Update DataBase & board when a player buys a house
Dim HousesOwned As Integer

Prop.Index = "Number"
Prop.Seek "=", Numb
PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
Prop.Edit
Prop.Fields("HousesOwned") = Prop.Fields("HousesOwned") + 1
Prop.Update
Call DrawHouses(Numb, HousesOwned)
    'Draw houses (HousesOwned) on Square (Numb)
Call PlyrMoney(CurPlayer, -PropSet.Fields("HousePrice"))
    'Reduce players' money
Call PlyrMoney(99, PropSet.Fields("HousePrice"))
    'Increase Banks' money

End Sub

Public Sub SellHouse(Numb)
    'Update DataBase & board when a player sells a house
Dim HousesOwned As Integer

Prop.Index = "Number"
Prop.Seek "=", Numb
PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
Prop.Edit
Prop.Fields("HousesOwned") = Prop.Fields("HousesOwned") - 1
Prop.Update
Call DrawHouses(Numb, HousesOwned)
    'Draw houses (HousesOwned) on Square (Numb)
Call PlyrMoney(CurPlayer, PropSet.Fields("HousePrice") / 2)
    'Increase players' money
Call PlyrMoney(99, -PropSet.Fields("HousePrice") / 2)
    'Reduce Banks' money
End Sub

Public Sub Stations(ByVal SetNum)
    'Set rent for stations owned by Player according to number owned
Dim RentOwed As Currency: Dim Player, i, Count As Integer

Plyr.Index = "Number"
Plyr.MoveFirst
Do Until Plyr.EOF
    Player = Plyr.Fields("Number")
    Count = 0
    If Player <> 0 Then
        Prop.Index = "Number"
        Prop.MoveFirst
        Do Until Prop.EOF   'Count number owned by Player
            If Prop.Fields("Set") = SetNum And Prop.Fields("OwnerNo") = Player Then
                Count = Count + 1
            End If
        Prop.MoveNext
        Loop

        If Count > 0 Then
        Prop.MoveFirst
        Do Until Prop.EOF   'Update Rent
            If Prop.Fields("Set") = SetNum And Prop.Fields("OwnerNo") = Player Then
                Prop.Edit
                Prop.Fields("HousesOwned") = Count - 1
                Prop.Update
            End If
            Prop.MoveNext
        Loop
        End If
    End If
    Plyr.MoveNext
Loop
End Sub

Public Sub MissTurn(ByVal Numb)
    'Set number (Numb) of turns to be missed by curent player
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
Plyr.Fields("MissTurns") = Numb
Plyr.Update
End Sub

Public Sub TurnMissed(ByVal Miss)
    'Reduced number of turns to be missed by curent player by 1
Dim n As Integer

Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Plyr.Edit
Plyr.Fields("MissTurns") = Miss - 1
Plyr.Update
n = MsgBox(Plyr.Fields("Name") & " to Miss " & Miss - 1 _
    & " more turns", vbInformation, "Miss a Turn")
Call NextPlayer 'Turn missed, Move to next player
End Sub

Public Sub LowMoney()
    'Current player can't afford rent owed
Dim n As Integer: Dim PlayerMoney, TotAssets As Currency

Prop.Index = "Number"
PropSet.Index = "Number"
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
PlayerMoney = GetPlayerMoney(CurPlayer)
TotAssets = PlayerMoney

Do Until Prop.EOF   'Check if player bankrupt - Can't raise enough money by selling assets
    If Prop.Fields("OwnerNo") = CurPlayer Then
        If Prop.Fields("HousesOwned") > 0 Then
            PropSet.Seek "=", Prop.Fields("Set")
            TotAssets = TotAssets + Prop.Fields("HousesOwned") * (PropSet.Fields("HousePrice") / 2)
        End If
        If Prop.Fields("Mortgaged") = True Then _
            TotAssets = TotAssets + ((Prop.Fields("Price") / 2) * 1.1)
    TotAssets = TotAssets + Prop.Fields("Price")
    End If
Prop.MoveNext
Loop

If TotAssets < AmountOwed Then  'Player is bankrupt
    n = MsgBox(Plyr.Fields("Name") & " is BANKRUPT", vbExclamation)
    Call RemovePlayer(CurPlayer)    'Player leaves game
    If TotPlayers < 2 Then
        Plyr.MoveFirst
        Plyr.MoveNext
        n = MsgBox(Plyr.Fields("Name") & " WINS", vbInformation, "Winner")
    End If
    Exit Sub
Else
LowMon = True
Call Trading    'Go back to trade options to sell more assets
End If
End Sub

Public Sub RemovePlayer(ByVal Player As Integer)
    'Bankrupt player leaves game
    
Dim n, s, Recipient, PropertySet As Integer

Plyr.Index = "Number"
Plyr.Seek "=", Player
Prop.Index = "Number"
PropSet.Index = "Number"
s = Plyr.Fields("Square")
Prop.Seek "=", s
Recipient = Prop.Fields("OwnerNo")
CChest.Index = "Number"
Chnce.Index = "Number"

CChest.MoveFirst
Do Until CChest.EOF
    If CChest.Fields("Action") = "Get Out of " & Jail Then
        CChest.Edit
        CChest.Fields("Owner") = Recipient
        CChest.Update
    End If
    CChest.MoveNext
Loop

Chnce.MoveFirst
Do Until Chnce.EOF
    If Chnce.Fields("Action") = "Get Out of " & Jail Then
        Chnce.Edit
        Chnce.Fields("Owner") = Recipient
        Chnce.Update
    End If
    Chnce.MoveNext
Loop

Prop.MoveFirst
Do Until Prop.EOF   'Property transferred to player/Bank who is owed money
    If Prop.Fields("OwnerNo") = Player Then
        PropertySet = Prop.Fields("Set")
        PropSet.Seek "=", PropertySet
        Prop.Edit
        If Recipient = 99 Then
            Prop.Fields("Mortgaged") = False
            Prop.Fields("HousesOwned") = 0
        End If
        Prop.Fields("OwnerNo") = Recipient
        Prop.Update
    End If
    Prop.MoveNext
Loop
Call PlyrMoney(Recipient, (Plyr.Fields("Money")))
    'Bankrupt palyers money transferred to owed player
Call Stations(9)
Call Stations(10)
Plyr.Index = "Number"
Plyr.Seek "=", (Player)
Plyr.Delete
FrmBoard.ImgCounter(Player).Visible = False
FrmBoard.CboViewPlayer.RemoveItem (Player - 1)
FrmBoard.CboViewPlayer.Refresh
Plyr.MoveNext
If Plyr.Fields("Number") = 99 Then
    Plyr.MoveFirst
    Plyr.MoveNext
End If
n = Plyr.Fields("Number")
SetCurPlayer (n)

CurPlayer = GetCurPlayer
ViewPlayer = CurPlayer
TotPlayers = TotPlayers - 1
Dice2 = 7
FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"
Call NextPlayer
Call UpdateBoard
End Sub

Public Sub EndGame()    'End Programme
If MsgBox("Are You Sure you want to quit?", 36, "Quit?") = 6 Then
    If MsgBox("Do you want to save your game?", 36, "Save Game?") = 6 Then
    End
    End If
    Call DBClearPlayers 'Reset DataBase
    End
End If
End Sub

Public Sub Duration()   'Update Elapsed Time
Secs = Secs + 1
If Secs = 60 Then
    Secs = 0
    Mins = Mins + 1
End If
If Mins = 60 Then
    Mins = 0
    Hrs = Hrs + 1
End If

FrmBoard.LblDuration.Caption = Hrs & ":" & Mins & ":" & Secs
End Sub

Public Function Random(ByVal Numb) As Integer    'Produce random numbers
Dim n As Integer
Randomize
    Random = Int(Numb * Rnd + 1)
End Function

Public Sub EnterDice()
Dim Values As String

Values = InputBox("Please enter TWO 1 digit numbers", "Enter Dice Numbers")
Dice1 = Val(Left$(Values, 1))
Dice2 = Val(Right$(Values, 1))
End Sub

