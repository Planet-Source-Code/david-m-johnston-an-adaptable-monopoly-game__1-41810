Attribute VB_Name = "ModGame"
Option Explicit

Public Sub Turn()   'Player takes turn - Has clicked "Roll Dice"
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer

Dim s, NewSqr, n, Miss, PropOwner, HOwned As Integer
Dim PlayerMoney As Currency
Dim OwnerName As String
LowMon = False

AmountOwed = 0
ViewPlayer = CurPlayer
Prop.Index = "Number"
s = PlayerSquare(CurPlayer)

Miss = Plyr.Fields("MissTurns")

If Miss > 0 And Plyr.Fields("Square") <> 11 Then    'Player missing a turn
    Call TurnMissed(Miss)                               'but not in jail
    Call EndTurn
Else

Prop.Seek "=", s
s = s + Dice1 + Dice2   'New square = old square + dice
FrmBoard.LblDice1.Caption = Dice1
FrmBoard.LblDice2.Caption = Dice2
FrmBoard.LblChance.Caption = ""
FrmBoard.LblComChest.Caption = ""

If Plyr.Fields("Square") = 11 And Miss > 0 Then 'In jail & to miss turn
    If Dice1 <> Dice2 Then      'Didn't shake a double
        Call TurnMissed(Miss)   'Reduce turns to miss by 1
        Exit Sub
    End If
    Plyr.Edit
    Plyr.Fields("MissTurns") = 0
    Plyr.Update
    n = MsgBox("You got a double, You can leave " & Jail, vbExclamation, "Leave " & Jail)
End If

If s > 40 Then      'Get £200 for passing "GO"
    s = s - 40
    Call PlyrMoney(CurPlayer, 200)
    Call PlyrMoney(99, -200)
End If
Call MovePlayer(s)

Select Case s
Case 8, 23, 37, 3, 18, 34

        Select Case s
            Case 8, 23, 37: Call Chance
            Case 3, 18, 34: Call CommChest
        End Select
        NewSqr = PlayerSquare(CurPlayer)
        If NewSqr = s Then
            Call EndTurn
            If Dice1 <> Dice2 Then Call NextPlayer
            Exit Sub
        End If
        s = NewSqr
End Select

PlayerMoney = GetPlayerMoney(CurPlayer) 'Update PlayerMoney after Chance,Com. Chest Cards
Call UpdateBoard
Prop.Seek "=", s

Select Case s
Case 5        'Income Tax
    AmountOwed = Prop.Fields("Rent")
    If PlayerMoney < AmountOwed Then
        n = MsgBox("You can't afford this tax" & vbLf & _
        "You must sell some property to raise £" & _
        AmountOwed - PlayerMoney, vbCritical, "Insufficient Funds")
        Call LowMoney
    Else
    Call EndTurn
    End If
    
Case 39       'Super Tax
    AmountOwed = Prop.Fields("Rent")
    If PlayerMoney < AmountOwed Then
        n = MsgBox("You can't afford this tax" & vbLf & _
        "You must sell some property to raise £" & _
        AmountOwed - PlayerMoney, vbCritical, "Insufficient Funds")
        Call LowMoney
    Else
    Call EndTurn
    End If
    
Case 31     'Go to Jail
    s = 11
    Plyr.Seek "=", CurPlayer
    Plyr.Edit
    Plyr.Fields("Square") = s
    Plyr.Update
    Call PositionPlayer(s)
    Call MissTurn(3)
    Call EndTurn
    Dice2 = 7       'Won't be a double at end of turn
    
Case 1, 11, 21
    If Dice1 <> Dice2 Then Call EndTurn
 
Case Else   'Set Rent Owed
    PropOwner = Prop.Fields("OwnerNo")
    Plyr.Seek "=", PropOwner
    HOwned = Prop.Fields("HousesOwned")
    AmountOwed = Prop.Fields(HOwned + 6)
    If Prop.Fields("Mortgaged") = True Then AmountOwed = 0
    If SetOwned(s) = True And HousesOnSet(s) = False And MortgInSetSet(s) = False Then AmountOwed = AmountOwed * 2
    Prop.Seek "=", s
    OwnerName = Plyr.Fields("Name")
    
    If PropOwner = 99 Then  'Unsold property
        If MsgBox("Would You Like to Buy " & vbLf & _
            Prop.Fields("Name") & vbLf & "For £" & _
            Prop.Fields("Price"), 36, "Buy?") = 6 Then
        Call BuyProperty(s)
    Else: Call EndTurn
        End If
    
    ElseIf Prop.Fields("Set") = 9 And PropOwner <> CurPlayer Then   'Company owned by another player
        n = MsgBox("You have landed on " & Prop.Fields("Name") & vbLf & _
        "Which is owned by " & OwnerName & vbLf & _
        "Pay £" & AmountOwed & " * Dice 1" & vbLf & _
        "£" & AmountOwed * Dice1, vbExclamation, "Pay " & Rent)
        
        If PlayerMoney < AmountOwed Then    'Can't afford Rent
            n = MsgBox("You can't afford this " & Rent & vbLf & _
                "You must sell some property to raise £" & _
                AmountOwed - PlayerMoney, vbCritical, "Insufficient Funds")
            Call LowMoney
        Else: Call EndTurn
        End If

    ElseIf PropOwner <> 0 And PropOwner <> 99 And _
        PropOwner <> CurPlayer And Prop.Fields("Mortgaged") = False Then
            'Pay Rent
        n = MsgBox("You have landed on " & Prop.Fields("Name") & vbLf & _
        "Which is owned by " & OwnerName & vbLf & _
        "Pay £" & AmountOwed, vbExclamation, "Pay " & Rent)
        
        If PlayerMoney < AmountOwed Then    'Cant afford Rent
            n = MsgBox("You can't afford this " & Rent & vbLf & _
                "You must sell some property to raise £" & _
                AmountOwed - PlayerMoney, vbCritical, "Insufficient Funds")
            Call LowMoney
        Else: Call EndTurn
        End If
    End If
   
End Select
End If
Call NextPlayer

End Sub

Public Sub NextPlayer() 'Move to next player
If CurPlayer <> 0 Then
    Plyr.Index = "Number"
    Plyr.Seek "=", CurPlayer

    If Dice1 = Dice2 Or AmountOwed > GetPlayerMoney(CurPlayer) Then Exit Sub
End If
Plyr.MoveNext
If Plyr.Fields("Number") = 99 Then
    Plyr.MoveFirst
    Plyr.MoveNext
End If
CurPlayer = Plyr.Fields("Number")
SetCurPlayer (CurPlayer)
ViewPlayer = CurPlayer
Call UpdateBoard    'Show new players property & Money
Plyr.Seek "=", CurPlayer
FrmBoard.LblInfo.Caption = Plyr.Fields("Name") & " To Go"

End Sub

