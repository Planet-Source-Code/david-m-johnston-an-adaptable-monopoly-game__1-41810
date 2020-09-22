Attribute VB_Name = "ModGameFunctions"
Option Explicit

Public Sub NamePriceClicked(ByVal Indx)
    'Show Deed or, if Jail clicked, option to use Get out of jail card
Dim Nme As String
Prop.Index = "Number"
If Indx = 11 Then
    Dim i As Integer

    Plyr.Index = "Number"
    Plyr.Seek "=", CurPlayer
    Chnce.Index = "Number"
    CChest.Index = "Number"

    For i = 41 To 42
        Prop.Seek "=", i
        Plyr.Edit
        Prop.Edit
        If Prop.Fields("OwnerNo") = CurPlayer Then  'If current player owns Get Out of Jail card
            If MsgBox("Are you Sure you want to use your " & vbCrLf & _
                "Get Out of " & Jail & " Free Card?", 36, "") = 6 Then
                Prop.Fields("OwnerNo") = 99
                Plyr.Fields("MissTurns") = 0
                Exit For
            End If
        End If
    Next i
Plyr.Update
Prop.Update
Else
    Prop.Seek "=", Indx
    Nme = Prop.Fields("Name")
    If Prop.Fields("OwnerNo") <> 0 Then Call Deed(Nme)  'Show Title Deed
End If
End Sub

Public Sub Deed(ByVal Name) 'Show Ttile Deed
Dim i, OwnersSet, OwnerNo, s As Integer
Dim Ctrl As Control: Dim HasSet, Houses, Motgd As Boolean

Prop.Index = "Name"
PropSet.Index = "Number"
Prop.Seek "=", Name
s = Prop.Fields("Number")
HasSet = SetOwned(s)
Houses = HousesOnSet(s)
Motgd = MortgInSetSet(s)
Prop.Index = "Name"
Prop.Seek "=", Name
If Prop.Fields("Set") = 0 Then Exit Sub
Plyr.Index = "Number"
Plyr.Seek "=", Prop.Fields("OwnerNo")
OwnerNo = Prop.Fields("OwnerNo")
PropSet.Seek "=", (Prop.Fields("Set"))
OwnersSet = Prop.Fields("Set")
With FrmProperty

.LblName.Caption = Prop.Fields("Name")

    If OwnersSet <> 9 Then  'Not a Company
        For i = 0 To 6
            .LblRentHouses(i).Visible = True
            .LblRent(i).Visible = True
        Next i
        .LblHPriceLab.Visible = True
        .LblHPrice.Visible = True
        .LblEachLab.Visible = True
        If OwnersSet < 9 And OwnerNo <> 99 And Houses = False And Motgd = False Then
        .LblRent(0) = "£" & Prop.Fields("Rent") * 2
        Else: .LblRent(0) = "£" & Prop.Fields("Rent")
        End If
        .LblRent(1) = "£" & Prop.Fields("Rent1")
        .LblRent(2) = "£" & Prop.Fields("Rent2")
        .LblRent(3) = "£" & Prop.Fields("Rent3")
        .LblSet9.Visible = False
    
    If OwnersSet <> 10 Then 'Not a Station
        .LblName.ForeColor = &HFFFFFF
        .LblName.BackColor = Val(PropSet.Fields("Colour"))
        If OwnersSet < 9 And HasSet = True And OwnerNo <> 99 And Houses = False And Motgd = False Then
        .LblRentHouses(0).Caption = Rent & " - Site only (Set Owned)"
        Else: .LblRentHouses(0).Caption = Rent & " - Site only"
        End If
        .LblRentHouses(1).Caption = "      " & Chr(34) & "      With 1 " & House
        .LblRentHouses(2).Caption = "      " & Chr(34) & "        " & "  " & Chr(34) & "  2 " & House
        .LblRentHouses(3).Caption = "      " & Chr(34) & "        " & "  " & Chr(34) & "  3 " & House
        .LblRentHouses(4).Caption = "      " & Chr(34) & "        " & "  " & Chr(34) & "  4 " & House
        .LblRentHouses(5).Caption = "      " & Chr(34) & "        " & "  " & Chr(34) & "  " & Hotel
        .LblRent(4) = "£" & Prop.Fields("Rent4")
        .LblRent(5) = "£" & Prop.Fields("Rent5")
        .LblHPrice.Caption = "£" & PropSet.Fields("HousePrice")

    Else
        .LblName.BackColor = &HFFFFFF
        .LblName.ForeColor = &H0&
        .LblRentHouses(0).Caption = "Rent"
        .LblRentHouses(1).Caption = "If 2 " & Station & "s are owned"
        .LblRentHouses(2).Caption = "If 3        " & Chr(34) & "       " & Chr(34) & "     " & Chr(34)
        .LblRentHouses(3).Caption = "If 4        " & Chr(34) & "       " & Chr(34) & "     " & Chr(34)
        .LblRentHouses(4).Visible = False
        .LblRentHouses(5).Visible = False
        .LblRent(4).Visible = False
        .LblRent(5).Visible = False
        .LblSet.Visible = False
        .LblHPriceLab.Visible = False
        .LblHPrice.Visible = False
        .LblEachLab.Visible = False
    End If

Else
    .LblName.BackColor = &HFFFFFF
    .LblName.ForeColor = &H0&
    For i = 0 To 5
        .LblRentHouses(i).Visible = False
        .LblRent(i).Visible = False
    Next i
    .LblSet.Visible = False
    .LblSet9.Caption = "If one " & Utility & " is owned rent is £" & Prop.Fields("Rent") & _
    " times amount shown on dice 1. If both " & Utility & "s are owned " & Rent & " is £" & _
    Prop.Fields("Rent1") & " times amount shown on dice 1."
    .LblSet9.Visible = True
    .LblHPriceLab.Visible = False
    .LblHPrice.Visible = False
    .LblEachLab.Visible = False
End If

.LblRent(6) = "£" & Prop.Fields("Price") / 2
.LblOwner.Caption = Plyr.Fields("Name")

    If Prop.Fields("Mortgaged") = True Then 'Pink text on grey background
                                                'if property mortgaged
        FrmProperty.BackColor = &HE0E0E0
        For Each Ctrl In FrmProperty.Controls
            If Ctrl.Name Like "Lbl*" And Ctrl.Name <> "LblName" Then
                Ctrl.ForeColor = &H8080FF
                Ctrl.BackColor = &HE0E0E0
            End If
        Next Ctrl
    Else
    FrmProperty.BackColor = &HFFFFFF
        For Each Ctrl In FrmProperty.Controls
            If Ctrl.Name Like "Lbl*" And Ctrl.Name <> "LblName" Then
                Ctrl.ForeColor = &H80000012
                Ctrl.BackColor = &HFFFFFF
            End If
        Next Ctrl
        If OwnerNo <> 99 Then .LblRentHouses(Prop.Fields("HousesOwned")).ForeColor = vbRed
        If OwnerNo <> 99 Then .LblRent(Prop.Fields("HousesOwned")).ForeColor = vbRed
    End If

.Show
End With

End Sub

Public Sub Cards(ByVal Action, ByVal Amount)
    'Receives Action & Amount from Chance or CommChest
    'Performs Action
        
Dim s, p, n As Integer: Dim Choice As String
Plyr.Index = "Number"
Prop.Index = "Number"
s = PlayerSquare(CurPlayer) 'Get square current player is on

Select Case Action

Case "Receive From Bank"
    Call PlyrMoney(CurPlayer, Amount)
    Call PlyrMoney(99, -Amount)
    Call EndTurn

Case "Receive From All Players"
    For p = 1 To TotPlayers
        If p <> CurPlayer Then
            Call PlyrMoney(p, -Amount)
            Call PlyrMoney(CurPlayer, Amount)
            Call EndTurn
        End If
    Next p
    
Case "Pay To Bank"
    Call PlyrMoney(CurPlayer, -Amount)
    Call PlyrMoney(99, Amount)
    Call EndTurn
    
Case "General Repairs"
Prop.MoveFirst
    Do Until Prop.EOF
        If Prop.Fields("OwnerNo") = CurPlayer And Prop.Fields("Set") < 9 Then
            If Prop.Fields("HousesOwned") = 5 Then      'Hotel
                Call PlyrMoney(CurPlayer, -Amount * 4)
                Call PlyrMoney(99, Amount * 4)
            Else                                        'Houses
            Call PlyrMoney(CurPlayer, -Amount * Prop.Fields("HousesOwned"))
            Call PlyrMoney(99, Amount * Prop.Fields("HousesOwned"))
            End If
        End If
    Prop.MoveNext
    Loop
    Call EndTurn

Case "Street Repairs"
    Prop.MoveFirst
    Do Until Prop.EOF
    If Prop.Fields("OwnerNo") = CurPlayer And Prop.Fields("Set") < 9 Then
        If Prop.Fields("HousesOwned") = 5 Then          'Hotel
            Call PlyrMoney(CurPlayer, -Amount * 3)
            Call PlyrMoney(99, Amount * 3)
        Else                                            'Houses
            Call PlyrMoney(CurPlayer, -(Amount * Prop.Fields("HousesOwned")))
            Call PlyrMoney(99, (Amount * Prop.Fields("HousesOwned")))
        End If
    End If
    Prop.MoveNext
    Loop
    Call EndTurn

Case "Advance To"
    If Amount < s Then  'Player gets £200 for passing "GO"
        Call PlyrMoney(CurPlayer, 200)
        Call PlyrMoney(99, -200)
    End If
    s = Amount
    Call MovePlayer(s)
    
Case "Back To"
    s = Amount
    Call MovePlayerBack(s)
    
Case "Go Back"
    s = s - Amount
    If s < 1 Then s = s + 40
    Call MovePlayerBack(s)
   
Case "Go Forward"
    s = s + Amount
        If s > 40 Then  '£200 for passing "GO"
            s = s - 40
            Call PlyrMoney(CurPlayer, 200)
            Call PlyrMoney(99, -200)
        End If
    Call MovePlayer(s)
    
Case "Fine or Chance"
    Choice = InputBox("Please Type 'F' for Fine or 'C' for Chance", "Fine or Chance", "C")
        If Choice = "F" Or Choice = "f" Then
            Call PlyrMoney(CurPlayer, -Amount)
            Call PlyrMoney(99, Amount)
            Call EndTurn
        ElseIf Choice = "C" Or Choice = "c" Then
            FrmBoard.LblComChest.Caption = ""
            Call Chance 'Chance Card
            Exit Sub
        Else
            n = MsgBox("Please type 'F' or 'C'", vbCritical, "Fine or Chance")
            Call Cards(Action, Amount)
        End If
        
Case "Goto " & Jail
    Plyr.Seek "=", CurPlayer
    Plyr.Edit
    Plyr.Fields("Square") = 11
    Plyr.Update
    PositionPlayer (11)
    Call Cards("Miss Turns", 3)
    Call EndTurn
    Dice2 = 7
    
Case "Miss Turns"
    Call MissTurn(Amount)
    Call EndTurn

End Select

End Sub

Public Sub CommChest()
Dim Action, Amount As Integer
Randomize
CChest.Index = "Number"
CChest.Seek "=", Random(16) 'Select Card at Random
Action = CChest.Fields("Action")
Amount = CChest.Fields("Amount")

If Action = "Get Out of " & Jail Then
    Prop.Index = "Number"
    Prop.Seek "=", 42
    If CChest.Fields("Owner") = 99 Then
        Prop.Edit
        Prop.Fields("OwnerNo") = CurPlayer
        Prop.Update
        FrmBoard.LblComChest.Caption = CChest.Fields("text")
        Call EndTurn
        Exit Sub
    Else
    Call CommChest  'If Get Out of Jail card held by another player
                        'select a different card
    Exit Sub
    End If
End If

FrmBoard.LblComChest.Caption = CChest.Fields("text")
Call Cards(Action, Amount)  'Perform action

End Sub

Public Sub Chance()
Dim Action, Amount As Integer
Randomize
Chnce.Index = "Number"
Chnce.Seek "=", Random(16) 'Select Card at Random
Action = Chnce.Fields("Action")
Amount = Chnce.Fields("Amount")

If Action = "Get Out of " & Jail Then
    Prop.Index = "Number"
    Prop.Seek "=", 41
    If Prop.Fields("OwnerNo") = 99 Then
        Prop.Edit
        Prop.Fields("OwnerNo") = CurPlayer
        Prop.Update
        FrmBoard.LblChance.Caption = Chnce.Fields("text")
        Call EndTurn
        Exit Sub
    Else
    Call Chance  'If Get Out of Jail card held by another player
                        'select a different card
    Exit Sub
    End If
End If

FrmBoard.LblChance.Caption = Chnce.Fields("text")
Call Cards(Action, Amount)  'Perform action

End Sub

Public Sub BuyProperty(ByVal s) 'Player buys a property
Dim i, n, Count, SetNo As Integer
Prop.Index = "Number"
Prop.Seek "=", s

SetNo = Prop.Fields("Set")

    If GetPlayerMoney(CurPlayer) - Prop.Fields("Price") < 0 Then    'Not enough money
        n = MsgBox("Sorry you only have £" & GetPlayerMoney(CurPlayer) & vbLf & _
            "You can't afford " & Prop.Fields("Name"), vbCritical, "Insufficient Funds")
        Exit Sub
    End If
Call PlyrMoney(CurPlayer, -Prop.Fields("Price"))    'Reduce players' money
Call PlyrMoney(99, Prop.Fields("Price"))    'Increase Banks' money
Prop.Edit
Prop.Fields("OwnerNo") = CurPlayer
Prop.Update

If SetNo = 10 Then Call Stations(10)
If SetNo = 9 Then Call Stations(9)

Call EndTurn
End Sub

Public Sub EndTurn()    'Complete players' turn

Dim s, PropOwner As Integer
Plyr.Index = "Number"
Plyr.Seek "=", CurPlayer
Prop.Index = "Number"
s = PlayerSquare(CurPlayer)
Prop.Seek "=", s
PropOwner = Prop.Fields("OwnerNo")
If GetPlayerMoney(CurPlayer) < AmountOwed Then Call LowMoney

Select Case s   'Update player's/Banks' money
Case 5, 39
    Call PlyrMoney(CurPlayer, -AmountOwed)
    Call PlyrMoney(99, AmountOwed)

Case Else
    If PropOwner <> 0 And PropOwner <> 99 And PropOwner <> CurPlayer And Prop.Fields("Mortgaged") = False Then
    If Prop.Fields("Set") = 9 Then
        Call PlyrMoney(CurPlayer, -AmountOwed * Dice1)
        Call PlyrMoney(PropOwner, AmountOwed * Dice1)
    Else
        Call PlyrMoney(CurPlayer, -AmountOwed)
        Call PlyrMoney(PropOwner, AmountOwed)
    End If
    End If
End Select
AmountOwed = 0
End Sub

Public Function GetCurPlayer()
Plyr.Index = "Number"
Plyr.MoveFirst
GetCurPlayer = 0
Do Until Plyr.EOF
    If Plyr.Fields("CurPlayer") = True Then
        GetCurPlayer = Plyr.Fields("Number")
        Exit Do
    End If
Plyr.MoveNext
Loop
If GetCurPlayer = 0 Then GetCurPlayer = 1
End Function

Public Sub SetCurPlayer(ByVal PlayerNo)
Plyr.Index = "Number"
Plyr.MoveFirst
Do Until Plyr.EOF
Plyr.Edit
    If Plyr.Fields("Number") = PlayerNo Then
    Plyr.Fields("CurPlayer") = True
    Else
    Plyr.Fields("CurPlayer") = False
    End If
Plyr.Update
Plyr.MoveNext
Loop
End Sub
