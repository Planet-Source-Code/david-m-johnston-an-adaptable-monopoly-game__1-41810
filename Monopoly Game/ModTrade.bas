Attribute VB_Name = "ModTrade"
Option Explicit

Public Sub Trading()    'Trade options selected
Dim i As Integer

Prop.Index = "Number"
FrmTrade.Show
FrmTrade.LstPlayerProp.Clear

For i = 1 To 42 'List all the current players' property
    Prop.Seek "=", i
    If Prop.Fields("OwnerNo") = CurPlayer Then _
        FrmTrade.LstPlayerProp.AddItem Prop.Fields("Name")
Next i
Call EnableTrade(False) 'disable trade options until a property is selected
End Sub

Public Sub EnableTrade(ByVal ED As Boolean)
    'Enable/disable trade options
Dim Ctrl As Control

If ED = True Then
    For Each Ctrl In FrmTrade.Controls
        If Ctrl.Name Like "LbC*" Then
            Ctrl.Enabled = True
            Ctrl.ForeColor = &H80000012
        End If
        If Ctrl.Name Like "Mnu*" Then Ctrl.Enabled = True
    Next Ctrl
Else
    For Each Ctrl In FrmTrade.Controls
        If Ctrl.Name Like "LbC*" And Ctrl.Name <> "LbCFinished" Then
            Ctrl.Enabled = False
            Ctrl.ForeColor = &H8000000F
        End If
        If Ctrl.Name Like "Mnu*" Then
            If Ctrl.Name <> "MnuGame" And Ctrl.Name <> "MnuActions" _
            And Ctrl.Name <> "MnuFinished" Then Ctrl.Enabled = False
        End If
    Next Ctrl
End If
End Sub

Public Sub SelectPlayer()   'Select player to sell to
Dim i, n, SetNo, Sqre As Integer: Dim PropName As String

FrmTrade.LstPlayers.Visible = True
FrmTrade.LstPlayers.Clear
PropName = FrmTrade.LstPlayerProp.Text
Plyr.Index = "Number"
Prop.Index = "Name"
Prop.Seek "=", PropName
Sqre = Prop.Fields("Number")
SetNo = Prop.Fields("Set")

If Prop.Fields("Mortgaged") = True Then 'Property is mortgaged
    n = MsgBox("Sorry you can't sell mortgaged property", vbCritical)
    Exit Sub
End If

'Set contains houses
If SetOwned(Sqre) = True Then
    If HousesOnSet(Sqre) = True And SetNo < 9 Then
        n = MsgBox("Sorry you must sell any " & House & "s/" & Hotel & "s on a set" & vbLf _
        & "before selling a property", vbCritical)
    Exit Sub
    End If
End If

Plyr.Move (1)
Do Until Plyr.EOF 'List all players & Bank except self
    If Plyr.Fields("Number") <> CurPlayer Then FrmTrade.LstPlayers.AddItem Plyr.Fields("Name")
    Plyr.MoveNext
Loop

FrmTrade.LbCSelectPlayer.ForeColor = &H8000000F
FrmTrade.MnuBuyer.Enabled = False
FrmTrade.LbCSelectPlayer.Enabled = False
FrmTrade.LbCSellProperty.ForeColor = &H80000012
FrmTrade.MnuSell.Enabled = True
FrmTrade.LbCSellProperty.Enabled = True
End Sub

Public Sub UpgradeProperty()    'Add houses/Hotels
Dim Name, PropName As String: Dim Money, HPrice As Currency
Dim n, Numb, SetNo, Buy, HousesOwned As Integer
Name = FrmTrade.LstPlayerProp.Text
Money = GetPlayerMoney(CurPlayer)
Buy = 7     '6= Yes, 7 = No

Prop.Index = "Name"
Prop.Seek "=", Name
PropSet.Index = "Number"

Numb = Prop.Fields("Number")
SetNo = Prop.Fields("Set")
PropSet.Seek "=", SetNo
HPrice = PropSet.Fields("HousePrice")
PropName = Prop.Fields("Name")
HousesOwned = Prop.Fields("HousesOwned")

If SetNo = 9 Or SetNo = 10 Or SetNo = 0 Then    'Companys,Stations,Non-propertys
n = MsgBox("Sorry you can't add " & House & "s to this property", vbInformation, "")
    
    ElseIf Name = "" Then n = MsgBox("Please select a property", vbCritical, "")
    
    'Can't affort house
    ElseIf Money < HPrice Then _
        n = MsgBox("Sorry you only have £" & Money & vbLf & _
        "You can't afford a " & House & " on " & Prop.Fields("Name"), vbInformation, "")
    
    'Mortgaged property
    ElseIf Prop.Fields("Mortgaged") = True Then n = MsgBox _
        ("Sorry, you can't upgrade mortgaged property", vbInformation, "")
    
    'Player doesn't own whole set
    ElseIf SetOwned(Numb) = False Then n = MsgBox("You Must Own The Whole Set " & vbCrLf & _
        "Bofore you can add a " & House, vbCritical, "")
    
    Else    'Option to buy
        Buy = MsgBox("A " & House & " on " & PropName & " costs £" & _
        HPrice & vbLf & "Do you want to buy Y/N", 36, "")
End If
        
If Buy = 6 Then 'Player wants to buy
    Select Case HousesOwned
        Case 0, 1, 2, 3, 4: Call BuyHouse(Numb)
        Case 5: MsgBox ("Sorry you already have a " & Hotel)
    End Select
    Call UpdateBoard
    Call UpdateHouses   'Re-Create houses
End If

End Sub

Public Sub SellProperty()   'Player selling a property
Dim Name, SellTo As String: Dim Price As Currency
Dim Buyer, n, SetNo As Integer

Name = FrmTrade.LstPlayerProp.Text
SellTo = FrmTrade.LstPlayers.Text
Plyr.Index = "Name"
Plyr.Seek "=", SellTo
Prop.Index = "Name"
Prop.Seek "=", Name
Prop.Edit
SetNo = Prop.Fields("Set")
Price = Prop.Fields("Price")
Buyer = Plyr.Fields("Number")

If GetPlayerMoney(Buyer) < Price Then
    n = MsgBox(SellTo & " can't afford " & Prop.Fields("Name") _
    & vbCrLf & "Please select another Player", vbCritical, "")
    Exit Sub
End If

Prop.Fields("OwnerNo") = Buyer
Prop.Update
n = MsgBox(Prop.Fields("Name") & " Sold", vbInformation, "SOLD")
Call PlyrMoney(CurPlayer, Price)
Call PlyrMoney(Buyer, -Price)
If SetNo = 9 Then Call Stations(9)
If SetNo = 10 Then Call Stations(10)
Call Trading
End Sub

Public Sub SellHouses() 'Player sells house/hotel
Dim Name As String: Dim HousePrice As Currency
Dim Numb, SetNo, n, Sell As Integer
Name = FrmTrade.LstPlayerProp.Text

Prop.Index = "Name"
Prop.Seek "=", Name
SetNo = Prop.Fields("Set")
PropSet.Index = "Number"
PropSet.Seek "=", SetNo
If SetNo = 9 Or SetNo = 10 Or SetNo = 0 Then    'No Houses to sell
    MsgBox ("Sorry there are no " & House & "s to to sell")

    'Property not selected
    ElseIf Name = "" Then MsgBox "Please select a property"
    
    Else
    n = MsgBox("A house on " & Prop.Fields("Name") & " is worth £" & _
    PropSet.Fields("HousePrice") / 2 & vbLf & "Do you want to sell one Y/N", 36, "sell " & House)
        
        If n = 6 Then   'Player wants to sell
            If Prop.Fields("HousesOwned") = 0 Then  'No houses to sell
                n = MsgBox("You don't have any houses or hotels to sell", vbInformation, "Sell " & House)
            Exit Sub
            
            Else
                Numb = Prop.Fields("Number")
                Call SellHouse(Numb)    'Perform transaction
        End If
        Call UpdateBoard
        Call UpdateHouses
    End If
End If
End Sub

Public Sub Mortgage(ByVal Name) 'Mortgage a property
Dim NewValue, Money As Currency: Dim YN As String
Dim Mtgaged As Boolean: Dim SetNo, Sqre, n As Integer

Prop.Index = "Name"
Prop.Seek "=", Name
Money = GetPlayerMoney(CurPlayer)   'Money = money owned by player
Mtgaged = Prop.Fields("Mortgaged")
SetNo = Prop.Fields("Set")
Sqre = Prop.Fields("Number")

If Prop.Fields("Set") = 0 Then
    MsgBox ("This can't be mortgaged")
    Exit Sub
End If
If Mtgaged = False Then
    If MsgBox("Are you sure you want to MORTGAGE " & Name, 36, "Mortgage") _
        = 6 Then
            
        'Set contains houses
        If HousesOnSet(Sqre) = True And SetNo < 9 Then
            n = MsgBox("Sorry you must sell any " & House & "s/" & Hotel & "s on a set" & vbLf _
            & "before selling a property", vbCritical)
            Exit Sub
        End If
           
    'Mortgage
    Prop.Index = "Name"
    Prop.Seek "=", Name
    Prop.Edit
    Prop.Fields("Mortgaged") = True
    NewValue = Prop.Fields("Price") / 2
    Prop.Update
    Call PlyrMoney(CurPlayer, NewValue)
    Call PlyrMoney(99, -NewValue)
    Else
        Exit Sub
    End If
Else
    If Money < Money - (Prop.Fields("Price") / 2) * 1.1 Then _
        MsgBox ("Sorry you only have " & Money & vbLf & _
            "You can't afford to unmortgage " & Name)

        'Unmortgage
        If MsgBox("Are you sure you want to UNMORTGAGE " & Name, 36, _
        "Unmortgage") = 6 Then
            Prop.Index = "Name"
            Prop.Seek "=", Name
            Prop.Edit
            Prop.Fields("Mortgaged") = False
            NewValue = (Prop.Fields("Price") / 2) * 1.1
            Prop.Update
            Call PlyrMoney(CurPlayer, -NewValue)
            Call PlyrMoney(99, NewValue)
        End If
    End If
End Sub

Public Sub SellProp()   'Player sells a property
Dim n As Integer

If FrmTrade.LstPlayerProp.Text = "" Then Exit Sub
If FrmTrade.LstPlayers.Text = "" Then
    n = MsgBox("Please choose a player to sell to", vbCritical, "Sell To?")
    Exit Sub
End If
Call SellProperty   'Perform transaction
End Sub

Public Sub FinishedTrade() 'Player finished trading
Dim n As Integer
If GetPlayerMoney(CurPlayer) < AmountOwed Then  'Player still needs to raise more money
    n = MsgBox("You must sell some more property " & vbCrLf & _
        "To raise another £" & AmountOwed - GetPlayerMoney(CurPlayer), vbCritical, "")
    Call Trading
    Exit Sub
End If

FrmTrade.LstPlayers.Visible = False
Call UpdateBoard
FrmTrade.Hide
If LowMon = True Then
    Call EndTurn
    Call NextPlayer
End If
End Sub

