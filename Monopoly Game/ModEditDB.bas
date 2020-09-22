Attribute VB_Name = "ModEditDB"
Option Explicit

Public Sub EditDB()
Call CreateLists    'Create list of properties on Edit DataBase form
End Sub

Public Sub CreateLists()
Dim i As Integer
FrmEditDB.TxtDBPath = DBPath

Prop.Index = "Number"
Chnce.Index = "Number"
CChest.Index = "Number"
PropSet.Index = ("Number")

With FrmEditDB

For i = 0 To 11     'Clear property List boxes
    If i <> 4 And i <> 5 Then .LstProperty(i).Clear
Next i

Prop.MoveFirst
Do Until Prop.EOF   'Add all properties to property list boxes
    For i = 0 To 11
        If i <> 4 And i <> 5 And i <> 12 Then
            .LstProperty(i).AddItem Prop.Fields(i)
        End If
    Next i
Prop.MoveNext
Loop
.LblPropNo.Caption = .LstProperty(0).Text

For i = 0 To 3      'Clear Chance & Community Chest List Boxes
    .LstChance(i).Clear
    .LstCChest(i).Clear
Next i

Chnce.MoveFirst
Do Until Chnce.EOF  'Create Chance List Boxes
    For i = 0 To 3
       .LstChance(i).AddItem Chnce.Fields(i)
    Next i
    Chnce.MoveNext
Loop

CChest.MoveFirst
Do Until CChest.EOF  'Create Community Chest List Boxes
    For i = 0 To 3
        .LstCChest(i).AddItem CChest.Fields(i)
    Next i
CChest.MoveNext
Loop

'Create action combo boxes for Chance & Community Chest Cards
For i = 0 To 1
FrmEditDB.CboAction(i).Clear
FrmEditDB.CboAction(i).AddItem "Receive From " & Bank
FrmEditDB.CboAction(i).AddItem "Receive From All Players"
FrmEditDB.CboAction(i).AddItem "Pay To " & Bank
FrmEditDB.CboAction(i).AddItem "General Repairs"
FrmEditDB.CboAction(i).AddItem "Street Repairs"
FrmEditDB.CboAction(i).AddItem "Advance To"
FrmEditDB.CboAction(i).AddItem "Back To"
FrmEditDB.CboAction(i).AddItem "Go Back"
FrmEditDB.CboAction(i).AddItem "Go Forward"
FrmEditDB.CboAction(i).AddItem "Fine or Chance"
FrmEditDB.CboAction(i).AddItem "Goto " & Jail
FrmEditDB.CboAction(i).AddItem "Miss Turns"
Next i

PropSet.MoveFirst
Do Until PropSet.EOF    'Set colours in Set Colour options
    PropSet.MoveNext
    i = PropSet.Fields("Number")
    If i > 8 Then Exit Do
    .LblSet(i - 1).BackColor = Val(PropSet.Fields("Colour"))
    .TxtHPrice(i - 1).Text = "£" & PropSet.Fields("HousePrice")
Loop

.TxtNamesName(0).Text = House
.TxtNamesName(1).Text = Hotel
.TxtNamesName(2).Text = Go
.TxtNamesName(3).Text = Jail
.TxtNamesName(4).Text = Bank
.TxtNamesName(5).Text = PropInfo
.TxtNamesName(6).Text = Rent
.TxtNamesName(7).Text = Utility
.TxtNamesName(8).Text = Station

End With
End Sub

Public Sub PropertySelected(ByVal Clicked)
    'Select all fields for chosen property
Dim i As Integer
For i = 0 To 11
    If i <> Clicked And i <> 4 And i <> 5 Then _
        FrmEditDB.LstProperty(i).ListIndex = FrmEditDB.LstProperty(Clicked).ListIndex
Next i

FrmEditDB.LblPropNo.Caption = FrmEditDB.LstProperty(0).Text
FrmEditDB.TxtName.Text = FrmEditDB.LstProperty(1).Text
FrmEditDB.CboSet.Text = FrmEditDB.LstProperty(2).Text
FrmEditDB.TxtPrice.Text = FrmEditDB.LstProperty(3).Text
For i = 6 To 11
    FrmEditDB.TxtRent(i).Text = FrmEditDB.LstProperty(i).Text
Next i
End Sub

Public Sub ChanceSelected(ByVal Clicked)
    'Select all fields for chosen Chance card
Dim i As Integer
For i = 0 To 3
    If i <> Clicked Then _
        FrmEditDB.LstChance(i).ListIndex = FrmEditDB.LstChance(Clicked).ListIndex
Next i
FrmEditDB.LblCardNo(0).Caption = FrmEditDB.LstChance(0).Text
FrmEditDB.TxtText(0).Text = FrmEditDB.LstChance(1).Text
FrmEditDB.CboAction(0).Text = FrmEditDB.LstChance(2).Text
FrmEditDB.TxtAmount(0).Text = FrmEditDB.LstChance(3).Text
End Sub

Public Sub CChestSelected(ByVal Clicked)
    'Select all fields for chosen Community Chest
Dim i As Integer
For i = 0 To 3
    If i <> Clicked Then _
        FrmEditDB.LstCChest(i).ListIndex = FrmEditDB.LstCChest(Clicked).ListIndex
Next i
FrmEditDB.LblCardNo(1).Caption = FrmEditDB.LstCChest(0).Text
FrmEditDB.TxtText(1).Text = FrmEditDB.LstCChest(1).Text
FrmEditDB.CboAction(1).Text = FrmEditDB.LstCChest(2).Text
FrmEditDB.TxtAmount(1).Text = FrmEditDB.LstCChest(3).Text
End Sub

Public Sub UpdateDBProperty()
    'Update DataBase with new data
Dim i, n As Integer

Prop.Index = "Number"
With FrmEditDB
If .TxtName = "" Or .TxtPrice = "" Or .TxtRent(6) = "" Or _
    .TxtRent(7) = "" Or .TxtRent(8) = "" Or .TxtRent(9) = "" Or _
    .TxtRent(10) = "" Or .TxtRent(11) = "" Or .CboSet.Text = "" Then
    n = MsgBox("Please Enter a Value for all fields", vbCritical, "Empty Field")
    Exit Sub
End If

Prop.Seek "=", FrmEditDB.LstProperty(0).Text
Prop.Edit
Prop.Fields("Name") = .TxtName.Text
Prop.Fields("Set") = Val(.CboSet.Text)
If .CboSet.Text = "0" Then Prop.Fields("OwnerNo") = 0
If .CboSet.Text <> "0" Then Prop.Fields("OwnerNo") = 99
Prop.Fields("Price") = .TxtPrice.Text
For i = 6 To 11
    Prop.Fields(i) = Val(.TxtRent(i).Text)
Next i
Prop.Fields("Mortgaged") = False
Prop.Fields("HousesOwned") = 0
Prop.Update
End With
n = MsgBox("Property Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBChance()
    'Update DataBase with new data
Dim Numb, Amount, n As Integer
With FrmEditDB
Numb = .LstChance(0).Text
Amount = .TxtAmount(0).Text
Chnce.Index = "Number"

Chnce.Seek "=", Numb
Chnce.Edit
Chnce.Fields("Text") = .TxtText(0).Text
Chnce.Fields("Action") = .CboAction(0).Text
Chnce.Fields("Amount") = Amount
End With
Chnce.Update
n = MsgBox("Chance Cards Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBCChest()
    'Update DataBase with new data
Dim Numb, Amount, n As Integer
With FrmEditDB
Numb = .LstCChest(0).Text
Amount = .TxtAmount(1).Text
CChest.Index = "Number"

CChest.Seek "=", Numb
CChest.Edit
CChest.Fields("Text") = .TxtText(1).Text
CChest.Fields("Action") = .CboAction(1).Text
CChest.Fields("Amount") = Amount
End With
CChest.Update
n = MsgBox("Community Chest Cards Updated", vbInformation, "Updated")
End Sub

Public Sub UpdateDBCols()
    'Update DataBase with new data
PropSet.Index = "Number"
Dim i, n As Integer

For i = 0 To 7
    PropSet.Seek "=", i + 1
    PropSet.Edit
    PropSet.Fields("Colour") = FrmEditDB.LblSet(i).BackColor
    PropSet.Fields("HousePrice") = FrmEditDB.TxtHPrice(i).Text
    PropSet.Update
Next i
n = MsgBox("Set Info Updated", vbInformation, "Updated")
End Sub
