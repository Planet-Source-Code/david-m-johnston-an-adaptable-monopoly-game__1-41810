Attribute VB_Name = "ModDrawBoard"
Option Explicit

Public Sub ResCheck()
'Detects the resolution in use & Draws board to fill screen
'Moves controls acordingly

'Rescomp used to keep control proportions the same
Dim Ctrl As Object: Dim s, n As Integer

FWidth = Screen.Width   'Find width of screen
FHeight = Screen.Height - (FrmBoard.SBar.Height * 4)    'Used for board hieght
LowRes = 0  'Set to 600 if Resolution = 800 * 600

Select Case Screen.Width
Case Is <= 9600 'Programme won't run in lower than 800 * 600
n = MsgBox("Sorry this program can't run in this resolution", vbCritical, "Resolution")
End

Case 12000  'Default Resolution
ResComp = 1 'Compensation not needed
LowRes = 600    'Used to increase width of side squares on board
                    'to allow more space for property names

For s = 1 To 40 'Make property names more readable on low res monitor
    FrmBoard.LblName(s).FontSize = 7
    FrmBoard.LblPrice(s).FontSize = 7
Next s

Case 15360
ResComp = 1.28
For s = 1 To 40
    FrmBoard.LblName(s).Font = "arial"
    FrmBoard.LblPrice(s).Font = "arial"
    FrmBoard.LblName(s).FontSize = 8
    FrmBoard.LblPrice(s).FontSize = 8
Next s

Case Else
ResComp = 1.44
For s = 1 To 40
    FrmBoard.LblName(s).Font = "arial"
    FrmBoard.LblPrice(s).Font = "arial"
    FrmBoard.LblName(s).FontSize = 8
    FrmBoard.LblPrice(s).FontSize = 8
Next s

End Select

'Alter position & size of controls using rescomp so that board looks
    'the same in any resolution
For Each Ctrl In FrmBoard.Controls
If Ctrl.Name Like "Lb*" Or Ctrl.Name Like "Cmd*" Or Ctrl.Name Like "Lst*" Then
    Ctrl.Left = Ctrl.Left * ResComp
    Ctrl.Top = Ctrl.Top * ResComp
    Ctrl.Width = Ctrl.Width * ResComp
    Ctrl.Height = Ctrl.Height * ResComp
End If
Next Ctrl
FrmBoard.CboViewPlayer.Left = FrmBoard.CboViewPlayer.Left * ResComp
FrmBoard.CboViewPlayer.Top = FrmBoard.CboViewPlayer.Top * ResComp

End Sub

Sub DrawBoard()
Dim SetCol As String: Dim s As Integer

'set sizes of squares according to resolution in use
Corner = (FWidth / 13) * 1.5
SqBShort = (FWidth - (Corner * 2)) / 9
SqSShort = (FHeight - (Corner * 2)) / 9

ViewPlayer = 1
Prop.Index = "Number"
PropSet.Index = "Number"

For s = 1 To 40     'Go through all squares
PosX (s)    'Sets XPos (X Position) according to square being drawn
PosY (s)    'Sets YPos (Y Position) according to square being drawn
Prop.Seek "=", s    'Move to Property for Square being drawn
PropSet.Seek "=", Prop.Fields("Set")    'Move to property set for Property being drawn
SetCol = Val(PropSet.Fields("Colour"))  'Colour of property
FrmBoard.BackColor = BrdColour      'Set colour of board

With FrmBoard

Select Case s       's = Square

Case 1, 11, 21, 31      'Corners
FrmBoard.Line (XPos, YPos)-Step(Corner, Corner), BrdColour, BF

Case 1 To 11, 21 To 31  'Top & Bottom
FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), BrdColour, BF

Case 11 To 21, 31 To 40 'Sides
FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), BrdColour, BF

End Select

Select Case s
Case 2, 4, 7, 9, 10     'Sets on bottom of board
FrmBoard.Line (XPos, YPos + 20)-Step(SqBShort - 10, 200), SetCol, BF

Case 12, 14, 15, 17, 19, 20 'Sets on Left of board
FrmBoard.Line (Corner + LowRes - 220, YPos + 10)-Step(200, SqSShort - 20), SetCol, BF

Case 22, 24, 25, 27, 28, 30 'Sets at Top of board
FrmBoard.Line (XPos + 10, Corner - 220)-Step(SqBShort - 30, 200), SetCol, BF

Case 32, 33, 35, 38, 40     'Sets on right of board
FrmBoard.Line (XPos + 10, YPos + 10)-Step(210, SqSShort - 20), SetCol, BF

End Select

Select Case s

Case 1      '"GO" Square
.LblName(s).Move XPos, (YPos + (Corner / 3)), Corner, Corner / 3    'Square Name
FrmBoard.Line (XPos, YPos)-Step(Corner, Corner), , B    'Border

Case 11, 21, 31 'Other Corners
.LblName(s).Move XPos, (YPos + 600), SqBShort
FrmBoard.Line (XPos, YPos)-Step(Corner, Corner), , B

Case 2 To 10, 22 To 30  'Bottom & Top
    .LblName(s).Move XPos, YPos + 300, SqBShort, (Corner / 2)
    .LblPrice(s).Move XPos, (YPos + 650), SqBShort, (Corner / 2)
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
    
Case 12 To 20   'Left
    .LblName(s).Move XPos, (YPos + 10), (Corner + LowRes - 200)
    .LblPrice(s).Move XPos, (YPos + 350), (Corner + LowRes - 200)
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B

Case 32 To 40   'Right
    .LblName(s).Move XPos, (YPos + 10), (Corner + LowRes + 200)
    .LblPrice(s).Move XPos, (YPos + 350), (Corner + LowRes + 200)
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
End Select

.LblName(s).Caption = Prop.Fields("Name")   'Property Name
If Prop.Fields("Price") <> "0" Then _
    .LblPrice(s).Caption = "£" & Prop.Fields("Price")   'Property Price

If Prop.Fields("Set") > 0 And Prop.Fields("Set") < 9 Then   'Set Colour
    .LblName(s).ToolTipText = "Click Here to View " & PropInfo
    .LblPrice(s).ToolTipText = "Click Here to View " & PropInfo
End If

If s = 11 Then 'Jail Square
    .LblName(11).ToolTipText = "Clik Here to use Get Out of " & Jail & " Free Card"
End If

End With

Next s
Call ChangeCols
End Sub

Public Sub ChangeCols()     'Set Board & Text(Squares only) colours
Dim Ctrl As Control

FrmBoard.BackColor = BrdColour
For Each Ctrl In FrmBoard.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = TextColour
End If
Next Ctrl

FrmEditDB.LbCFinished.BackColor = BrdColour
FrmEditDB.BackColor = BrdColour

FrmOptions.BackColor = BrdColour
For Each Ctrl In FrmOptions.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = TextColour
End If
Next Ctrl

FrmPlayers.BackColor = BrdColour
For Each Ctrl In FrmPlayers.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = TextColour
End If
Next Ctrl

FrmTrade.BackColor = BrdColour
For Each Ctrl In FrmTrade.Controls
If Ctrl.Name Like "Lb*" Then
    Ctrl.BackColor = BrdColour
    Ctrl.ForeColor = TextColour
End If
Next Ctrl

With FrmBoard   'GO Square doesn't change
.LblName(1).FontSize = 25
.LblName(1).ForeColor = &HFF&
.LblName(1).Alignment = 2
.LblChance.BackColor = &H80C0FF
.LblComChest.BackColor = &HFFD0FF
End With
End Sub

Public Sub UpdateBoard()    'Update Property Lists & Current Player
Dim Square As Integer

Plyr.Index = "Number"
Plyr.Seek "=", ViewPlayer
Prop.Index = "Number"
Square = 1

Prop.MoveFirst
Do Until Prop.EOF
If Prop.Fields("Set") <> 0 Then
Square = Prop.Fields("Number")
    If Prop.Fields("Mortgaged") = True Then 'Pink text on grey background
                                                'if property mortgaged
        FrmBoard.LblName(Square).ForeColor = &H8080FF
        FrmBoard.LblPrice(Square).ForeColor = &H8080FF
    Else
        FrmBoard.LblName(Square).ForeColor = TextColour
        FrmBoard.LblPrice(Square).ForeColor = TextColour
    End If
End If
Prop.MoveNext
Loop

Call BankProperty   'Update list of property in bank
PlayerProperty (ViewPlayer) 'Update list of property held by player being viewed
With FrmBoard
.LblOwner.Caption = Plyr.Fields("Name")
.CboViewPlayer.Text = Plyr.Fields("Name")
.LblMoney = "£" & GetPlayerMoney(CurPlayer) 'Money Owned by vurrent player
.LblBankMoney.Caption = "£" & GetPlayerMoney(99)    'Money in Bank
End With

End Sub

Public Sub ShowCounters()
Dim i As Integer

For i = 1 To TotPlayers
    FrmBoard.ImgCounter(i).Move XPos, YPos + Corner / 2
    FrmBoard.ImgCounter(i).Visible = True
    XPos = XPos + (Corner / TotPlayers)
Next i
End Sub

Public Sub UpdateHouses()   'Re-Draw Houses/Hotels
Dim i, Houses, PropSet As Integer

Prop.Index = "Number"
For i = 1 To 40
Prop.Seek "=", i
PropSet = Prop.Fields("Set")
Houses = Prop.Fields("HousesOwned")
If Houses > 0 And PropSet > 0 And PropSet < 9 Then Call DrawHouses(i, Houses) 'Draw Houses
Next i
End Sub

Public Sub ClearHouses(ByVal Numb)
'Remove Houses/Hotels from Square (Numb) at position XPos,YPos
Dim SetCol As String

PropSet.Index = "Number"
PropSet.Seek "=", Prop.Fields("Set")
SetCol = Val(PropSet.Fields("Colour"))  'Set Colour
PosX (Numb)
PosY (Numb)
Select Case Numb
Case 2 To 10
    FrmBoard.Line (XPos, YPos + 20)-Step(SqBShort - 10, 200), SetCol, BF
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
Case 22 To 30:
    FrmBoard.Line (XPos + 10, Corner - 220)-Step(SqBShort - 30, 200), SetCol, BF
    FrmBoard.Line (XPos, YPos)-Step(SqBShort, Corner), , B
Case 12 To 20:
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
    FrmBoard.Line (Corner + LowRes - 220, YPos + 10)-Step(200, SqSShort - 20), SetCol, BF
Case 32 To 40:
    FrmBoard.Line (XPos, YPos)-Step(Corner + LowRes, SqSShort), , B
    FrmBoard.Line (XPos + 10, YPos + 10)-Step(210, SqSShort - 20), SetCol, BF
End Select

End Sub

Public Sub DrawHouses(ByVal Numb, ByVal HousesOwned)
'Draw HousesOwned Houses/Hotels on Square Numb
Dim i As Integer

Call ClearHouses(Numb)  'Remove Houses alredy on square Numb
PosX (Numb)
PosY (Numb)

For i = 1 To HousesOwned
Select Case Numb

Case 2 To 10    'Bottom
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + ((SqBShort / 2) - SqBShort / 4)), YPos + 40)-Step(SqBShort / 2, 150), &H80&, BF
        Exit Sub
    End If
    If i = 1 Then XPos = XPos + (SqBShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + ((SqBShort / 6) + 40) * (i - 1)), YPos + 40)-Step(SqBShort / 6, 150), &HFF00&, BF

Case 22 To 30   'Top
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + ((SqBShort / 2) - SqBShort / 4)), YPos + Corner - 190)-Step(SqBShort / 2, 150), &H80&, BF
        Exit Sub
    End If
    If i = 1 Then XPos = XPos + (SqBShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + ((SqBShort / 6) + 40) * (i - 1)), YPos + Corner - 190)-Step(SqBShort / 6, 150), &HFF00&, BF

Case 12 To 20   'Left
    If HousesOwned = 5 Then
        FrmBoard.Line ((XPos + LowRes + Corner - 190), YPos + (SqSShort / 2) - (SqSShort / 4))-Step(150, SqSShort / 2), &H80&, BF
        Exit Sub
    End If
    If i = 1 Then YPos = YPos + (SqSShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line ((XPos + LowRes + Corner - 200), YPos + ((SqSShort / 6) + 40) * (i - 1))-Step(150, SqSShort / 6), &HFF00&, BF

Case 32 To 40   'Right
    If HousesOwned = 5 Then
    FrmBoard.Line ((XPos + 40), YPos + ((SqSShort / 2) - SqSShort / 4))-Step(150, SqSShort / 2), &H80&, BF
    Exit Sub
    End If
    If i = 1 Then YPos = YPos + (SqSShort / 12) + 20
    If i > 0 And i < 5 Then FrmBoard.Line (XPos + 40, YPos + ((SqSShort / 6) + 40) * (i - 1))-Step(150, SqSShort / 6), &HFF00&, BF

End Select

Next i
End Sub

Public Sub BankProperty()   'Clear & Re-Create list of Properties in Bank
Dim i As Integer

FrmBoard.LstBankProp.Clear
Prop.MoveFirst
Do Until Prop.EOF   'Go throug all properties
    If Prop.Fields("OwnerNo") = 99 And Prop.Fields("Set") <> 0 Then _
        FrmBoard.LstBankProp.AddItem Prop.Fields("Name")
Prop.MoveNext
Loop

End Sub

Public Function PlayerProperty(ByVal Player)
'Clear & Re-Create list of all properties held by Player
Dim i As Integer

FrmBoard.LstPlayerProp.Clear

Prop.MoveFirst
Do Until Prop.EOF   'Go throug all properties
    If Prop.Fields("OwnerNo") = Player Then _
    FrmBoard.LstPlayerProp.AddItem Prop.Fields("Name")
Prop.MoveNext
Loop
End Function
