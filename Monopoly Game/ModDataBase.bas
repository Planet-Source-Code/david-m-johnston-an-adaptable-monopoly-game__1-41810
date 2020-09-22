Attribute VB_Name = "ModDataBase"
Option Explicit

Private WrkJet As Workspace

Sub LoadDatabase(ByVal PathToDatabase As String)    'Load the DataBase
    On Error GoTo LoadError
    Set WrkJet = CreateWorkspace("", "admin", "")
    Set DB = WrkJet.OpenDatabase(PathToDatabase)
    Exit Sub
LoadError:
    MsgBox Err.Description, vbCritical, Err.Number
    End
    
End Sub

Public Function SetRecordSets()     'Set RecordSet Variables

Set Prop = DB.OpenRecordset("Property")
Set PropSet = DB.OpenRecordset("PropertySet")
Set Plyr = DB.OpenRecordset("Player")
Set Counter = DB.OpenRecordset("Counter")
Set Chnce = DB.OpenRecordset("Chance")
Set CChest = DB.OpenRecordset("Com_Chest")
Set Vers = DB.OpenRecordset("Version")

End Function

Public Sub DBAddPlayer(ByVal CounterNumb)       'Add a Player to DataBase
On Error GoTo AddError
Counter.Index = "Number"
Counter.Seek "=", CounterNumb

With FrmPlayers
Plyr.AddNew
Plyr.Fields("Number") = PlyrAdd
Plyr.Fields("Name") = .TxtPlayerName.Text
Plyr.Fields("CounterPath") = Counter.Fields("FilePath")
Plyr.Fields("Square") = 1
Plyr.Fields("Money") = 1500
Plyr.Update
Exit Sub
End With

AddError:
    MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Public Sub DBClearPlayers()         'Re-Set DataBase
Dim i As Integer
Plyr.Index = "Number"
Prop.Index = "Number"
Chnce.Index = "Number"
CChest.Index = "Number"
Counter.Index = "Number"

On Error GoTo DeleteError

Prop.MoveFirst
Do Until Prop.EOF       'Go through all Records in Properties Table
    Prop.Edit
    Select Case Prop.Fields("Number")
        Case 1, 3, 8, 11, 18, 21, 23, 31, 34, 37
        Prop.Fields("OwnerNo") = 0      'Set Owner of non-property squares (Go, tax etc.)
                                         'to Player 0
        Case Else
        Prop.Fields("OwnerNo") = 99     'Set Owner of Properties to Bank
        Prop.Fields("HousesOwned") = 0  'Set Number of Houses Owned to 0
     Prop.Fields("Mortgaged") = False   'Set all properties to unmortgaged
    End Select
Prop.Update
Prop.MoveNext
Loop

Plyr.Seek "=", 99           'Set Banks' money to £13080
Plyr.Edit
Plyr.Fields("Money") = 13080
Plyr.Update

Chnce.MoveFirst
Do Until Chnce.EOF      'Go through all Chance cards
    Chnce.Edit
    Chnce.Fields("Owner") = 99  'Set Owner to 0
    Chnce.Update
    Chnce.MoveNext
Loop

CChest.MoveFirst
Do Until CChest.EOF      'Go through all Community Chest cards
    CChest.Edit
    CChest.Fields("Owner") = 99  'Set Owner to 0
    CChest.Update
    CChest.MoveNext
Loop

Plyr.MoveFirst
Do Until Plyr.EOF        'Go through all records in Players Table
    If Plyr.Fields("Number") <> 0 And Plyr.Fields("Number") <> 99 Then
        Plyr.Delete
    End If
Plyr.MoveNext
Loop

Exit Sub
DeleteError:
    MsgBox Err.Description, vbExclamation, Err.Number
End Sub

Public Function GetPlayerMoney(ByVal Player) As Currency
    'Receives Player number & Returns Money Owned
Plyr.Index = "Number"
Plyr.Seek "=", Player
GetPlayerMoney = Plyr.Fields("Money")
End Function

Public Function PlyrMoney(ByVal Player, ByVal Cash)
    'Receives Player Number & Money to be added (-ve amount if to be deducted)
        'and updates DataBase
Plyr.Index = "Number"
Plyr.Seek "=", Player
Plyr.Edit
Plyr.Fields("Money") = Plyr.Fields("Money") + Cash
Plyr.Update
End Function

Public Function PlayerSquare(ByVal Player) As Integer
    'Receives Player Number & Returns the square they are on
Plyr.Index = "Number"
Plyr.Seek "=", Player
PlayerSquare = Plyr.Fields("Square")
End Function

