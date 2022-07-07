Attribute VB_Name = "GenericModuleName"
Option Explicit

'sub that cycles through mails and folders
Sub accTopFolder()

Dim oAccount As Account
Dim ns As NameSpace
Dim fldr As Folder
Dim item As Object
Dim inbx As Folder

Set ns = GetNamespace("MAPI")

For Each oAccount In Session.Accounts

    Debug.Print vbCr & "oAccount: " & oAccount
    ' Shows all accounts
    For Each fldr In ns.Folders
    ' Shows all the names so you can replace "test"
        Debug.Print " top folder: " & fldr.Name
        If fldr = "test" Then
            Set inbx = fldr.Folders("Inbox")
            'inbx.Display
            For Each item In inbx.Items
                Debug.Print "  item .Subject: " & item.Subject
            Next
            Exit For
        End If
    Next
Next

Set inbx = Nothing
Set ns = Nothing

End Sub

'-------------------------------------------------------------------------------------------------------------------------------------
'sub that connects to OracleDB
'Must include following libraries in Outlook VBA (Tools -> References):
'   - Microsoft ActiveX Data Objects 2.8 Library
'   - Microsoft ActiveX Data Objects Recordset 2.8 Library
Sub TestOracleDBAConnection()

    Dim strConnection As String
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
                strConnection = "Provider=OraOLEDB.Oracle;Data Source=xxx.xxx.xxx:xxxx/xxx;User ID=myUserName;Password=myPassword;"
    
    cn.Open (strConnection)
    
    If cn.State = adStateOpen Then
     cn.Close
     MsgBox "Completed!"
    Else
     MsgBox "Connection failed!"
    End If
   
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------
'Progress Bar module; use VBA Progress Bar User Form.png for guidelines
Dim progress As Double, maxProgress As Double, maxWidth As Long, startTime As Double
Public Sub Initialize(title As String, Optional max As Long = 100)
'Initialize and shor progress bar
    Me.Caption = title
    maxProgress = max:  maxWidth = lBar.Width:    lBar.Width = 0
    lProgress.Caption = "0%"
    Me.Show False
    startTime = Time
End Sub
Public Sub AddProgress(Optional inc As Long = 1)
'Increase progress by an increment
    Dim tl As Double, tlMin As Long, tlSec As Long, tlHour As Long, tlTotal As Long, tlTotalSec, tlTotalMin, tlTotalHour
    progress = progress + inc
    If progress > maxProgress Then progress = maxProgress
    lBar.Width = CLng(CDbl(progress) / maxProgress * maxWidth)
    DoEvents
    tl = Time - startTime
    tlSec = Second(tl) + Minute(tl) * 60 + Hour(tl) * 3600
    tlTotal = tlSec
    If progress = 0 Then
        tlSec = 0
    Else
        tlSec = (tlSec / progress) * (maxProgress - progress)
    End If
    tlHour = Floor(tlSec / 3600)
    tlTotalHour = Floor(tlTotal / 3600)
    tlSec = tlSec - 3600 * tlHour
    tlTotal = tlTotal - 3600 * tlTotalHour
    tlMin = Floor(tlSec / 60)
    tlTotalMin = Floor(tlTotal / 60)
    tlSec = tlSec - 60 * tlMin
    tlTotal = tlTotal - 60 * tlTotalMin
    If tlSec > 0 Then
        tlMin = tlMin + 1
    End If
    'Captions
    lProgress.Caption = "" & CLng(CDbl(progress) / maxProgress * 100) & "%"
    lTimeLeft.Caption = "" & tlHour & " hours, " & tlMin & " minutes"
    lTimePassed.Caption = "" & tlTotalHour & " hours, " & tlTotalMin & " minutes, " & tlTotal & " seconds"
    'Hide if finished
    If progress = maxProgress Then Me.Hide
End Sub
Public Function Floor(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    Floor = Int(x / Factor) * Factor
End Function

'*********************************ProgressBar example***************************************************
'Declare and Initialize the ProgressBar UserForm
Dim pb As ProgressBar
Set pb = New ProgressBar
 
'Set the Title for the ProgressBar and the Maximum Value making the UserForm visible
pb.Initialize "My title", 100
 
'Use the ProgressBar to track macro execution by running the For loop 100 times
For i = 0 to 99 
    pb.AddProgress 1 'Add 1% progress
    '...
Next i
 
'Clean-up: Hide the ProgressBar
pb.Hide
'Free up memory
Set pb = Nothing

