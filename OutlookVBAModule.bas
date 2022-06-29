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

