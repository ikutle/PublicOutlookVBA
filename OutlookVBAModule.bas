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

