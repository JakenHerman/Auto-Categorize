
Public Sub autocategorize()
    Set Items = Session.GetDefaultFolder(olFolderInbox).Items
    
    Dim sn As Outlook.MailItem
    
    Set sn = Application.ActiveInspector.CurrentItem
    If sn.SenderName = "Keith Ahee" Then
        sn.Categories = "Work;" & myItem.Categories
    End If
    
End Sub
