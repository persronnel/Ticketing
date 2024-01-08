Sub ChangeSubjectsInInbox()
    Dim inbox As Outlook.MAPIFolder
    Dim item As Object
    Dim conversation As Outlook.Conversation
    
    ' Get the Inbox folder
    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    ' Loop through all items in the Inbox
    For Each item In inbox.Items
        ' Check if the item is a mail item
        If TypeOf item Is Outlook.MailItem Then
            ' Check if the mail item is part of a conversation
            If Not item.Conversation Is Nothing Then
                ' Change the subject of the first email in the conversation
                item.Conversation.GetTable.GetNextRow.FieldAccessor.SetProperty "Subject", "New Subject"
                item.Save
            End If
        End If
    Next item
End Sub
