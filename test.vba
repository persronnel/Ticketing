Sub UpdateSubjectInConversation()
    Dim ActWind As Object
    Dim MyItem As Object
    Dim Conversation As Object
    Dim Item As Object
    
    ' Get the currently active window
    Set ActWind = Application.ActiveWindow
    
    ' Check if the active window is an Inspector or if there is a selection
    If ActWind.Class = olInspector Then
        Set MyItem = ActWind.CurrentItem
    ElseIf ActWind.Selection.Count > 0 Then
        Set MyItem = ActWind.Selection(1)
    End If
    
    ' Check if the item is part of a conversation
    If Not MyItem Is Nothing Then
        If HasGetConversationMethod(MyItem) And MyItem.ConversationIndex <> vbNullString Then
            ' Get the conversation of the item
            Set Conversation = MyItem.GetConversation
            
            ' Update subject for all items in the conversation
            For Each Item In Conversation.GetAssociatedItems
                Item.Subject = "id number 12" & Item.Subject
            Next Item
        Else
            ' If the item is not part of a conversation or doesn't support GetConversation, update only the current item
            MyItem.Subject = "id number 12" & MyItem.Subject
        End If
    End If
End Sub

Function HasGetConversationMethod(obj As Object) As Boolean
    On Error Resume Next
    HasGetConversationMethod = Not obj.GetConversation Is Nothing
    On Error GoTo 0
End Function
