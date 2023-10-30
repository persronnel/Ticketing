Function ParseJson(jsonString As String, key As String) As String
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = """" & key & """:\s*""([^""]*)"""
    End With
    
    If regex.test(jsonString) Then
        Set matches = regex.Execute(jsonString)
        ParseJson = matches(0).SubMatches(0)
    Else
        ParseJson = "Key not found"
    End If
End Function


Function SanitizeJSONString(ByVal input As String) As String
    Dim sanitized As String
    Dim char As String
    sanitized = ""
    For i = 1 To Len(input)
        char = Mid(input, i, 1)
        If (AscW(char) >= 32 And AscW(char) <= 126) Or char = vbCr Or char = vbLf Or char = vbTab Then
            sanitized = sanitized & char
        End If
    Next i
    SanitizeJSONString = sanitized
End Function

Sub SendToAPI()

    Dim xml As Object
    Set xml = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' API endpoint URL
    Dim apiUrl As String
    apiUrl = "https://your-ticketing-system-api-endpoint.com/create-ticket"
    
    Dim authToken As String
    Dim tokenFilePath As String
    tokenFilePath = "C:\myfilepath\myfile.txt" 

    ' Check if the file exists
    If Dir(tokenFilePath) <> "" Then
        ' Open the file for reading
        Open tokenFilePath For Input As #1
        ' Read the token from the file
        Line Input #1, authToken
        ' Close the file
        Close #1
    Else
        ' Handle the case where the file doesn't exist
        MsgBox "OAuth token file not found."
        ' You might want to exit or handle this error in a different way
        Exit Sub ' Exit the subroutine if the token file is not found
    End If

    ' Now, authToken contains the token read from the file
    Debug.Print "OAuth Token: " & authToken

    xml.Open "POST", apiUrl, False
    xml.setRequestHeader "Authorization", "Bearer " & authToken
    xml.setRequestHeader "Content-Type", "application/json"

    ' Data to be included in the ticket
    Dim ticketData As String
    ticketData = "{
        ""AssignmentGroup"": ""WW_Team"",
        ""Urgency"": ""UrgencyVal"",
        ""AssignedToFullName"": ""VALUE"",
        ""ImpactOnTheService"": ""No Impact"",
        ""Impact"": ""ImpactVal"",
        ""Description"": ""DescriptionVal"",
        ""Environment"": ""EnvironmentVal"",
        ""ITServiceName"": ""TheservicenName"",
        ""ShortDescription"": ""SDesc"",
        ""IntegrationReference"": ""FFF"",
        ""EndUserFullName"": ""EndUserFullNameVal"",
        ""Purpose"": ""PurposeVal"",
        ""PeopleToNotify"": ""ronnel.me@gmail.com"",
        ""WishedDueDate"": ""WishedDueDateVal""
    }"

    Dim olApp As Object
    Dim olItem As Object
    Set olApp = GetObject(, "Outlook.Application")

    If Not olApp Is Nothing Then
        On Error Resume Next
        Set olItem = olApp.ActiveInspector.CurrentItem
        On Error GoTo 0
    End If

    ' Use the email subject as ShortDescription
    Dim SDesc As String
    If Not olItem Is Nothing Then
        SDesc = SanitizeJSONString(Replace(olItem.Subject, vbCrLf, "\n"))
    End If

    ' Use the email body as Description
    Dim Description As String
    If Not olItem Is Nothing Then
        Description = SanitizeJSONString(Replace(olItem.Body, vbCrLf, "\n"))
    End If

    ' Replace "VALUE" with the user-provided values
    ticketData = Replace(ticketData, "ShortDescriptionVal", SDesc)
    ticketData = Replace(ticketData, "DescriptionVal", Description)
    ticketData = Replace(ticketData, "UrgencyVal", Urgency)
    ticketData = Replace(ticketData, "PurposeVal", Purpose)
    ticketData = Replace(ticketData, "AssignedToFullNameVal", AssignedToFullName)
    ticketData = Replace(ticketData, "EnvironmentVal", Environment)
    ticketData = Replace(ticketData, "ImpactVal", Impact)
    ticketData = Replace(ticketData, "WishedDueDateVal", WishedDueDate)
    ticketData = Replace(ticketData, "EndUserFullNameVal", EndUserFullName)

    ' Send the request
    xml.send ticketData
    
    If xml.Status = 200 Then
            Dim response As String
            response = xml.responseText
    
            Dim ticketNumber As String
            ticketNumber = ParseJson(response, "id")
    
            ' Update the email subject with the ticket number
            If Not olItem Is Nothing Then
                olItem.Subject = "Ticket #" & ticketNumber & " - " & olItem.Subject
                olItem.Save ' Save the email with the updated subject
            End If
    
            ' Show the ticket number
            MsgBox "Ticket Number: " & ticketNumber
        Else
            ' Handle API request failure
            MsgBox "API request failed. Ticket creation was not successful."
        End If
End Sub


//

Private Sub SubmitButton_Click()
    ' Collect user inputs from the UserForm controls
    Dim Urgency As String
    Urgency = UrgencyComboBox.Value

    Dim AssignedToFullName As String
    AssignedToFullName = AssignedToFullNameComboBox.Value

    Dim Impact As String
    Impact = ImpactComboBox.Value

    Dim Environment As String
    Environment = EnvironmentComboBox.Value

    Dim EndUserFullName As String
    EndUserFullName = EndUserFullNameComboBox.Value

    Dim Purpose As String
    Purpose = PurposeComboBox.Value

    Dim WishedDueDate As String
    WishedDueDate = WishedDueDateComboBox.Value

    ' Call the SendToAPI function with the collected data
    SendToAPI Urgency, AssignedToFullName, Impact, Environment, EndUserFullName, Purpose, WishedDueDate

    ' Close the UserForm
    Unload Me
End Sub


// mmodule
Sub SendToAPI(Urgency As String, AssignedToFullName As String, Impact As String, Environment As String, EndUserFullName As String, Purpose As String, WishedDueDate As String)


ticketData = "{
    ""AssignmentGroup"": ""WW_Team"",
    ""Urgency"": """ & Urgency & """",
    ""AssignedToFullName"": """ & AssignedToFullName & """",
    ""ImpactOnTheService"": ""No Impact"",
    ""Impact"": ""ImpactVal"",
    ""Description"": ""DescriptionVal"",
    ""Environment"": """ & Environment & """",
    ""ITServiceName"": ""TheservicenName"",
    ""ShortDescription"": """ & SDesc & """",
    ""IntegrationReference"": ""FFF"",
    ""EndUserFullName"": """ & EndUserFullName & """",
    ""Purpose"": """ & Purpose & """",
    ""PeopleToNotify"": ""ronnel.me@gmail.com"",
    ""WishedDueDate"": """ & WishedDueDate & """
}"
