Sub CreateTicketInTicketingSystem()
    Dim xml As Object
    Set xml = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' API endpoint URL
    Dim apiUrl As String
    apiUrl = "https://your-ticketing-system-api-endpoint.com/create-ticket"
    
    Dim authToken As String
    Dim tokenFilePath As String
    tokenFilePath = "C:\myfilepath\myfile.txt" ' Correct path format

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
        ""Urgency"": ""VALUE"",
        ""AssignedToFullName"": ""VALUE"",
        ""ImpactOnTheService"": ""No Impact"",
        ""Impact"": ""VALUE"",
        ""Description"": ""VALUE"",
        ""Environment"": ""VALUE"",
        ""ITServiceName"": ""TheservicenName"",
        ""ShortDescription"": ""VALUE"",
        ""IntegrationReference"": ""FFF"",
        ""EndUserFullName"": ""VALUE"",
        ""Purpose"": ""VALUE"",
        ""PeopleToNotify"": ""ronnel.me@gmail.com"",
        ""WishedDueDate"": ""VALUE""
    }"

    ' Prompt the user to enter values for specific fields
    Dim Urgency As String
    Urgency = InputBox("Enter Urgency")
    Dim Purpose As String
    Purpose = InputBox("Enter Purpose")
    Dim AssignedToFullName As String
    AssignedToFullName = InputBox("Enter AssignedToFullName")
    Dim Environment As String
    Environment = InputBox("Enter Environment")
    Dim Impact As String
    Impact = InputBox("Enter Impact")
    Dim WishedDueDate As String
    WishedDueDate = InputBox("Enter WishedDueDate")
    Dim EndUserFullName As String
    EndUserFullName = InputBox("Enter EndUserFullName")

    ' Get the currently selected Outlook email
    Dim olApp As Object
    Dim olItem As Object
    Set olApp = GetObject(, "Outlook.Application")

    If Not olApp Is Nothing Then
        On Error Resume Next
        Set olItem = olApp.ActiveInspector.CurrentItem
        On Error GoTo 0
    End If

    ' Use the email subject as ShortDescription
    Dim ShortDescription As String
    If Not olItem Is Nothing Then
        ShortDescription = olItem.Subject
    End If

    ' Use the email body as Description
    Dim Description As String
    If Not olItem Is Nothing Then
        Description = olItem.Body
    End If

    ' Replace "VALUE" with the user-provided values
    ticketData = Replace(ticketData, "VALUE", ShortDescription)
    ticketData = Replace(ticketData, "VALUE", Description)
    ticketData = Replace(ticketData, "VALUE", Urgency)
    ticketData = Replace(ticketData, "VALUE", Purpose)
    ticketData = Replace(ticketData, "VALUE", AssignedToFullName)
    ticketData = Replace(ticketData, "VALUE", Environment)
    ticketData = Replace(ticketData, "VALUE", Impact)
    ticketData = Replace(ticketData, "VALUE", WishedDueDate)
    ticketData = Replace(ticketData, "VALUE", EndUserFullName)

    ' Send the request
    xml.send ticketData

    ' Handle the response (you can add your own logic here)
    Dim response As String
    response = xml.responseText
    MsgBox "Response: " & response
End Sub
