Sub CreateTicketInTicketingSystem()
    Dim xml As Object
    Set xml = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' API endpoint URL
    Dim apiUrl As String
    apiUrl = "https://your-ticketing-system-api-endpoint.com/create-ticket"

    ' Authentication with OAuth token
    Dim authToken As String
    authToken = "YOUR_OAUTH_TOKEN_HERE"

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
End If

' Now, authToken contains the token read from the file
Debug.Print "OAuth Token: " & authToken

    
    xml.Open "POST", apiUrl, False
    xml.setRequestHeader "Authorization", "Bearer " & authToken
    xml.setRequestHeader "Content-Type", "application/json"

    ' Data to be included in the ticket
    Dim ticketData As String
    ticketData = "{
        ""AssignmentGroup"": ""VALUE"",
        ""Urgency"": ""VALUE"",
        ""AssignedToFullName"": ""VALUE"",
        ""ImpactOnTheService"": ""VALUE"",
        ""Description"": ""VALUE"",
        ""Environment"": ""VALUE"",
        ""ITServiceName"": ""VALUE"",
        ""ShortDescription"": ""VALUE"",
        ""IntegrationReference"": ""VALUE"",
        ""EndUserFullName"": ""VALUE"",
        ""Purpose"": ""VALUE"",
        ""PeopleToNotify"": ""VALUE"",
        ""WishedDueDate"": ""VALUE""
    }"

    ' Prompt the user to enter values for specific fields
    Dim assignmentGroup As String
    assignmentGroup = InputBox("Enter AssignmentGroup")
    Dim urgency As String
    urgency = InputBox("Enter Urgency")
    Dim assignedToFullName As String
    assignedToFullName = InputBox("Enter AssignedToFullName")
    Dim impactOnTheService As String
    impactOnTheService = InputBox("Enter ImpactOnTheService")

    ' Replace "VALUE" with the user-provided values
    ticketData = Replace(ticketData, "VALUE", assignmentGroup)
    ticketData = Replace(ticketData, "VALUE", urgency)
    ticketData = Replace(ticketData, "VALUE", assignedToFullName)
    ticketData = Replace(ticketData, "VALUE", impactOnTheService)

    ' Send the request
    xml.send ticketData

    ' Handle the response (you can add your own logic here)
    Dim response As String
    response = xml.responseText
    MsgBox "Response: " & response
End Sub




Dim olApp As Object
    Dim olItem As Object
    Set olApp = GetObject(, "Outlook.Application")
    
    If Not olApp Is Nothing Then
        On Error Resume Next
        Set olItem = olApp.ActiveInspector.CurrentItem
        On Error GoTo 0
    End If

    ' Use the email subject as ShortDescription
    Dim shortDescription As String
    If Not olItem Is Nothing Then
        shortDescription = olItem.Subject
    End If

    ' Use the email body as Description
    Dim description As String
    If Not olItem Is Nothing Then
        description = olItem.Body
    End If
