Sub SendEmailToTicketingSystem(Item As Outlook.MailItem)
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Replace with your API endpoint and authentication details
    Dim url As String
    url = "https://api.example.com/tickets"
    Dim authHeader As String
    authHeader = "Bearer YourAuthTokenHere"

    ' Prompt for AssignmentGroup
    Dim assignmentGroup As String
    assignmentGroup = InputBox("Enter Assignment Group:")

    ' Extract relevant email data
    Dim emailSubject As String
    Dim emailBody As String
    Dim endUserFullName As String

    emailSubject = Item.Subject
    emailBody = Item.Body
    endUserFullName = Item.SenderName

    ' Construct the API request
    Dim requestData As String
    requestData = "{""AssignmentGroup"": """ & assignmentGroup & """, " & _
                   """Description"": """ & emailBody & """, " & _
                   """Environment"": ""HOMOLOGATION"", " & _
                   """Impact"": """", " & _
                   """Urgency"": """", " & _
                   """ITServiceName"": """", " & _
                   """ShortDescription"": """ & emailSubject & """, " & _
                   """EndUserFullName"": """ & endUserFullName & """, " & _
                   """AssignedToFullName"": ""}"

    ' Set the request headers, including the Authorization header
    http.Open "POST", url, False
    http.setRequestHeader "Authorization", authHeader
    http.setRequestHeader "Content-Type", "application/json"

    ' Send the request
    http.send requestData

    ' Handle the response (if needed)
    Dim responseText As String
    responseText = http.responseText

    ' Log the response or take further actions
    Debug.Print responseText
End Sub
