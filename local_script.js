function submitForm(event) {
    event.preventDefault(); // Prevent the form from submitting by default
    console.log('Submit button clicked');
    var formData = new FormData(document.getElementById('sampleForm'));
    var apiUrl = 'https://area109.com/requests';
    var httpMethod = 'POST';
    var accessToken = 'test'; // Replace with your actual access token
    var bearerToken = 'Bearer ' + accessToken;
    var xhr = new XMLHttpRequest();
    xhr.open(httpMethod, apiUrl, true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.setRequestHeader('Authorization', bearerToken);
    xhr.onload = function () {
        if (xhr.status === 200) {
            console.log('Request was successful');
        } else {
            console.error('Request failed');
        }
    };
    var jsonData = {};
    formData.forEach(function (value, key) {
        jsonData[key] = value;
    });
    var jsonPayload = JSON.stringify(jsonData);
    xhr.send(jsonPayload);
}
