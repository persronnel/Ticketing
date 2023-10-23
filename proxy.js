const express = require('express');
const request = require('request');

const app = express();
const port = 3000; // You can use any available port

// Set up a middleware to handle all incoming requests
app.use('/', (req, res) => {
  const url = 'https://g-itsmylink.group.com' + req.url; // Change this URL to the external server you want to access

  // Pipe the request from the client to the external server
  req.pipe(request(url)).pipe(res);
});

app.listen(port, () => {
  console.log(`Proxy server is running on port ${port}`);
});
