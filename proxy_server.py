import http.server
import http.client
from http.server import SimpleHTTPRequestHandler
from io import BytesIO

# Define the target server and port
target_host = 'area109.com'
target_port = 443  # Use 443 for HTTPS

# Create a custom request handler
class ProxyHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        self.proxy_request('GET')

    def do_POST(self):
        self.proxy_request('POST')

    def proxy_request(self, method):
        # Create a connection to the target server
        target_conn = http.client.HTTPSConnection(target_host, target_port)  # Use HTTPSConnection for HTTPS

        # Prepare the request to the target server
        target_conn.request(method, self.path, body=self.rfile, headers=self.headers)  # Pass request body and headers
        response = target_conn.getresponse()

        # Send the target server's response back to the client
        self.send_response(response.status)
        self.send_header('Content-type', response.getheader('Content-type'))
        self.end_headers()
        self.wfile.write(response.read())

if __name__ == '__main__':
    server_address = ('', 8000)  # You can change the port as needed
    httpd = http.server.HTTPServer(server_address, ProxyHandler)
    print('Proxy server is running on port 8000')
    httpd.serve_forever()
