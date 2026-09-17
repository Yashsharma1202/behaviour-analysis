import http.server
import socketserver
import socket
import sys
import os

PORT = 8080
DIRECTORY = r"D:\behaviour analysis"

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

def run():
    hostname = socket.gethostname()
    local_ip = socket.gethostbyname(hostname)
    
    print("==========================================================================================")
    print("F&O INTELLIGENCE & HOLIDAY TRADING WEB DASHBOARD SERVER")
    print("==========================================================================================")
    print(f"Local Host URL       : http://localhost:{PORT}/")
    print(f"IP Host URL          : http://{local_ip}:{PORT}/")
    print(f"IP Dashboard Link    : http://{local_ip}:{PORT}/fo_intelligence_dashboard.html")
    print("==========================================================================================")
    
    with socketserver.TCPServer(("0.0.0.0", PORT), Handler) as httpd:
        print(f"Server is LIVE & Running on 0.0.0.0:{PORT} ... Press Ctrl+C to stop.")
        sys.stdout.flush()
        httpd.serve_forever()

if __name__ == '__main__':
    run()
