import http.server
import socketserver
import socket
import sys
import os
import webbrowser

PORT = 8085
DIRECTORY = r"D:\behaviour analysis"

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def do_GET(self):
        if self.path in ['/', '/index.html']:
            self.path = '/event_trading_dashboard.html'
        return super().do_GET()

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def run():
    ip = get_ip()
    local_url = f"http://localhost:{PORT}/event_trading_dashboard.html"
    ip_url = f"http://{ip}:{PORT}/event_trading_dashboard.html"
    
    print("==========================================================================================")
    print("NIFTY 50 EVENT INTELLIGENCE & POSITION TAKING HOST SERVER")
    print("==========================================================================================")
    print(f"Local Host URL       : {local_url}")
    print(f"Network IP URL       : {ip_url}")
    print("==========================================================================================")
    print(f"Serving D:\\behaviour analysis on port {PORT}...")
    sys.stdout.flush()

    # Automatically open browser
    try:
        webbrowser.open(local_url)
    except Exception:
        pass

    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("0.0.0.0", PORT), Handler) as httpd:
        httpd.serve_forever()

if __name__ == '__main__':
    run()
