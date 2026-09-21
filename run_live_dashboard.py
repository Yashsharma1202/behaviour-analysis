"""
run_live_dashboard.py
===============================================================================
Launches the Nifty 50 Event Execution Dashboard locally in your default browser.
===============================================================================
"""
import http.server
import socketserver
import webbrowser
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

ROOT = Path(__file__).resolve().parent
PORT = 8550

class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

def main():
    html_file = ROOT / "simple_execution_portal.html"
    if not html_file.exists():
        print("Error: simple_execution_portal.html not found.")
        return
    
    url = f"http://localhost:8550/simple_execution_portal.html"
    print("=" * 80)
    print("⚡ LAUNCHING NIFTY 50 LIVE EXECUTION DASHBOARD")
    print("=" * 80)
    print(f"Server Running at: {url}")
    print("Press Ctrl+C in terminal to stop server.")
    print("=" * 80)
    
    webbrowser.open(url)
    
    with socketserver.TCPServer(("", PORT), QuietHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nServer stopped.")

if __name__ == "__main__":
    main()
