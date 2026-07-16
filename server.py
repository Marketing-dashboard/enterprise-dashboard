"""Local preview server — python server.py  →  http://localhost:3030"""
import http.server, os
PORT, BASE = 3030, os.path.dirname(os.path.abspath(__file__))
class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self,*a,**k): super().__init__(*a,directory=BASE,**k)
    def log_message(self,*a): pass
print(f'Dashboard at http://localhost:{PORT}  (Ctrl+C to stop)')
http.server.HTTPServer(('',PORT),H).serve_forever()
