import os
from http.server import HTTPServer, BaseHTTPRequestHandler
from api.export_chat import handler

# Класс-обертка для перенаправления запросов в ваш handler из Vercel
class ProxyHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        # Передаем управление вашему существующему handler
        return handler(self)

    def do_GET(self):
        # На случай, если нужно проверить, жив ли сервер
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Server is running")

if __name__ == "__main__":
    # Railway автоматически назначает порт через переменную окружения
    port = int(os.environ.get('PORT', 8080))
    server_address = ('0.0.0.0', port)
    
    httpd = HTTPServer(server_address, ProxyHandler)
    print(f"Сервер запущен на порту {port}")
    httpd.serve_forever()
