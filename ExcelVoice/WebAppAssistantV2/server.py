import http.server
import ssl

# Configurações do servidor
server_address = ('', 8082)  # Porta 4443 para HTTPS
httpd = http.server.HTTPServer(server_address, http.server.SimpleHTTPRequestHandler)

# Configuração SSL
httpd.socket = ssl.wrap_socket(httpd.socket,
                               server_side=True,
                               certfile='cert.pem',
                               keyfile='key.pem',
                               ssl_version=ssl.PROTOCOL_TLS)

print("Servidor HTTPS na porta 8082...")
httpd.serve_forever()