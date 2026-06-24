import http.server
import socketserver
import json
import requests
import os
import sys

PORT = 8000

class MyHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/api/baixar':
            # 1. Obter o tamanho do corpo da requisição
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            
            try:
                data = json.loads(post_data.decode('utf-8'))
            except Exception as e:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'JSON inválido'}).encode('utf-8'))
                return

            usuario = data.get('usuario')
            senha = data.get('senha')
            report_id = data.get('id')

            if not usuario or not senha or not report_id:
                self.send_response(400)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': 'Dados obrigatórios ausentes.'}).encode('utf-8'))
                return

            # Obter o domínio da variável de ambiente ou usar o padrão de contingência
            domain = os.environ.get('REPORT_SYSTEM_DOMAIN', 'reportview.cocal.com.br')

            try:
                # 2. Iniciar sessão do requests
                session = requests.Session()
                session.headers.update({
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
                })
                
                # 3. Efetuar Login
                login_url = f"http://{domain}/login.php"
                login_data = {'usuario': usuario, 'senha': senha}
                
                # Permite redirects para capturar todo o fluxo do PHP
                response_login = session.post(login_url, data=login_data, allow_redirects=True)

                if not session.cookies.get('PHPSESSID'):
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': 'Falha na autenticação do Interview. Verifique usuário e senha.'}).encode('utf-8'))
                    return

                # 4. Buscar Dados do Relatório
                dados_url = f"http://{domain}/relatorio/dados.php?id={report_id}"
                response_dados = session.get(dados_url)

                if response_dados.status_code != 200:
                    self.send_response(response_dados.status_code)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': f'O servidor de relatórios retornou o status: {response_dados.status_code}'}).encode('utf-8'))
                    return

                text = response_dados.text

                # Se retornou a página de login, o acesso foi negado
                if 'Login - Interview' in text or text.strip().startswith('<!doctype html'):
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': 'Acesso negado. Credenciais incorretas ou sem permissão.'}).encode('utf-8'))
                    return

                # Decodificar e retornar JSON
                try:
                    # Remove o UTF-8 BOM se presente
                    clean_text = text.replace('\ufeff', '')
                    json_data = json.loads(clean_text)
                    
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps(json_data).encode('utf-8'))
                except Exception as parse_err:
                    self.send_response(500)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        'error': 'Falha ao processar os dados do relatório. O servidor não retornou um JSON válido.',
                        'preview': text[:150]
                    }).encode('utf-8'))

            except Exception as conn_err:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({'error': f'Erro de conexão interna: {str(conn_err)}'}).encode('utf-8'))
        else:
            # Qualquer outra chamada de POST não mapeada
            self.send_response(404)
            self.end_headers()

# Iniciar o servidor
if __name__ == '__main__':
    # Configurar para servir os arquivos do diretório atual
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    
    try:
        # Permite reutilizar a mesma porta sem erros de TIME_WAIT
        socketserver.TCPServer.allow_reuse_address = True
        with socketserver.TCPServer(("", PORT), MyHandler) as httpd:
            print(f"Servidor local rodando em http://localhost:{PORT}")
            print("Pressione Ctrl+C para parar o servidor.")
            httpd.serve_forever()
    except OSError as e:
        # Se a porta 8000 já estiver em uso, sai silenciosamente (já está rodando)
        print(f"Porta {PORT} já ocupada. O servidor local já está ativo.")
        sys.exit(0)
    except KeyboardInterrupt:
        print("\nServidor parado pelo usuário.")
        sys.exit(0)
