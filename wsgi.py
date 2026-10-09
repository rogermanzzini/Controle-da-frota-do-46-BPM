"""
wsgi.py - Ponto de entrada de Produção WSGI
Compatível com Waitress (Windows/Linux) e Gunicorn (Linux/Containers na Nuvem)
"""
import os
import sys
from app import app

# Objeto WSGI padrão para Gunicorn / uWSGI
application = app

if __name__ == '__main__':
    from waitress import serve
    porta = int(os.environ.get('PORT', 5000))
    threads = int(os.environ.get('WAITRESS_THREADS', 8))
    
    print("=" * 65)
    print("  46º BATALHÃO DE POLÍCIA MILITAR - BAGÉ / RS")
    print("  Servidor WSGI de Produção (Waitress)")
    print(f"  Disponível na rede em: http://0.0.0.0:{porta}")
    print(f"  Pool de threads ativas: {threads}")
    print("=" * 65)
    
    serve(app, host='0.0.0.0', port=porta, threads=threads)
