"""
Entry point for IIS via HttpPlatformHandler.

IIS injects the port through the HTTP_PLATFORM_PORT environment variable.
Waitress listens on 127.0.0.1:<port>.

DispatcherMiddleware strips the /freight-bill prefix so Flask only
sees routes starting at /.

For local development with auto-reload:
    python run_server.py --dev
"""
import os
import sys
from pathlib import Path

# IIS uses the system Python (uv trampoline .venv/Scripts/python.exe won't work).
# Add .venv site-packages to sys.path so all dependencies are available regardless
# of which Python binary IIS is configured to use.
_site_pkgs = Path(__file__).parent / '.venv' / 'Lib' / 'site-packages'
if _site_pkgs.exists() and str(_site_pkgs) not in sys.path:
    sys.path.insert(0, str(_site_pkgs))

from waitress import serve
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.exceptions import NotFound

from config import APPLICATION_ROOT
from app import app

port = int(os.environ.get('HTTP_PLATFORM_PORT', 8080))

if __name__ == '__main__':
    if '--dev' in sys.argv:
        from werkzeug.serving import run_simple
        from werkzeug.wrappers import Response

        def _redirect_root(environ, start_response):
            path = environ.get('PATH_INFO', '/')
            if path == '/' or path == '':
                res = Response('', status=302, headers={'Location': APPLICATION_ROOT + '/'})
                return res(environ, start_response)
            return NotFound()(environ, start_response)

        dev_port = int(os.environ.get('DEV_PORT', 5000))
        app.config['TEMPLATES_AUTO_RELOAD'] = True
        wsgi_app = DispatcherMiddleware(_redirect_root, {APPLICATION_ROOT: app})
        print(f'[DEV] Freight Bill running on http://127.0.0.1:{dev_port}{APPLICATION_ROOT}')
        print('[DEV] Auto-reload enabled — save any file and the server restarts automatically.')
        run_simple('127.0.0.1', dev_port, wsgi_app, use_reloader=True, use_debugger=False)
    else:
        wsgi_app = DispatcherMiddleware(NotFound(), {APPLICATION_ROOT: app})
        print(f'Freight Bill Processor running on http://127.0.0.1:{port}{APPLICATION_ROOT}')
        serve(wsgi_app, host='127.0.0.1', port=port, threads=4)
