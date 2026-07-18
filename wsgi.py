"""
Production WSGI entry point.

Flask's built-in dev server (`python app.py`) is single-threaded and not
meant to be reachable from the network -- it's for local development only.
In production, gunicorn imports the Flask `app` object through this file
instead and serves it with multiple worker processes.

Run with:
    gunicorn -c gunicorn.conf.py wsgi:app

See DEPLOY.md for the full setup (systemd + nginx + HTTPS).
"""

from app import app  # noqa: F401  (gunicorn imports `app` from this module)
