"""
Gunicorn configuration for production.

Bind to localhost only -- nginx sits in front and is the only thing that
should ever be reachable from the internet, terminating HTTPS and proxying
to gunicorn over plain HTTP on the loopback interface.
"""

bind = "127.0.0.1:5000"

# SQLite + a single shared file means concurrent writers can hit "database
# is locked". Keep worker count modest for this app's traffic (an internal
# accounts tool, not a public high-traffic site) rather than the usual
# (2 x CPU + 1) formula -- 2-3 workers is plenty and reduces write contention.
workers = 3
worker_class = "sync"
timeout = 30
graceful_timeout = 30
keepalive = 5

accesslog = "-"   # stdout -- captured by systemd/journald
errorlog = "-"
loglevel = "info"
