# Gunicorn configuration for Render deployment
# Automatically loaded by Gunicorn on startup, ensuring 300s timeout even if start command lacks flags

timeout = 300
workers = 1
threads = 8
worker_class = 'gthread'
keepalive = 5
max_requests = 100
max_requests_jitter = 10
