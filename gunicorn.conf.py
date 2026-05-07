# Render free tier: ~512MB RAM
# With 1 worker + 2 threads we stay safely under the limit
workers = 1
threads = 2
timeout = 480           # 8 minutes; covers large PDFs with OCR
keepalive = 5           # keep idle connections alive for 5s
worker_class = "gthread"
max_requests = 50       # recycle worker after 50 requests to prevent memory creep
max_requests_jitter = 10
preload_app = False     # don't preload; saves ~30MB on startup
