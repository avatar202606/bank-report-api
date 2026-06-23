# Procfile for Render.com deployment
# 告诉 Render 如何启动服务

web: gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
