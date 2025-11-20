#!/bin/bash
echo "Installing Japanese fonts..."
apt-get update
apt-get install -y fonts-ipafont-gothic fonts-ipafont-mincho fonts-noto-cjk

# フォントキャッシュを更新（これがないと matplotlib が使えない）
fc-cache -fv

echo "Available fonts (debug):"
fc-list :lang=ja

echo "Starting Streamlit app on port $PORT..."
streamlit run chatbot_app_ver2.0.py --server.port $PORT --server.address 0.0.0.0 --server.enableCORS false
