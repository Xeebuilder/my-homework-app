#!/bin/bash
cd "$(dirname "$0")"
echo "🚀 Iniciando servidor local..."
echo "📄 Abre tu navegador en: http://localhost:8080/index.html"
echo "   (Se abrirá automáticamente en 2 segundos)"
sleep 2
open "http://localhost:8080/index.html"
python3 -m http.server 8080
