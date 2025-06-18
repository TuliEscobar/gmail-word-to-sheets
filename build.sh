#!/usr/bin/env bash
# build.sh - Script para instalar dependencias en Render

# Detener en caso de error
set -o errexit

echo "🔄 Actualizando lista de paquetes..."
apt-get update

echo "📦 Instalando LibreOffice (requerido para conversión de documentos)..."
DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-base \
    libreoffice-math \
    libreoffice-common \
    libreoffice-core \
    libreoffice-java-common \
    fonts-opensymbol \
    uno-libs3 \
    ure \
    ure-dbg \
    libuno-cil-dev \
    libuno-cil-doc \
    libuno-cil-java \
    libuno-purpenvhelpergcc3-3 \
    libuno-sal3 \
    libuno-salhelpergcc3-3 \
    libunoloader \
    python3-uno \
    python3-uno-dbg \
    python3-uno-doc \
    unoconv

echo "🧹 Limpiando caché de paquetes para reducir el tamaño de la imagen..."
apt-get clean
rm -rf /var/lib/apt/lists/*

echo "🐍 Instalando dependencias de Python..."
pip install --upgrade pip
pip install -r requirements.txt

echo "✅ Configuración completada con éxito"