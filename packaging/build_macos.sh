#!/bin/bash
# ============================================================
#  NeuroTrace — Build para macOS (.app + .dmg)
# ============================================================
set -e

# Navegar para a raiz do projeto
cd "$(dirname "$0")/.."

APP_NAME="NeuroTrace"
VERSION="2.0.0"
DMG_NAME="${APP_NAME}_macOS_v${VERSION}"
DIST_DIR="dist"
DMG_DIR="dmg_staging"
DMG_OUTPUT="installer_output/${DMG_NAME}.dmg"

echo "============================================"
echo "  NeuroTrace — Build macOS"
echo "  (.app bundle + .dmg installer)"
echo "============================================"
echo ""

# ─── Etapa 1: Gerar o .app com PyInstaller ───
echo "[1/2] Gerando .app bundle com PyInstaller..."
echo ""

python3 -m PyInstaller packaging/main_macos.spec --clean --noconfirm

if [ ! -d "${DIST_DIR}/${APP_NAME}.app" ]; then
    echo ""
    echo "[ERRO] Falha ao criar o .app bundle."
    exit 1
fi

echo ""
echo "[OK] .app criado em ${DIST_DIR}/${APP_NAME}.app"
echo ""

# ─── Etapa 2: Criar o .dmg ───
echo "[2/2] Criando .dmg installer..."
echo ""

# Limpar staging anterior
rm -rf "${DMG_DIR}"
mkdir -p "${DMG_DIR}"
mkdir -p "installer_output"

# Copiar o .app para o staging
cp -R "${DIST_DIR}/${APP_NAME}.app" "${DMG_DIR}/"

# Criar symlink para /Applications (arrastar e soltar)
ln -s /Applications "${DMG_DIR}/Applications"

# Criar .dmg (remover anterior se existir)
rm -f "${DMG_OUTPUT}"

hdiutil create \
    -volname "${APP_NAME}" \
    -srcfolder "${DMG_DIR}" \
    -ov \
    -format UDZO \
    -imagekey zlib-level=9 \
    "${DMG_OUTPUT}"

# Limpar staging
rm -rf "${DMG_DIR}"

echo ""
echo "============================================"
echo "  Build macOS completo!"
echo ""
echo "  .app bundle: ${DIST_DIR}/${APP_NAME}.app"
echo "  Instalador:  ${DMG_OUTPUT}"
echo "============================================"
echo ""
