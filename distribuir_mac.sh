#!/bin/bash
# distribuir_mac.sh — Gera DMG portatil do SIGEP S-1210 Extrator para macOS
# Execute este script em um Mac com Java 17+ (Temurin recomendado) instalado.
# Requisito: JAR ja compilado (mvn package) na pasta target/
set -e
cd "$(dirname "$0")"

# ── Localiza JAVA_HOME com jpackage ─────────────────────────────────────────
find_java() {
    for candidate in \
        "$(/usr/libexec/java_home -v 17+ 2>/dev/null)" \
        "/Library/Java/JavaVirtualMachines/temurin-21.jdk/Contents/Home" \
        "/Library/Java/JavaVirtualMachines/jdk-21.jdk/Contents/Home" \
        "/Library/Java/JavaVirtualMachines/jdk-17.jdk/Contents/Home"
    do
        if [ -n "$candidate" ] && [ -x "$candidate/bin/jpackage" ]; then
            echo "$candidate"
            return
        fi
    done
    echo ""
}

JAVA_HOME="$(find_java)"
if [ -z "$JAVA_HOME" ]; then
    echo "ERRO: Java 17+ com jpackage nao encontrado."
    echo "Instale o Eclipse Temurin 21: https://adoptium.net/"
    exit 1
fi
echo "Java: $JAVA_HOME"

# ── Localiza JAR ─────────────────────────────────────────────────────────────
JAR=$(ls target/s1210-extrator-*-jar-with-dependencies.jar 2>/dev/null | head -1)
if [ -z "$JAR" ]; then
    echo "ERRO: JAR nao encontrado. Execute: mvn package"
    exit 1
fi

FNAME=$(basename "$JAR")
VERSAO="${FNAME#s1210-extrator-}"
VERSAO="${VERSAO%-jar-with-dependencies.jar}"
NOME="SIGEP S-1210 Extrator"
DEST="dist-mac"

echo "Versao: $VERSAO"
echo "Removendo distribuicao anterior..."
rm -rf "$DEST"
mkdir -p "$DEST"

# ── Gera icone ICNS a partir do PNG (usa ferramentas nativas do macOS) ───────
ICNS="src/main/resources/icon.icns"
if [ ! -f "$ICNS" ]; then
    PNG="src/main/resources/logo_sigep.png"
    if [ -f "$PNG" ]; then
        echo "Convertendo PNG para ICNS..."
        ICONSET=$(mktemp -d /tmp/sigep_XXXX.iconset)
        sips -z 16   16   "$PNG" --out "$ICONSET/icon_16x16.png"    2>/dev/null
        sips -z 32   32   "$PNG" --out "$ICONSET/icon_16x16@2x.png" 2>/dev/null
        sips -z 32   32   "$PNG" --out "$ICONSET/icon_32x32.png"    2>/dev/null
        sips -z 64   64   "$PNG" --out "$ICONSET/icon_32x32@2x.png" 2>/dev/null
        sips -z 128  128  "$PNG" --out "$ICONSET/icon_128x128.png"  2>/dev/null
        sips -z 256  256  "$PNG" --out "$ICONSET/icon_128x128@2x.png" 2>/dev/null
        sips -z 256  256  "$PNG" --out "$ICONSET/icon_256x256.png"  2>/dev/null
        sips -z 512  512  "$PNG" --out "$ICONSET/icon_256x256@2x.png" 2>/dev/null
        sips -z 512  512  "$PNG" --out "$ICONSET/icon_512x512.png"  2>/dev/null
        iconutil -c icns "$ICONSET" -o "$ICNS" 2>/dev/null && echo "ICNS gerado: $ICNS"
        rm -rf "$ICONSET"
    fi
fi

# ── jpackage: gera .app bundle ───────────────────────────────────────────────
echo "Gerando .app bundle..."
ICON_OPT=""
[ -f "$ICNS" ] && ICON_OPT="--icon $ICNS"

"$JAVA_HOME/bin/jpackage" \
    --type app-image \
    --input target \
    --main-jar "$(basename "$JAR")" \
    --name "$NOME" \
    --app-version "$VERSAO" \
    --dest "$DEST" \
    $ICON_OPT \
    --java-options "-Xmx512m"

APP_PATH="$DEST/$NOME.app"
if [ ! -d "$APP_PATH" ]; then
    echo "ERRO: .app nao foi gerado."
    exit 1
fi

# ── Cria DMG ─────────────────────────────────────────────────────────────────
DMG_NAME="SIGEP_S1210_Extrator_v${VERSAO}.dmg"
echo "Gerando DMG..."
hdiutil create \
    -volname "$NOME v$VERSAO" \
    -srcfolder "$APP_PATH" \
    -ov -format UDZO \
    -o "$DEST/$DMG_NAME" \
    2>/dev/null

echo ""
echo "=== DMG gerado ==="
echo "$DEST/$DMG_NAME"
echo ""
echo "=== App bundle ==="
echo "$APP_PATH"
