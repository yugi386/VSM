#!/bin/bash

# === CONFIGURAÇÃO ===
# 2 frames por segundo
FRAMES_POR_SEGUNDO=1

# Habilita nullglob para não falhar se não houver arquivos
shopt -s nullglob

VIDEOS=(*.mp4 *.avi *.mkv *.mov)

if [ ${#VIDEOS[@]} -eq 0 ]; then
    echo "[Bash] Nenhum arquivo de vídeo encontrado na pasta atual."
    exit 0
fi

echo "[Bash] Encontrados ${#VIDEOS[@]} vídeos para processar."
echo "[Bash] Taxa de extração: $FRAMES_POR_SEGUNDO fps."

# Verifica se img2pdf está instalado
if ! command -v img2pdf &> /dev/null; then
    echo "[ERRO] O comando 'img2pdf' não foi encontrado."
    echo "Instale usando: pip install img2pdf"
    exit 1
fi

for video in "${VIDEOS[@]}"; do
    echo "=================================================="
    echo "[Bash] Processando: $video"
    
    FILENAME="${video%.*}"
    PDF_OUTPUT="${FILENAME}.pdf"
    DIR_TEMP="temp_${FILENAME}"

    if [ -f "$PDF_OUTPUT" ]; then
        echo "[Bash] O arquivo $PDF_OUTPUT já existe. Pulando..."
        continue
    fi

    # 1. Cria diretório
    rm -rf "$DIR_TEMP"
    mkdir -p "$DIR_TEMP"

    # 2. Extrai frames (com numeração ordenada %08d)
    echo "[Bash] Extraindo frames com FFmpeg..."
    ffmpeg -i "$video" -vf fps=$FRAMES_POR_SEGUNDO "$DIR_TEMP/frame_%08d.jpg" -loglevel error

    # 3. Converte para PDF usando img2pdf
    echo "[Bash] Gerando PDF com img2pdf (Tamanho Original)..."
    
    # CORREÇÃO: Removemos flags de orientação/tamanho. 
    # O padrão é usar o tamanho exato da imagem (Paisagem se o vídeo for paisagem).
    find "$DIR_TEMP" -name "*.jpg" | sort | xargs img2pdf -o "$PDF_OUTPUT"

    # 4. Limpeza
    echo "[Bash] Limpando arquivos temporários..."
    rm -rf "$DIR_TEMP"
    
    echo "[Bash] Concluído: $video -> $PDF_OUTPUT"
done

echo "=================================================="
echo "[Bash] Processamento finalizado!"
