from gtts import gTTS

# Carregar o conteúdo do ficheiro
with open("03-02-2025 cleaned_transcription.txt", "r", encoding="utf-8") as f:
    texto = f.read()

# Dividir o texto em blocos de ~1400 palavras (cerca de 10 minutos)
def dividir_texto(texto, tamanho=1400):
    palavras = texto.split()
    return [' '.join(palavras[i:i + tamanho]) for i in range(0, len(palavras), tamanho)]

blocos = dividir_texto(texto)

# Gerar ficheiros de áudio para cada bloco
for i, bloco in enumerate(blocos):
    tts = gTTS(text=bloco, lang='pt', tld='pt')  # 'pt' com tld='pt' para português europeu
    nome_ficheiro = f"bloco_audio_{i+1}.mp3"
    tts.save(nome_ficheiro)
    print(f"Ficheiro guardado: {nome_ficheiro}")
