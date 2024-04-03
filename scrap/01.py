import re
from googletrans import Translator

# Lista de links (substitua pelos seus próprios links)
links = [
    "https://www.vevor.de/elektrische-seilwinde-c_11304/vevor-elektrische-seilwinde-12v-13000lbs-5897kg-offroad-motorwinde-seilzug-elektrowinde-12-straengiges-nylonseil-mit-2-in-1-controller-schwarz-geeignet-fuer-mittelgrosse-grosse-suvs-lkws-und-sogar-yachten-p_010365902640",
    "https://www.vevor.de/elektrische-seilwinde-c_11304/vevor-elektrische-seilwinde-12v-4500lbs-2041kg-offroad-motorwinde-seilzug-elektrowinde-nylonseil-mit-kabelloser-fernbedienung-schwarz-ideal-fuer-mittelgrosse-grosse-suvs-lkws-und-sogar-yachten-p_010877720730",
    # Adicione mais links aqui
]


# Função para extrair e traduzir o título do link
def extrair_e_traduzir_titulo_do_link(link):
    padrao = r"/([^/]+)$"
    correspondencia = re.search(padrao, link)
    if correspondencia:
        titulo = correspondencia.group(1)
        # Realize o processamento necessário para formatar o título como desejado
        titulo_formatado = titulo.replace("-", " ").replace("p_", "").capitalize()

        # Use a biblioteca googletrans para traduzir o título para o português
        translator = Translator()
        titulo_traduzido = translator.translate(
            titulo_formatado, src="auto", dest="pt"
        ).text

        return titulo_traduzido
    else:
        return "Título não encontrado"


# Itera sobre os links e imprime os títulos traduzidos
for link in links:
    titulo_traduzido = extrair_e_traduzir_titulo_do_link(link)
    print(titulo_traduzido)
