# C:\Users\Eliezer\Desktop\WS_Imoveis\venv\Scripts\activate.ps1
import pandas as pd
import numpy as np

from urllib.request import urlopen as uReq
import requests
from bs4 import BeautifulSoup
import json
import time
from datetime import date

import re
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

listaJson = []
options = Options()
options.add_argument("--disable-notifications")
options.add_argument("--mute-audio")
driver = webdriver.Chrome(
    options=options, executable_path='./chromedriver.exe')


def definirParams(auth, path, referer):
    PARAMS = {
        "authority": auth,
        "method": "GET",
        "path": path,
        "scheme": "https",
        "referer": referer,
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36',
    }

    return PARAMS


def criarJson(dataPostagem, titulo, preco, url, regiao, regiaoCidade, resumo, metragem):

    print("titulo: " + str(titulo))
    print("Preco: " + str(preco))
    print(str(metragem)+ "m²")
    print(regiao)
    print(dataPostagem)
    print(url)
    print(resumo)
    print("----")

    json = {
        "preco": preco,
        "titulo": titulo,
        "metragem": metragem,
        "regiao": regiao,
        "dataPostagem": dataPostagem,
        "regiaoCidade": regiaoCidade,
        "resumo": resumo,
        "url": url, }

    # adiciona na lista GERAL de JSON
    listaJson.append(json)


def retornarSoupSimples(url):
    auth = url.split("/")[2]
    path = url.split("/")[3]
    referer = url
    PARAMS = definirParams(auth=auth, path=path, referer=url)
    page = requests.get(url=url, headers=PARAMS)
    soup = BeautifulSoup(page.content, "lxml")

    return soup


def buscarDadosOlx(pages):

    for x in range(1, pages):

        try:
            print("Página número: " + str(x))
            #url = "https://sp.olx.com.br/grande-campinas/imoveis?o="+str(x)
            #url = "https://sp.olx.com.br/grande-campinas/imoveis?f=p&o="+str(x)
            url = "https://sp.olx.com.br/grande-campinas/regiao-de-campinas/imoveis?f=p&o="+str(x)

            soup = retornarSoupSimples(url)

            itens = soup.find_all('li', {'class': 'sc-1fcmfeb-2'})

            for item in itens:
                try:
                    titulo = item.find_all('a')[0]['title']
                    url = item.find_all('a')[0]['href']
                    preco = item.find_all('span', {"aria-label": re.compile("Preço")})[0].text
                    preco = preco.replace('R$ ', '')
                    preco = preco.replace('.', '')
                    preco = float(preco)
                    resumo = item.get_text()
                    metragem = item.find_all('span', {'aria-label': re.compile("m²")})[0].text
                    metragem = metragem.replace('m²', '')
                    metragem = int(metragem)
                    dataPostagem = item.find_all('span', {"aria-label": re.compile("Anúncio")})[0].text
                    regiao = item.find_all('span', {"aria-label": re.compile("Localização")})[0].text
                    regiao = regiao.strip()
                    try:
                        regiaoCidade = regiao.split(',')[0]
                    except:
                        regiaoCidade = regiao

                    #da print e adiciona na lista geral
                    criarJson(dataPostagem, titulo, preco, url,
                            regiao, regiaoCidade, resumo, metragem)

                except Exception as error:
                    print(url)
                    print('Erro na OLX', error)
                    pass

        except Exception as error:
            print("Erro na URL da OLX" + url)
            print(error)
            pass
            
        
buscarDadosOlx(101)

print(len(listaJson))
df = pd.DataFrame(listaJson)
df.tail()

with pd.ExcelWriter("Particular_Campinas.xlsx") as writer:
    df.to_excel(writer, 'Teste')
