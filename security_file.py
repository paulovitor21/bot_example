from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Configurar opções do Chrome
chrome_options = Options()
chrome_options.add_argument('--disable-web-security')  # Desativa a segurança da web
chrome_options.add_argument('--allow-running-insecure-content')  # Permite conteúdo não seguro
chrome_options.add_argument('--disable-notifications')  # Desativa notificações
chrome_options.add_argument('--disable-popup-blocking')  # Desativa bloqueio de pop-ups

# Inicializar o driver
service = Service('/path/to/chromedriver')
driver = webdriver.Chrome(service=service, options=chrome_options)

# Continuar com suas operações
driver.get('http://example.com')


# --------------------------------------------------------------------------------------------
from botcity.web import WebBot

def start_browser():
    webbot = WebBot()
    # Configurações do navegador, se possível
    # Por exemplo:
    # webbot.set_option('chrome', '--disable-web-security')
    # webbot.set_option('chrome', '--allow-running-insecure-content')
    # webbot.set_option('chrome', '--disable-notifications')
    return webbot
