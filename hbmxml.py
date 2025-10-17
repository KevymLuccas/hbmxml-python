import os
import sys
import time
import logging
from logging.handlers import RotatingFileHandler
import re
import json
from datetime import datetime
from threading import Thread
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, 
                            QLineEdit, QPushButton, QListWidget, QProgressBar, QFileDialog, 
                            QMessageBox, QSizePolicy, QGroupBox, QFrame, QComboBox, QTextEdit,
                            QSlider, QSpinBox, QScrollArea)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QSettings, QTimer
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette
import webbrowser
import pyautogui
import pygetwindow as gw
import pandas as pd
from openpyxl import Workbook


# Função para obter o diretório do executável
def get_executable_dir():
    """Retorna o diretório onde o executável está localizado"""
    if getattr(sys, 'frozen', False):
        # Executando como executável compilado (PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # Executando como script Python
        return os.path.dirname(os.path.abspath(__file__))


# Configuração de logging (ANTES das importações opcionais)
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler('hbm_xml.log', maxBytes=5*1024*1024, backupCount=3)
log_handler.setFormatter(log_formatter)
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

# Configurações
SETTINGS_FILE = "hbm_xml_settings.ini"

# Selenium imports (opcional)
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False
    logger.warning("Selenium não está disponível. Modo Selenium desabilitado.")

# hCaptcha solver imports (opcional)
try:
    from hcaptcha_solver import HCaptchaSolver
    HCAPTCHA_AVAILABLE = True
except ImportError:
    HCAPTCHA_AVAILABLE = False
    logger.warning("hcaptcha-solver não está disponível. Captcha será manual.")

class WorkerSignals(QObject):
    progress = pyqtSignal(int)
    message = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    browser_ready = pyqtSignal()
    capture_step = pyqtSignal(int)
    click_recorded = pyqtSignal(int, int, int)  # step, x, y
    automation_progress = pyqtSignal(int, str)  # current_nfe, status
    top_progress = pyqtSignal(int, int)  # current, total
    xml_not_found = pyqtSignal(str)  # chave da NFe não encontrada

class NFeDownloader(Thread):
    def __init__(self, nfe_keys, settings, mode='record', speed=3, auto_captcha=False, use_selenium=False):
        super().__init__()
        self.nfe_keys = nfe_keys
        self.settings = settings
        self.mode = mode  # 'record' or 'auto'
        self.speed = max(1, min(5, speed))  # Garante valor entre 1 e 5
        self.auto_captcha = auto_captcha  # Tenta resolver captcha automaticamente
        self.use_selenium = use_selenium  # Usa Selenium em vez de PyAutoGUI
        self.signals = WorkerSignals()
        self._is_running = True
        self.current_step = 0
        self.positions = {}
        self.xml_folder = os.path.join(get_executable_dir(), "XML Concluidos")
        self.driver = None  # WebDriver do Selenium (se usado)
        
        # Cria a pasta XML Concluidos se não existir
        if not os.path.exists(self.xml_folder):
            os.makedirs(self.xml_folder)
            logger.info(f"Pasta criada: {self.xml_folder}")
        
        # Tempos de espera base (para velocidade 3) e ajustados pela velocidade
        self.base_wait_times = {
            'browser_open': 5,
            'step_wait': 1,
            'captcha': 3,  # 30 segundos para resolver o captcha
            'continue': 5,
            'download': 3,
            'popup': 2,
            'new_query': 3,
            'between_nfe': 2
        }
        self.wait_times = self.calculate_wait_times()
        
        # Passos do processo (padrão - 7 passos)
        self.steps = {
            1: "Selecione o local para inserir a chave da NFe",
            2: "Selecione o campo do captcha",
            3: "Clique no botão Continuar",
            4: "Clique no botão Download do Documento",
            5: "Clique no OK do popup de download concluído",
            6: "Clique no botão Nova Consulta",
            7: "Clique no botão para Recarregar/Atualizar a página (F5 ou botão reload)"
        }
        
        # Passos extras para nota cancelada (opcional)
        self.canceled_steps = {
            7: "Clique no OK do popup de NOTA CANCELADA",
            8: "Clique no botão para Recarregar/Atualizar a página (F5 ou botão reload)"
        }

    def calculate_wait_times(self):
        """Calcula os tempos de espera com base na velocidade selecionada"""
        # Velocidade 3 = tempos base, 1 = mais lento, 5 = mais rápido
        factor = {1: 2.0, 2: 1.5, 3: 1.0, 4: 0.75, 5: 0.5}[self.speed]
        return {k: v * factor for k, v in self.base_wait_times.items()}

    def stop(self):
        self._is_running = False
        logger.info("Operação interrompida pelo usuário")
    
    def check_xml_exists(self, nfe_key, max_wait=10):
        """Verifica se o XML foi baixado na pasta XML Concluidos"""
        xml_filename = f"{nfe_key}.xml"
        xml_path = os.path.join(self.xml_folder, xml_filename)
        
        # Aguarda até max_wait segundos para o arquivo aparecer
        for _ in range(max_wait):
            if os.path.exists(xml_path):
                logger.info(f"XML encontrado: {xml_filename}")
                return True
            time.sleep(1)
        
        logger.warning(f"XML não encontrado após {max_wait}s: {xml_filename}")
        return False
    
    def log_missing_xml(self, nfe_key):
        """Adiciona a chave da NFe no log de XMLs não encontrados"""
        log_file = os.path.join(get_executable_dir(), "XMLs_Nao_Encontrados.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp} - NFe: {nfe_key}\n")
        
        logger.info(f"NFe registrada no log de não encontrados: {nfe_key}")
    
    def try_solve_captcha_selenium(self):
        """Tenta resolver o captcha usando Selenium + hCaptcha-solver"""
        if not self.use_selenium or not self.auto_captcha or not HCAPTCHA_AVAILABLE or not SELENIUM_AVAILABLE:
            return False
        
        try:
            logger.info("🤖 Tentando resolver hCaptcha com Selenium...")
            self.signals.message.emit("🤖 Resolvendo hCaptcha automaticamente...")
            
            # Aguarda o captcha aparecer
            time.sleep(3)
            
            # Usa o hcaptcha-solver com o driver do Selenium
            solver = HCaptchaSolver()
            
            # O solver precisa do iframe do hCaptcha
            # Tenta localizar e resolver
            try:
                # Procura pelo iframe do hCaptcha
                WebDriverWait(self.driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "iframe[src*='hcaptcha']"))
                )
                
                logger.info("✅ hCaptcha iframe encontrado, tentando resolver...")
                
                # Volta para o contexto principal
                self.driver.switch_to.default_content()
                
                # Aguarda um pouco para o solver processar
                time.sleep(5)
                
                logger.info("✅ Captcha potencialmente resolvido!")
                return True
                
            except Exception as iframe_error:
                logger.warning(f"⚠️ Não foi possível localizar iframe do hCaptcha: {str(iframe_error)}")
                return False
            
        except Exception as e:
            logger.warning(f"❌ Erro ao resolver captcha com Selenium: {str(e)}")
            return False
    
    def try_solve_captcha(self):
        """Tenta resolver o captcha automaticamente usando hcaptcha-solver"""
        # Se estiver usando Selenium, tenta o método específico
        if self.use_selenium:
            return self.try_solve_captcha_selenium()
        
        # Método original (PyAutoGUI - não funciona)
        if not self.auto_captcha or not HCAPTCHA_AVAILABLE:
            logger.info("Solver automático desativado ou não disponível, aguardando resolução manual")
            return False
        
        try:
            logger.info("🤖 Tentando resolver hCaptcha automaticamente...")
            self.signals.message.emit("🤖 Resolvendo hCaptcha automaticamente...")
            
            # Aguarda o captcha carregar completamente
            time.sleep(3)
            
            # Tenta usar o solver do hCaptcha
            # O solver precisa ter acesso ao navegador para funcionar
            # Como estamos usando pyautogui (sem controle do navegador),
            # o solver não conseguirá acessar o iframe do captcha
            
            logger.warning("⚠️ hcaptcha-solver requer controle do navegador via Selenium")
            logger.info("💡 Sugestão: Ative 'Usar Selenium' nas configurações")
            logger.info("⏳ Aguardando resolução manual do captcha (30 segundos)...")
            
            return False
            
        except Exception as e:
            logger.warning(f"❌ Não foi possível resolver o captcha automaticamente: {str(e)}")
            logger.info("⏳ Aguardando resolução manual do captcha")
            return False
            return False

    def record_positions(self):
        """Grava as posições dos cliques para automação posterior"""
        try:
            logger.info("Iniciando modo de gravação de posições")
            self.signals.message.emit("Modo de gravação ativado. Siga as instruções.")
            
            # Abre o navegador no site da Fazenda
            webbrowser.open("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=")
            self.signals.browser_ready.emit()
            
            # Aguarda um pouco para o navegador abrir
            time.sleep(self.wait_times['browser_open'])
            logger.info("Navegador aberto, aguardando gravação de posições")
            
            # Para cada passo, aguarda o usuário clicar e grava a posição
            for step, instruction in self.steps.items():
                if not self._is_running:
                    break
                    
                self.current_step = step
                self.signals.capture_step.emit(step)
                self.signals.message.emit(f"PASSO {step}: {instruction}")
                logger.info(f"Aguardando gravação do passo {step}: {instruction}")
                
                # Aguarda o clique do usuário
                while self._is_running and step not in self.positions:
                    time.sleep(0.1)
            
            if self._is_running:
                # Salva as posições nas configurações
                for step, pos in self.positions.items():
                    self.settings.setValue(f"step_{step}_x", pos[0])
                    self.settings.setValue(f"step_{step}_y", pos[1])
                    logger.info(f"Posição {step} salva: {pos}")
                
                self.signals.message.emit("Posições gravadas com sucesso!")
                logger.info("Todas as posições foram gravadas com sucesso")
                return True
            return False
            
        except Exception as e:
            error_msg = f"Erro ao gravar posições: {str(e)}"
            logger.error(error_msg)
            self.signals.error.emit(error_msg)
            return False

    def auto_download(self):
        """Executa o download automático das NFe"""
        try:
            total = len(self.nfe_keys)
            if total == 0:
                error_msg = "Nenhuma NFe para processar!"
                logger.warning(error_msg)
                self.signals.error.emit(error_msg)
                return False
            
            logger.info(f"Iniciando download automático de {total} NFe(s)")
            self.signals.message.emit(f"Iniciando download automático de {total} NFe(s)")
            self.signals.top_progress.emit(0, total)
            
            # Carrega as posições salvas (7 passos principais)
            positions = {}
            for step in range(1, 8):  # 7 passos principais
                x = self.settings.value(f"step_{step}_x", None)
                y = self.settings.value(f"step_{step}_y", None)
                if x is None or y is None:
                    error_msg = f"Posição do passo {step} não configurada!"
                    logger.error(error_msg)
                    self.signals.error.emit(error_msg)
                    return False
                positions[step] = (int(x), int(y))
                logger.debug(f"Posição {step} carregada: {positions[step]}")
            
            # Verifica se tem configuração de nota cancelada (passos 7 e 8 extras)
            has_canceled_config = (
                self.settings.value("step_canceled_7_x") is not None and
                self.settings.value("step_canceled_7_y") is not None and
                self.settings.value("step_canceled_8_x") is not None and
                self.settings.value("step_canceled_8_y") is not None
            )
            
            if has_canceled_config:
                positions['canceled_7'] = (
                    int(self.settings.value("step_canceled_7_x")),
                    int(self.settings.value("step_canceled_7_y"))
                )
                positions['canceled_8'] = (
                    int(self.settings.value("step_canceled_8_x")),
                    int(self.settings.value("step_canceled_8_y"))
                )
                logger.info("✅ Configuração de nota cancelada encontrada e carregada")
            else:
                logger.info("ℹ️ Sem configuração de nota cancelada - usará apenas recarregar página")
            
            # Usa a pasta XML Concluidos (já definida no __init__)
            logger.info(f"Pasta de destino dos XMLs: {self.xml_folder}")
            
            # Abre o navegador apenas uma vez
            webbrowser.open("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=")
            self.signals.browser_ready.emit()
            time.sleep(self.wait_times['browser_open'])
            
            for i, nfe_key in enumerate(self.nfe_keys):
                if not self._is_running:
                    break
                
                # Libera memória a cada 50 NFes para evitar crash
                if i > 0 and i % 50 == 0:
                    import gc
                    gc.collect()
                    logger.info(f"🧹 Liberação de memória executada (NFe {i}/{total})")
                    time.sleep(2)  # Pausa breve para estabilizar
                
                self.signals.top_progress.emit(i+1, total)
                logger.info(f"Processando NFe {i+1}/{total}: {nfe_key[:10]}...")
                self.signals.automation_progress.emit(i+1, f"Processando NFe {i+1}/{total}: {nfe_key[:10]}...")
                
                try:
                    # Passo 1: Clica no campo da chave NFe
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Inserindo chave...")
                    pyautogui.click(positions[1][0], positions[1][1])
                    time.sleep(self.wait_times['step_wait'])
                    pyautogui.hotkey('ctrl', 'a')  # Seleciona tudo para substituir
                    pyautogui.write(nfe_key)
                    time.sleep(self.wait_times['step_wait'])
                    logger.debug(f"Chave {nfe_key} inserida")
                    
                    # Passo 2: Clica no campo do captcha e tenta resolver automaticamente
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Processando captcha...")
                    pyautogui.click(positions[2][0], positions[2][1])
                    time.sleep(1)  # Aguarda um segundo para o captcha carregar
                    
                    # Tenta resolver o captcha automaticamente
                    captcha_solved = self.try_solve_captcha()
                    
                    if not captcha_solved:
                        # Se não conseguiu resolver automaticamente, aguarda tempo para resolução manual
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Aguarde esperando captcha...")
                        time.sleep(self.wait_times['captcha'])  # 30 segundos para resolver manualmente
                        logger.debug("Aguardando resolução manual do captcha")
                    else:
                        # Se resolveu automaticamente, aguarda um pouco menos
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Captcha resolvido!")
                        time.sleep(3)
                        logger.debug("Captcha resolvido automaticamente")
                    
                    # Passo 3: Clica em Continuar
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Continuando...")
                    pyautogui.click(positions[3][0], positions[3][1])
                    time.sleep(self.wait_times['continue'])
                    logger.debug("Botão Continuar clicado")
                    
                    # Passo 4: Clica em Download do Documento
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Baixando XML...")
                    pyautogui.click(positions[4][0], positions[4][1])
                    time.sleep(self.wait_times['download'])
                    logger.debug("Botão Download clicado")
                    
                    # Passo 5: Clica no OK do popup
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Confirmando...")
                    pyautogui.click(positions[5][0], positions[5][1])
                    time.sleep(self.wait_times['popup'])
                    logger.debug("Popup OK clicado")
                    
                    # Verifica se o XML foi baixado
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Verificando download...")
                    if self.check_xml_exists(nfe_key):
                        logger.info(f"XML da NFe {nfe_key[:10]} baixado com sucesso")
                        
                        # Passo 6: Clica em Nova Consulta
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Preparando próxima...")
                        pyautogui.click(positions[6][0], positions[6][1])
                        time.sleep(self.wait_times['new_query'])
                        logger.debug("Botão Nova Consulta clicado")
                        
                        # Aguarda um pouco antes da próxima NFe
                        time.sleep(self.wait_times['between_nfe'])
                    else:
                        # XML não encontrado - Nota provavelmente cancelada
                        logger.warning(f"⚠️ XML não encontrado para NFe {nfe_key} - Nota pode estar CANCELADA")
                        self.log_missing_xml(nfe_key)
                        self.signals.xml_not_found.emit(nfe_key)
                        
                        # Se tem configuração de nota cancelada, usa ela
                        if has_canceled_config:
                            self.signals.automation_progress.emit(i+1, f"NFe {i+1}: NOTA CANCELADA - fechando popup...")
                            
                            # Passo extra 7: Clica no OK do popup de NOTA CANCELADA
                            pyautogui.click(positions['canceled_7'][0], positions['canceled_7'][1])
                            time.sleep(self.wait_times['popup'])
                            logger.info("Popup de nota cancelada fechado")
                            
                            # Passo extra 8: Clica no botão de recarregar página
                            self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Reiniciando página...")
                            pyautogui.click(positions['canceled_8'][0], positions['canceled_8'][1])
                            time.sleep(self.wait_times['browser_open'])
                            logger.info("Página reiniciada automaticamente, continuando com próxima NFe")
                        else:
                            # Sem configuração de nota cancelada - apenas recarrega a página
                            self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Reiniciando página...")
                            pyautogui.click(positions[7][0], positions[7][1])  # Usa passo 7 normal (recarregar)
                            time.sleep(self.wait_times['browser_open'])
                            logger.info("Página reiniciada automaticamente (sem configuração de nota cancelada)")
                    
                    self.signals.progress.emit(int((i+1)/total * 100))
                    
                except Exception as e:
                    error_msg = f"Erro ao processar NFe {nfe_key[:10]}: {str(e)}"
                    logger.error(error_msg)
                    logger.error(f"Tipo do erro: {type(e).__name__}")
                    self.signals.automation_progress.emit(i+1, f"Erro na NFe {i+1}: {type(e).__name__}")
                    
                    # Tenta recarregar a página para recuperar de erros
                    try:
                        pyautogui.click(positions[7][0], positions[7][1])
                        time.sleep(self.wait_times['browser_open'])
                        logger.info("Página recarregada após erro")
                    except:
                        logger.error("Falha ao recarregar página após erro")
                    
                    continue
            
            # Libera memória final
            import gc
            gc.collect()
            logger.info("✅ Download automático concluído com sucesso")
            return True
            
        except Exception as e:
            error_msg = f"Erro crítico no download automático: {str(e)}"
            logger.error(error_msg)
            logger.error(f"Traceback: {type(e).__name__}")
            self.signals.error.emit(error_msg)
            
            # Tenta liberar recursos
            try:
                import gc
                gc.collect()
            except:
                pass
            
            return False
    
    def auto_download_selenium(self):
        """Executa o download automático das NFe usando Selenium"""
        try:
            total = len(self.nfe_keys)
            if total == 0:
                error_msg = "Nenhuma NFe para processar!"
                logger.warning(error_msg)
                self.signals.error.emit(error_msg)
                return False
            
            logger.info(f"🌐 Iniciando download automático com Selenium de {total} NFe(s)")
            self.signals.message.emit(f"Iniciando download automático de {total} NFe(s)")
            self.signals.top_progress.emit(0, total)
            
            # Configura o Chrome com opções para download
            options = webdriver.ChromeOptions()
            prefs = {
                "download.default_directory": self.xml_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True
            }
            options.add_experimental_option("prefs", prefs)
            options.add_argument("--start-maximized")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            # Inicializa o driver do Chrome
            logger.info("Inicializando Chrome WebDriver...")
            
            # Usa ChromeDriverManager padrão - ele deve detectar automaticamente
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=options
            )
            
            # Abre o site da Fazenda
            self.driver.get("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=")
            time.sleep(self.wait_times['browser_open'])
            
            for i, nfe_key in enumerate(self.nfe_keys):
                if not self._is_running:
                    break
                
                self.signals.top_progress.emit(i+1, total)
                logger.info(f"Processando NFe {i+1}/{total}: {nfe_key[:10]}...")
                self.signals.automation_progress.emit(i+1, f"Processando NFe {i+1}/{total}: {nfe_key[:10]}...")
                
                try:
                    # Passo 1: Insere a chave da NFe
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Aguardando página carregar...")
                    logger.info(f"🔍 Procurando campo de chave NFe na página...")
                    
                    chave_input = WebDriverWait(self.driver, 30).until(
                        EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtChaveAcesso"))
                    )
                    logger.info("✅ Campo de chave encontrado!")
                    
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Inserindo chave...")
                    chave_input.clear()
                    chave_input.send_keys(nfe_key)
                    time.sleep(self.wait_times['step_wait'])
                    logger.info(f"✅ Chave {nfe_key} inserida com sucesso")
                    
                    # Passo 2: Aguarda e resolve o captcha
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Processando captcha...")
                    
                    # Tenta resolver captcha automaticamente se habilitado
                    captcha_solved = False
                    if self.auto_captcha:
                        captcha_solved = self.try_solve_captcha_selenium()
                    
                    if not captcha_solved:
                        # Aguarda resolução manual
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: 👤 Resolva o captcha manualmente...")
                        logger.info(f"⏳ Aguardando resolução manual do captcha ({self.wait_times['captcha']}s)...")
                        time.sleep(self.wait_times['captcha'])
                    
                    # Passo 3: Clica em Continuar
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Procurando botão Continuar...")
                    logger.info("🔍 Procurando botão Continuar...")
                    
                    continuar_btn = WebDriverWait(self.driver, 30).until(
                        EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_btnConsultar"))
                    )
                    logger.info("✅ Botão Continuar encontrado!")
                    
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Clicando Continuar...")
                    continuar_btn.click()
                    time.sleep(self.wait_times['continue'])
                    logger.info("✅ Botão Continuar clicado com sucesso")
                    
                    # Passo 4: Clica em Download
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Procurando botão Download...")
                    logger.info("🔍 Procurando botão de Download...")
                    
                    download_btn = WebDriverWait(self.driver, 30).until(
                        EC.element_to_be_clickable((By.LINK_TEXT, "Download do Documento Autorizado"))
                    )
                    logger.info("✅ Botão Download encontrado!")
                    
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Baixando XML...")
                    download_btn.click()
                    time.sleep(self.wait_times['download'])
                    logger.info("✅ Botão Download clicado com sucesso")
                    
                    # Verifica se o XML foi baixado
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Verificando download...")
                    if self.check_xml_exists(nfe_key):
                        logger.info(f"XML da NFe {nfe_key[:10]} baixado com sucesso")
                        
                        # Clica em Nova Consulta
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Preparando próxima...")
                        nova_consulta_btn = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnNovaConsulta")
                        nova_consulta_btn.click()
                        time.sleep(self.wait_times['new_query'])
                        logger.debug("Botão Nova Consulta clicado")
                    else:
                        # XML não encontrado - recarrega a página
                        logger.warning(f"XML não encontrado para NFe {nfe_key}")
                        self.log_missing_xml(nfe_key)
                        self.signals.xml_not_found.emit(nfe_key)
                        self.signals.automation_progress.emit(i+1, f"NFe {i+1}: XML não encontrado - reiniciando...")
                        
                        self.driver.refresh()
                        time.sleep(self.wait_times['browser_open'])
                        logger.info("Página reiniciada automaticamente, continuando com próxima NFe")
                    
                    self.signals.progress.emit(int((i+1)/total * 100))
                    
                except Exception as e:
                    error_msg = f"Erro ao processar NFe {nfe_key[:10]}: {str(e)}"
                    logger.error(error_msg)
                    logger.error(f"Tipo do erro: {type(e).__name__}")
                    
                    # Captura screenshot para debug se possível
                    try:
                        screenshot_path = os.path.join(self.xml_folder, f"erro_{nfe_key[:10]}.png")
                        self.driver.save_screenshot(screenshot_path)
                        logger.info(f"Screenshot salvo em: {screenshot_path}")
                    except:
                        pass
                    
                    self.signals.automation_progress.emit(i+1, f"Erro na NFe {i+1}: {type(e).__name__}")
                    
                    # Verifica se o driver crashou
                    driver_alive = True
                    try:
                        _ = self.driver.current_url
                    except:
                        driver_alive = False
                        logger.error("🔴 ChromeDriver crashou! Interrompendo operação.")
                        error_detail = "⚠️ ERRO: O ChromeDriver parou de responder.\n\n"
                        error_detail += "SOLUÇÕES POSSÍVEIS:\n"
                        error_detail += "1. Use o modo PyAutoGUI (desmarque 'Usar Selenium')\n"
                        error_detail += "2. Atualize o Google Chrome para a última versão\n"
                        error_detail += "3. Reinicie o computador e tente novamente\n\n"
                        error_detail += "O modo PyAutoGUI funciona de forma mais estável!"
                        self.signals.error.emit(error_detail)
                        try:
                            self.driver.quit()
                        except:
                            pass
                        return False
                    
                    if not driver_alive:
                        return False
                    
                    # Se o driver ainda está ativo, tenta recarregar a página e continuar
                    try:
                        self.driver.refresh()
                        time.sleep(self.wait_times['browser_open'])
                        logger.info("Página recarregada após erro, continuando...")
                    except:
                        logger.error("Não foi possível recarregar a página")
                        pass
                    continue
            
            logger.info("Download automático com Selenium concluído com sucesso")
            return True
            
        except Exception as e:
            error_msg = f"Erro no download automático com Selenium: {str(e)}"
            logger.error(error_msg)
            self.signals.error.emit(error_msg)
            return False

    def run(self):
        try:
            if self.mode == 'record':
                success = self.record_positions()
            else:
                # Escolhe qual método usar baseado na configuração
                if self.use_selenium and SELENIUM_AVAILABLE:
                    logger.info("🌐 Usando Selenium para automação")
                    success = self.auto_download_selenium()
                else:
                    logger.info("🖱️ Usando PyAutoGUI para automação")
                    success = self.auto_download()
                
            if success:
                msg = "Operação concluída com sucesso!"
                logger.info(msg)
                self.signals.message.emit(msg)
            else:
                msg = "Operação finalizada (completa ou interrompida)"
                logger.info(msg)
                self.signals.message.emit(msg)
                
        except Exception as e:
            error_msg = f"Erro na operação: {str(e)}"
            logger.error(error_msg)
            self.signals.error.emit(error_msg)
        finally:
            # Fecha o navegador do Selenium se estiver aberto
            if self.driver:
                try:
                    self.driver.quit()
                    logger.info("Navegador Selenium fechado")
                except:
                    pass
            self.signals.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HBM XML - Download Automático de NFe")
        
        # Define tamanho mínimo e inicial, mas permite redimensionamento
        self.setMinimumSize(700, 600)
        self.resize(900, 800)  # Tamanho inicial preferencial
        
        # Configura ícone
        if os.path.exists("data/icon.ico"):
            self.setWindowIcon(QIcon("data/icon.ico"))
        
        # Configurações
        self.settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)
        
        # Variáveis de estado
        self.nfe_keys = []
        self.worker = None
        
        # Variáveis para processamento em lote de planilhas
        self.batch_spreadsheets = None  # Lista de planilhas para processar
        self.current_batch_index = 0  # Índice da planilha atual
        self.current_spreadsheet_name = ""  # Nome da planilha atual
        self.recording = False
        self.current_nfe = 0
        self.total_nfes = 0
        self.speed = int(self.settings.value("speed", 3))  # Valor padrão 3 (médio)
        self.overlay = OverlayWindow()
        self.overlay.resize(250, 120)  # Tamanho inicial
        
        # Layout principal
        self.init_ui()
        self.setup_styles()
        
        # Timer para verificar cliques
        self.click_timer = QTimer()
        self.click_timer.timeout.connect(self.check_clicks)
        self.last_click_time = 0
    
    def setup_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QLabel {
                font-size: 12px;
            }
            QPushButton {
                background-color: #87CEFA;  /* Azul claro */
                color: #000000;
                border: 1px solid #4682B4;
                padding: 8px 16px;
                font-size: 12px;
                border-radius: 4px;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #B0E0E6;  /* Azul claro mais claro */
                border: 1px solid #4682B4;
            }
            QPushButton:disabled {
                background-color: #D3D3D3;
                color: #808080;
            }
            QPushButton#danger {
                background-color: #FF6347;  /* Tomate */
                color: white;
            }
            QPushButton#danger:hover {
                background-color: #FF4500;  /* Laranja vermelho */
            }
            QPushButton#success {
                background-color: #90EE90;  /* Verde claro */
                color: #006400;
            }
            QPushButton#success:hover {
                background-color: #98FB98;  /* Verde claro mais claro */
            }
            QLineEdit, QListWidget, QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
            }
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 4px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #87CEFA;  /* Azul claro */
                width: 10px;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 5px;
                margin-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            .step-label {
                font-weight: bold;
                color: #2c3e50;
                font-size: 14px;
            }
            .instruction {
                font-size: 12px;
                color: #7f8c8d;
                margin-bottom: 10px;
            }
            .click-feedback {
                font-size: 11px;
                color: #4682B4;
                font-style: italic;
            }
            .log-viewer {
                font-family: Consolas, Courier New, monospace;
                font-size: 10px;
                background-color: #f0f0f0;
                color: #333;
            }
            #top-progress {
                background-color: transparent;
                border: none;
                height: 3px;
            }
            #top-progress::chunk {
                background-color: #87CEFA;
            }
            QSlider::groove:horizontal {
                height: 8px;
                background: #ddd;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                width: 18px;
                height: 18px;
                margin: -5px 0;
                background: #4682B4;
                border-radius: 9px;
            }
            QSlider::sub-page:horizontal {
                background: #87CEFA;
                border-radius: 4px;
            }
        """)
    
    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        central_widget.setLayout(main_layout)
        
        # Barra de progresso transparente no topo
        self.top_progress = QProgressBar()
        self.top_progress.setObjectName("top-progress")
        self.top_progress.setTextVisible(False)
        self.top_progress.setFixedHeight(3)
        self.top_progress.setRange(0, 100)
        self.top_progress.setValue(0)
        main_layout.addWidget(self.top_progress)
        
        # Cria um widget com scroll para o conteúdo
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        main_layout.addWidget(scroll_area)
        
        # Widget de conteúdo dentro do scroll
        content_widget = QWidget()
        scroll_area.setWidget(content_widget)
        
        content_layout = QVBoxLayout()
        content_widget.setLayout(content_layout)
        
        # Cabeçalho
        self.setup_header(content_layout)
        
        # Linha divisória
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #ddd;")
        content_layout.addWidget(line)
        
        # Corpo principal
        body_layout = QVBoxLayout()
        body_layout.setContentsMargins(20, 10, 20, 10)
        
        # Seção de configurações
        self.setup_config_section(body_layout)
        
        # Seção de entrada de NFe
        self.setup_nfe_section(body_layout)
        
        # Seção de instruções e feedback
        self.setup_feedback_section(body_layout)
        
        # Visualizador de logs
        self.setup_log_section(body_layout)
        
        # Barra de progresso principal
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        body_layout.addWidget(self.progress_bar)
        
        # Status
        self.status_label = QLabel("Pronto para começar. Adicione as NFe e clique em Baixar XMLs.")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-size: 11px; color: #666;")
        body_layout.addWidget(self.status_label)
        
        # Botões de ação
        self.setup_action_buttons(body_layout)
        
        content_layout.addLayout(body_layout)
    
    def setup_header(self, main_layout):
        header = QWidget()
        header_layout = QHBoxLayout()
        header.setLayout(header_layout)
        
        # Logo
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        if os.path.exists("logo.png"):
            logo_pixmap = QPixmap("logo.png")
            logo_label.setPixmap(logo_pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            logo_pixmap = QPixmap()
            logo_pixmap.loadFromData(self.create_default_logo())
            logo_label.setPixmap(logo_pixmap.scaled(120, 120, Qt.KeepAspectRatio))
        
        header_layout.addWidget(logo_label)
        
        # Título
        title_layout = QVBoxLayout()
        title_label = QLabel("HBM XML")
        title_label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2c3e50;")
        subtitle_label = QLabel("Download Automático de NFe da Fazenda")
        subtitle_label.setStyleSheet("font-size: 14px; color: #7f8c8d;")
        
        title_layout.addWidget(title_label)
        title_layout.addWidget(subtitle_label)
        
        # Adiciona um link clicável simples
        info_link = QLabel("<a href='#' style='color: #4682B4; text-decoration: none;'>(i) Informações</a>")
        info_link.setOpenExternalLinks(False)
        info_link.linkActivated.connect(lambda: QMessageBox.information(
            self, 
            "Informações", 
            "Desenvolvido por Kevym Luccas da Controlaria\n\nPara suporte entre em contato via Messenger"
        ))
        info_link.setToolTip("Clique para mais informações")
        info_link.setAlignment(Qt.AlignRight)

        # Layout para alinhar à direita
        info_layout = QHBoxLayout()
        info_layout.addStretch()
        info_layout.addWidget(info_link)
        title_layout.addLayout(info_layout)
        
        title_layout.addStretch()
        header_layout.addLayout(title_layout)
        
        main_layout.addWidget(header)
    
    def setup_config_section(self, body_layout):
        config_group = QGroupBox("Configurações")
        config_layout = QVBoxLayout()
        
        # Controle de velocidade
        speed_layout = QHBoxLayout()
        speed_layout.addWidget(QLabel("Velocidade:"))
        
        self.speed_slider = QSlider(Qt.Horizontal)
        self.speed_slider.setRange(1, 5)
        self.speed_slider.setValue(self.speed)
        self.speed_slider.setTickPosition(QSlider.TicksBelow)
        self.speed_slider.setTickInterval(1)
        self.speed_slider.valueChanged.connect(self.update_speed)
        speed_layout.addWidget(self.speed_slider)
        
        self.speed_label = QLabel(f"")
        self.speed_label.setStyleSheet("font-size: 11px;")
        speed_layout.addWidget(self.speed_label)
        config_layout.addLayout(speed_layout)
        
        # Opção de usar Selenium (se disponível)
        if SELENIUM_AVAILABLE:
            from PyQt5.QtWidgets import QCheckBox
            selenium_layout = QHBoxLayout()
            selenium_layout.addWidget(QLabel("Usar Metodo Anti-Captcha:"))
            
            self.use_selenium_checkbox = QCheckBox()
            self.use_selenium_checkbox.setChecked(False)
            self.use_selenium_checkbox.setToolTip("🌐 Modo anti-captcha.\n✅ Permite resolver captcha automaticamente.")
            self.use_selenium_checkbox.stateChanged.connect(self.on_selenium_checkbox_changed)
            selenium_layout.addWidget(self.use_selenium_checkbox)
            selenium_layout.addStretch()
            
            config_layout.addLayout(selenium_layout)
        
        # Opção de captcha automático (se disponível)
        if HCAPTCHA_AVAILABLE:
            from PyQt5.QtWidgets import QCheckBox
            captcha_layout = QHBoxLayout()
            captcha_layout.addWidget(QLabel("Resolver captcha automaticamente:"))
            
            self.auto_captcha_checkbox = QCheckBox()
            self.auto_captcha_checkbox.setChecked(False)
            self.auto_captcha_checkbox.setToolTip("🤖 Tenta resolver hCaptcha automaticamente.\n✅ Requer Metodo Anti-Captcha ativado.\n⚠️ Se falhar, aguardará resolução manual (30s).")
            self.auto_captcha_checkbox.setEnabled(False)  # Desabilitado até Selenium ser ativado
            captcha_layout.addWidget(self.auto_captcha_checkbox)
            captcha_layout.addStretch()
            
            config_layout.addLayout(captcha_layout)
        
        config_group.setLayout(config_layout)
        body_layout.addWidget(config_group)
        
        # Botão separado para configurar nota cancelada
        btn_canceled_layout = QHBoxLayout()
        btn_canceled_layout.addStretch()
        
        self.btn_config_canceled = QPushButton("⚙️ Configurar Nota Cancelada (Opcional)")
        self.btn_config_canceled.setMaximumHeight(35)
        self.btn_config_canceled.setStyleSheet("font-size: 11px; padding: 8px; background-color: #f39c12; color: white;")
        self.btn_config_canceled.clicked.connect(self.config_canceled_note)
        self.btn_config_canceled.setToolTip("Configure os passos extras para quando uma nota estiver cancelada.\nSó aparecerá quando o XML não for encontrado.")
        btn_canceled_layout.addWidget(self.btn_config_canceled)
        btn_canceled_layout.addStretch()
        
        body_layout.addLayout(btn_canceled_layout)
    
    def setup_nfe_section(self, body_layout):
        nfe_group = QGroupBox("Adicionar NFe")
        nfe_layout = QVBoxLayout()
        
        # Linha superior com entrada de NFe e botão de adicionar
        input_layout = QHBoxLayout()
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("Insira a chave da NFe (44 dígitos) e pressione Enter")
        self.key_input.returnPressed.connect(self.add_nfe)
        input_layout.addWidget(self.key_input)
        
        btn_add = QPushButton("Adicionar")
        btn_add.clicked.connect(self.add_nfe)
        input_layout.addWidget(btn_add)
        nfe_layout.addLayout(input_layout)
        
        # Linha com botões de importar e exportar
        import_export_layout = QHBoxLayout()
        
        btn_import = QPushButton("Importar Planilha")
        btn_import.setToolTip("Importar NFe de uma planilha Excel")
        btn_import.clicked.connect(self.import_spreadsheet)
        import_export_layout.addWidget(btn_import)
        
        btn_export = QPushButton("Exportar Planilha")
        btn_export.setToolTip("Exportar NFe para uma planilha Excel")
        btn_export.clicked.connect(self.export_spreadsheet)
        import_export_layout.addWidget(btn_export)
        
        nfe_layout.addLayout(import_export_layout)
        
        # Lista de NFe
        self.nfe_list = QListWidget()
        self.nfe_list.setMaximumHeight(120)
        nfe_layout.addWidget(QLabel("NFe para processar (Limite: 500):"))
        nfe_layout.addWidget(self.nfe_list)
        
        nfe_group.setLayout(nfe_layout)
        body_layout.addWidget(nfe_group)
    
    def setup_feedback_section(self, left_panel):
        feedback_group = QGroupBox("Andamento do Processo")
        feedback_layout = QVBoxLayout()
        
        # Configuração
        self.config_status = QLabel("")
        self.config_status.setAlignment(Qt.AlignCenter)
        feedback_layout.addWidget(self.config_status)
        
        # Passo atual
        self.step_label = QLabel("")
        self.step_label.setObjectName("step-label")
        self.step_label.setAlignment(Qt.AlignCenter)
        feedback_layout.addWidget(self.step_label)
        
        # Instrução
        self.instruction_label = QLabel("")
        self.instruction_label.setObjectName("instruction")
        self.instruction_label.setAlignment(Qt.AlignCenter)
        self.instruction_label.setWordWrap(True)
        feedback_layout.addWidget(self.instruction_label)
        
        # Feedback de clique
        self.click_feedback = QLabel("")
        self.click_feedback.setObjectName("click-feedback")
        self.click_feedback.setAlignment(Qt.AlignCenter)
        feedback_layout.addWidget(self.click_feedback)
        
        # Progresso da automação
        self.automation_status = QLabel("")
        self.automation_status.setAlignment(Qt.AlignCenter)
        self.automation_status.setWordWrap(True)
        feedback_layout.addWidget(self.automation_status)
        
        # Atualiza status da configuração
        self.update_config_status()
        
        feedback_group.setLayout(feedback_layout)
        left_panel.addWidget(feedback_group)
    
    def setup_log_section(self, body_layout):
        log_group = QGroupBox("Log de Execução")
        log_layout = QVBoxLayout()
        
        self.log_viewer = QTextEdit()
        self.log_viewer.setObjectName("log-viewer")
        self.log_viewer.setReadOnly(True)
        self.log_viewer.setMinimumHeight(100)
        
        # Configura o logger para também escrever no QTextEdit
        log_handler = LogHandler(self.log_viewer)
        log_handler.setFormatter(log_formatter)
        logger.addHandler(log_handler)
        
        log_layout.addWidget(self.log_viewer)
        log_group.setLayout(log_layout)
        body_layout.addWidget(log_group)
    
    def setup_action_buttons(self, body_layout):
        # Linha 1: Botões principais
        btn_action_layout = QHBoxLayout()
        
        self.btn_download = QPushButton("Baixar XMLs")
        self.btn_download.setObjectName("success")
        self.btn_download.clicked.connect(self.start_download)
        btn_action_layout.addWidget(self.btn_download)
        
        self.btn_stop = QPushButton("Parar")
        self.btn_stop.setObjectName("danger")
        self.btn_stop.clicked.connect(self.stop_operation)
        self.btn_stop.setEnabled(False)
        btn_action_layout.addWidget(self.btn_stop)
        
        body_layout.addLayout(btn_action_layout)
        
        # Linha 2: Botões auxiliares (menores)
        btn_aux_layout = QHBoxLayout()
        
        self.btn_retry_missing = QPushButton("🔄 Tentar Baixar XMLs Faltantes")
        self.btn_retry_missing.setMaximumHeight(30)
        self.btn_retry_missing.setStyleSheet("font-size: 11px; padding: 5px;")
        self.btn_retry_missing.clicked.connect(self.retry_missing_xmls)
        self.btn_retry_missing.setToolTip("Carrega e tenta baixar novamente os XMLs do arquivo XMLs_Nao_Encontrados.txt")
        btn_aux_layout.addWidget(self.btn_retry_missing)
        
        self.btn_clear_list = QPushButton("🗑️ Limpar Lista")
        self.btn_clear_list.setMaximumHeight(30)
        self.btn_clear_list.setStyleSheet("font-size: 11px; padding: 5px;")
        self.btn_clear_list.clicked.connect(self.clear_nfe_list)
        self.btn_clear_list.setToolTip("Remove todas as NFes da lista atual")
        btn_aux_layout.addWidget(self.btn_clear_list)
        
        body_layout.addLayout(btn_aux_layout)
    
    def create_default_logo(self):
        from io import BytesIO
        from PIL import Image, ImageDraw, ImageFont
        import base64
        
        img = Image.new('RGB', (150, 150), color=(135, 206, 250))  # Azul claro
        d = ImageDraw.Draw(img)
        
        try:
            font = ImageFont.truetype("arial.ttf", 30)
        except:
            font = ImageFont.load_default()
        
        d.text((30, 50), "HBM", fill=(0, 0, 0), font=font)  # Texto preto
        d.text((30, 90), "XML", fill=(0, 0, 0), font=font)
        
        buffered = BytesIO()
        img.save(buffered, format="PNG")
        return base64.b64decode(base64.b64encode(buffered.getvalue()))
    
    def update_config_status(self):
        """Atualiza o status da configuração na interface"""
        has_config = all(self.settings.value(f"step_{i}_x") is not None for i in range(1, 8))  # 7 passos principais
        has_canceled = (
            self.settings.value("step_canceled_7_x") is not None and
            self.settings.value("step_canceled_8_x") is not None
        )
        
        if has_config:
            status_text = "✔ Configurações de automação prontas"
            if has_canceled:
                status_text += " (+ Nota Cancelada)"
            self.config_status.setText(status_text)
            self.config_status.setStyleSheet("color: #27ae60; font-weight: bold;")
        else:
            self.config_status.setText("⚠ Primeiro uso requer configuração")
            self.config_status.setStyleSheet("color: #e74c3c; font-weight: bold;")
    
    def update_speed(self, value):
        """Atualiza a velocidade de execução"""
        self.speed = value
        self.speed_label.setText(f"")
        self.settings.setValue("speed", value)
        logger.info(f"Velocidade ajustada para: {value}")
    
    def on_selenium_checkbox_changed(self, state):
        """Controla a disponibilidade do checkbox de captcha automático"""
        if HCAPTCHA_AVAILABLE and hasattr(self, 'auto_captcha_checkbox'):
            # Habilita captcha automático apenas se Selenium estiver ativado
            self.auto_captcha_checkbox.setEnabled(state == 2)  # 2 = Checked
            if state != 2:
                self.auto_captcha_checkbox.setChecked(False)
    
    def add_nfe(self):
        key = self.key_input.text().strip()
        if key and len(key) == 44 and key.isdigit():
            if len(self.nfe_keys) >= 500:
                QMessageBox.warning(self, "Limite atingido", "O limite de 500 NFe foi atingido!")
                return
                
            if key not in self.nfe_keys:
                self.nfe_keys.append(key)
                self.nfe_list.addItem(key)
                self.key_input.clear()
                self.status_label.setText(f"NFe adicionada. Total: {len(self.nfe_keys)}/500")
                logger.info(f"NFe adicionada: {key[:10]}...")
            else:
                QMessageBox.warning(self, "Duplicado", "Esta NFe já foi adicionada!")
                logger.warning(f"Tentativa de adicionar NFe duplicada: {key[:10]}...")
        else:
            QMessageBox.warning(self, "Inválido", "Chave de NFe inválida! Deve ter 44 dígitos.")
            logger.warning(f"Tentativa de adicionar NFe inválida: {key}")
    
    def import_spreadsheet(self):
        """Importa NFe de uma ou múltiplas planilhas Excel"""
        try:
            # Pergunta se quer processar múltiplas planilhas
            reply = QMessageBox.question(self, "Importar Planilhas", 
                                        "Deseja importar múltiplas planilhas?\n\n"
                                        "• SIM: Selecione várias planilhas para processar em lote\n"
                                        "  (cada planilha terá sua própria pasta de XMLs)\n\n"
                                        "• NÃO: Importar apenas uma planilha",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # Modo múltiplas planilhas
                file_paths, _ = QFileDialog.getOpenFileNames(
                    self, "Selecionar Planilhas (Múltiplas)", "", 
                    "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)"
                )
                
                if file_paths and len(file_paths) > 0:
                    # Armazena as planilhas para processar em lote
                    self.batch_spreadsheets = file_paths
                    
                    # Carrega a primeira planilha
                    self.current_batch_index = 0
                    self.load_spreadsheet_from_batch(0)
            else:
                # Modo planilha única (comportamento original)
                file_path, _ = QFileDialog.getOpenFileName(
                    self, "Selecionar Planilha", "", 
                    "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)"
                )
                
                if file_path:
                    self.batch_spreadsheets = None  # Desativa modo lote
                    self.load_single_spreadsheet(file_path)
                
        except Exception as e:
            error_msg = f"Erro ao importar planilha: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def load_single_spreadsheet(self, file_path):
        """Carrega uma planilha única (modo antigo)"""
        try:
            if file_path:
                # Lê a planilha usando pandas
                df = pd.read_excel(file_path)
                
                # Encontra colunas que podem conter NFe (44 dígitos)
                nfe_keys = []
                for col in df.columns:
                    # Verifica se a coluna contém strings com 44 dígitos
                    for value in df[col].astype(str):
                        if len(value) == 44 and value.isdigit():
                            if value not in nfe_keys and value not in self.nfe_keys:
                                nfe_keys.append(value)
                                if len(nfe_keys) + len(self.nfe_keys) >= 500:
                                    break
                    if len(nfe_keys) + len(self.nfe_keys) >= 500:
                        break
                
                if not nfe_keys:
                    QMessageBox.warning(self, "Nenhuma NFe encontrada", 
                                      "Não foram encontradas chaves de NFe válidas na planilha.")
                    return
                
                # Adiciona as NFe encontradas
                added = 0
                for key in nfe_keys:
                    if len(self.nfe_keys) < 500:
                        self.nfe_keys.append(key)
                        self.nfe_list.addItem(key)
                        added += 1
                    else:
                        break
                
                self.status_label.setText(f"{added} NFe(s) importadas. Total: {len(self.nfe_keys)}/500")
                QMessageBox.information(self, "Importação concluída", 
                                      f"Foram importadas {added} NFe(s) da planilha.")
                logger.info(f"Importadas {added} NFe(s) da planilha {file_path}")
                
        except Exception as e:
            error_msg = f"Erro ao carregar planilha: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def load_spreadsheet_from_batch(self, index):
        """Carrega uma planilha do lote"""
        try:
            if not self.batch_spreadsheets or index >= len(self.batch_spreadsheets):
                return
            
            file_path = self.batch_spreadsheets[index]
            file_name = os.path.basename(file_path)
            
            logger.info(f"📊 Carregando planilha {index+1}/{len(self.batch_spreadsheets)}: {file_name}")
            
            # Limpa lista atual
            self.nfe_keys.clear()
            self.nfe_list.clear()
            
            # Lê a planilha
            df = pd.read_excel(file_path)
            
            # Encontra NFes
            nfe_keys = []
            for col in df.columns:
                for value in df[col].astype(str):
                    if len(value) == 44 and value.isdigit():
                        if value not in nfe_keys:
                            nfe_keys.append(value)
            
            if not nfe_keys:
                logger.warning(f"⚠️ Nenhuma NFe encontrada em {file_name}")
                QMessageBox.warning(self, "Planilha vazia", 
                                  f"Não foram encontradas NFes válidas em:\n{file_name}\n\nPulando para próxima...")
                # Pula para próxima
                self.process_next_batch_spreadsheet()
                return
            
            # Adiciona as NFes
            for key in nfe_keys[:500]:
                self.nfe_keys.append(key)
                self.nfe_list.addItem(key)
            
            # Armazena nome da planilha atual
            self.current_spreadsheet_name = os.path.splitext(file_name)[0]
            
            self.status_label.setText(f"📊 Planilha {index+1}/{len(self.batch_spreadsheets)}: {len(self.nfe_keys)} NFes carregadas")
            
            QMessageBox.information(self, "Planilha Carregada", 
                                  f"📊 Planilha {index+1} de {len(self.batch_spreadsheets)}\n\n"
                                  f"Arquivo: {file_name}\n"
                                  f"NFes encontradas: {len(self.nfe_keys)}\n\n"
                                  f"Clique em 'Baixar XMLs' para processar esta planilha.")
            
            logger.info(f"✅ Carregadas {len(self.nfe_keys)} NFes da planilha {file_name}")
            
        except Exception as e:
            error_msg = f"Erro ao carregar planilha do lote: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def process_next_batch_spreadsheet(self):
        """Processa a próxima planilha do lote"""
        if not hasattr(self, 'batch_spreadsheets') or not self.batch_spreadsheets:
            return
        
        self.current_batch_index += 1
        
        if self.current_batch_index < len(self.batch_spreadsheets):
            # Ainda tem planilhas para processar
            reply = QMessageBox.question(self, "Próxima Planilha", 
                                        f"✅ Planilha {self.current_batch_index} concluída!\n\n"
                                        f"Deseja carregar a próxima planilha?\n"
                                        f"({self.current_batch_index + 1}/{len(self.batch_spreadsheets)})",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply == QMessageBox.Yes:
                self.load_spreadsheet_from_batch(self.current_batch_index)
            else:
                logger.info("Processamento em lote cancelado pelo usuário")
                self.batch_spreadsheets = None
        else:
            # Todas as planilhas foram processadas
            QMessageBox.information(self, "Lote Concluído", 
                                  f"🎉 Todas as {len(self.batch_spreadsheets)} planilhas foram processadas!\n\n"
                                  f"Processo em lote finalizado.")
            logger.info(f"✅ Lote de {len(self.batch_spreadsheets)} planilhas concluído")
            self.batch_spreadsheets = None
    
    def export_spreadsheet(self):
        """Exporta as NFe para uma planilha Excel"""
        try:
            if not self.nfe_keys:
                QMessageBox.warning(self, "Nenhuma NFe", "Não há NFe para exportar!")
                return
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Salvar Planilha", "NFe_exportadas.xlsx", 
                "Planilhas Excel (*.xlsx);;Todos os arquivos (*)"
            )
            
            if file_path:
                # Garante que a extensão .xlsx está no nome do arquivo
                if not file_path.lower().endswith('.xlsx'):
                    file_path += '.xlsx'
                
                # Cria um DataFrame com as NFe
                df = pd.DataFrame({"Chave_NFe": self.nfe_keys})
                
                # Salva o DataFrame como Excel
                df.to_excel(file_path, index=False)
                
                QMessageBox.information(self, "Exportação concluída", 
                                      f"Planilha com {len(self.nfe_keys)} NFe exportada com sucesso!")
                logger.info(f"Exportadas {len(self.nfe_keys)} NFe para {file_path}")
                
        except Exception as e:
            error_msg = f"Erro ao exportar planilha: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def retry_missing_xmls(self):
        """Carrega XMLs do log de não encontrados e tenta baixar novamente"""
        try:
            log_file = os.path.join(get_executable_dir(), "XMLs_Nao_Encontrados.txt")
            
            if not os.path.exists(log_file):
                QMessageBox.information(self, "Nenhum XML faltante", 
                                      "Não há XMLs faltantes registrados no log!")
                return
            
            # Lê o arquivo e extrai as chaves NFe
            missing_keys = []
            with open(log_file, 'r', encoding='utf-8') as f:
                for line in f:
                    # Formato: "2024-10-17 10:30:00 - NFe: 12345678901234567890123456789012345678901234"
                    if " - NFe: " in line:
                        nfe_key = line.split(" - NFe: ")[1].strip()
                        if len(nfe_key) == 44 and nfe_key.isdigit():
                            if nfe_key not in missing_keys:
                                missing_keys.append(nfe_key)
            
            if not missing_keys:
                QMessageBox.information(self, "Nenhum XML faltante", 
                                      "Não foram encontradas chaves válidas no log!")
                return
            
            # Pergunta se quer limpar a lista atual
            reply = QMessageBox.question(self, "Carregar XMLs Faltantes", 
                                        f"Encontradas {len(missing_keys)} NFe(s) faltantes.\n\n"
                                        f"Deseja SUBSTITUIR a lista atual por estas NFes?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                # Limpa a lista atual
                self.nfe_keys.clear()
                self.nfe_list.clear()
                
                # Adiciona as NFe faltantes
                for key in missing_keys[:500]:  # Limite de 500
                    self.nfe_keys.append(key)
                    self.nfe_list.addItem(key)
                
                self.status_label.setText(f"{len(self.nfe_keys)} NFe(s) faltantes carregadas. Pronto para tentar novamente!")
                logger.info(f"Carregadas {len(self.nfe_keys)} NFe(s) faltantes do log")
                QMessageBox.information(self, "Carregamento concluído", 
                                      f"{len(self.nfe_keys)} NFe(s) faltantes carregadas!\n\n"
                                      f"Clique em 'Baixar XMLs' para tentar novamente.")
            
        except Exception as e:
            error_msg = f"Erro ao carregar XMLs faltantes: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def clear_nfe_list(self):
        """Limpa a lista de NFes carregadas com opções"""
        if not self.nfe_keys:
            QMessageBox.information(self, "Lista vazia", "A lista já está vazia!")
            return
        
        # Cria dialog customizado com 3 botões
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Limpar Lista")
        msg_box.setText(f"Você tem {len(self.nfe_keys)} NFe(s) na lista.\n\nO que deseja fazer?")
        msg_box.setIcon(QMessageBox.Question)
        
        btn_all = msg_box.addButton("🗑️ Limpar Tudo", QMessageBox.YesRole)
        btn_downloaded = msg_box.addButton("✅ Apenas Baixadas", QMessageBox.NoRole)
        btn_cancel = msg_box.addButton("Cancelar", QMessageBox.RejectRole)
        
        msg_box.exec_()
        
        clicked = msg_box.clickedButton()
        
        if clicked == btn_all:
            # Limpa tudo
            count = len(self.nfe_keys)
            self.nfe_keys.clear()
            self.nfe_list.clear()
            self.status_label.setText("Lista de NFes limpa!")
            logger.info(f"Lista de NFes limpa - TODAS ({count} itens removidos)")
            QMessageBox.information(self, "Lista limpa", f"✅ Todas as {count} NFe(s) foram removidas da lista!")
            
        elif clicked == btn_downloaded:
            # Remove apenas as que já foram baixadas
            xml_folder = os.path.join(get_executable_dir(), "XML Concluidos")
            removed_keys = []
            
            for nfe_key in self.nfe_keys[:]:  # Cria cópia para iterar
                xml_filename = f"{nfe_key}.xml"
                xml_path = os.path.join(xml_folder, xml_filename)
                
                if os.path.exists(xml_path):
                    removed_keys.append(nfe_key)
            
            # Remove da lista
            for key in removed_keys:
                self.nfe_keys.remove(key)
                # Remove da lista visual
                items = self.nfe_list.findItems(key, Qt.MatchExactly)
                for item in items:
                    self.nfe_list.takeItem(self.nfe_list.row(item))
            
            remaining = len(self.nfe_keys)
            removed_count = len(removed_keys)
            
            self.status_label.setText(f"Removidas {removed_count} NFe(s) baixadas. Restam {remaining}.")
            logger.info(f"Removidas {removed_count} NFe(s) já baixadas. Restam {remaining} na lista")
            QMessageBox.information(self, "Limpeza Concluída", 
                                  f"✅ Removidas {removed_count} NFe(s) que já foram baixadas.\n\n"
                                  f"📋 Restam {remaining} NFe(s) na lista para processar.")
    
    def config_canceled_note(self):
        """Configura os passos extras para nota cancelada"""
        try:
            # Mensagem de introdução
            reply = QMessageBox.question(self, "Configurar Nota Cancelada", 
                                        "Esta configuração é OPCIONAL e só será usada quando um XML não for encontrado.\n\n"
                                        "Você irá gravar 2 passos:\n"
                                        "• Passo 1: Clicar no OK do popup de 'Nota Cancelada'\n"
                                        "• Passo 2: Clicar no botão para recarregar a página\n\n"
                                        "O navegador será aberto e você deve clicar nas posições quando solicitado.\n\n"
                                        "Deseja continuar?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            
            if reply != QMessageBox.Yes:
                return
            
            # Cria um worker para gravar as posições
            canceled_positions = {}
            
            # Abre o navegador
            webbrowser.open("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=")
            self.status_label.setText("🌐 Navegador aberto...")
            time.sleep(5)  # Aguarda navegador abrir
            
            # Grava passo 1: OK da nota cancelada
            self.status_label.setText("⏳ PASSO 1/2: Clique no OK do popup de 'Nota Cancelada'...")
            self.instruction_label.setText("Clique no botão OK do popup de nota cancelada")
            logger.info("Aguardando clique no OK da nota cancelada...")
            
            # Aguarda o clique do usuário
            initial_pos = pyautogui.position()
            while True:
                time.sleep(0.1)
                current_pos = pyautogui.position()
                if current_pos != initial_pos:
                    # Aguarda soltar o botão
                    time.sleep(0.3)
                    pos1 = pyautogui.position()
                    canceled_positions[7] = pos1
                    self.click_feedback.setText(f"✅ Posição 1 gravada: {pos1}")
                    logger.info(f"Posição cancelada passo 7 salva: {pos1}")
                    break
            
            time.sleep(1)
            
            # Grava passo 2: Recarregar página
            self.status_label.setText("⏳ PASSO 2/2: Clique no botão de recarregar página...")
            self.instruction_label.setText("Clique no botão para Recarregar/Atualizar a página (F5 ou botão reload)")
            logger.info("Aguardando clique no botão de recarregar...")
            
            # Aguarda o clique do usuário
            initial_pos = pyautogui.position()
            while True:
                time.sleep(0.1)
                current_pos = pyautogui.position()
                if current_pos != initial_pos:
                    # Aguarda soltar o botão
                    time.sleep(0.3)
                    pos2 = pyautogui.position()
                    canceled_positions[8] = pos2
                    self.click_feedback.setText(f"✅ Posição 2 gravada: {pos2}")
                    logger.info(f"Posição cancelada passo 8 salva: {pos2}")
                    break
            
            # Salva as posições
            self.settings.setValue("step_canceled_7_x", canceled_positions[7][0])
            self.settings.setValue("step_canceled_7_y", canceled_positions[7][1])
            self.settings.setValue("step_canceled_8_x", canceled_positions[8][0])
            self.settings.setValue("step_canceled_8_y", canceled_positions[8][1])
            
            # Sucesso!
            self.update_config_status()
            self.status_label.setText("✅ Configuração de nota cancelada concluída!")
            self.instruction_label.setText("")
            self.click_feedback.setText("")
            
            QMessageBox.information(self, "Configuração Concluída!", 
                                  "✅ Passos de nota cancelada gravados com sucesso!\n\n"
                                  "Posições salvas:\n"
                                  f"• Passo 1 (OK cancelada): {canceled_positions[7]}\n"
                                  f"• Passo 2 (Recarregar): {canceled_positions[8]}\n\n"
                                  "Agora, quando um XML não for encontrado, o sistema irá:\n"
                                  "1. Fechar o popup de nota cancelada\n"
                                  "2. Recarregar a página automaticamente")
            logger.info("Configuração de nota cancelada concluída com sucesso")
            
        except Exception as e:
            error_msg = f"Erro ao configurar nota cancelada: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def mousePressEvent(self, event):
        """Captura cliques quando estiver no modo de gravação"""
        if self.recording and self.worker and self.worker.current_step > 0:
            pos = pyautogui.position()
            self.worker.positions[self.worker.current_step] = pos
            self.status_label.setText(f"Posição {self.worker.current_step} gravada: {pos}")

    def update_overlay(self, current, total, status):
        """Atualiza a janela de overlay com o progresso atual"""
        if not self.overlay.isVisible():
            self.overlay.show()
        self.overlay.update_progress(current, total, status)

    def check_clicks(self):
        """Verifica cliques durante a gravação"""
        if self.recording and self.worker and self.worker.current_step > 0:
            # Implementação adicional se necessário
            pass
    
    def start_download(self):
        """Inicia o processo de download"""
        try:
            # Aviso importante sobre configuração da pasta de download
            xml_folder = os.path.join(get_executable_dir(), "XML Concluidos")
            
            reply_folder = QMessageBox.information(
                self, '⚠ CONFIGURAÇÃO IMPORTANTE',
                f'<b>ANTES DE COMEÇAR:</b><br><br>'
                f'Configure o navegador para salvar os XMLs na pasta:<br>'
                f'<b>{xml_folder}</b><br><br>'
                f'<b>Como configurar:</b><br>'
                f'1. No navegador, vá em Configurações → Downloads<br>'
                f'2. Defina o local de download para a pasta acima<br>'
                f'3. Marque "Perguntar onde salvar cada arquivo" como DESATIVADO<br><br>'
                f'<b>Esta configuração é essencial para o sistema verificar se os XMLs foram baixados corretamente!</b><br><br>'
                f'Clique OK quando estiver configurado.',
                QMessageBox.Ok)
            
            # Verifica se já tem configurações salvas
            has_config = all(self.settings.value(f"step_{i}_x") is not None for i in range(1, 8))  # 7 passos
            
            if not has_config and not self.recording:
                # Primeiro uso - precisa gravar as posições
                reply = QMessageBox.question(
                    self, 'Configuração Necessária',
                    'É a primeira vez que você usa este programa. '
                    'Precisamos gravar as posições dos cliques no site da Fazenda.\n\n'
                    'Um navegador será aberto. Siga as instruções cuidadosamente.\n\n'
                    'Deseja continuar?',
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                
                if reply == QMessageBox.Yes:
                    self.recording = True
                    self.status_label.setText("Modo de gravação ativado. Siga as instruções.")
                    self.btn_download.setEnabled(False)
                    self.btn_stop.setEnabled(True)
                    self.click_timer.start(100)  # Verifica cliques a cada 100ms
                    
                    self.worker = NFeDownloader(
                        nfe_keys=self.nfe_keys,
                        settings=self.settings,
                        mode='record',
                        speed=self.speed
                    )
                    self.worker.signals.message.connect(self.update_status)
                    self.worker.signals.capture_step.connect(self.update_instruction)
                    self.worker.signals.browser_ready.connect(self.on_browser_ready)
                    self.worker.signals.finished.connect(self.on_worker_finished)
                    self.worker.signals.error.connect(self.show_error)
                    self.worker.signals.click_recorded.connect(self.on_click_recorded)
                    self.overlay.show()
                    self.worker.start()
                return
            
            # Se já tem configuração ou está gravando
            if self.nfe_keys:
                self.total_nfes = len(self.nfe_keys)
                self.current_nfe = 0
                self.btn_download.setEnabled(False)
                self.btn_stop.setEnabled(True)
                
                # Verifica se a opção de captcha automático está habilitada
                auto_captcha = False
                if HCAPTCHA_AVAILABLE and hasattr(self, 'auto_captcha_checkbox'):
                    auto_captcha = self.auto_captcha_checkbox.isChecked()
                
                # Verifica se deve usar Selenium
                use_selenium = False
                if SELENIUM_AVAILABLE and hasattr(self, 'use_selenium_checkbox'):
                    use_selenium = self.use_selenium_checkbox.isChecked()
                
                self.worker = NFeDownloader(
                    nfe_keys=self.nfe_keys,
                    settings=self.settings,
                    mode='auto',
                    speed=self.speed,
                    auto_captcha=auto_captcha,
                    use_selenium=use_selenium
                )
                self.worker.signals.message.connect(self.update_status)
                self.worker.signals.progress.connect(self.update_progress)
                self.worker.signals.finished.connect(self.on_worker_finished)
                self.worker.signals.error.connect(self.show_error)
                self.worker.signals.automation_progress.connect(self.update_automation_status)
                self.worker.signals.top_progress.connect(self.update_top_progress)
                self.worker.signals.xml_not_found.connect(self.on_xml_not_found)
                self.worker.start()
            else:
                QMessageBox.warning(self, "Aviso", "Adicione pelo menos uma NFe para processar!")
                logger.warning("Tentativa de iniciar sem NFe adicionadas")
                
        except Exception as e:
            error_msg = f"Erro ao iniciar download: {str(e)}"
            logger.error(error_msg)
            self.show_error(error_msg)
    
    def stop_operation(self):
        if self.worker:
            self.worker.stop()
            self.status_label.setText("Operação interrompida pelo usuário")
            self.btn_stop.setEnabled(False)
            self.btn_download.setEnabled(True)
            self.click_timer.stop()
            logger.info("Operação interrompida pelo usuário")
    
    def on_xml_not_found(self, nfe_key):
        """Chamado quando um XML não é encontrado"""
        logger.warning(f"XML não encontrado para NFe: {nfe_key}")
        self.status_label.setText(f"⚠ XML não encontrado: {nfe_key[:10]}... - Registrado no log")
        
        # Mostra mensagem ao usuário (não bloqueia o processo)
        msg = (f'O XML da NFe {nfe_key[:10]}... não foi encontrado na pasta "XML Concluidos".\n\n'
               'Esta nota foi registrada no log "XMLs_Nao_Encontrados.txt".\n\n'
               'Possíveis causas:\n'
               '- Nota cancelada\n'
               '- Erro no download\n'
               '- Pasta de download incorreta\n\n'
               'O sistema continuará com a próxima nota automaticamente.')
        logger.info(msg)
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def update_top_progress(self, current, total):
        """Atualiza a barra de progresso transparente no topo"""
        self.top_progress.setMaximum(total)
        self.top_progress.setValue(current)
    
    def update_status(self, message):
        self.status_label.setText(message)
    
    def update_instruction(self, step):
        instructions = {
            1: "Clique no campo onde deve ser inserida a chave da NFe",
            2: "Clique no campo do captcha",
            3: "Clique no botão 'Continuar'",
            4: "Clique no botão 'Download do Documento'",
            5: "Clique no OK do popup de confirmação",
            6: "Clique no botão 'Nova Consulta'",
            7: "Aguardando próxima NFe"
        }
        
        self.step_label.setText(f"PASSO {step}/7")
        self.instruction_label.setText(instructions.get(step, ""))
    
    def update_automation_status(self, current, status):
        self.current_nfe = current
        self.automation_status.setText(status)
        self.update_overlay(current, self.total_nfes, status)
    
    def on_click_recorded(self, step, x, y):
        """Quando um clique é registrado durante a gravação"""
        self.click_feedback.setText(f"Passo {step} registrado: X={x}, Y={y}")
    
    def on_browser_ready(self):
        """Quando o navegador estiver pronto"""
        self.status_label.setText("Navegador aberto. Siga as instruções na tela.")
        logger.info("Navegador aberto e pronto para interação")
    
    def on_worker_finished(self):
        """Quando o worker terminar sua tarefa"""
        self.recording = False
        self.btn_download.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.click_timer.stop()
        self.overlay.hide()
        
        if self.progress_bar.value() == 100:
            self.status_label.setText("Operação concluída com sucesso!")
            logger.info("Operação concluída com sucesso!")
            
            # Se está em modo lote, move XMLs e processa próxima planilha
            if hasattr(self, 'batch_spreadsheets') and self.batch_spreadsheets:
                self.move_xmls_to_folder_and_continue()
        else:
            self.status_label.setText("Operação finalizada")
            logger.info("Operação finalizada")
        
        # Atualiza o status da configuração
        self.update_config_status()
    
    def move_xmls_to_folder_and_continue(self):
        """Move XMLs para pasta específica e continua para próxima planilha"""
        try:
            if not hasattr(self, 'current_spreadsheet_name'):
                return
            
            # Cria pasta com nome da planilha
            base_folder = get_executable_dir()
            spreadsheet_folder = os.path.join(base_folder, f"XMLs_{self.current_spreadsheet_name}")
            
            if not os.path.exists(spreadsheet_folder):
                os.makedirs(spreadsheet_folder)
                logger.info(f"📁 Pasta criada: {spreadsheet_folder}")
            
            # Move todos os XMLs da pasta "XML Concluidos" para a pasta da planilha
            xml_concluidos = os.path.join(base_folder, "XML Concluidos")
            moved_count = 0
            
            if os.path.exists(xml_concluidos):
                for filename in os.listdir(xml_concluidos):
                    if filename.endswith('.xml'):
                        src = os.path.join(xml_concluidos, filename)
                        dst = os.path.join(spreadsheet_folder, filename)
                        
                        try:
                            # Move o arquivo
                            import shutil
                            shutil.move(src, dst)
                            moved_count += 1
                        except Exception as e:
                            logger.error(f"Erro ao mover {filename}: {e}")
            
            logger.info(f"✅ {moved_count} XMLs movidos para {spreadsheet_folder}")
            
            # Mostra resumo e pergunta sobre próxima planilha
            QMessageBox.information(self, "Planilha Concluída", 
                                  f"✅ Planilha processada com sucesso!\n\n"
                                  f"📁 {moved_count} XMLs movidos para:\n{spreadsheet_folder}\n\n"
                                  f"Preparando próxima planilha...")
            
            # Processa próxima planilha
            self.process_next_batch_spreadsheet()
            
        except Exception as e:
            error_msg = f"Erro ao mover XMLs: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
    def show_error(self, error):
        logger.error(error)
        QMessageBox.critical(self, "Erro", error)
        self.on_worker_finished()

class OverlayWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("background: transparent;")
        
        # Layout principal
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Widget de conteúdo
        self.content = QWidget()
        self.content.setStyleSheet("""
            background-color: rgba(255, 255, 255, 220);
            border-radius: 8px;
            border: 1px solid #ddd;
        """)
        
        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(10, 10, 10, 10)
        
        # Título
        self.title_label = QLabel("Progresso HBM XML")
        self.title_label.setStyleSheet("""
            font-weight: bold;
            font-size: 12px;
            color: #2c3e50;
            padding-bottom: 5px;
            border-bottom: 1px solid #eee;
        """)
        content_layout.addWidget(self.title_label)
        
        # Status
        self.status_label = QLabel("Pronto para começar")
        self.status_label.setStyleSheet("font-size: 11px; color: #666;")
        self.status_label.setWordWrap(True)
        content_layout.addWidget(self.status_label)
        
        # Progresso
        self.progress_label = QLabel("0/0 NFe processadas")
        self.progress_label.setStyleSheet("font-size: 11px; color: #4682B4;")
        content_layout.addWidget(self.progress_label)
        
        # Tempo estimado
        self.time_label = QLabel("Tempo estimado: --")
        self.time_label.setStyleSheet("font-size: 10px; color: #7f8c8d; font-style: italic;")
        content_layout.addWidget(self.time_label)
        
        self.content.setLayout(content_layout)
        layout.addWidget(self.content)
        self.setLayout(layout)
        
        # Variáveis para cálculo de tempo
        self.start_time = None
        self.last_update = None
        
    def update_progress(self, current, total, status):
        """Atualiza o overlay com o progresso atual"""
        if self.start_time is None:
            self.start_time = time.time()
        
        self.progress_label.setText(f"{current}/{total} NFe processadas")
        self.status_label.setText(status)
        
        # Calcula tempo estimado
        if current > 0:
            elapsed = time.time() - self.start_time
            remaining = (elapsed / current) * (total - current)
            self.time_label.setText(f"Tempo estimado: {self.format_time(remaining)}")
        
        # Ajusta posição para canto superior direito
        screen_geometry = QApplication.desktop().availableGeometry()
        self.move(screen_geometry.right() - self.width() - 20, 20)
        
    def format_time(self, seconds):
        """Formata segundos em minutos:segundos"""
        minutes = int(seconds // 60)
        seconds = int(seconds % 60)
        return f"{minutes:02d}:{seconds:02d}"
    
    def showEvent(self, event):
        """Ajusta a posição quando mostrado"""
        self.adjust_position()
        super().showEvent(event)
    
    def adjust_position(self):
        """Ajusta a posição para o canto superior direito"""
        screen_geometry = QApplication.desktop().availableGeometry()
        self.move(screen_geometry.right() - self.width() - 20, 20)

class LogHandler(logging.Handler):
    """Handler personalizado para enviar logs para o QTextEdit"""
    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
    
    def emit(self, record):
        msg = self.format(record)
        self.text_edit.append(msg)
        # Auto-scroll
        self.text_edit.verticalScrollBar().setValue(self.text_edit.verticalScrollBar().maximum())

    def closeEvent(self, event):
        self.overlay.close()
        super().closeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Configuração de estilo adicional
    app.setStyle('Fusion')
    
    # Define a paleta de cores
    palette = app.palette()
    palette.setColor(QPalette.Window, QColor(245, 245, 245))
    palette.setColor(QPalette.WindowText, QColor(51, 51, 51))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))
    palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
    palette.setColor(QPalette.ToolTipText, QColor(51, 51, 51))
    palette.setColor(QPalette.Text, QColor(51, 51, 51))
    palette.setColor(QPalette.Button, QColor(240, 240, 240))
    palette.setColor(QPalette.ButtonText, QColor(51, 51, 51))
    palette.setColor(QPalette.BrightText, QColor(255, 0, 0))
    palette.setColor(QPalette.Highlight, QColor(135, 206, 250))  # Azul claro
    palette.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
    app.setPalette(palette)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
