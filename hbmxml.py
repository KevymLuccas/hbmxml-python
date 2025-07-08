import os
import sys
import time
import logging
from logging.handlers import RotatingFileHandler
import re
import json
from threading import Thread
from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, 
                            QLineEdit, QPushButton, QListWidget, QProgressBar, QFileDialog, 
                            QMessageBox, QSizePolicy, QGroupBox, QFrame, QComboBox, QTextEdit,
                            QSlider, QSpinBox)
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QSettings, QTimer
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette
import webbrowser
import pyautogui
import pygetwindow as gw
import pandas as pd
from openpyxl import Workbook


# Configuração de logging
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler = RotatingFileHandler('hbm_xml.log', maxBytes=5*1024*1024, backupCount=3)
log_handler.setFormatter(log_formatter)
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

# Configurações
SETTINGS_FILE = "hbm_xml_settings.ini"

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

class NFeDownloader(Thread):
    def __init__(self, nfe_keys, settings, mode='record', speed=3):
        super().__init__()
        self.nfe_keys = nfe_keys
        self.settings = settings
        self.mode = mode  # 'record' or 'auto'
        self.speed = max(1, min(5, speed))  # Garante valor entre 1 e 5
        self.signals = WorkerSignals()
        self._is_running = True
        self.current_step = 0
        self.positions = {}
        
        # Tempos de espera base (para velocidade 3) e ajustados pela velocidade
        self.base_wait_times = {
            'browser_open': 5,
            'step_wait': 1,
            'captcha': 3,
            'continue': 5,
            'download': 3,
            'popup': 2,
            'new_query': 3,
            'between_nfe': 2
        }
        self.wait_times = self.calculate_wait_times()
        
        # Passos do processo
        self.steps = {
            1: "Selecione o local para inserir a chave da NFe",
            2: "Selecione o campo do captcha",
            3: "Clique no botão Continuar",
            4: "Clique no botão Download do Documento",
            5: "Clique no OK do popup",
            6: "Clique no botão Nova Consulta",
            7: "Aguardando próxima NFe"
        }

    def calculate_wait_times(self):
        """Calcula os tempos de espera com base na velocidade selecionada"""
        # Velocidade 3 = tempos base, 1 = mais lento, 5 = mais rápido
        factor = {1: 2.0, 2: 1.5, 3: 1.0, 4: 0.75, 5: 0.5}[self.speed]
        return {k: v * factor for k, v in self.base_wait_times.items()}

    def stop(self):
        self._is_running = False
        logger.info("Operação interrompida pelo usuário")

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
                
                # Salva também a pasta de download se estiver configurada
                if hasattr(self, 'download_folder'):
                    self.settings.setValue("download_folder", self.download_folder)
                
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
            
            # Carrega as posições salvas
            positions = {}
            for step in range(1, 8):  # Temos 7 passos
                x = self.settings.value(f"step_{step}_x", None)
                y = self.settings.value(f"step_{step}_y", None)
                if x is None or y is None:
                    error_msg = f"Posição do passo {step} não configurada!"
                    logger.error(error_msg)
                    self.signals.error.emit(error_msg)
                    return False
                positions[step] = (int(x), int(y))
                logger.debug(f"Posição {step} carregada: {positions[step]}")
            
            # Carrega a pasta de download salva
            self.download_folder = self.settings.value("download_folder", os.path.expanduser("~\\Downloads"))
            logger.info(f"Pasta de download definida como: {self.download_folder}")
            
            # Abre o navegador apenas uma vez
            webbrowser.open("https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx?tipoConsulta=resumo&tipoConteudo=7PhJ+gAVw2g=")
            self.signals.browser_ready.emit()
            time.sleep(self.wait_times['browser_open'])
            
            for i, nfe_key in enumerate(self.nfe_keys):
                if not self._is_running:
                    break
                
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
                    
                    # Passo 2: Clica no campo do captcha
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Aguardando captcha...")
                    pyautogui.click(positions[2][0], positions[2][1])
                    time.sleep(self.wait_times['captcha'])  # Tempo para resolver o captcha manualmente
                    logger.debug("Campo captcha clicado, aguardando resolução")
                    
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
                    
                    # Passo 6: Clica em Nova Consulta
                    self.signals.automation_progress.emit(i+1, f"NFe {i+1}: Preparando próxima...")
                    pyautogui.click(positions[6][0], positions[6][1])
                    time.sleep(self.wait_times['new_query'])
                    logger.debug("Botão Nova Consulta clicado")
                    
                    # Passo 7: Aguarda um pouco antes da próxima NFe
                    time.sleep(self.wait_times['between_nfe'])
                    
                    self.signals.progress.emit(int((i+1)/total * 100))
                    
                except Exception as e:
                    error_msg = f"Erro ao processar NFe {nfe_key[:10]}: {str(e)}"
                    logger.error(error_msg)
                    self.signals.automation_progress.emit(i+1, f"Erro: {error_msg}")
                    continue
            
            logger.info("Download automático concluído com sucesso")
            return True
            
        except Exception as e:
            error_msg = f"Erro no download automático: {str(e)}"
            logger.error(error_msg)
            self.signals.error.emit(error_msg)
            return False

    def run(self):
        try:
            if self.mode == 'record':
                success = self.record_positions()
            else:
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
            self.signals.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HBM XML - Download Automático de NFe")
        self.setFixedSize(900, 800)  # Aumentado para acomodar novos controles
        
        # Configura ícone
        if os.path.exists("data/icon.ico"):
            self.setWindowIcon(QIcon("data/icon.ico"))
        
        # Configurações
        self.settings = QSettings(SETTINGS_FILE, QSettings.IniFormat)
        
        # Variáveis de estado
        self.nfe_keys = []
        self.worker = None
        self.recording = False
        self.current_nfe = 0
        self.total_nfes = 0
        self.download_folder = self.settings.value("download_folder", os.path.expanduser("~\\Downloads"))
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
        central_widget.setLayout(main_layout)
        
        # Barra de progresso transparente no topo
        self.top_progress = QProgressBar()
        self.top_progress.setObjectName("top-progress")
        self.top_progress.setTextVisible(False)
        self.top_progress.setFixedHeight(3)
        self.top_progress.setRange(0, 100)
        self.top_progress.setValue(0)
        main_layout.addWidget(self.top_progress)
        
        # Cabeçalho
        self.setup_header(main_layout)
        
        # Linha divisória
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #ddd;")
        main_layout.addWidget(line)
        
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
        
        main_layout.addLayout(body_layout)
    
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
        
        config_group.setLayout(config_layout)
        body_layout.addWidget(config_group)
    
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
        has_config = all(self.settings.value(f"step_{i}_x") is not None for i in range(1, 8))  # 7 passos
        
        if has_config:
            self.config_status.setText("✔ Configurações de automação prontas")
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
        """Importa NFe de uma planilha Excel"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Selecionar Planilha", "", 
                "Planilhas Excel (*.xlsx *.xls);;Todos os arquivos (*)"
            )
            
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
            error_msg = f"Erro ao importar planilha: {str(e)}"
            logger.error(error_msg)
            QMessageBox.critical(self, "Erro", error_msg)
    
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
                    self.worker.download_folder = self.download_folder  # Passa a pasta de download
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
                
                self.worker = NFeDownloader(
                    nfe_keys=self.nfe_keys,
                    settings=self.settings,
                    mode='auto',
                    speed=self.speed
                )
                self.worker.download_folder = self.download_folder  # Passa a pasta de download
                self.worker.signals.message.connect(self.update_status)
                self.worker.signals.progress.connect(self.update_progress)
                self.worker.signals.finished.connect(self.on_worker_finished)
                self.worker.signals.error.connect(self.show_error)
                self.worker.signals.automation_progress.connect(self.update_automation_status)
                self.worker.signals.top_progress.connect(self.update_top_progress)
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
            logger.info("Operação concluída com sucesso")
        else:
            self.status_label.setText("Operação finalizada")
            logger.info("Operação finalizada")
        
        # Atualiza o status da configuração
        self.update_config_status()
    
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
