import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QProgressBar, QTextEdit, QHBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
import pdfplumber
import openpyxl
import os
import config

# Função para extrair texto do PDF
def extrair_texto_pdf(pdf_path):
    try:
        texto_total = ""
        with pdfplumber.open(pdf_path) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                texto_total += texto
        return texto_total
    except Exception as e:
        print(f"Erro ao extrair texto do PDF: {e}")
        return ""

# Função para formatar os dados como uma tabela
def formatar_como_tabela(dados):
    tabela = [config.CONFIG["colunas"]]  # Cabeçalhos
    for valor in dados:
        valores_colunas = valor.split()  # Divida as linhas em colunas
        if valores_colunas:
            tabela.append(valores_colunas)  # Adiciona a linha formatada
    return tabela

# Função para exibir popups de sucesso ou erro
def exibir_popup(mensagem, titulo, app):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(mensagem)
    msg.setWindowTitle(titulo)
    msg.exec_()

# Função para processar o Excel e lidar com células mescladas
def processar_excel(dados, molde_path, progress_bar, progress_label, app):
    try:
        if not os.path.exists(molde_path):
            raise Exception(f"Arquivo de molde não encontrado: {molde_path}")
        
        wb = openpyxl.load_workbook(molde_path)
        sheet = wb.active

        # Preencher dados a partir da linha 62 e coluna D (linha 62, coluna 4)
        linha_inicio = 62
        coluna_inicio = 4

        # Preencher os dados sem alterar a formatação das células existentes
        for row_index, linha in enumerate(dados, start=linha_inicio):
            for col_index, valor in enumerate(linha, start=coluna_inicio):
                cell = sheet.cell(row=row_index, column=col_index)
                # Atribuir valor à célula, sem modificar a formatação
                cell.value = valor

        # Salvar o arquivo Excel
        excel_saida_path = app.excel_save_path
        if not excel_saida_path:
            exibir_popup("Erro: Nenhum local selecionado para salvar o arquivo.", "Erro", app)
            return

        wb.save(excel_saida_path)

        # Notificar sucesso
        exibir_popup(f"Arquivo salvo em: {excel_saida_path}", "Sucesso", app)
        progress_bar.setValue(100)
        progress_label.setText("Concluído")
    except Exception as e:
        erro_msg = f"Erro ao salvar o arquivo Excel: {str(e)}"
        print(erro_msg)
        exibir_popup(erro_msg, "Erro", app)

# Classe principal do aplicativo
class MyApp(QWidget):
    def __init__(self):
        super().__init__()

        self.pdf_path = None
        self.molde_path = None
        self.excel_save_path = None

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Extrator e Processador de Dados PDF para Excel') #titulo do app
        self.setGeometry(100, 100, 500, 400)

        # Definir o fundo laranja da janela
        self.setStyleSheet("background-color: orange;")

        layout = QVBoxLayout()

        # Título com texto em branco
        self.title_label = QLabel("Aut 3º ", self)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: white;")
        layout.addWidget(self.title_label)

        # Instruções com texto em branco
        self.instructions_label = QLabel("Escolha um PDF, extraia os dados e depois salve no molde Excel.", self)
        self.instructions_label.setAlignment(Qt.AlignCenter)
        self.instructions_label.setStyleSheet("color: white;")
        layout.addWidget(self.instructions_label)

        # Progresso com texto em branco
        self.progress_label = QLabel("0% Concluído", self)
        self.progress_label.setStyleSheet("color: white;")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Usando QTextEdit em vez de QLabel para permitir edição
        self.text_input = QTextEdit(self)
        self.text_input.setPlaceholderText("Texto extraído do PDF")
        self.text_input.setReadOnly(False)  # Agora pode ser editado
        layout.addWidget(self.text_input)

        # Botões
        buttons_layout = QHBoxLayout()

        self.escolher_pdf_button = QPushButton("Escolher PDF", self)
        self.escolher_pdf_button.clicked.connect(self.escolher_pdf)
        buttons_layout.addWidget(self.escolher_pdf_button)

        self.escolher_molde_button = QPushButton("Escolher Molde Excel", self)
        self.escolher_molde_button.clicked.connect(self.escolher_molde)
        buttons_layout.addWidget(self.escolher_molde_button)

        self.confirmar_button = QPushButton("Confirmar Extração", self)
        self.confirmar_button.setEnabled(False)
        self.confirmar_button.clicked.connect(self.confirmar_extracao)
        buttons_layout.addWidget(self.confirmar_button)

        self.inserir_no_molde_button = QPushButton("Inserir no Molde Excel", self)
        self.inserir_no_molde_button.setEnabled(False)
        self.inserir_no_molde_button.clicked.connect(self.inserir_no_molde)
        buttons_layout.addWidget(self.inserir_no_molde_button)

        layout.addLayout(buttons_layout)

        self.extrair_novamente_button = QPushButton("Extrair Novamente?", self)
        self.extrair_novamente_button.setEnabled(False)
        self.extrair_novamente_button.clicked.connect(self.resetar_processo)
        layout.addWidget(self.extrair_novamente_button)

        self.toggle_table_button = QPushButton("Alternar para Visualizar Tabela", self)
        self.toggle_table_button.clicked.connect(self.toggle_table)
        layout.addWidget(self.toggle_table_button)

        self.setLayout(layout)

    def escolher_pdf(self):
        """Escolher o arquivo PDF usando o QFileDialog"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "Arquivos PDF (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.confirmar_button.setEnabled(True)

    def escolher_molde(self):
        """Escolher o arquivo de Molde Excel usando o QFileDialog"""
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolher Molde Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.molde_path = file_path
            self.inserir_no_molde_button.setEnabled(True)

    def confirmar_extracao(self):  # Alteração aqui
        """Método para confirmar a extração do texto do PDF.""" 
        if not self.pdf_path:
            exibir_popup("Erro: Por favor, selecione um arquivo PDF.", "Erro", self)
            return
        texto_extraído = extrair_texto_pdf(self.pdf_path)
        self.text_input.setText(texto_extraído)
        self.confirmar_button.setEnabled(False)
        self.inserir_no_molde_button.setEnabled(True)
        self.extrair_novamente_button.setEnabled(True)

    def inserir_no_molde(self):
        """Método para inserir os dados extraídos no molde Excel.""" 
        if not self.molde_path or not self.pdf_path:
            exibir_popup("Erro: Por favor, selecione tanto o arquivo PDF quanto o molde Excel.", "Erro", self)
            return

        texto_extraído = self.text_input.toPlainText()
        dados_extraídos = texto_extraído.split("\n")  # Dividir por linhas
        dados_formatados = formatar_como_tabela(dados_extraídos)

        # Atualiza a barra de progresso
        self.progress_bar.setValue(50)
        self.progress_label.setText("50% Concluído")

        # Solicitar o caminho para salvar o arquivo
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.excel_save_path = file_path
            processar_excel(dados_formatados, self.molde_path, self.progress_bar, self.progress_label, self)

    def resetar_processo(self):  
        """Método para reiniciar a extração.""" 
        self.text_input.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("0% Concluído")
        self.confirmar_button.setEnabled(True)
        self.inserir_no_molde_button.setEnabled(False)
        self.extrair_novamente_button.setEnabled(False)

    def toggle_table(self):
        """Alternar entre visualizar texto ou tabela.""" 
        tabela = formatar_como_tabela(self.text_input.toPlainText().split("\n"))
        self.text_input.setText("\n".join(["\t".join(linha) for linha in tabela]))  # Mostrar como texto com separação
        self.toggle_table_button.setText("Alternar para Visualizar Texto")

#inicia operação
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myapp = MyApp()
    myapp.show()
    sys.exit(app.exec_())
