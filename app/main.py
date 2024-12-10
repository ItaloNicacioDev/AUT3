import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QFileDialog, QProgressBar, QTextEdit, QHBoxLayout, QMessageBox, QDialog, QScrollArea
from PyQt5.QtCore import Qt
import pdfplumber
import openpyxl
import config


def exibir_popup(mensagem, titulo, app, tipo=QMessageBox.Information):
    """Exibe uma janela popup com a mensagem."""
    msg = QMessageBox()
    msg.setIcon(tipo)
    msg.setText(mensagem)
    msg.setWindowTitle(titulo)
    msg.exec_()


def extrair_texto_pdf(pdf_path):
    """Extrai o texto de um arquivo PDF usando pdfplumber."""
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


def formatar_como_tabela(dados):
    """Formata os dados extraídos em uma tabela."""
    tabela = [config.CONFIG["colunas"]]  # Cabeçalhos
    for valor in dados:
        valores_colunas = valor.split()  # Divide as linhas em colunas
        if valores_colunas:
            tabela.append(valores_colunas)
    return tabela


def processar_excel(dados, molde_path, progress_bar, progress_label, app):
    """Processa o arquivo Excel e preenche os dados no molde."""
    try:
        if not os.path.exists(molde_path):
            raise FileNotFoundError(f"Arquivo de molde não encontrado: {molde_path}")
        
        wb = openpyxl.load_workbook(molde_path)
        sheet = wb.active

        linha_inicio = 62
        coluna_inicio = 4

        for row_index, linha in enumerate(dados, start=linha_inicio):
            for col_index, valor in enumerate(linha, start=coluna_inicio):
                cell = sheet.cell(row=row_index, column=col_index)
                cell.value = valor  # Atribui o valor sem alterar formatação

        excel_saida_path = app.excel_save_path
        if not excel_saida_path:
            exibir_popup("Erro: Nenhum local selecionado para salvar o arquivo.", "Erro", app, QMessageBox.Critical)
            return

        wb.save(excel_saida_path)

        exibir_popup(f"Arquivo salvo em: {excel_saida_path}", "Sucesso", app)
        progress_bar.setValue(100)
        progress_label.setText("Concluído")
    except Exception as e:
        exibir_popup(f"Erro ao salvar o arquivo Excel: {str(e)}", "Erro", app, QMessageBox.Critical)


class MyApp(QWidget):
    def __init__(self):
        super().__init__()

        self.pdf_path = None
        self.molde_path = None
        self.excel_save_path = None

        self.initUI()

    def initUI(self):
        """Inicializa a interface gráfica."""
        self.setWindowTitle('Extrator e Processador de Dados PDF para Excel')
        self.setGeometry(100, 100, 500, 400)
        self.setStyleSheet("background-color: orange;")

        layout = QVBoxLayout()

        # Título e instruções
        self.title_label = QLabel("Aut 3º ", self)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold; color: white;")
        layout.addWidget(self.title_label)

        self.instructions_label = QLabel("Escolha um PDF, extraia os dados e depois salve no molde Excel.", self)
        self.instructions_label.setAlignment(Qt.AlignCenter)
        self.instructions_label.setStyleSheet("color: white;")
        layout.addWidget(self.instructions_label)

        self.progress_label = QLabel("0% Concluído", self)
        self.progress_label.setStyleSheet("color: white;")
        layout.addWidget(self.progress_label)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Campo de texto para exibir e editar o texto extraído
        self.text_input = QTextEdit(self)
        self.text_input.setPlaceholderText("Texto extraído do PDF")
        self.text_input.setReadOnly(False)
        layout.addWidget(self.text_input)

        # Layout de botões
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

        # Layout para posicionar o botão de Ajuda no canto superior direito
        ajuda_button_layout = QHBoxLayout()
        ajuda_button_layout.addStretch()  # Adiciona um espaçador antes do botão
        self.ajuda_button = QPushButton("Ajuda", self)
        self.ajuda_button.setFixedSize(80, 30)  # Definir tamanho menor
        self.ajuda_button.clicked.connect(self.abrir_ajuda)
        ajuda_button_layout.addWidget(self.ajuda_button)

        layout.addLayout(ajuda_button_layout)

        self.setLayout(layout)

    def escolher_pdf(self):
        """Escolhe o arquivo PDF."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolher PDF", "", "Arquivos PDF (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.confirmar_button.setEnabled(True)

    def escolher_molde(self):
        """Escolhe o arquivo de Molde Excel."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolher Molde Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.molde_path = file_path
            self.inserir_no_molde_button.setEnabled(True)

    def confirmar_extracao(self):
        """Confirma a extração de texto do PDF."""
        if not self.pdf_path:
            exibir_popup("Erro: Por favor, selecione um arquivo PDF.", "Erro", self, QMessageBox.Critical)
            return
        texto_extraído = extrair_texto_pdf(self.pdf_path)
        self.text_input.setText(texto_extraído)
        self.confirmar_button.setEnabled(False)
        self.inserir_no_molde_button.setEnabled(True)
        self.extrair_novamente_button.setEnabled(True)

    def inserir_no_molde(self):
        """Insere os dados extraídos no molde Excel."""
        if not self.molde_path or not self.pdf_path:
            exibir_popup("Erro: Por favor, selecione tanto o arquivo PDF quanto o molde Excel.", "Erro", self, QMessageBox.Critical)
            return

        texto_extraído = self.text_input.toPlainText()
        dados_extraídos = texto_extraído.split("\n")  # Divide por linhas
        dados_formatados = formatar_como_tabela(dados_extraídos)

        self.progress_bar.setValue(50)
        self.progress_label.setText("50% Concluído")

        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Arquivo Excel", "", "Arquivos Excel (*.xlsx)")
        if file_path:
            self.excel_save_path = file_path
            processar_excel(dados_formatados, self.molde_path, self.progress_bar, self.progress_label, self)

    def resetar_processo(self):
        """Reseta o processo de extração.""" 
        self.text_input.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("0% Concluído")
        self.confirmar_button.setEnabled(True)
        self.inserir_no_molde_button.setEnabled(False)
        self.extrair_novamente_button.setEnabled(False)

    def toggle_table(self):
        """Alterna entre exibir texto ou tabela."""
        tabela = formatar_como_tabela(self.text_input.toPlainText().split("\n"))
        self.text_input.setText("\n".join(["\t".join(linha) for linha in tabela]))
        self.toggle_table_button.setText("Alternar para Visualizar Texto")

    def exibir_popup_ajuda(self, ajuda_texto):
        """Exibe a ajuda em uma janela pop-up ajustável."""
        ajuda_dialog = QDialog(self)
        ajuda_dialog.setWindowTitle("Ajuda")
        ajuda_dialog.setMinimumWidth(400)
        ajuda_dialog.setMinimumHeight(300)

        layout = QVBoxLayout()
        scroll_area = QScrollArea(ajuda_dialog)
        scroll_area.setWidgetResizable(True)

        ajuda_text_edit = QTextEdit(ajuda_dialog)
        ajuda_text_edit.setText(ajuda_texto)
        ajuda_text_edit.setReadOnly(True)
        ajuda_text_edit.setWordWrapMode(True)

        scroll_area.setWidget(ajuda_text_edit)
        layout.addWidget(scroll_area)

        close_button = QPushButton("Fechar", ajuda_dialog)
        close_button.clicked.connect(ajuda_dialog.close)
        layout.addWidget(close_button)

        ajuda_dialog.setLayout(layout)
        ajuda_dialog.exec_()

    def abrir_ajuda(self):
        """Abre o arquivo de ajuda em uma janela ajustável."""
        try:
            ajuda_path = os.path.join(os.path.dirname(__file__), "docs", "C:\\Users\\usuario\\Desktop\\automações\\Final 3\\app\\Ajuda\\ajuda.txt")

            with open(ajuda_path, "r", encoding="utf-8") as file:
                ajuda_texto = file.read()

            self.exibir_popup_ajuda(ajuda_texto)

        except FileNotFoundError:
            exibir_popup("Erro: O arquivo 'ajuda.txt' não foi encontrado.", "Erro", self, QMessageBox.Critical)
        except Exception as e:
            exibir_popup(f"Erro ao abrir o arquivo de ajuda: {str(e)}", "Erro", self, QMessageBox.Critical)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myapp = MyApp()
    myapp.show()
    sys.exit(app.exec_())
