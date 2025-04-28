import sys
import xlwings as xw
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QLineEdit, QComboBox
from PyQt5.QtCore import Qt


class XlsConverter(QWidget):
    def __init__(self):
        super().__init__()

        # Configurações iniciais da janela
        self.setWindowTitle('Conversor de .xls para .xlsx / .xlsm')
        self.setGeometry(400, 200, 600, 200)

        # Layout
        layout = QVBoxLayout()

        # Instruções
        self.label_instrucoes = QLabel("Selecione o arquivo .xls para converter:", self)
        layout.addWidget(self.label_instrucoes)

        # Campo de texto para exibir o caminho do arquivo
        self.entry_arquivo_xls = QLineEdit(self)
        self.entry_arquivo_xls.setPlaceholderText("Caminho do arquivo .xls")
        layout.addWidget(self.entry_arquivo_xls)

        # Botão para abrir o arquivo .xls
        self.botao_abrir_xls = QPushButton('Abrir Arquivo .xls', self)
        self.botao_abrir_xls.clicked.connect(self.abrir_arquivo_xls)
        layout.addWidget(self.botao_abrir_xls)

        # Opção de seleção do formato de destino (.xlsx ou .xlsm)
        self.combo_formatos = QComboBox(self)
        self.combo_formatos.addItem(".xlsx")
        self.combo_formatos.addItem(".xlsm")
        layout.addWidget(self.combo_formatos)

        # Botão para converter o arquivo
        self.botao_converter = QPushButton('Converter Arquivo', self)
        self.botao_converter.clicked.connect(self.converter_xls)
        layout.addWidget(self.botao_converter)

        # Label para mostrar o status
        self.label_status = QLabel('', self)
        self.label_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label_status)

        # Configurar o layout na janela
        self.setLayout(layout)

    def abrir_arquivo_xls(self):
        # Abre a janela de seleção de arquivo .xls
        caminho_arquivo, _ = QFileDialog.getOpenFileName(self, "Abrir Arquivo .xls", "", "Excel Files (*.xls)")
        if caminho_arquivo:
            self.entry_arquivo_xls.setText(caminho_arquivo)

    def converter_xls(self):
        caminho_xls = self.entry_arquivo_xls.text()
        formato_selecionado = self.combo_formatos.currentText()

        if caminho_xls:
            try:
                # Chamar a função para converter o arquivo .xls
                novo_arquivo = self.converter(caminho_xls, formato_selecionado)
                self.label_status.setText(f"Arquivo convertido para: {novo_arquivo}")
            except Exception as e:
                self.label_status.setText(f"Erro ao converter: {e}")
        else:
            self.label_status.setText("Por favor, selecione um arquivo .xls.")

    def converter(self, caminho_arquivo_xls, formato_selecionado):
        # Abre o arquivo .xls com xlwings
        with xw.App(visible=False) as app:
            # Abrir o arquivo Excel .xls
            wb = app.books.open(caminho_arquivo_xls)
            
            # Gerar o novo caminho para salvar como .xlsx ou .xlsm
            if formato_selecionado == ".xlsx":
                novo_arquivo = caminho_arquivo_xls.replace('.xls', '.xlsx')
            elif formato_selecionado == ".xlsm":
                novo_arquivo = caminho_arquivo_xls.replace('.xls', '.xlsm')
            
            # Salvar o arquivo no formato selecionado
            wb.save(novo_arquivo)
            
            # Fechar o arquivo
            wb.close()
        
        return novo_arquivo


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = XlsConverter()
    window.show()
    sys.exit(app.exec_())
