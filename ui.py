# ui.py
import qtawesome as qta
from PyQt5.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QWidget, QTableWidget, QTableWidgetItem,
    QFileDialog, QMessageBox, QProgressDialog, QFormLayout, QGroupBox, QDateEdit, QLineEdit, QDialog, QScrollArea, QComboBox
)
from PyQt5.QtCore import Qt, QDate, QRegExp
from PyQt5.QtGui import QRegExpValidator
from processing import analyze_file
from export import export_to_pdf, export_to_csv, export_to_txt


class NFCeAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.last_file_path = None
        self.last_report = None
        self.filtered_report = None
        self.initUI()

    def initUI(self):
        """Inicializa a interface gráfica do usuário."""
        self.setWindowTitle("Analisador de NFC-e")
        self.setGeometry(100, 100, 1400, 800)  # Ajuste de tamanho para acomodar novos filtros

        # Estilização do aplicativo
        self.setStyleSheet("""
            QMainWindow {
                background-color: #FDF6E3;
            }
            QLabel {
                color: #D9534F;
                font-size: 14px;
            }
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E6E6E6;
                font-family: Consolas;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #FAD7A0;
                border: 1px solid #E6E6E6;
                padding: 5px;
                font-weight: bold;
            }
            QPushButton {
                background-color: #FFA500;
                color: white;
                border: none;
                padding: 10px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #FF8C00;
            }
            QGroupBox {
                border: 1px solid #FFA500;
                border-radius: 5px;
                margin-top: 10px;
                padding: 10px;
            }
            QLineEdit, QDateEdit, QComboBox {
                border: 1px solid #FFA500;
                border-radius: 4px;
                padding: 5px;
            }
            QDialog {
                background-color: #FFFFFF;
                border-radius: 10px;
                padding: 15px;
            }
            QScrollArea {
                border: none;
                background-color: #FFFFFF;
            }
        """)

        main_layout = QVBoxLayout()

        # Título do aplicativo
        title_label = QLabel("Analisador de NFC-e")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; text-align: center; margin-bottom: 20px; color: #FF8C00;")
        main_layout.addWidget(title_label, alignment=Qt.AlignCenter)

        # Botões de ação
        button_layout = QHBoxLayout()

        analyze_button = QPushButton(" Analisar Arquivo")
        analyze_button.setIcon(qta.icon('fa.file-o'))
        analyze_button.clicked.connect(self.on_analyze)
        button_layout.addWidget(analyze_button)

        reanalyze_button = QPushButton(" Reanalisar Arquivo")
        reanalyze_button.setIcon(qta.icon('fa.refresh'))
        reanalyze_button.clicked.connect(self.on_reanalyze)
        reanalyze_button.setEnabled(False)
        self.reanalyze_button = reanalyze_button
        button_layout.addWidget(reanalyze_button)

        main_layout.addLayout(button_layout)

        # Grupo de filtros
        filters_group = QGroupBox("Filtros")
        filters_layout = QFormLayout()

        # Filtro por Status (Substituído QLineEdit por QComboBox)
        self.status_filter = QComboBox()
        self.status_filter.addItem("Todos")  # Opção para não filtrar
        self.status_filter.addItems(["Autorizada", "Cancelada", "Desconhecido"])
        filters_layout.addRow("Filtrar por Status:", self.status_filter)

        # Filtro por Produto
        self.product_filter = QLineEdit()
        self.product_filter.setPlaceholderText("Digite o nome do produto")
        filters_layout.addRow("Filtrar por Produto:", self.product_filter)

        # Filtro por CFOP
        self.cfop_filter = QLineEdit()
        self.cfop_filter.setPlaceholderText("Digite o código CFOP (ex: 5102)")
        filters_layout.addRow("Filtrar por CFOP:", self.cfop_filter)

        # Adicionar validação para CFOP (apenas 4 dígitos numéricos)
        regex = QRegExp(r"^\d{0,4}$")  # Permite até 4 dígitos
        validator = QRegExpValidator(regex)
        self.cfop_filter.setValidator(validator)

        # Filtro por Número da NFC-e (nNF)
        self.nNF_filter = QLineEdit()
        self.nNF_filter.setPlaceholderText("Digite o Número da NFC-e (nNF)")
        # Adicionar validação para nNF (apenas números)
        nnf_regex = QRegExp(r"^\d+$")
        nnf_validator = QRegExpValidator(nnf_regex)
        self.nNF_filter.setValidator(nnf_validator)
        filters_layout.addRow("Filtrar por Número NFC-e (nNF):", self.nNF_filter)

        # Filtro por Data de Autorização
        self.start_date_edit = QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))
        filters_layout.addRow("Data Início (Autorização):", self.start_date_edit)

        self.end_date_edit = QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        filters_layout.addRow("Data Fim (Autorização):", self.end_date_edit)

        # Filtro por Valor
        self.min_value_filter = QLineEdit()
        self.min_value_filter.setPlaceholderText("Valor Mínimo")
        # Adicionar validação para números decimais
        min_value_regex = QRegExp(r"^\d+(\.\d{1,2})?$")
        min_value_validator = QRegExpValidator(min_value_regex)
        self.min_value_filter.setValidator(min_value_validator)
        filters_layout.addRow("Valor Mínimo (R$):", self.min_value_filter)

        self.max_value_filter = QLineEdit()
        self.max_value_filter.setPlaceholderText("Valor Máximo")
        # Adicionar validação para números decimais
        max_value_regex = QRegExp(r"^\d+(\.\d{1,2})?$")
        max_value_validator = QRegExpValidator(max_value_regex)
        self.max_value_filter.setValidator(max_value_validator)
        filters_layout.addRow("Valor Máximo (R$):", self.max_value_filter)

        # Botão para aplicar filtros
        apply_filters_button = QPushButton("Aplicar Filtros")
        apply_filters_button.setIcon(qta.icon('fa.filter'))
        apply_filters_button.clicked.connect(self.apply_filters)
        filters_layout.addRow(apply_filters_button)

        filters_group.setLayout(filters_layout)
        main_layout.addWidget(filters_group)

        # Tabela para exibição do relatório
        self.report_table = QTableWidget()
        self.report_table.setColumnCount(6)  # Número da NFC-e, Nome da Nota, Valor, Status, Data de Emissão, Data de Autorização
        self.report_table.setHorizontalHeaderLabels(["Número NFC-e", "Nome da Nota", "Valor (R$)", "Status", "Data de Emissão", "Data de Autorização"])
        self.report_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.report_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.report_table.cellDoubleClicked.connect(self.on_note_double_click)
        main_layout.addWidget(self.report_table)

        # Botões de exportação
        export_button_layout = QHBoxLayout()

        export_pdf_button = QPushButton(" Exportar PDF")
        export_pdf_button.setIcon(qta.icon('fa.file-pdf-o'))
        export_pdf_button.clicked.connect(lambda: self.export_report("pdf"))
        export_button_layout.addWidget(export_pdf_button)

        export_csv_button = QPushButton(" Exportar CSV")
        export_csv_button.setIcon(qta.icon('fa.file-text-o'))
        export_csv_button.clicked.connect(lambda: self.export_report("csv"))
        export_button_layout.addWidget(export_csv_button)

        export_txt_button = QPushButton(" Exportar TXT")
        export_txt_button.setIcon(qta.icon('fa.file'))
        export_txt_button.clicked.connect(lambda: self.export_report("txt"))
        export_button_layout.addWidget(export_txt_button)

        main_layout.addLayout(export_button_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def on_analyze(self):
        """Manipula a ação de análise de arquivo."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo XML ou ZIP", "", "Arquivos (*.xml *.zip)")
        if not file_path:
            return

        self.last_file_path = file_path
        self.analyze(file_path)

    def on_reanalyze(self):
        """Reanalisa o último arquivo selecionado."""
        if self.last_file_path:
            self.analyze(self.last_file_path)
        else:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo para reanalisar.")

    def analyze(self, file_path):
        """Executa a análise de arquivo."""
        progress_dialog = QProgressDialog("Analisando arquivo...", "Cancelar", 0, 100, self)
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setMinimumDuration(0)

        try:
            self.last_report = analyze_file(file_path, progress_dialog)
            self.filtered_report = self.last_report
            self.display_report(self.last_report)
            self.reanalyze_button.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao analisar o arquivo: {str(e)}")
        finally:
            progress_dialog.close()

    def apply_filters(self):
        """Aplica os filtros selecionados ao relatório atual."""
        if not self.last_report:
            QMessageBox.warning(self, "Aviso", "Nenhum relatório disponível para filtrar.")
            return

        # Filtros
        selected_status = self.status_filter.currentText().strip().lower()
        product_filter = self.product_filter.text().strip().lower()
        cfop_filter = self.cfop_filter.text().strip().lower()
        nNF_filter = self.nNF_filter.text().strip().lower()
        # Removido cNF_filter
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()
        min_value = float(self.min_value_filter.text()) if self.min_value_filter.text() else None
        max_value = float(self.max_value_filter.text()) if self.max_value_filter.text() else None

        filtered_notas = []

        # Filtrar as notas
        for nota in self.last_report.get("notas", []):
            # Filtro por Status
            if selected_status != "todos" and selected_status != nota.get("status", "").lower():
                continue

            # Filtro por CFOP: verificar se algum produto tem o CFOP especificado
            if cfop_filter:
                produtos = nota.get("produtos", [])
                if not any(cfop_filter == produto.get("cfop", "").lower() for produto in produtos):
                    continue

            # Filtro por Número da NFC-e (nNF)
            if nNF_filter:
                if nNF_filter not in nota.get("nNF", "").lower():
                    continue

            # Filtro por Data de Autorização
            if nota.get("autorizada"):
                try:
                    auth_date = QDate.fromString(nota["autorizada"], "yyyy-MM-dd").toPyDate()
                    if not (start_date <= auth_date <= end_date):
                        continue
                except Exception:
                    # Caso a data esteja em formato diferente ou inválida
                    continue

            # Filtro por Valor
            if min_value is not None and nota.get("valor", 0) < min_value:
                continue
            if max_value is not None and nota.get("valor", 0) > max_value:
                continue

            # Filtro por Produto
            if product_filter:
                produtos = nota.get("produtos", [])
                if not any(product_filter in (prod.get("nome", "") or "").lower() for prod in produtos):
                    continue

            filtered_notas.append(nota)

        # Recalcular resumo
        filtered_resumo = {
            "total_notas": len(filtered_notas),
            "valor_total": sum(nota["valor"] for nota in filtered_notas),
        }

        self.filtered_report = {
            "resumo": filtered_resumo,
            "notas": filtered_notas,
        }
        self.display_report(self.filtered_report)

        # Verificar se há resultados
        if not filtered_notas:
            QMessageBox.information(self, "Nenhum Resultado", "Nenhum resultado encontrado com os critérios de filtro aplicados.")

    def export_report(self, format_type):
        """Exporta o relatório para o formato especificado."""
        output_file, _ = QFileDialog.getSaveFileName(self, f"Salvar como {format_type.upper()}", "", f"Arquivos {format_type.upper()} (*.{format_type})")
        if not output_file:
            return

        try:
            content = self.generate_report_text(self.filtered_report)
            if format_type == "pdf":
                export_to_pdf(content, output_file)
            elif format_type == "csv":
                export_to_csv(content, output_file)
            elif format_type == "txt":
                export_to_txt(content, output_file)
            QMessageBox.information(self, "Sucesso", f"Relatório exportado como {format_type.upper()}.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao exportar: {str(e)}")

    def display_report(self, report):
        """Exibe o relatório processado na tabela."""
        notas = report.get("notas", [])
        self.report_table.setRowCount(len(notas))
        for row, nota in enumerate(notas):
            # Número da NFC-e
            self.report_table.setItem(row, 0, QTableWidgetItem(nota.get("nNF", "N/A")))
            # Nome da Nota
            self.report_table.setItem(row, 1, QTableWidgetItem(nota.get("nome", "")))
            # Valor
            self.report_table.setItem(row, 2, QTableWidgetItem(f"R$ {nota.get('valor', 0):.2f}"))
            # Status
            self.report_table.setItem(row, 3, QTableWidgetItem(nota.get("status", "")))
            # Data de Emissão
            self.report_table.setItem(row, 4, QTableWidgetItem(nota.get("emitida", "")))
            # Data de Autorização
            self.report_table.setItem(row, 5, QTableWidgetItem(nota.get("autorizada", "")))

    def on_note_double_click(self, row, column):
        """Abre um modal para exibir detalhes da nota clicada."""
        nota = self.filtered_report["notas"][row]
        self.show_note_details(nota)

    def show_note_details(self, nota):
        """Exibe os detalhes completos de uma nota em um modal."""
        modal = QDialog(self)
        modal.setWindowTitle(f"Detalhes da Nota - {nota.get('nome', 'N/A')}")
        modal.setGeometry(300, 150, 800, 600)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # Cabeçalho
        header = QLabel(f"<h2>Detalhes da Nota: {nota.get('nome', 'N/A')}</h2>")
        header.setStyleSheet("color: #D9534F; margin-bottom: 15px;")
        layout.addWidget(header)

        # Informações da Nota (cNF removido)
        info_label = QLabel(
            f"<b>Número NFC-e:</b> {nota.get('nNF', 'N/A')}<br>"
            # f"<b>Código Numérico (cNF):</b> {nota.get('cNF', 'N/A')}<br>"  # Removido
            f"<b>Valor Total:</b> R$ {nota.get('valor', 0):.2f}<br>"
            f"<b>Status:</b> {nota.get('status', 'N/A')}<br>"
            f"<b>Data de Emissão:</b> {nota.get('emitida', 'N/A')}<br>"
            f"<b>Data de Autorização:</b> {nota.get('autorizada', 'N/A')}<br>"
        )
        info_label.setStyleSheet("font-size: 14px; margin-bottom: 10px;")
        layout.addWidget(info_label)

        # Área de rolagem para os produtos
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        # Produtos
        produtos_group = QGroupBox("Produtos")
        produtos_layout = QVBoxLayout()

        for produto in nota.get("produtos", []):
            produto_label = QLabel(
                f"<b>Nome:</b> {produto.get('nome', 'Desconhecido')}<br>"
                f"<b>Código:</b> {produto.get('codigo', 'N/A')}<br>"
                f"<b>CFOP:</b> {produto.get('cfop', 'N/A')}<br>"
                f"<b>Quantidade:</b> {produto.get('quantidade', 0):.2f} {produto.get('unidade', '')}<br>"
                f"<b>Valor Unitário:</b> R$ {produto.get('valor_unitario', 0):,.2f}<br>"
                f"<b>Valor Total:</b> R$ {produto.get('valor_total', 0):,.2f}<br>"
                f"<hr>"
            )
            produto_label.setStyleSheet("font-size: 14px;")
            produtos_layout.addWidget(produto_label)

        produtos_group.setLayout(produtos_layout)
        scroll_layout.addWidget(produtos_group)

        scroll_area.setWidget(scroll_content)
        layout.addWidget(scroll_area)

        # Botão de fechar
        close_button = QPushButton("Fechar")
        close_button.setStyleSheet("""
            QPushButton {
                background-color: #FFA500;
                color: white;
                border: none;
                padding: 10px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #FF8C00;
            }
        """)
        close_button.clicked.connect(modal.close)
        layout.addWidget(close_button, alignment=Qt.AlignRight)

        modal.setLayout(layout)
        modal.exec_()

    def generate_report_text(self, report):
        """
        Gera o texto do relatório para exportação.
        Este método precisa ser implementado conforme a estrutura do seu relatório.
        """
        # Exemplo simplificado:
        resumo = report.get("resumo", {})
        notas = report.get("notas", [])

        report_text = f"Resumo:\nTotal de Notas: {resumo.get('total_notas', 0)}\nValor Total: R$ {resumo.get('valor_total', 0):.2f}\n\nDetalhes das Notas:\n"

        for nota in notas:
            report_text += f"Número NFC-e: {nota.get('nNF', 'N/A')}\n"
            # report_text += f"Código Numérico (cNF): {nota.get('cNF', 'N/A')}\n"  # Removido
            report_text += f"Nome da Nota: {nota.get('nome', 'N/A')}\n"
            report_text += f"Valor: R$ {nota.get('valor', 0):.2f}\n"
            report_text += f"Status: {nota.get('status', 'N/A')}\n"
            report_text += f"Data de Emissão: {nota.get('emitida', 'N/A')}\n"
            report_text += f"Data de Autorização: {nota.get('autorizada', 'N/A')}\n"
            report_text += "Produtos:\n"
            for produto in nota.get("produtos", []):
                report_text += f"  - Nome: {produto.get('nome', 'N/A')}, Código: {produto.get('codigo', 'N/A')}, CFOP: {produto.get('cfop', 'N/A')}, Quantidade: {produto.get('quantidade', 0)} {produto.get('unidade', '')}, Valor Unitário: R$ {produto.get('valor_unitario', 0):.2f}, Valor Total: R$ {produto.get('valor_total', 0):.2f}\n"
            report_text += "\n"

        return report_text

    def closeEvent(self, event):
        """
        Garante que a aplicação seja fechada corretamente.
        """
        reply = QMessageBox.question(
            self,
            'Sair',
            'Tem certeza de que deseja sair?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()
