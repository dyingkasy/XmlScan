import qtawesome as qta
from PyQt6.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QWidget,
    QFileDialog, QMessageBox, QProgressDialog, QFormLayout, QGroupBox, QDateEdit,
    QLineEdit, QDialog, QScrollArea, QComboBox, QTextEdit, QTableView
)
from PyQt6.QtCore import Qt, QDate, QRegularExpression, QObject, pyqtSignal, QRunnable, QThreadPool, QModelIndex, QAbstractTableModel
from PyQt6.QtGui import QRegularExpressionValidator, QBrush, QColor
import datetime
import os
import locale

from processing import analyze_file
from export import export_to_pdf, export_to_txt, export_to_csv, export_to_excel

locale.setlocale(locale.LC_ALL, '')

def format_currency(value: float) -> str:
    """
    Formata um valor monetário conforme a localidade.
    """
    try:
        return locale.currency(value, grouping=True)
    except Exception:
        return f"R$ {value:,.2f}"

class WorkerSignals(QObject):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

class AnalyzeWorker(QRunnable):
    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path
        self.signals = WorkerSignals()

    def run(self):
        try:
            report = analyze_file(self.file_path, progress_dialog=None)
            self.signals.finished.emit(report)
        except Exception as e:
            self.signals.error.emit(str(e))

class NotasTableModel(QAbstractTableModel):
    def __init__(self, notas: list, parent=None):
        super().__init__(parent)
        self._notas = notas
        # Cabeçalho padrão; será atualizado em display_report conforme o modelo
        self._headers = ["Número NFC-e", "Chave", "Valor (R$)", "Status", "Data de Emissão", "Data de Autorização"]

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._notas)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return len(self._headers)

    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        nota = self._notas[index.row()]
        col = index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            if col == 0:
                return nota.get("nNF", "N/A")
            elif col == 1:
                return nota.get("chNFe") or "N/A"
            elif col == 2:
                return format_currency(nota.get("valor", 0))
            elif col == 3:
                return nota.get("status", "")
            elif col == 4:
                return nota.get("emitida", "")
            elif col == 5:
                return nota.get("autorizada", "")
        elif role == Qt.ItemDataRole.BackgroundRole:
            status_lower = (nota.get("status") or "").lower()
            if status_lower == "autorizada":
                return QBrush(QColor(200, 255, 200))
            elif status_lower == "cancelada":
                return QBrush(QColor(255, 200, 200))
            elif status_lower == "sem protocolo":
                return QBrush(QColor(255, 255, 200))
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self._headers[section]
        return None

    def updateData(self, notas: list):
        self.beginResetModel()
        self._notas = notas
        self.endResetModel()

class NFCeAnalyzerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.last_file_path = None
        self.last_report = None
        self.filtered_report = None
        self.threadpool = QThreadPool()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Analisador de NFC-e")
        self.setGeometry(100, 100, 1400, 800)

        main_layout = QVBoxLayout()

        title_label = QLabel("Analisador de NFC-e")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; text-align: center; margin-bottom: 20px; color: #FF8C00;")
        main_layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignCenter)

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

        compare_pdf_button = QPushButton(" Comparar PDF")
        compare_pdf_button.setIcon(qta.icon('fa.file-pdf-o'))
        compare_pdf_button.clicked.connect(self.on_compare_pdf)
        button_layout.addWidget(compare_pdf_button)

        compare_excel_button = QPushButton(" Comparar Excel")
        compare_excel_button.setIcon(qta.icon('fa.file-excel-o'))
        compare_excel_button.clicked.connect(self.on_compare_excel)
        button_layout.addWidget(compare_excel_button)

        main_layout.addLayout(button_layout)

        filters_group = QGroupBox("Filtros")
        filters_layout = QFormLayout()

        self.status_filter = QComboBox()
        self.status_filter.addItem("Todos")
        self.status_filter.addItems(["Autorizada", "Cancelada", "Desconhecido", "Sem Protocolo"])
        filters_layout.addRow("Filtrar por Status:", self.status_filter)

        self.product_filter = QLineEdit()
        self.product_filter.setPlaceholderText("Nome do Produto")
        filters_layout.addRow("Filtrar por Produto:", self.product_filter)

        self.cfop_filter = QLineEdit()
        self.cfop_filter.setPlaceholderText("Código CFOP (ex: 5102)")
        self.cfop_filter.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d{0,4}$")))
        filters_layout.addRow("Filtrar por CFOP:", self.cfop_filter)

        self.nNF_filter = QLineEdit()
        self.nNF_filter.setPlaceholderText("Número NFC-e (nNF)")
        self.nNF_filter.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d+$")))
        filters_layout.addRow("Filtrar por Número NFC-e (nNF):", self.nNF_filter)

        self.start_date_edit = QDateEdit(calendarPopup=True)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-1))
        filters_layout.addRow("Data Início (Autorização):", self.start_date_edit)

        self.end_date_edit = QDateEdit(calendarPopup=True)
        self.end_date_edit.setDate(QDate.currentDate())
        filters_layout.addRow("Data Fim (Autorização):", self.end_date_edit)

        self.min_value_filter = QLineEdit()
        self.min_value_filter.setPlaceholderText("Valor Mínimo (R$)")
        self.min_value_filter.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d+(\.\d{1,2})?$")))
        filters_layout.addRow("Valor Mínimo (R$):", self.min_value_filter)

        self.max_value_filter = QLineEdit()
        self.max_value_filter.setPlaceholderText("Valor Máximo (R$)")
        self.max_value_filter.setValidator(QRegularExpressionValidator(QRegularExpression(r"^\d+(\.\d{1,2})?$")))
        filters_layout.addRow("Valor Máximo (R$):", self.max_value_filter)

        apply_filters_button = QPushButton("Aplicar Filtros")
        apply_filters_button.setIcon(qta.icon('fa.filter'))
        apply_filters_button.clicked.connect(self.apply_filters)
        filters_layout.addRow(apply_filters_button)

        filters_group.setLayout(filters_layout)
        main_layout.addWidget(filters_group)

        self.table_view = QTableView()
        self.model = NotasTableModel([])
        self.table_view.setModel(self.model)
        self.table_view.doubleClicked.connect(self.on_note_double_click)
        main_layout.addWidget(self.table_view)

        self.summary_label = QLabel("")
        main_layout.addWidget(self.summary_label)

        export_layout = QHBoxLayout()

        btn_pdf = QPushButton(" Exportar PDF")
        btn_pdf.setIcon(qta.icon('fa.file-pdf-o'))
        btn_pdf.clicked.connect(lambda: self.export_report("pdf"))
        export_layout.addWidget(btn_pdf)

        btn_txt = QPushButton(" Exportar TXT")
        btn_txt.setIcon(qta.icon('fa.file-text-o'))
        btn_txt.clicked.connect(lambda: self.export_report("txt"))
        export_layout.addWidget(btn_txt)

        btn_csv = QPushButton(" Exportar CSV")
        btn_csv.setIcon(qta.icon('fa.file-text-o'))
        btn_csv.clicked.connect(lambda: self.export_report("csv"))
        export_layout.addWidget(btn_csv)

        btn_xlsx = QPushButton(" Exportar Excel")
        btn_xlsx.setIcon(qta.icon('fa.file-excel-o'))
        btn_xlsx.clicked.connect(lambda: self.export_report("xlsx"))
        export_layout.addWidget(btn_xlsx)

        btn_csv_excel = QPushButton(" Exportar CSV/Excel")
        btn_csv_excel.setIcon(qta.icon('fa.file-excel-o'))
        btn_csv_excel.clicked.connect(self.export_csv_or_excel)
        export_layout.addWidget(btn_csv_excel)

        main_layout.addLayout(export_layout)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def on_analyze(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo XML ou ZIP", "", "Arquivos (*.xml *.zip)"
        )
        if not file_path:
            return
        self.last_file_path = file_path
        self.start_analysis(file_path)

    def on_reanalyze(self):
        if self.last_file_path:
            self.start_analysis(self.last_file_path)
        else:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo para reanalisar.")

    def start_analysis(self, file_path: str):
        self.progress_dialog = QProgressDialog("Analisando arquivo...", "Cancelar", 0, 0, self)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.show()

        worker = AnalyzeWorker(file_path)
        worker.signals.finished.connect(self.analysis_finished)
        worker.signals.error.connect(self.analysis_error)
        self.threadpool.start(worker)

    def analysis_finished(self, report: dict):
        self.progress_dialog.close()
        self.last_report = report
        self.filtered_report = report
        self.display_report(report)
        errors = report.get("errors", [])
        if errors:
            self.show_errors_dialog(errors)
        duplicates = report.get("duplicates", [])
        if duplicates:
            self.show_duplicates_dialog(duplicates)
        missing = report.get("missing_keys", [])
        if missing:
            self.show_missing_keys_dialog(missing)
        self.reanalyze_button.setEnabled(True)

    def analysis_error(self, error_msg: str):
        self.progress_dialog.close()
        QMessageBox.critical(self, "Erro", f"Erro ao analisar: {error_msg}")

    def apply_filters(self):
        if not self.last_report:
            QMessageBox.warning(self, "Aviso", "Nenhum relatório carregado.")
            return

        sel_status = self.status_filter.currentText().strip().lower()
        product_filter = self.product_filter.text().strip().lower()
        cfop_filter = self.cfop_filter.text().strip().lower()
        nNF_filter = self.nNF_filter.text().strip().lower()
        start_date = self.start_date_edit.date().toPyDate()
        end_date = self.end_date_edit.date().toPyDate()

        min_val = float(self.min_value_filter.text()) if self.min_value_filter.text() else None
        max_val = float(self.max_value_filter.text()) if self.max_value_filter.text() else None

        filtered_notas = []
        for nota in self.last_report.get("notas", []):
            if sel_status != "todos" and sel_status != (nota.get("status") or "").lower():
                continue
            if cfop_filter:
                if not any(cfop_filter == p.get("cfop", "").lower() for p in nota.get("produtos", [])):
                    continue
            if nNF_filter:
                if nNF_filter not in (nota.get("nNF") or "").lower():
                    continue
            if nota.get("autorizada"):
                try:
                    auth_date = datetime.datetime.strptime(nota["autorizada"], "%Y-%m-%d").date()
                    if not (start_date <= auth_date <= end_date):
                        continue
                except:
                    pass
            val = nota.get("valor", 0)
            if min_val is not None and val < min_val:
                continue
            if max_val is not None and val > max_val:
                continue
            if product_filter:
                if not any(product_filter in p.get("nome", "").lower() for p in nota.get("produtos", [])):
                    continue
            filtered_notas.append(nota)

        filtered_resumo = {
            "total_notas": len(filtered_notas),
            "valor_total": sum(n.get("valor", 0) for n in filtered_notas)
        }

        self.filtered_report = {
            "resumo": filtered_resumo,
            "notas": filtered_notas,
            "errors": self.last_report.get("errors", []),
            "duplicates": self.last_report.get("duplicates", []),
            "missing_keys": self.last_report.get("missing_keys", [])
        }
        self.display_report(self.filtered_report)
        if not filtered_notas:
            QMessageBox.information(self, "Sem resultados", "Nenhuma nota encontrada com esses filtros.")

    def display_report(self, report: dict):
        notas = report.get("notas", [])
        if notas:
            modelo = notas[0].get("modelo", "NFC-E")
            self.model._headers[0] = "Número NF-e" if modelo.upper() == "NFE" else "Número NFC-e"
        else:
            self.model._headers[0] = "Número NFC-e"
        self.model.updateData(notas)
    
        total_notas = len(notas)
        total_geral = sum(n.get("valor", 0) for n in notas)
        notas_autorizadas = [n for n in notas if (n.get("status") or "").lower() == "autorizada"]
        total_autorizadas = len(notas_autorizadas)
        valor_autorizadas = sum(n.get("valor", 0) for n in notas_autorizadas)
        txt = (
            f"<b>Total de Notas:</b> {total_notas} "
            f"| <b>Valor Total:</b> {format_currency(total_geral)}<br>"
            f"<b>Notas Autorizadas:</b> {total_autorizadas} "
            f"| <b>Valor Autorizadas:</b> {format_currency(valor_autorizadas)}"
        )
        self.summary_label.setText(txt)

    def export_report(self, format_type: str):
        if not self.filtered_report:
            QMessageBox.warning(self, "Aviso", "Nenhum relatório para exportar.")
            return
        if format_type.lower() == "pdf":
            filename, _ = QFileDialog.getSaveFileName(self, "Salvar PDF", "", "Arquivos PDF (*.pdf)")
            if filename:
                export_to_pdf(self.filtered_report, filename)
        elif format_type.lower() == "txt":
            filename, _ = QFileDialog.getSaveFileName(self, "Salvar TXT", "", "Arquivos TXT (*.txt)")
            if filename:
                export_to_txt(self.filtered_report, filename)
        elif format_type.lower() == "csv":
            filename, _ = QFileDialog.getSaveFileName(self, "Salvar CSV", "", "Arquivos CSV (*.csv)")
            if filename:
                export_to_csv(self.filtered_report, filename)
        elif format_type.lower() == "xlsx":
            filename, _ = QFileDialog.getSaveFileName(self, "Salvar Excel", "", "Arquivos Excel (*.xlsx)")
            if filename:
                export_to_excel(self.filtered_report, filename)
        QMessageBox.information(self, "Sucesso", f"Relatório exportado como {format_type.upper()}.")

    def export_csv_or_excel(self):
        if not self.filtered_report:
            QMessageBox.warning(self, "Aviso", "Nenhum relatório para exportar.")
            return
        file_filter = "CSV/Excel (*.csv *.xlsx);;Todos Arquivos (*.*)"
        out_file, _ = QFileDialog.getSaveFileName(self, "Salvar CSV ou Excel", "", file_filter)
        if not out_file:
            return
        import os
        _, ext = os.path.splitext(out_file)
        ext = ext.lower()
        if ext == ".csv":
            export_to_csv(self.filtered_report, out_file)
            QMessageBox.information(self, "Sucesso", "Exportado como CSV.")
        elif ext == ".xlsx":
            export_to_excel(self.filtered_report, out_file)
            QMessageBox.information(self, "Sucesso", "Exportado como Excel.")
        else:
            QMessageBox.warning(self, "Extensão inválida", "Escolha .csv ou .xlsx.")

    def on_note_double_click(self, index: QModelIndex):
        row = index.row()
        nota = self.filtered_report.get("notas", [])[row]
        self.show_note_details(nota)

    def show_note_details(self, nota: dict):
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Detalhes da Nota - {nota.get('nNF','N/A')} (Chave: {nota.get('chNFe','N/A')})")
        dlg.resize(700, 500)
        ly = QVBoxLayout(dlg)
        hdr = QLabel(f"<h2>Detalhes da Nota {nota.get('nNF','N/A')} - Chave: {nota.get('chNFe','N/A')}</h2>")
        ly.addWidget(hdr)
        info = QLabel(
            f"<b>Número:</b> {nota.get('nNF','N/A')}<br>"
            f"<b>Chave:</b> {nota.get('chNFe','N/A')}<br>"
            f"<b>Valor:</b> {format_currency(nota.get('valor',0))}<br>"
            f"<b>Status:</b> {nota.get('status','N/A')}<br>"
            f"<b>Emissão:</b> {nota.get('emitida','N/A')}<br>"
            f"<b>Autorização:</b> {nota.get('autorizada','N/A')}<br>"
        )
        ly.addWidget(info)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        sc_cont = QWidget()
        sc_ly = QVBoxLayout(sc_cont)
        group_prod = QGroupBox("Produtos")
        g_ly = QVBoxLayout(group_prod)
        for prod in nota.get("produtos", []):
            lbl_p = QLabel(
                f"<b>Nome:</b> {prod.get('nome','N/A')}<br>"
                f"<b>Código:</b> {prod.get('codigo','N/A')}<br>"
                f"<b>CFOP:</b> {prod.get('cfop','N/A')}<br>"
                f"<b>Quantidade:</b> {prod.get('quantidade',0)} {prod.get('unidade','')}<br>"
                f"<b>Valor Unitário:</b> {format_currency(prod.get('valor_unitario',0))}<br>"
                f"<b>Valor Total:</b> {format_currency(prod.get('valor_total',0))}<br>"
                f"<hr>"
            )
            g_ly.addWidget(lbl_p)
        group_prod.setLayout(g_ly)
        sc_ly.addWidget(group_prod)
        scroll.setWidget(sc_cont)
        ly.addWidget(scroll)
        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(dlg.close)
        ly.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def show_errors_dialog(self, errors: list):
        dlg = QDialog(self)
        dlg.setWindowTitle("Erros de Processamento")
        dlg.resize(600, 400)
        ly = QVBoxLayout(dlg)
        lbl = QLabel("<b>Alguns arquivos tiveram problemas:</b>")
        ly.addWidget(lbl)
        txt = QTextEdit()
        txt.setReadOnly(True)
        ly.addWidget(txt)
        for e in errors:
            txt.append(e)
        btn = QPushButton("Fechar")
        btn.clicked.connect(dlg.close)
        ly.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def show_duplicates_dialog(self, duplicates: list):
        dlg = QDialog(self)
        dlg.setWindowTitle("Notas Duplicadas")
        dlg.resize(600, 400)
        ly = QVBoxLayout(dlg)
        lbl = QLabel("<b>(nNF, cNF, cnpj) repetidos:</b>")
        ly.addWidget(lbl)
        txt = QTextEdit()
        txt.setReadOnly(True)
        ly.addWidget(txt)
        for d in duplicates:
            txt.append(str(d))
        btn = QPushButton("Fechar")
        btn.clicked.connect(dlg.close)
        ly.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def show_missing_keys_dialog(self, missing: list):
        dlg = QDialog(self)
        dlg.setWindowTitle("Chaves Oficiais Ausentes")
        dlg.resize(600, 400)
        ly = QVBoxLayout(dlg)
        lbl = QLabel("<b>Algumas chaves oficiais não foram encontradas:</b>")
        ly.addWidget(lbl)
        txt = QTextEdit()
        txt.setReadOnly(True)
        ly.addWidget(txt)
        for m in missing:
            txt.append(str(m))
        btn = QPushButton("Fechar")
        btn.clicked.connect(dlg.close)
        ly.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self, 'Sair', 'Deseja realmente sair?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()

    def on_compare_pdf(self):
        # Método implementado para diferenciar ou chamar funcionalidade de comparação de PDF
        QMessageBox.information(self, "Comparar PDF", "Funcionalidade de comparação de PDF não implementada.")

    def on_compare_excel(self):
        # Método implementado para diferenciar ou chamar funcionalidade de comparação de Excel
        QMessageBox.information(self, "Comparar Excel", "Funcionalidade de comparação de Excel não implementada.")
