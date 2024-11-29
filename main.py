# main.py
import sys
import qtawesome as qta
from PyQt5.QtWidgets import QApplication
from ui import NFCeAnalyzerApp
from export import export_to_excel  # Importar a função export_to_excel

def apply_styles(app):
    """
    Aplica os estilos globais ao aplicativo.
    """
    app.setStyleSheet("""
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
        QLineEdit, QDateEdit {
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


def main():
    """
    Ponto de entrada principal do aplicativo.
    """
    app = QApplication(sys.argv)

    # Aplica os estilos globais ao aplicativo
    apply_styles(app)

    # Configura o ícone global da aplicação usando Font Awesome
    app.setWindowIcon(qta.icon('fa.file-o'))  # Ícone para arquivos

    # Inicializa a interface principal do aplicativo
    main_window = NFCeAnalyzerApp()

    # Aqui, você pode conectar o sinal de exportação da interface com a função export_to_excel
    # Supondo que na sua interface exista um botão que, ao ser clicado, chama a função de exportação
    # Por exemplo:
    # main_window.export_excel_button.clicked.connect(lambda: export_to_excel(report_text, 'relatorio.xlsx'))

    main_window.show()

    # Inicia o loop de eventos da aplicação
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
