import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon

# Crie o QApplication antes de importar módulos que usam qtawesome!
app = QApplication(sys.argv)

import qtawesome as qta  # Agora é seguro usar qtawesome
from ui import NFCeAnalyzerApp

def main() -> None:
    # Cria um QIcon a partir do pixmap do ícone retornado pelo qtawesome
    icon = QIcon(qta.icon('fa.file-o').pixmap(64, 64))
    app.setWindowIcon(icon)

    main_window = NFCeAnalyzerApp()
    main_window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()