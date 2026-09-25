"""VetGestión: abre la ventana del programa y conecta la interfaz con el Excel."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import webview  # noqa: E402

import db  # noqa: E402
from api import Api  # noqa: E402


def main():
    db.inicializar()
    index = os.path.join(db._recursos_dir(), "ui", "index.html")
    webview.create_window(
        "VetGestión — Control de medicamentos",
        url=index,
        js_api=Api(),
        width=1360,
        height=860,
        min_size=(1000, 640),
        background_color="#EEF1EE",
    )
    webview.start()


if __name__ == "__main__":
    main()
