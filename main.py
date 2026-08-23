from app.localization import t
import sys

from PyQt6.QtWidgets import QApplication, QMessageBox

from app.db.database import initialize_database
from app.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(t("app_title"))
    app.setOrganizationName("JI JEONG GEUN")
    app.setQuitOnLastWindowClosed(True)

    try:
        initialize_database()
    except Exception as exc:
        QMessageBox.critical(
            None,
            t("database_error"),
            f"{t('database_init_failed')}\n\n{exc}"
        )
        return 1

    print("[INFO] Creating MainWindow...")
    window = MainWindow()

    # Keep a strong reference on QApplication as well.
    app.main_window = window

    print("[INFO] Showing MainWindow...")
    window.show()
    window.showMaximized()
    window.raise_()
    window.activateWindow()

    print("[INFO] Entering QApplication event loop...")
    exit_code = app.exec()
    print(f"[INFO] QApplication exited with code: {exit_code}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
