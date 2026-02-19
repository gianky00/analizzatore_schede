import logging
import os
import sys
import tkinter as tk
import traceback
from tkinter import messagebox

# Ensure imports work correctly
if getattr(sys, 'frozen', False):
    # Running as compiled exe
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.abspath(os.path.dirname(__file__))

sys.path.insert(0, APP_DIR)


def setup_logging() -> None:
    """Configura il logging."""
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        logger.handlers.clear()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    ))
    console_handler.setLevel(logging.INFO)
    logger.addHandler(console_handler)


def main() -> None:
    """Punto di ingresso principale."""
    setup_logging()

    try:
        from analyzer_app.gui import App

        logging.info("Moduli caricati correttamente")
        logging.info("Avvio interfaccia grafica...")

        root = tk.Tk()
        App(root)
        root.mainloop()


    except Exception as e:
        error_msg = f"Errore critico:\n{type(e).__name__}: {e}"
        logging.critical(error_msg, exc_info=True)

        # Save error to file
        error_file = os.path.join(APP_DIR, "errore_avvio.txt")
        with open(error_file, "w", encoding='utf-8') as f:
            f.write(f"{error_msg}\n\n{'='*50}\nTRACEBACK:\n{'='*50}\n")
            traceback.print_exc(file=f)

        _show_error(f"{error_msg}\n\nDettagli salvati in 'errore_avvio.txt'")

    finally:
        logging.info("Applicazione terminata.")
        logging.shutdown()


def _show_error(message: str) -> None:
    """Mostra errore in una finestra."""
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Errore Applicazione", message)
        root.destroy()
    except Exception as e:
        print(f"\nERRORE: {e}")  # noqa: T201



if __name__ == "__main__":
    main()
