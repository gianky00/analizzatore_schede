import tkinter as tk
from tkinter import messagebox
import logging
import sys
import os

# Aggiungi la directory root al path per permettere l'import di analyzer_app
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from analyzer_app import config
from analyzer_app.gui import App

def setup_logging():
    """Configura il logging su file e console."""
    logger = logging.getLogger() # Root logger
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        logger.handlers.clear()

    formatter_file = logging.Formatter('%(asctime)s - %(levelname)s - [%(name)s:%(funcName)s:%(lineno)d] - %(message)s')
    formatter_console = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter_console)
    console_handler.setLevel(logging.INFO)
    logger.addHandler(console_handler)

    if config.LOG_FILEPATH:
        try:
            file_handler = logging.FileHandler(config.LOG_FILEPATH, encoding='utf-8', mode='w')
            file_handler.setFormatter(formatter_file)
            logger.addHandler(file_handler)
            logging.info(f"Logging configurato. Log su file: {config.LOG_FILEPATH}")
        except Exception as e:
            logging.error(f"Impossibile creare FileHandler per il log in '{config.LOG_FILEPATH}'. Errore: {e}")
    else:
        logging.warning("Percorso del file di log non determinato. Il log andrà solo su console.")

def main():
    """Punto di ingresso principale dell'applicazione."""
    try:
        config.load_config()
        setup_logging()

        root = tk.Tk()
        app = App(root)
        root.mainloop()

    except Exception as e:
        logging.critical(f"Errore fatale durante l'avvio dell'applicazione: {e}", exc_info=True)
        temp_root = tk.Tk()
        temp_root.withdraw()
        messagebox.showerror(
            "Errore Fatale Applicazione",
            f"Si è verificato un errore critico durante l'inizializzazione:\n\n{e}\n\n"
            f"L'applicazione sarà chiusa. Controllare il log per i dettagli:\n"
            f"{config.LOG_FILEPATH or 'Output console'}",
            parent=temp_root
        )
        temp_root.destroy()
    finally:
        logging.info("Applicazione terminata.")
        if logging.getLogger().hasHandlers():
            logging.shutdown()

if __name__ == "__main__":
    main()
