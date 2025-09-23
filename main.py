import tkinter as tk
from tkinter import messagebox
import logging
import sys
import os
import traceback

# This needs to be at the very top to ensure imports from the app folder work
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

def setup_logging(log_path=None):
    """Configura il logging su file e console."""
    # Get the root logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    # Clean up existing handlers to avoid duplicate logs
    if logger.hasHandlers():
        logger.handlers.clear()

    # Formatter for file logs (more detailed)
    formatter_file = logging.Formatter(
        '%(asctime)s - %(levelname)s - [%(name)s:%(funcName)s:%(lineno)d] - %(message)s'
    )
    # Formatter for console logs (less verbose)
    formatter_console = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter_console)
    console_handler.setLevel(logging.INFO) # Show only INFO and above on console
    logger.addHandler(console_handler)

    # File handler (only if a path is provided)
    if log_path:
        try:
            # Ensure the directory for the log file exists
            log_dir = os.path.dirname(log_path)
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)

            file_handler = logging.FileHandler(log_path, encoding='utf-8', mode='w')
            file_handler.setFormatter(formatter_file)
            file_handler.setLevel(logging.DEBUG) # Log everything to the file
            logger.addHandler(file_handler)
            logging.info(f"Logging su file abilitato: {log_path}")
        except Exception as e:
            # This error will go to the console handler
            logging.error(f"Impossibile creare FileHandler per il log in '{log_path}'. Errore: {e}")
    else:
        logging.warning("Percorso del file di log non fornito. Il log andrà solo su console.")

def main():
    """Punto di ingresso principale dell'applicazione."""
    # Set up basic console logging immediately to catch early errors
    setup_logging()

    try:
        # --- Imports are moved inside the try block ---
        # This ensures that if any of them fail (e.g., missing dependency),
        # the exception is caught and reported.
        from analyzer_app import config
        from analyzer_app.gui import App

        logging.info("Importazioni dei moduli dell'applicazione riuscite.")

        # Now, load the configuration, which might fail if the file is missing/corrupt
        config.load_config()
        logging.info("Configurazione caricata da 'parametri.xlsm'.")

        # Re-configure logging to include the file path from the now-loaded config
        setup_logging(log_path=config.LOG_FILEPATH)

        logging.info("Avvio dell'interfaccia grafica (GUI)...")
        root = tk.Tk()
        app = App(root)
        root.mainloop()

    except Exception as e:
        # This is the last line of defense. It catches any error during startup.
        error_message = f"Si è verificato un errore critico durante l'avvio:\n\n{type(e).__name__}: {e}"
        logging.critical(error_message, exc_info=True) # Log the full traceback

        # Also write the error to an emergency file, in case logging to file failed
        emergency_file = "startup_fatal_error.txt"
        with open(emergency_file, "w", encoding='utf-8') as f:
            f.write(error_message + "\n\n")
            f.write("------ TRACEBACK ------\n")
            traceback.print_exc(file=f)

        # Try to show a message box to the user
        try:
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror(
                "Errore Fatale Applicazione",
                f"{error_message}\n\nL'applicazione non può continuare."
                f"\nDettagli dell'errore sono stati salvati nel file '{emergency_file}'"
            )
            root_err.destroy()
        except tk.TclError:
            # If tkinter itself is broken, we can't show a GUI error.
            # The error is already logged and saved to the text file.
            print("Impossibile mostrare la finestra di dialogo di errore. Controllare i log.")

    finally:
        logging.info("Applicazione terminata.")
        logging.shutdown()

if __name__ == "__main__":
    # Add a print statement right at the beginning to confirm the script is running
    print("Esecuzione di main.py avviata...")
    main()
    print("Esecuzione di main.py terminata.")
