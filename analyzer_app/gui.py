import os
import tkinter as tk
from tkinter import ttk, messagebox, font as tkFont
from functools import partial
import logging
import threading
import queue
from collections import Counter, defaultdict
from datetime import datetime

from . import config
from . import excel_io
from . import analysis
from . import reporting
from .data_models import InstrumentSheet

logger = logging.getLogger(__name__)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Analisi Schede Taratura - v2.0 Refactored")
        self.root.geometry("1750x980")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.analysis_queue = queue.Queue()
        self.analysis_results = []
        self.all_cert_usages = []
        self.schede_info_list_for_compilation = []
        self.human_errors_details = []
        self.candidate_files_count = 0
        self.validated_file_count = 0
        self.strumenti_campione = []

        self._setup_styles()
        self.create_widgets()

        # Inizia l'analisi automaticamente all'avvio
        self.start_analysis()

    def _setup_styles(self):
        self.style = ttk.Style(self.root)
        try:
            theme = 'vista' if 'vista' in self.style.theme_names() else 'clam'
            self.style.theme_use(theme)
            logger.info(f"Tema ttk impostato su: {theme}")
        except tk.TclError:
            logger.warning("Tema 'vista' o 'clam' non trovato.")

        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), relief="groove")
        self.style.configure("Treeview", rowheight=28, font=('Segoe UI', 9))
        self.style.configure("TNotebook.Tab", font=('Segoe UI', 10, 'bold'), padding=[12, 6])
        self.style.configure("TLabelframe.Label", font=('Segoe UI', 11, 'bold'))
        self.style.configure("Accent.TButton", font=('Segoe UI', 10, 'bold'), padding=8)
        self.style.configure("Hyperlink.TLabel", foreground="blue", font=('Segoe UI', 9, 'underline'))

    def create_widgets(self):
        # Frame principale
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(expand=True, fill=tk.BOTH)

        # Notebook per i tab
        self.notebook = ttk.Notebook(main_frame, style="TNotebook")
        self.notebook.pack(expand=True, fill='both', pady=(0, 10))

        # --- Tab di Progresso ---
        self.progress_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.progress_tab, text=' Progresso Analisi ')

        log_frame = ttk.LabelFrame(self.progress_tab, text="Log di Analisi", padding=10)
        log_frame.pack(expand=True, fill=tk.BOTH)

        log_v_scroll = ttk.Scrollbar(log_frame)
        log_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state=tk.DISABLED, yscrollcommand=log_v_scroll.set, font=("Consolas", 10))
        self.log_text.pack(expand=True, fill=tk.BOTH)
        log_v_scroll.config(command=self.log_text.yview)

        self.progress_bar = ttk.Progressbar(self.progress_tab, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.progress_label = ttk.Label(self.progress_tab, text="In attesa di iniziare l'analisi...")
        self.progress_label.pack(fill=tk.X)

        # --- Tab Cruscotto (inizialmente vuoto) ---
        self.cruscotto_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.cruscotto_tab, text=' Cruscotto Riepilogativo ', state=tk.DISABLED)

        # --- Tab Dettaglio Certificati (inizialmente vuoto) ---
        self.cert_details_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.cert_details_tab, text=' Dettaglio Utilizzo Certificati ', state=tk.DISABLED)

        # --- Tab Compilatore (inizialmente vuoto) ---
        self.compilatore_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.compilatore_tab, text=' Compilatore Automatico ', state=tk.DISABLED)

        # --- Tab Suggerimenti (inizialmente vuoto) ---
        self.suggerimenti_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.suggerimenti_tab, text=' Suggerimenti Strumenti ', state=tk.DISABLED)

    def _log_message(self, message, level="INFO"):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{level}] {message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        logger.log(logging.getLevelName(level), message)

    def start_analysis(self):
        self._log_message("Avvio del thread di analisi...")
        self.progress_bar['value'] = 0

        # Crea e avvia il thread di analisi
        self.analysis_thread = threading.Thread(target=self._analysis_worker, daemon=True)
        self.analysis_thread.start()

        # Inizia a controllare la coda per gli aggiornamenti
        self.root.after(100, self._check_analysis_queue)

    def _analysis_worker(self):
        """Funzione eseguita nel thread separato per non bloccare la GUI."""
        try:
            # 1. Leggi registro strumenti
            self.analysis_queue.put(('log', "Lettura registro strumenti..."))
            self.strumenti_campione = excel_io.leggi_registro_strumenti()
            if self.strumenti_campione is None:
                raise ValueError("Errore critico nel caricamento del registro strumenti.")
            self.analysis_queue.put(('log', f"Letti {len(self.strumenti_campione)} strumenti validi dal registro."))

            # 2. Trova file candidati
            folder_path = config.FOLDER_PATH_DEFAULT
            if not folder_path or not os.path.isdir(folder_path):
                raise NotADirectoryError(f"La cartella delle schede non è valida: {folder_path}")

            candidate_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~')]
            self.candidate_files_count = len(candidate_files)
            self.analysis_queue.put(('log', f"Trovati {self.candidate_files_count} file candidati da analizzare."))
            self.analysis_queue.put(('total_files', self.candidate_files_count))

            # 3. Analizza ogni file
            results = []
            for i, filename in enumerate(candidate_files):
                file_path = os.path.join(folder_path, filename)
                self.analysis_queue.put(('progress', (i + 1, f"Analisi di: {filename}")))
                try:
                    raw_data = excel_io.read_instrument_sheet_raw_data(file_path)
                    sheet_result = analysis.analyze_sheet_data(raw_data, self.strumenti_campione)
                    results.append(sheet_result)
                except Exception as e:
                    logger.error(f"Errore durante l'analisi del file {filename}: {e}", exc_info=True)
                    results.append(InstrumentSheet(file_path=file_path, base_filename=filename, status=f"Errore: {e}", is_valid=False))

            # 4. Invia i risultati finali alla coda
            self.analysis_queue.put(('done', results))

        except Exception as e:
            logger.critical(f"Errore fatale nel thread di analisi: {e}", exc_info=True)
            self.analysis_queue.put(('error', e))

    def _check_analysis_queue(self):
        """Controlla la coda per i messaggi dal thread di analisi e aggiorna la GUI."""
        try:
            while not self.analysis_queue.empty():
                msg_type, data = self.analysis_queue.get_nowait()

                if msg_type == 'log':
                    self._log_message(data)
                elif msg_type == 'total_files':
                    self.progress_bar['maximum'] = data
                elif msg_type == 'progress':
                    count, message = data
                    self.progress_bar['value'] = count
                    self.progress_label['text'] = message
                elif msg_type == 'done':
                    self.analysis_results = data
                    self.progress_label['text'] = "Analisi completata. Elaborazione risultati..."
                    self._process_final_results()
                    self._populate_results_ui()
                    return # Termina il polling
                elif msg_type == 'error':
                    self.progress_label['text'] = f"Errore durante l'analisi: {data}"
                    messagebox.showerror("Errore di Analisi", f"Si è verificato un errore: {data}")
                    return # Termina il polling

        except queue.Empty:
            pass
        finally:
            # Continua a controllare finché il thread non è morto
            if self.analysis_thread.is_alive():
                self.root.after(100, self._check_analysis_queue)

    def _process_final_results(self):
        """Aggrega i dati dall'analisi per popolare l'interfaccia utente."""
        self.validated_file_count = sum(1 for res in self.analysis_results if res.is_valid)
        self.all_cert_usages = [usage for res in self.analysis_results if res.is_valid for usage in res.certificate_usages]
        self.schede_info_list_for_compilation = [res.compilation_data for res in self.analysis_results if res.is_valid and res.compilation_data]
        self.human_errors_details = [{'file': res.base_filename, 'key': key, 'path': res.file_path} for res in self.analysis_results for key in res.human_error_keys]

        self._log_message(f"Elaborazione completata. Schede validate: {self.validated_file_count}/{self.candidate_files_count}")

    def _populate_results_ui(self):
        """Popola tutti i tab dei risultati dopo che l'analisi è terminata."""
        # Abilita i tab dei risultati
        self.notebook.tab(self.cruscotto_tab, state=tk.NORMAL)
        self.notebook.tab(self.cert_details_tab, state=tk.NORMAL)
        self.notebook.tab(self.compilatore_tab, state=tk.NORMAL)
        self.notebook.tab(self.suggerimenti_tab, state=tk.NORMAL)

        # Popola il cruscotto (versione semplificata)
        self._populate_cruscotto_tab()

        # Seleziona il tab del cruscotto
        self.notebook.select(self.cruscotto_tab)

    def _populate_cruscotto_tab(self):
        # Pulisci il tab se viene ripopolato
        for widget in self.cruscotto_tab.winfo_children():
            widget.destroy()

        stats_frame = ttk.LabelFrame(self.cruscotto_tab, text="Statistiche Generali", padding=10)
        stats_frame.pack(fill=tk.X, pady=5)

        ttk.Label(stats_frame, text=f"File analizzati: {self.candidate_files_count}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Schede validate: {self.validated_file_count}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Utilizzi certificati totali: {len(self.all_cert_usages)}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Errori di compilazione trovati: {len(self.human_errors_details)}").pack(anchor=tk.W)

        action_frame = ttk.LabelFrame(self.cruscotto_tab, text="Azioni", padding=10)
        action_frame.pack(fill=tk.X, pady=5)

        btn_report = ttk.Button(action_frame, text="Stampa Report Anomalie (Word)", command=self._generate_report_word, style="Accent.TButton")
        btn_report.pack(side=tk.LEFT)

    def _generate_report_word(self):
        # Questa funzione verrà chiamata dal pulsante nel cruscotto
        # La logica per preparare le liste (temporal_list, incongruent_list) va qui
        logger.info("Preparazione dati per il report Word...")

        # Esempio di preparazione dati (da adattare dalla logica originale)
        temporal_list = []
        incongruent_list = []
        for usage in self.all_cert_usages:
            item = usage.__dict__.copy() # Converti dataclass a dict
            item['card_date_str'] = usage.card_date.strftime('%d/%m/%Y') if usage.card_date else 'N/D'
            if usage.used_before_emission:
                item['alert_type'] = 'premature_emission'
                temporal_list.append(item)
            elif usage.is_expired_at_use:
                item['alert_type'] = 'expired_at_use'
                temporal_list.append(item)

            if usage.is_congruent is False and not usage.used_before_emission:
                incongruent_list.append(item)

        file_path = reporting.crea_e_apri_report_anomalie_word(
            self.human_errors_details,
            temporal_list,
            incongruent_list,
            self.candidate_files_count,
            self.validated_file_count
        )
        if file_path:
            messagebox.showinfo("Report Generato", f"Report Word generato e aperto:\n{file_path}", parent=self.root)
        else:
            messagebox.showwarning("Report non Generato", "Nessuna anomalia significativa trovata o si è verificato un errore.", parent=self.root)

    def _on_close(self):
        if messagebox.askokcancel("Chiudi", "Vuoi davvero chiudere l'applicazione?"):
            self.root.destroy()
            logger.info("Applicazione chiusa dall'utente.")
