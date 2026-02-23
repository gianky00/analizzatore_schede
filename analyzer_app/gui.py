"""
GUI Module - Analizzatore Schede Taratura v8.1
Modern Light Theme UI with improved responsiveness and colored logs
"""
import contextlib
import logging
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from collections import Counter, defaultdict
from datetime import datetime
from functools import partial
from tkinter import filedialog, messagebox, ttk

import pyperclip  # type: ignore

from . import analysis, config, excel_io, reporting, services
from .data_models import CertificateUsage, InstrumentSheet

logger = logging.getLogger(__name__)


# ============================================================================
# THEME COLORS - Modern Light Theme
# ============================================================================
class ThemeColors:
    """Colori tema chiaro moderno per l'applicazione."""
    # Backgrounds
    BG_PRIMARY = "#ffffff"      # White
    BG_SECONDARY = "#f8fafc"    # Very light gray
    BG_TERTIARY = "#f1f5f9"     # Light gray
    BG_CARD = "#ffffff"         # Card background
    BG_SIDEBAR = "#e2e8f0"      # Sidebar

    # Accent colors
    PRIMARY = "#2563eb"         # Blue
    PRIMARY_HOVER = "#1d4ed8"   # Darker blue
    PRIMARY_LIGHT = "#dbeafe"   # Light blue
    SUCCESS = "#16a34a"         # Green
    SUCCESS_LIGHT = "#dcfce7"   # Light green
    WARNING = "#d97706"         # Orange
    WARNING_LIGHT = "#fef3c7"   # Light orange
    ERROR = "#dc2626"           # Red
    ERROR_LIGHT = "#fee2e2"     # Light red
    INFO = "#0891b2"            # Cyan

    # Text colors
    TEXT_PRIMARY = "#0f172a"    # Very dark
    TEXT_SECONDARY = "#475569"  # Medium gray
    TEXT_MUTED = "#94a3b8"      # Light gray
    TEXT_ON_PRIMARY = "#ffffff" # White text on primary

    # Borders
    BORDER = "#e2e8f0"          # Light border
    BORDER_FOCUS = "#2563eb"    # Focus border

    # Log colors
    LOG_DEBUG = "#7c3aed"       # Purple
    LOG_INFO = "#16a34a"        # Green
    LOG_WARNING = "#d97706"     # Orange
    LOG_ERROR = "#dc2626"       # Red
    LOG_FILE = "#2563eb"        # Blue
    LOG_SUCCESS = "#059669"     # Emerald


class App:
    """Applicazione principale per l'analisi delle schede di taratura."""

    VERSION = "8.1"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Analizzatore Schede Taratura v{self.VERSION}")

        # Maximize window
        try:
            self.root.state('zoomed')
        except tk.TclError:
            w, h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            self.root.geometry(f"{w}x{h}+0+0")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.configure(bg=ThemeColors.BG_SECONDARY)

        # Data structures
        self.analysis_queue = queue.Queue()
        self.analysis_results: list[InstrumentSheet] = []
        self.all_cert_usages: list[CertificateUsage] = []
        self.human_errors_details: list[dict] = []
        self.candidate_files_count = 0
        self.validated_file_count = 0
        self.strumenti_campione: list[config.CalibrationStandard] = []
        self.cert_details_map = defaultdict(lambda: {
            'id': "", 'utilizzi': 0, 'date_utilizzo_obj_set': set(),
            'range_su_scheda_counter': Counter(), 'tipologie_scheda_associate_counter': Counter(),
            'usi_congrui': 0, 'usi_total_incongrui': 0, 'usi_prima_emissione': 0, 'usi_scaduti_puri': 0,
            'dettaglio_usi_list': []
        })
        self.last_clicked_item_id_for_toggle = [None]
        self.analysis_thread: threading.Thread | None = None
        self.analysis_service = services.AnalysisService(self.analysis_queue)

        # Setup
        self._setup_styles()
        self._create_widgets()

    def _setup_styles(self):
        """Configura gli stili ttk per un look moderno chiaro."""
        self.style = ttk.Style(self.root)

        with contextlib.suppress(tk.TclError):
            self.style.theme_use('clam')


        # Base configuration
        self.style.configure(".",
            background=ThemeColors.BG_SECONDARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            font=('Segoe UI', 10))

        # Notebook (tabs)
        self.style.configure("TNotebook",
            background=ThemeColors.BG_SECONDARY,
            borderwidth=0,
            tabmargins=[0, 5, 0, 0])
        self.style.configure("TNotebook.Tab",
            background=ThemeColors.BG_TERTIARY,
            foreground=ThemeColors.TEXT_SECONDARY,
            padding=[18, 10],
            font=('Segoe UI', 10, 'bold'))
        self.style.map("TNotebook.Tab",
            background=[("selected", ThemeColors.BG_PRIMARY)],
            foreground=[("selected", ThemeColors.PRIMARY)])

        # Frames
        self.style.configure("TFrame", background=ThemeColors.BG_SECONDARY)
        self.style.configure("Card.TFrame", background=ThemeColors.BG_CARD)

        # Labels
        self.style.configure("TLabel",
            background=ThemeColors.BG_SECONDARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            font=('Segoe UI', 10))
        self.style.configure("Title.TLabel",
            font=('Segoe UI', 18, 'bold'),
            foreground=ThemeColors.TEXT_PRIMARY,
            background=ThemeColors.BG_SECONDARY)
        self.style.configure("Subtitle.TLabel",
            font=('Segoe UI', 11),
            foreground=ThemeColors.TEXT_SECONDARY,
            background=ThemeColors.BG_SECONDARY)
        self.style.configure("CardTitle.TLabel",
            font=('Segoe UI', 12, 'bold'),
            foreground=ThemeColors.TEXT_PRIMARY,
            background=ThemeColors.BG_CARD)

        # LabelFrames
        self.style.configure("TLabelframe",
            background=ThemeColors.BG_CARD,
            bordercolor=ThemeColors.BORDER,
            relief="solid",
            borderwidth=1)
        self.style.configure("TLabelframe.Label",
            background=ThemeColors.BG_CARD,
            foreground=ThemeColors.PRIMARY,
            font=('Segoe UI', 11, 'bold'))

        # Buttons
        self.style.configure("TButton",
            background=ThemeColors.BG_TERTIARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            padding=[14, 8],
            font=('Segoe UI', 10),
            borderwidth=1)
        self.style.map("TButton",
            background=[("active", ThemeColors.BORDER)])

        self.style.configure("Accent.TButton",
            background=ThemeColors.PRIMARY,
            foreground=ThemeColors.TEXT_ON_PRIMARY,
            padding=[18, 10],
            font=('Segoe UI', 11, 'bold'))
        self.style.map("Accent.TButton",
            background=[("active", ThemeColors.PRIMARY_HOVER), ("disabled", ThemeColors.BG_TERTIARY)],
            foreground=[("disabled", ThemeColors.TEXT_MUTED)])

        self.style.configure("Success.TButton",
            background=ThemeColors.SUCCESS,
            foreground=ThemeColors.TEXT_ON_PRIMARY)

        # Entry
        self.style.configure("TEntry",
            fieldbackground=ThemeColors.BG_PRIMARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            bordercolor=ThemeColors.BORDER,
            lightcolor=ThemeColors.BORDER,
            insertcolor=ThemeColors.TEXT_PRIMARY,
            padding=8)
        self.style.map("TEntry",
            bordercolor=[("focus", ThemeColors.BORDER_FOCUS)])

        # Treeview
        self.style.configure("Treeview",
            background=ThemeColors.BG_PRIMARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            fieldbackground=ThemeColors.BG_PRIMARY,
            rowheight=36,
            font=('Segoe UI', 10),
            borderwidth=0)
        self.style.configure("Treeview.Heading",
            background=ThemeColors.BG_TERTIARY,
            foreground=ThemeColors.TEXT_PRIMARY,
            font=('Segoe UI', 10, 'bold'),
            padding=[8, 6])
        self.style.map("Treeview",
            background=[("selected", ThemeColors.PRIMARY_LIGHT)],
            foreground=[("selected", ThemeColors.PRIMARY)])

        # Progressbar
        self.style.configure("TProgressbar",
            background=ThemeColors.PRIMARY,
            troughcolor=ThemeColors.BG_TERTIARY,
            borderwidth=0,
            thickness=10)

        # Scrollbar
        self.style.configure("TScrollbar",
            background=ThemeColors.BG_TERTIARY,
            troughcolor=ThemeColors.BG_SECONDARY,
            borderwidth=0,
            arrowsize=14)

        # PanedWindow
        self.style.configure("TPanedwindow", background=ThemeColors.BG_SECONDARY)

    def _create_widgets(self):
        """Crea tutti i widget dell'interfaccia."""
        # Main container
        main_frame = ttk.Frame(self.root, style="TFrame")
        main_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=15)

        # Notebook (tabs)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(expand=True, fill='both')

        # Create tabs
        self.progress_tab = ttk.Frame(self.notebook)
        self.cruscotto_tab = ttk.Frame(self.notebook)
        self.cert_details_tab = ttk.Frame(self.notebook)
        self.correction_tab = ttk.Frame(self.notebook)
        self.suggerimenti_tab = ttk.Frame(self.notebook)
        self.autofill_tab = ttk.Frame(self.notebook)
        self.config_tab = ttk.Frame(self.notebook)

        # Add tabs
        self.notebook.add(self.progress_tab, text='  Analisi  ')
        self.notebook.add(self.cruscotto_tab, text='  Cruscotto  ', state=tk.DISABLED)
        self.notebook.add(self.cert_details_tab, text='  Certificati  ', state=tk.DISABLED)
        self.notebook.add(self.correction_tab, text='  Correzioni  ', state=tk.DISABLED)
        self.notebook.add(self.suggerimenti_tab, text='  Suggerimenti  ', state=tk.DISABLED)
        self.notebook.add(self.autofill_tab, text='  Auto-Compila  ', state=tk.DISABLED)
        self.notebook.add(self.config_tab, text='  Configurazione  ')

        # Populate tabs
        self._populate_progress_tab()
        self._populate_config_tab()

        # Check if config is valid, otherwise go to config tab
        if not config.is_config_valid():
            self.notebook.select(self.config_tab)
            messagebox.showinfo(
                "Configurazione Richiesta",
                "Benvenuto! Prima di iniziare, configura i percorsi dei file nella scheda Configurazione.",
                parent=self.root
            )
        else:
            self.notebook.select(self.progress_tab)

    def _populate_progress_tab(self):
        """Popola la tab di progresso."""
        container = ttk.Frame(self.progress_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Header
        header_frame = ttk.Frame(container, style="TFrame")
        header_frame.pack(fill=tk.X, pady=(0, 20))

        ttk.Label(header_frame,
                  text="Analisi Schede di Taratura",
                  style="Title.TLabel").pack(side=tk.LEFT)

        self.start_button = ttk.Button(
            header_frame,
            text="Avvia Analisi",
            command=self.start_analysis,
            style="Accent.TButton")
        self.start_button.pack(side=tk.RIGHT)

        # Progress section
        progress_card = ttk.LabelFrame(container, text="  Progresso  ", padding=20)
        progress_card.pack(fill=tk.X, pady=(0, 15))

        self.progress_bar = ttk.Progressbar(
            progress_card,
            orient='horizontal',
            mode='determinate',
            length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 12))

        self.progress_label = ttk.Label(
            progress_card,
            text="Pronto. Configura i percorsi e premi 'Avvia Analisi'.",
            style="Subtitle.TLabel")
        self.progress_label.pack(fill=tk.X)

        # Log section
        log_card = ttk.LabelFrame(container, text="  Log di Analisi  ", padding=15)
        log_card.pack(expand=True, fill=tk.BOTH)

        log_container = ttk.Frame(log_card)
        log_container.pack(expand=True, fill=tk.BOTH)

        log_scrollbar = ttk.Scrollbar(log_container)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text = tk.Text(
            log_container,
            wrap=tk.WORD,
            state=tk.DISABLED,
            yscrollcommand=log_scrollbar.set,
            font=("Consolas", 10),
            bg=ThemeColors.BG_PRIMARY,
            fg=ThemeColors.TEXT_PRIMARY,
            insertbackground=ThemeColors.TEXT_PRIMARY,
            selectbackground=ThemeColors.PRIMARY_LIGHT,
            selectforeground=ThemeColors.PRIMARY,
            borderwidth=1,
            relief="solid",
            highlightthickness=0,
            padx=12,
            pady=12)
        self.log_text.pack(expand=True, fill=tk.BOTH)
        log_scrollbar.config(command=self.log_text.yview)

        # Configure log tags
        self.log_text.tag_configure("DEBUG", foreground=ThemeColors.LOG_DEBUG)
        self.log_text.tag_configure("INFO", foreground=ThemeColors.LOG_INFO)
        self.log_text.tag_configure("WARNING", foreground=ThemeColors.LOG_WARNING)
        self.log_text.tag_configure("ERROR", foreground=ThemeColors.LOG_ERROR, font=("Consolas", 10, "bold"))
        self.log_text.tag_configure("SUCCESS", foreground=ThemeColors.LOG_SUCCESS, font=("Consolas", 10, "bold"))
        self.log_text.tag_configure("FILE", foreground=ThemeColors.LOG_FILE)
        self.log_text.tag_configure("TIMESTAMP", foreground=ThemeColors.TEXT_MUTED)
        self.log_text.tag_configure("SEPARATOR", foreground=ThemeColors.BORDER)

    def _log_message(self, message: str, level: str = "INFO"):
        """Aggiunge un messaggio colorato al log."""
        self.root.after(0, self._log_message_impl, message, level)

    def _log_message_impl(self, message: str, level: str):
        """Implementazione del log con colori."""
        try:
            self.log_text.config(state=tk.NORMAL)

            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{timestamp}] ", "TIMESTAMP")

            level_symbols = {
                "DEBUG": "[D]", "INFO": "[i]", "WARNING": "[!]",
                "ERROR": "[X]", "SUCCESS": "[OK]", "FILE": "[>]"
            }
            symbol = level_symbols.get(level, "[*]")

            if "---" in message and ("INIZIO" in message or "FINE" in message or "ERRORE" in message):
                self.log_text.insert(tk.END, f"\n{'─' * 70}\n", "SEPARATOR")
                tag = "SUCCESS" if "FINE" in message else ("ERROR" if "ERRORE" in message else "FILE")
                self.log_text.insert(tk.END, f"{symbol} {message}\n", tag)
                self.log_text.insert(tk.END, f"{'─' * 70}\n", "SEPARATOR")
            else:
                self.log_text.insert(tk.END, f"{symbol} {message}\n", level)

            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)

            log_level = getattr(logging, level if level in ["DEBUG", "INFO", "WARNING", "ERROR"] else "INFO")
            logger.log(log_level, message)
        except tk.TclError:
            pass

    def start_analysis(self):
        """Avvia l'analisi delegando ad AnalysisService."""
        # Verifica configurazione
        if not config.is_config_valid():
            messagebox.showerror(
                "Configurazione Mancante",
                "Configura i percorsi obbligatori nella scheda Configurazione prima di avviare l'analisi.",
                parent=self.root
            )
            self.notebook.select(self.config_tab)
            return

        self.start_button.config(state=tk.DISABLED)

        for i in range(1, 6):
            self.notebook.tab(i, state=tk.DISABLED)

        self.notebook.select(self.progress_tab)

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)

        self._log_message("Avvio analisi schede...", "INFO")
        self.progress_bar['value'] = 0

        # Deleghiamo al servizio
        self.strumenti_campione = excel_io.leggi_registro_strumenti() or []
        self.analysis_service.start_analysis(config.FOLDER_PATH_DEFAULT, self.strumenti_campione)

        self.root.after(50, self._check_analysis_queue)

    def _check_analysis_queue(self):
        """Controlla la coda messaggi per aggiornare la UI."""
        try:
            messages_processed = 0
            while not self.analysis_queue.empty() and messages_processed < 15:
                msg_type, data = self.analysis_queue.get_nowait()
                messages_processed += 1

                if msg_type == 'log':
                    message, level = data if isinstance(data, tuple) else (data, "INFO")
                    self._log_message(message, level)
                elif msg_type == 'total_files':
                    self.candidate_files_count = data
                    self.progress_bar['maximum'] = data
                elif msg_type == 'progress':
                    count, message = data
                    self.progress_bar['value'] = count
                    self.progress_label['text'] = message
                elif msg_type == 'done':
                    self.analysis_results = data
                    self.progress_label['text'] = "Analisi completata! Elaborazione risultati..."
                    self._log_message("Analisi completata con successo!", "SUCCESS")
                    self.root.update_idletasks()
                    self._process_final_results()
                    self._populate_results_ui()
                    self.start_button.config(state=tk.NORMAL)
                    return
                elif msg_type == 'error':
                    self.progress_label['text'] = f"Errore: {data}"
                    self._log_message(f"Errore fatale: {data}", "ERROR")
                    messagebox.showerror("Errore", f"Si e verificato un errore:\n{data}")
                    self.start_button.config(state=tk.NORMAL)
                    return

            if messages_processed > 0:
                self.root.update_idletasks()
        except queue.Empty:
            pass
        finally:
            # Controllo se il thread del servizio è ancora attivo
            if self.analysis_service._thread and self.analysis_service._thread.is_alive():
                self.root.after(50, self._check_analysis_queue)

    def _process_final_results(self):
        """Elabora i risultati finali."""
        self.validated_file_count = sum(1 for res in self.analysis_results if res.is_valid)
        self.all_cert_usages = [
            usage for res in self.analysis_results
            if res.is_valid
            for usage in res.certificate_usages
        ]
        self.human_errors_details = [
            {'file': res.base_filename, 'key': error.key, 'path': res.file_path}
            for res in self.analysis_results
            if res.human_errors
            for error in res.human_errors
        ]

        self._log_message(
            f"Riepilogo: {self.validated_file_count}/{self.candidate_files_count} schede valide, "
            f"{len(self.all_cert_usages)} utilizzi certificati, "
            f"{len(self.human_errors_details)} errori trovati",
            "SUCCESS"
        )
        self._update_cert_details_map()

    def _populate_results_ui(self):
        """Popola tutte le tab dei risultati."""
        for i in range(1, 7):
            self.notebook.tab(i, state=tk.NORMAL)

        self.progress_label['text'] = "Caricamento cruscotto..."
        self.root.update_idletasks()
        self._populate_cruscotto_tab()

        self.progress_label['text'] = "Caricamento certificati..."
        self.root.update_idletasks()
        self._populate_cert_details_tab()

        self.progress_label['text'] = "Caricamento correzioni..."
        self.root.update_idletasks()
        self._populate_correction_tab()

        self.progress_label['text'] = "Caricamento suggerimenti..."
        self.root.update_idletasks()
        self._populate_suggerimenti_tab()

        self.progress_label['text'] = "Caricamento auto-compilatore..."
        self.root.update_idletasks()
        self._populate_autofill_tab()

        self._populate_config_tab()

        self.progress_label['text'] = "Tutto pronto!"
        self.notebook.select(self.cruscotto_tab)

    def _populate_cruscotto_tab(self):
        """Popola la tab cruscotto."""
        for widget in self.cruscotto_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.cruscotto_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Header
        ttk.Label(container, text="Cruscotto Riepilogativo", style="Title.TLabel").pack(anchor='w', pady=(0, 20))

        # Stats grid
        stats_frame = ttk.Frame(container, style="TFrame")
        stats_frame.pack(fill=tk.X, pady=(0, 20))

        stats = [
            ("File Analizzati", str(self.candidate_files_count), ThemeColors.INFO, ThemeColors.BG_SECONDARY),
            ("Schede Valide", str(self.validated_file_count), ThemeColors.SUCCESS, ThemeColors.SUCCESS_LIGHT),
            ("Utilizzi Certificati", str(len(self.all_cert_usages)), ThemeColors.PRIMARY, ThemeColors.PRIMARY_LIGHT),
            ("Errori Trovati", str(len(self.human_errors_details)),
             ThemeColors.ERROR if self.human_errors_details else ThemeColors.SUCCESS,
             ThemeColors.ERROR_LIGHT if self.human_errors_details else ThemeColors.SUCCESS_LIGHT),
        ]

        for i, (label, value, fg_color, bg_color) in enumerate(stats):
            card = tk.Frame(stats_frame, bg=bg_color, padx=25, pady=20, highlightbackground=ThemeColors.BORDER, highlightthickness=1)
            card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0 if i == 0 else 10, 0))

            tk.Label(card, text=value, font=('Segoe UI', 32, 'bold'), bg=bg_color, fg=fg_color).pack()
            tk.Label(card, text=label, font=('Segoe UI', 11), bg=bg_color, fg=ThemeColors.TEXT_SECONDARY).pack()

        # Actions
        actions_frame = ttk.LabelFrame(container, text="  Azioni Rapide  ", padding=15)
        actions_frame.pack(fill=tk.X)

        ttk.Button(actions_frame, text="Genera Report Word", command=self._generate_report_word, style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(actions_frame, text="Apri Cartella Schede", command=lambda: self._open_path(config.FOLDER_PATH_DEFAULT)).pack(side=tk.LEFT)

    def _populate_cert_details_tab(self):
        """Popola la tab certificati."""
        for widget in self.cert_details_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.cert_details_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # Header
        header = ttk.Frame(container, style="TFrame")
        header.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(header, text="Dettaglio Utilizzo Certificati", style="Title.TLabel").pack(side=tk.LEFT)
        ttk.Label(header, text="Doppio click per aprire | Click per espandere", style="Subtitle.TLabel").pack(side=tk.RIGHT)

        # Treeview
        tree_frame = ttk.Frame(container)
        tree_frame.pack(expand=True, fill=tk.BOTH)

        cols = ["ID Certificato", "Utilizzi", "Tipologia", "Congrui", "Non Congrui",
                "Prima Emiss.", "Scaduti", "Data Recente", "Range"]

        self.tree_cert = ttk.Treeview(tree_frame, columns=cols, show='headings')

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_cert.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree_cert.xview)
        self.tree_cert.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree_cert.pack(fill='both', expand=True)

        col_widths = {"ID Certificato": 160, "Utilizzi": 70, "Tipologia": 140,
                      "Congrui": 70, "Non Congrui": 90, "Prima Emiss.": 90,
                      "Scaduti": 70, "Data Recente": 110, "Range": 180}

        for col in cols:
            self.tree_cert.heading(col, text=col, anchor=tk.W, command=partial(self._sort_treeview, self.tree_cert, col, False))
            self.tree_cert.column(col, width=col_widths.get(col, 100), minwidth=50, anchor=tk.W)

        self.tree_cert.tag_configure('child_base', font=('Consolas', 9), background=ThemeColors.BG_TERTIARY)
        self.tree_cert.tag_configure('parent_has_issues', foreground=ThemeColors.ERROR)
        self.tree_cert.tag_configure('child_error', foreground=ThemeColors.ERROR)

        data_for_tree = self._prepare_data_for_treeview()
        child_counter = 0

        for idx, row_data in enumerate(data_for_tree):
            tags = []
            if row_data["Non Congrui"] > 0 or row_data["Prima Emiss."] > 0:
                tags.append('parent_has_issues')

            parent_id = self.tree_cert.insert("", "end", values=[row_data.get(col, "") for col in cols], tags=tags)

            cert_id = row_data["ID Certificato"]
            details = self.cert_details_map.get(cert_id, {})
            usi_dett = details.get('dettaglio_usi_list', [])

            usi_sorted = sorted(usi_dett, key=lambda x: x.card_date if x.card_date else datetime.min, reverse=True)

            for uso in usi_sorted:
                child_vals = [""] * len(cols)
                date_str = uso.card_date.strftime('%d/%m/%Y') if uso.card_date else 'N/D'
                child_vals[0] = f"  > {uso.file_name} ({date_str})"
                child_vals[2] = f"{uso.tipologia_strumento_scheda}"
                congr = "OK" if uso.is_congruent is True else ("NO" if uso.is_congruent is False else "?")
                child_vals[3] = f"{congr} {uso.congruency_notes[:30]}..."

                child_tags = ['child_base']
                if uso.is_congruent is False or uso.used_before_emission or uso.is_expired_at_use:
                    child_tags.append('child_error')
                child_tags.append(uso.file_path)

                self.tree_cert.insert(parent_id, "end", values=child_vals, tags=tuple(child_tags), iid=f"child_{child_counter}")
                child_counter += 1

            if idx % 50 == 0:
                self.root.update_idletasks()

        self.tree_cert.bind("<Double-1>", self._on_tree_item_double_click)
        self.tree_cert.bind("<Button-1>", self._on_tree_item_single_click)

    def _populate_correction_tab(self):
        """Popola la tab correzioni."""
        for widget in self.correction_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.correction_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        ttk.Label(container, text="Correzione Schede con Errori", style="Title.TLabel").pack(anchor='w', pady=(0, 15))

        pane = ttk.PanedWindow(container, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True)

        # Left panel
        files_frame = ttk.Frame(pane)
        pane.add(files_frame, weight=1)

        xlsx_frame = ttk.LabelFrame(files_frame, text="  Correggibili (.xlsx)  ", padding=10)
        xlsx_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        cols = ("File", "Errori")
        self.xlsx_files_tree = ttk.Treeview(xlsx_frame, columns=cols, show='headings', height=8)
        self.xlsx_files_tree.heading("File", text="File")
        self.xlsx_files_tree.heading("Errori", text="N")
        self.xlsx_files_tree.column("File", width=200)
        self.xlsx_files_tree.column("Errori", width=50, anchor='center')
        self.xlsx_files_tree.pack(fill=tk.BOTH, expand=True)

        xls_frame = ttk.LabelFrame(files_frame, text="  Manuali (.xls)  ", padding=10)
        xls_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.xls_files_tree = ttk.Treeview(xls_frame, columns=cols, show='headings', height=8)
        self.xls_files_tree.heading("File", text="File")
        self.xls_files_tree.heading("Errori", text="N")
        self.xls_files_tree.column("File", width=200)
        self.xls_files_tree.column("Errori", width=50, anchor='center')
        self.xls_files_tree.pack(fill=tk.BOTH, expand=True)

        # Right panel
        details_frame = ttk.Frame(pane)
        pane.add(details_frame, weight=2)

        self.errors_frame = ttk.LabelFrame(details_frame, text="  Dettaglio Errori  ", padding=10)
        self.errors_frame.pack(fill=tk.BOTH, expand=True)

        self.correction_panel = ttk.LabelFrame(details_frame, text="  Correzione  ", padding=10)
        self.correction_panel.pack(fill=tk.X, pady=(10, 0))
        self.correction_panel.grid_columnconfigure(1, weight=1)

        files_with_errors = [res for res in self.analysis_results if not res.is_valid and res.human_errors]

        for res in files_with_errors:
            target = self.xlsx_files_tree if res.file_path.lower().endswith('.xlsx') else self.xls_files_tree
            target.insert("", "end", iid=res.file_path, values=(res.base_filename, len(res.human_errors)))

        self.xlsx_files_tree.bind("<<TreeviewSelect>>", self._on_file_error_select)
        self.xls_files_tree.bind("<<TreeviewSelect>>", self._on_file_error_select)

    def _on_file_error_select(self, event):
        """Gestisce selezione file con errori."""
        for widget in self.errors_frame.winfo_children():
            widget.destroy()
        for widget in self.correction_panel.winfo_children():
            widget.destroy()

        tree = event.widget
        selected = tree.focus()
        if not selected:
            return

        sheet_result = next((res for res in self.analysis_results if res.file_path == selected), None)
        if not sheet_result:
            return

        cols = ("Descrizione", "Cella")
        errors_tree = ttk.Treeview(self.errors_frame, columns=cols, show='headings', height=6)
        errors_tree.heading("Descrizione", text="Descrizione Errore")
        errors_tree.heading("Cella", text="Cella")
        errors_tree.column("Descrizione", width=350)
        errors_tree.column("Cella", width=70, anchor='center')
        errors_tree.pack(fill=tk.BOTH, expand=True)

        for i, error in enumerate(sheet_result.human_errors):
            errors_tree.insert("", "end", iid=str(i), values=(error.description, error.cell or 'N/A'))

        errors_tree.bind("<<TreeviewSelect>>", partial(self._on_error_detail_select, sheet_result, errors_tree))

    def _on_error_detail_select(self, sheet_result, errors_tree, event):
        """Gestisce selezione errore specifico."""
        for widget in self.correction_panel.winfo_children():
            widget.destroy()

        selected_id = errors_tree.focus()
        if not selected_id:
            return

        error = sheet_result.human_errors[int(selected_id)]
        is_xlsx = sheet_result.file_path.lower().endswith('.xlsx')

        ttk.Button(self.correction_panel, text="Apri Scheda", command=lambda: self._open_file(sheet_result.file_path)).grid(row=0, column=0, sticky='w', pady=5)

        if is_xlsx and error.cell:
            ttk.Label(self.correction_panel, text=f"Cella: {error.cell}").grid(row=0, column=1, sticky='w', padx=10)
            ttk.Label(self.correction_panel, text="Nuovo valore:").grid(row=1, column=0, sticky='w', pady=5)
            entry = ttk.Entry(self.correction_panel, width=40)
            if error.suggestion:
                entry.insert(0, error.suggestion)
            entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
            ttk.Button(self.correction_panel, text="Correggi e Rianalizza", style="Accent.TButton",
                       command=lambda: self._apply_correction(sheet_result.file_path, error.cell, entry.get())).grid(row=1, column=2, padx=5, pady=5)
        else:
            msg = "Correzione automatica disponibile solo per .xlsx" if not is_xlsx else "Nessuna cella specificata"
            ttk.Label(self.correction_panel, text=msg, foreground=ThemeColors.WARNING).grid(row=1, column=0, columnspan=3, sticky='w')

    def _apply_correction(self, file_path: str, cell: str, value: str):
        """Applica correzione."""
        if not cell:
            messagebox.showerror("Errore", "Nessuna cella specificata.", parent=self.root)
            return
        if excel_io.write_cell(file_path, cell, value):
            messagebox.showinfo("Successo", "Correzione applicata. Rianalisi in corso...", parent=self.root)
            self._reanalyze_single_file(file_path)
        else:
            messagebox.showerror("Errore", "Impossibile applicare la correzione.", parent=self.root)

    def _reanalyze_single_file(self, file_path: str):
        """Rianalizza singolo file."""
        self.progress_label['text'] = f"Rianalisi {os.path.basename(file_path)}..."
        self.root.update_idletasks()
        try:
            raw_data = excel_io.read_instrument_sheet_raw_data(file_path)
            new_result = analysis.analyze_sheet_data(raw_data, self.strumenti_campione)
            for i, res in enumerate(self.analysis_results):
                if res.file_path == file_path:
                    self.analysis_results[i] = new_result
                    break
            else:
                self.analysis_results.append(new_result)
        except Exception as e:
            logger.error(f"Errore rianalisi: {e}")
            messagebox.showerror("Errore", f"Impossibile rianalizzare: {e}", parent=self.root)
            return
        self._process_final_results()
        self._populate_results_ui()
        self.progress_label['text'] = "Rianalisi completata!"
        messagebox.showinfo("Completato", "Rianalisi completata.", parent=self.root)

    def _populate_suggerimenti_tab(self):
        """Popola tab suggerimenti."""
        for widget in self.suggerimenti_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.suggerimenti_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        ttk.Label(container, text="Suggerimenti Strumenti Alternativi", style="Title.TLabel").pack(anchor='w', pady=(0, 15))

        search_frame = ttk.LabelFrame(container, text="  Parametri Ricerca  ", padding=15)
        search_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(search_frame, text="ID Certificato (opz.):").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.cert_id_sugg_entry = ttk.Entry(search_frame, width=25)
        self.cert_id_sugg_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        ttk.Label(search_frame, text="Range Richiesto:").grid(row=0, column=2, sticky='w', padx=(20, 5), pady=5)
        self.range_sugg_entry = ttk.Entry(search_frame, width=25)
        self.range_sugg_entry.grid(row=0, column=3, sticky='ew', padx=5, pady=5)

        ttk.Label(search_frame, text="Data Riferimento:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.date_sugg_entry = ttk.Entry(search_frame, width=15)
        self.date_sugg_entry.insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.date_sugg_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)

        ttk.Button(search_frame, text="Cerca Alternative", command=self._search_suggestions, style="Accent.TButton").grid(row=1, column=2, columnspan=2, padx=5, pady=5)
        search_frame.columnconfigure(1, weight=1)
        search_frame.columnconfigure(3, weight=1)

        results_frame = ttk.LabelFrame(container, text="  Risultati  ", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)

        self.sugg_results_text = tk.Text(results_frame, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 10),
                                          bg=ThemeColors.BG_PRIMARY, fg=ThemeColors.TEXT_PRIMARY, borderwidth=1, relief="solid", padx=10, pady=10)
        self.sugg_results_text.pack(fill='both', expand=True)

    def _search_suggestions(self):
        """Cerca alternative."""
        cert_id = self.cert_id_sugg_entry.get().strip()
        range_req = self.range_sugg_entry.get().strip()
        date_str = self.date_sugg_entry.get().strip()

        date_ref = excel_io.parse_date_robust(date_str)
        if not date_ref:
            messagebox.showerror("Errore", "Formato data non valido. Usare gg/mm/aaaa.", parent=self.root)
            return

        results = analysis.trova_strumenti_alternativi(range_req, date_ref, self.strumenti_campione)

        self.sugg_results_text.config(state=tk.NORMAL)
        self.sugg_results_text.delete("1.0", tk.END)

        if not results:
            self.sugg_results_text.insert(tk.END, "Nessuna alternativa trovata.\n\nSuggerimenti:\n- Verifica il formato del range\n- Prova con una data diversa\n")
        else:
            count = 0
            for res in results:
                if res.id_certificato == cert_id:
                    continue
                count += 1
                scad_str = res.scadenza.strftime('%d/%m/%Y') if res.scadenza else 'N/D'
                self.sugg_results_text.insert(tk.END, f"[OK] {res.id_certificato}\n")
                self.sugg_results_text.insert(tk.END, f"     Modello: {res.modello_strumento}\n")
                self.sugg_results_text.insert(tk.END, f"     Range: {res.range}\n")
                self.sugg_results_text.insert(tk.END, f"     Scadenza: {scad_str}\n\n")
            if count == 0:
                self.sugg_results_text.insert(tk.END, "Nessuna alternativa (escluso certificato corrente).\n")

        self.sugg_results_text.config(state=tk.DISABLED)

    def _populate_autofill_tab(self):
        """Popola tab auto-compilazione."""
        for widget in self.autofill_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.autofill_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        ttk.Label(container, text="Compilatore Automatico Schede", style="Title.TLabel").pack(anchor='w', pady=(0, 15))

        action_frame = ttk.LabelFrame(container, text="  Azione  ", padding=20)
        action_frame.pack(fill=tk.X, pady=(0, 15))

        self.autofill_button = ttk.Button(action_frame, text="Avvia Compilazione Automatica", command=self._run_autofill, style="Accent.TButton")
        self.autofill_button.pack(pady=10)

        if not config.FILE_DATI_COMPILAZIONE_SCHEDE:
            self.autofill_button.config(state=tk.DISABLED)
            ttk.Label(action_frame, text="Funzione disabilitata: 'File Dati Compilazione' non configurato.", foreground=ThemeColors.WARNING).pack(pady=5)

        info_frame = ttk.LabelFrame(container, text="  Come Funziona  ", padding=15)
        info_frame.pack(fill=tk.BOTH, expand=True)

        info_text = """COMPILAZIONE AUTOMATICA SCHEDE

Questa funzione compila automaticamente i campi anagrafici mancanti
nelle schede (ODC, Data, PDL, Esecutore, Supervisore, Contratto).


PROCESSO:
1. Identifica le schede con errori di compilazione (COMP_*)
2. Per ogni scheda, cerca corrispondenza nel file dati compilazione
3. Matching basato su PDL o ODC
4. Compila i campi mancanti nel file .xlsx

ATTENZIONE:
- Verranno modificati i file .xlsx nella cartella analizzata
- I file .xls richiedono conversione manuale
- Fare sempre un backup prima di procedere

FILE RICHIESTI:
- File Dati Compilazione (nella configurazione)
- Foglio: RIASSUNTO
- Colonne: Data, Esecutore, Supervisore, ODC, PDL"""

        info_label = tk.Text(info_frame, wrap=tk.WORD, font=("Segoe UI", 10), bg=ThemeColors.BG_PRIMARY,
                             fg=ThemeColors.TEXT_PRIMARY, borderwidth=1, relief="solid", padx=15, pady=15, height=16)
        info_label.insert("1.0", info_text)
        info_label.config(state=tk.DISABLED)
        info_label.pack(fill=tk.BOTH, expand=True)

    def _run_autofill(self):
        """Esegue compilazione automatica delegando al Service Layer."""
        self._log_message("Avvio compilazione automatica...", "INFO")

        if not config.FILE_DATI_COMPILAZIONE_SCHEDE or not os.path.exists(config.FILE_DATI_COMPILAZIONE_SCHEDE):
            msg = f"File dati compilazione non trovato: {config.FILE_DATI_COMPILAZIONE_SCHEDE}"
            self._log_message(msg, "ERROR")
            messagebox.showerror("Errore", msg, parent=self.root)
            return

        try:
            modifiche = services.AutofillService.run_autofill(
                self.analysis_results,
                config.FILE_DATI_COMPILAZIONE_SCHEDE,
                self._log_message
            )

            if modifiche > 0:
                self._log_message(f"Compilazione completata: {modifiche} schede modificate.", "SUCCESS")
                messagebox.showinfo("Completato", f"{modifiche} schede sono state aggiornate.\n\nRianalizzare per verificare le modifiche.", parent=self.root)
            else:
                self._log_message("Nessuna scheda modificata.", "WARNING")
                messagebox.showinfo("Completato", "Nessuna scheda e stata modificata.", parent=self.root)

        except Exception as e:
            self._log_message(f"Errore durante l'autofill: {e}", "ERROR")
            messagebox.showerror("Errore", f"Impossibile completare l'autofill:\n{e}", parent=self.root)


    def _populate_config_tab(self):
        """Popola la tab configurazione."""
        for widget in self.config_tab.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.config_tab, style="TFrame")
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        ttk.Label(container, text="Configurazione", style="Title.TLabel").pack(anchor='w', pady=(0, 5))
        ttk.Label(container, text="I percorsi vengono salvati automaticamente", style="Subtitle.TLabel").pack(anchor='w', pady=(0, 20))

        # Required paths
        req_frame = ttk.LabelFrame(container, text="  Percorsi Obbligatori  ", padding=15)
        req_frame.pack(fill=tk.X, pady=(0, 15))

        self.config_entries = {}

        req_items = [
            ("FILE_REGISTRO_STRUMENTI", "Registro Strumenti:", False),
            ("FOLDER_PATH_DEFAULT", "Cartella Schede:", True),
        ]

        for row, (key, label, is_folder) in enumerate(req_items):
            ttk.Label(req_frame, text=label).grid(row=row, column=0, sticky='w', padx=5, pady=8)
            entry = ttk.Entry(req_frame, width=70)
            entry.grid(row=row, column=1, sticky='ew', padx=5, pady=8)
            current = getattr(config, key, "") or ""
            if current:
                entry.insert(0, current)
            self.config_entries[key] = entry
            cmd = partial(self._browse_folder, entry) if is_folder else partial(self._browse_file, entry)
            ttk.Button(req_frame, text="Sfoglia", command=cmd).grid(row=row, column=2, padx=5, pady=8)

            # Button "Apri Cartella"
            def open_dir(e=entry, folder=is_folder):
                path = e.get()
                if path:
                    target = path if folder else os.path.dirname(path)
                    self._open_path(target)

            ttk.Button(req_frame, text="Apri Cartella", command=open_dir).grid(row=row, column=3, padx=5, pady=8)

        req_frame.columnconfigure(1, weight=1)

        # Optional paths
        opt_frame = ttk.LabelFrame(container, text="  Percorsi Opzionali  ", padding=15)
        opt_frame.pack(fill=tk.X, pady=(0, 15))

        opt_items = [
            ("FILE_DATI_COMPILAZIONE_SCHEDE", "Dati Compilazione:", False),
            ("FILE_MASTER_DIGITALE_XLSX", "Master Digitale:", False),
            ("FILE_MASTER_ANALOGICO_XLSX", "Master Analogico:", False),
        ]

        opt_descriptions = {
            "FILE_DATI_COMPILAZIONE_SCHEDE": "File Excel (RIASSUNTO) per auto-compilazione dati.",
            "FILE_MASTER_DIGITALE_XLSX": "Template Excel per schede Digitali.",
            "FILE_MASTER_ANALOGICO_XLSX": "Template Excel per schede Analogiche."
        }

        for row, (key, label, is_folder) in enumerate(opt_items):
            ttk.Label(opt_frame, text=label).grid(row=row, column=0, sticky='w', padx=5, pady=8)
            entry = ttk.Entry(opt_frame, width=70)
            entry.grid(row=row, column=1, sticky='ew', padx=5, pady=8)
            current = getattr(config, key, "") or ""
            if current:
                entry.insert(0, current)
            self.config_entries[key] = entry
            cmd = partial(self._browse_folder, entry) if is_folder else partial(self._browse_file, entry)
            ttk.Button(opt_frame, text="Sfoglia", command=cmd).grid(row=row, column=2, padx=5, pady=8)

            # Suggestion label
            desc = opt_descriptions.get(key, "")
            if desc:
                ttk.Label(opt_frame, text=desc, font=('Segoe UI', 9), foreground=ThemeColors.TEXT_SECONDARY).grid(row=row, column=3, sticky='w', padx=10, pady=8)

        opt_frame.columnconfigure(1, weight=1)

        # Buttons
        btn_frame = ttk.Frame(container, style="TFrame")
        btn_frame.pack(fill=tk.X, pady=20)

        ttk.Button(btn_frame, text="Salva Configurazione", command=self._save_config, style="Accent.TButton").pack(side=tk.LEFT)
        ttk.Label(btn_frame, text="Le modifiche saranno applicate immediatamente", style="Subtitle.TLabel").pack(side=tk.LEFT, padx=20)

    def _update_cert_details_map(self):
        """Aggiorna mappa dettagli certificati delegando al Service Layer."""
        self.cert_details_map = services.StatisticsService.calculate_cert_details(self.analysis_results)

    def _prepare_data_for_treeview(self) -> list[dict]:
        """Prepara dati per treeview."""
        tree_data = []
        for cert_id, details in self.cert_details_map.items():
            valid_dates = [d for d in details.get('date_utilizzo_obj_set', set()) if d]
            try:
                date_rec = max(valid_dates).strftime('%d/%m/%Y') if valid_dates else "N/D"
            except Exception:
                date_rec = "N/D"

            range_counter = details.get('range_su_scheda_counter', Counter())
            range_p = range_counter.most_common(1)[0][0] if range_counter else "N/D"
            tip_counter = details.get('tipologie_scheda_associate_counter', Counter())
            tip_p = tip_counter.most_common(1)[0][0] if tip_counter else "N/D"
            tree_data.append({"ID Certificato": cert_id, "Utilizzi": details.get('utilizzi', 0), "Tipologia": tip_p, "Congrui": details.get('usi_congrui', 0), "Non Congrui": details.get('usi_total_incongrui', 0), "Prima Emiss.": details.get('usi_prima_emissione', 0), "Scaduti": details.get('usi_scaduti_puri', 0), "Data Recente": date_rec, "Range": range_p})
        return sorted(tree_data, key=lambda x: (-x.get("Prima Emiss.", 0), -x.get("Non Congrui", 0), -x.get("Utilizzi", 0)))

    def _on_tree_item_single_click(self, event):
        """Gestisce click singolo."""
        item_id = self.tree_cert.identify_row(event.y)
        if item_id and not self.tree_cert.parent(item_id):
            if item_id == self.last_clicked_item_id_for_toggle[0]:
                self.tree_cert.item(item_id, open=not self.tree_cert.item(item_id, 'open'))
                self.last_clicked_item_id_for_toggle[0] = None
            else:
                if self.last_clicked_item_id_for_toggle[0] and self.tree_cert.exists(self.last_clicked_item_id_for_toggle[0]):
                    self.tree_cert.item(self.last_clicked_item_id_for_toggle[0], open=False)
                self.tree_cert.item(item_id, open=True)
                self.last_clicked_item_id_for_toggle[0] = item_id

    def _on_tree_item_double_click(self, event):
        """Gestisce doppio click."""
        item_id = self.tree_cert.identify_row(event.y)
        if not item_id:
            return
        if self.tree_cert.parent(item_id):
            tags = self.tree_cert.item(item_id, 'tags')
            for tag in tags:
                if isinstance(tag, str) and (tag.lower().endswith('.xls') or tag.lower().endswith('.xlsx')):
                    self._open_file(tag)
                    return
        else:
            values = self.tree_cert.item(item_id, 'values')
            cert_id, range_val = values[0], values[8]
            self.notebook.select(self.suggerimenti_tab)
            self.cert_id_sugg_entry.delete(0, tk.END)
            self.cert_id_sugg_entry.insert(0, cert_id)
            self.range_sugg_entry.delete(0, tk.END)
            self.range_sugg_entry.insert(0, range_val)
            self._search_suggestions()

    def _generate_report_word(self):
        """Genera report Word."""
        self._log_message("Generazione report Word...", "INFO")
        temporal_list, incongruent_list = [], []
        for usage in self.all_cert_usages:
            item = usage.__dict__.copy()
            item['card_date_str'] = usage.card_date.strftime('%d/%m/%Y') if usage.card_date else 'N/D'
            if usage.used_before_emission:
                item['alert_type'] = 'premature_emission'
                temporal_list.append(item)
            elif usage.is_expired_at_use:
                item['alert_type'] = 'expired_at_use'
                temporal_list.append(item)
            if usage.is_congruent is False and not usage.used_before_emission:
                incongruent_list.append(item)
        file_path = reporting.crea_e_apri_report_anomalie_word(self.human_errors_details, temporal_list, incongruent_list, self.candidate_files_count, self.validated_file_count)
        if file_path:
            self._log_message(f"Report generato: {file_path}", "SUCCESS")
            messagebox.showinfo("Report Generato", f"Report salvato e aperto:\n{file_path}", parent=self.root)
        else:
            messagebox.showwarning("Attenzione", "Nessuna anomalia da riportare o errore nella generazione.", parent=self.root)

    def _open_file(self, file_path: str):
        """Apre un file."""
        try:
            normalized = os.path.normpath(file_path)
            if normalized.startswith('\\') and not normalized.startswith('\\\\'):
                normalized = '\\' + normalized
            pyperclip.copy(normalized)
            if messagebox.askyesno("Conferma", f"Percorso copiato:\n{normalized}\n\nAprire il file?", parent=self.root):
                if sys.platform == "win32":
                    os.startfile(normalized)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", normalized])
                else:
                    subprocess.Popen(["xdg-open", normalized])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire:\n{e}", parent=self.root)

    def _open_path(self, path: str):
        """Apre percorso."""
        try:
            normalized = os.path.normpath(path) if path else ""
            if normalized and os.path.exists(normalized):
                if sys.platform == "win32":
                    os.startfile(normalized)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", normalized])
                else:
                    subprocess.Popen(["xdg-open", normalized])
            else:
                messagebox.showerror("Errore", f"Percorso non trovato:\n{normalized}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire:\n{e}", parent=self.root)

    def _browse_file(self, entry_widget):
        """Apre dialogo file."""
        filepath = filedialog.askopenfilename(title="Seleziona File", filetypes=(("Excel", "*.xlsx *.xlsm *.xls"), ("Tutti", "*.*")))
        if filepath:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filepath)

    def _browse_folder(self, entry_widget):
        """Apre dialogo cartella."""
        folderpath = filedialog.askdirectory(title="Seleziona Cartella")
        if folderpath:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folderpath)

    def _save_config(self):
        """Salva configurazione delegando al Service Layer."""
        new_config = {key: entry.get() for key, entry in self.config_entries.items()}
        if services.ConfigService.save_and_reload(new_config):
            messagebox.showinfo("Successo", "Configurazione salvata!", parent=self.root)
        else:
            messagebox.showerror("Errore", "Impossibile salvare la configurazione.", parent=self.root)

    def _sort_treeview(self, tree: ttk.Treeview, col: str, reverse: bool):
        """Ordina treeview."""
        try:
            items = [(tree.set(item, col), item) for item in tree.get_children('')]
            def sort_key(item_tuple):
                value = item_tuple[0]
                if value in ("", "N/D"):
                    return (1, "")
                try:
                    return (0, float(value))
                except ValueError:
                    return (0, str(value).lower())

            items.sort(key=sort_key, reverse=reverse)
            for idx, (_, item) in enumerate(items):
                tree.move(item, '', idx)
            for c in tree['columns']:
                tree.heading(c, text=c, command=partial(self._sort_treeview, tree, c, False))
            arrow = " v" if reverse else " ^"
            tree.heading(col, text=col + arrow, command=partial(self._sort_treeview, tree, col, not reverse))
        except Exception as e:
            logger.warning(f"Errore ordinamento: {e}")

    def _on_close(self):
        """Gestisce chiusura."""
        if messagebox.askokcancel("Chiudi", "Vuoi chiudere l'applicazione?", parent=self.root):
            self.root.destroy()
            logger.info("Applicazione chiusa.")
