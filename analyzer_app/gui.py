import os
import tkinter as tk
from tkinter import ttk, messagebox, font as tkFont, filedialog
from functools import partial
import logging
import threading
import queue
import pyperclip # type: ignore
import re
import subprocess
import sys
import multiprocessing
from collections import Counter, defaultdict
from datetime import datetime
from typing import List, Dict

from . import config
from . import excel_io
from . import analysis
from . import reporting
from .data_models import InstrumentSheet, CertificateUsage, SheetError

logger = logging.getLogger(__name__)

def read_file_worker(q, file_path):
    try:
        raw_data = excel_io.read_instrument_sheet_raw_data(file_path)
        q.put(('success', raw_data))
    except Exception as e:
        q.put(('error', e))

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Analisi Schede Taratura - v6.0 Optimized Re-analysis")
        self.root.geometry("1750x980")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.analysis_queue = queue.Queue()
        self.analysis_results: List[InstrumentSheet] = []
        self.all_cert_usages: List[CertificateUsage] = []
        self.human_errors_details: List[Dict] = []
        self.candidate_files_count = 0
        self.validated_file_count = 0
        self.strumenti_campione: List[config.CalibrationStandard] = []
        self.cert_details_map = defaultdict(lambda: {
            'id': "", 'utilizzi': 0, 'date_utilizzo_obj_set': set(),
            'range_su_scheda_counter': Counter(), 'tipologie_scheda_associate_counter': Counter(),
            'usi_congrui': 0, 'usi_total_incongrui': 0, 'usi_prima_emissione': 0, 'usi_scaduti_puri': 0,
            'dettaglio_usi_list': []
        })
        self.last_clicked_item_id_for_toggle = [None]

        self._setup_styles()
        self.create_widgets()

    def _setup_styles(self):
        self.style = ttk.Style(self.root)
        try:
            theme = 'vista' if 'vista' in self.style.theme_names() else 'clam'
            self.style.theme_use(theme)
        except tk.TclError: logger.warning("Tema 'vista' o 'clam' non trovato.")
        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), relief="groove")
        self.style.configure("Treeview", rowheight=28, font=('Segoe UI', 9))
        self.style.configure("TNotebook.Tab", font=('Segoe UI', 10, 'bold'), padding=[12, 6])
        self.style.configure("TLabelframe.Label", font=('Segoe UI', 11, 'bold'))
        self.style.configure("Accent.TButton", font=('Segoe UI', 10, 'bold'), padding=8)
        self.style.configure("Hyperlink.TLabel", foreground="blue", font=('Segoe UI', 9, 'underline'))

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(expand=True, fill=tk.BOTH)
        self.notebook = ttk.Notebook(main_frame, style="TNotebook")
        self.notebook.pack(expand=True, fill='both', pady=(0, 10))

        self.progress_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.progress_tab, text=' Progresso Analisi ')
        self.start_button = ttk.Button(self.progress_tab, text="Avvia Analisi", command=self.start_analysis, style="Accent.TButton")
        self.start_button.pack(pady=10)
        log_frame = ttk.LabelFrame(self.progress_tab, text="Log di Analisi", padding=10)
        log_frame.pack(expand=True, fill=tk.BOTH)
        log_v_scroll = ttk.Scrollbar(log_frame); log_v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state=tk.DISABLED, yscrollcommand=log_v_scroll.set, font=("Consolas", 10))
        self.log_text.pack(expand=True, fill=tk.BOTH); log_v_scroll.config(command=self.log_text.yview)
        self.progress_bar = ttk.Progressbar(self.progress_tab, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)
        self.progress_label = ttk.Label(self.progress_tab, text="Pronto per iniziare l'analisi. Modificare la configurazione o premere 'Avvia Analisi'.")
        self.progress_label.pack(fill=tk.X)

        self.cruscotto_tab = ttk.Frame(self.notebook, padding=10)
        self.cert_details_tab = ttk.Frame(self.notebook, padding=10)
        self.correction_tab = ttk.Frame(self.notebook, padding=10)
        self.suggerimenti_tab = ttk.Frame(self.notebook, padding=10)
        self.config_tab = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.cruscotto_tab, text=' Cruscotto Riepilogativo ', state=tk.DISABLED)
        self.notebook.add(self.cert_details_tab, text=' Dettaglio Utilizzo Certificati ', state=tk.DISABLED)
        self.notebook.add(self.correction_tab, text=' Correzione Schede ', state=tk.DISABLED)
        self.notebook.add(self.suggerimenti_tab, text=' Suggerimenti Strumenti ', state=tk.DISABLED)
        self.notebook.add(self.config_tab, text=' Configurazione ')
        self._populate_config_tab()

    def _log_message(self, message, level="INFO"):
        self.root.after(0, self.__log_message_thread_safe, message, level)

    def __log_message_thread_safe(self, message, level):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{level}] {message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        logger.log(logging.getLevelName(level), message)

    def start_analysis(self):
        self.start_button.config(state=tk.DISABLED)
        for i in self.notebook.tabs():
            if self.notebook.index(i) > 0: self.notebook.tab(i, state=tk.DISABLED)
        self.notebook.select(self.progress_tab)
        self.log_text.config(state=tk.NORMAL); self.log_text.delete('1.0', tk.END); self.log_text.config(state=tk.DISABLED)
        self._log_message("Avvio del thread di analisi...")
        self.progress_bar['value'] = 0
        self.analysis_thread = threading.Thread(target=self._analysis_worker, daemon=True)
        self.analysis_thread.start()
        self.root.after(100, self._check_analysis_queue)

    def _analysis_worker(self):
        try:
            config.load_config()
            self.analysis_queue.put(('log', "Configurazione ricaricata."))
            self.analysis_queue.put(('log', "Lettura registro strumenti..."))
            self.strumenti_campione = excel_io.leggi_registro_strumenti() or []
            self.analysis_queue.put(('log', f"Letti {len(self.strumenti_campione)} strumenti validi dal registro."))
            folder_path = config.FOLDER_PATH_DEFAULT
            if not folder_path or not os.path.isdir(folder_path): raise NotADirectoryError(f"Cartella schede non valida: {folder_path}")
            candidate_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xls', '.xlsx')) and not f.startswith('~')]
            self.candidate_files_count = len(candidate_files)
            self.analysis_queue.put(('log', f"Trovati {self.candidate_files_count} file candidati."))
            self.analysis_queue.put(('total_files', self.candidate_files_count))
            results = []
            for i, filename in enumerate(candidate_files):
                file_path = os.path.join(folder_path, filename)
                self.analysis_queue.put(('log', f"--- INIZIO elaborazione file {i+1}/{self.candidate_files_count}: {filename} ---"))
                self.analysis_queue.put(('progress', (i + 1, f"Analisi di: {filename}")))
                try:
                    self.analysis_queue.put(('log', f"Fase 1: Lettura dati da {filename} (con timeout di 30s)"))
                    q = multiprocessing.Queue()
                    p = multiprocessing.Process(target=read_file_worker, args=(q, file_path))
                    p.start(); p.join(30)
                    if p.is_alive():
                        p.terminate(); p.join()
                        raise TimeoutError("La lettura del file ha superato i 30 secondi.")
                    status, result = q.get()
                    if status == 'error': raise result
                    raw_data = result
                    self.analysis_queue.put(('log', f"Fase 2: Analisi logica per {filename}"))
                    sheet_result = analysis.analyze_sheet_data(raw_data, self.strumenti_campione)
                    results.append(sheet_result)
                    self.analysis_queue.put(('log', f"--- FINE elaborazione file: {sheet_result.status}"))
                except Exception as e:
                    logger.error(f"Errore durante l'analisi del file {filename}: {e}", exc_info=True)
                    results.append(InstrumentSheet(file_path=file_path, base_filename=filename, status=f"Errore: {e}", is_valid=False))
                    self.analysis_queue.put(('log', f"--- ERRORE elaborazione file: {filename} ---"))
            self.analysis_queue.put(('done', results))
        except Exception as e:
            logger.critical(f"Errore fatale nel thread di analisi: {e}", exc_info=True)
            self.analysis_queue.put(('error', e))

    def _check_analysis_queue(self):
        try:
            while not self.analysis_queue.empty():
                msg_type, data = self.analysis_queue.get_nowait()
                if msg_type == 'log': self._log_message(data)
                elif msg_type == 'total_files': self.progress_bar['maximum'] = data
                elif msg_type == 'progress':
                    count, message = data
                    self.progress_bar['value'] = count
                    self.progress_label['text'] = message
                elif msg_type == 'done':
                    self.analysis_results = data
                    self.progress_label['text'] = "Analisi completata. Elaborazione risultati..."
                    self._process_final_results()
                    self._populate_results_ui()
                    self.start_button.config(state=tk.NORMAL)
                    return
                elif msg_type == 'error':
                    self.progress_label['text'] = f"Errore durante l'analisi: {data}"
                    messagebox.showerror("Errore di Analisi", f"Si è verificato un errore: {data}")
                    self.start_button.config(state=tk.NORMAL)
                    return
        except queue.Empty: pass
        finally:
            if self.analysis_thread.is_alive(): self.root.after(100, self._check_analysis_queue)

    def _process_final_results(self):
        self.validated_file_count = sum(1 for res in self.analysis_results if res.is_valid)
        self.all_cert_usages = [usage for res in self.analysis_results if res.is_valid for usage in res.certificate_usages]
        self.human_errors_details = [{'file': res.base_filename, 'key': error.key, 'path': res.file_path} for res in self.analysis_results if res.is_valid for error in res.human_errors]
        self._log_message(f"Elaborazione completata. Schede validate: {self.validated_file_count}/{self.candidate_files_count}")
        self._update_cert_details_map()

    def _populate_results_ui(self):
        for tab in [self.cruscotto_tab, self.cert_details_tab, self.correction_tab, self.suggerimenti_tab, self.config_tab]:
            self.notebook.tab(tab, state=tk.NORMAL)
        self._populate_cruscotto_tab()
        self._populate_cert_details_tab()
        self._populate_correction_tab()
        self._populate_suggerimenti_tab()
        self._populate_config_tab()
        self.notebook.select(self.cruscotto_tab)

    def _populate_cruscotto_tab(self):
        for widget in self.cruscotto_tab.winfo_children(): widget.destroy()
        stats_frame = ttk.LabelFrame(self.cruscotto_tab, text="Statistiche Generali", padding=10)
        stats_frame.pack(fill=tk.X, pady=5, anchor='n')
        ttk.Label(stats_frame, text=f"File analizzati: {self.candidate_files_count}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Schede validate: {self.validated_file_count}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Utilizzi certificati totali: {len(self.all_cert_usages)}").pack(anchor=tk.W)
        ttk.Label(stats_frame, text=f"Errori di compilazione trovati: {len(self.human_errors_details)}").pack(anchor=tk.W)
        action_frame = ttk.LabelFrame(self.cruscotto_tab, text="Azioni", padding=10)
        action_frame.pack(fill=tk.X, pady=5, anchor='n')
        btn_report = ttk.Button(action_frame, text="Stampa Report Anomalie (Word)", command=self._generate_report_word, style="Accent.TButton")
        btn_report.pack(side=tk.LEFT)

    def _populate_cert_details_tab(self):
        for widget in self.cert_details_tab.winfo_children(): widget.destroy()
        cols = ["ID Certificato", "Utilizzi", "Tipologia Principale", "Congrui", "Non Congrui", "Prima Emiss.", "Scaduti", "Scadenza Recente", "Range Principale"]
        self.tree_cert = ttk.Treeview(self.cert_details_tab, columns=cols, show='headings')
        vsb = ttk.Scrollbar(self.cert_details_tab, orient="vertical", command=self.tree_cert.yview)
        hsb = ttk.Scrollbar(self.cert_details_tab, orient="horizontal", command=self.tree_cert.xview)
        self.tree_cert.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x'); self.tree_cert.pack(fill='both', expand=True)
        col_widths = {"ID Certificato":180, "Utilizzi":60, "Tipologia Principale":170, "Congrui":70, "Non Congrui":90, "Prima Emiss.":90, "Scaduti":70, "Scadenza Recente":120, "Range Principale":200}
        for col_name in cols:
            self.tree_cert.heading(col_name, text=col_name, anchor=tk.W, command=partial(self._sort_treeview, self.tree_cert, col_name, False))
            self.tree_cert.column(col_name, width=col_widths.get(col_name, 120), minwidth=60, anchor=tk.W)
        self.tree_cert.tag_configure('child_base', font=tkFont.Font(family='Consolas', size=8), background='#FAFAFA')
        self.tree_cert.tag_configure('parent_has_issues', foreground='red', font=tkFont.Font(weight='bold'))
        self.tree_cert.tag_configure('child_error', foreground='red')
        data_for_tree = self._prepare_data_for_treeview()
        child_item_counter = 0
        self.tree_cert.delete(*self.tree_cert.get_children())
        for row_data in data_for_tree:
            tags = []
            if row_data["Non Congrui"] > 0 or row_data["Prima Emiss."] > 0: tags.append('parent_has_issues')
            parent_item_id = self.tree_cert.insert("", "end", values=[row_data.get(col, "") for col in cols], tags=tags)
            cert_id = row_data["ID Certificato"]
            usi_dett = sorted(self.cert_details_map.get(cert_id, {}).get('dettaglio_usi_list', []), key=lambda x: x.card_date, reverse=True)
            for uso_info in usi_dett:
                child_vals = [""] * len(cols)
                child_vals[0] = f"  └─File: {uso_info.file_name} (Scheda: {uso_info.card_date.strftime('%d/%m/%Y') if uso_info.card_date else 'N/D'})"
                child_vals[2] = f"Tip.Strum: {uso_info.tipologia_strumento_scheda}, Modello: {uso_info.modello_L9_scheda}"
                congr_str = "Congruo" if uso_info.is_congruent else "NON Congruo" if uso_info.is_congruent is False else "N/V"
                child_vals[3] = f"{congr_str} ({uso_info.congruency_notes})"
                child_tags = ['child_base']
                if not uso_info.is_congruent or uso_info.used_before_emission or uso_info.is_expired_at_use: child_tags.append('child_error')
                unique_iid = f"{uso_info.file_path}_{child_item_counter}"
                self.tree_cert.insert(parent_item_id, "end", values=child_vals, tags=tuple(child_tags), iid=unique_iid)
                child_item_counter += 1
        self.tree_cert.bind("<Double-1>", self._on_tree_item_double_click)
        self.tree_cert.bind("<Button-1>", self._on_tree_item_single_click)

    def _populate_correction_tab(self):
        for widget in self.correction_tab.winfo_children(): widget.destroy()
        pane = ttk.PanedWindow(self.correction_tab, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True)

        files_frame = ttk.Frame(pane)
        pane.add(files_frame, weight=1)

        # XLSX Tree (Modifiable)
        xlsx_frame = ttk.LabelFrame(files_frame, text="Correggibili Automaticamente (.xlsx)", padding=5)
        xlsx_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        cols = ("File", "Errori")
        self.xlsx_files_tree = ttk.Treeview(xlsx_frame, columns=cols, show='headings')
        self.xlsx_files_tree.heading("File", text="File"); self.xlsx_files_tree.heading("Errori", text="N. Errori")
        self.xlsx_files_tree.column("File", width=250); self.xlsx_files_tree.column("Errori", width=50, anchor='center')
        self.xlsx_files_tree.pack(fill=tk.BOTH, expand=True)

        # XLS Tree (Manual)
        xls_frame = ttk.LabelFrame(files_frame, text="Da Aprire Manualmente (.xls)", padding=5)
        xls_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.xls_files_tree = ttk.Treeview(xls_frame, columns=cols, show='headings')
        self.xls_files_tree.heading("File", text="File"); self.xls_files_tree.heading("Errori", text="N. Errori")
        self.xls_files_tree.column("File", width=250); self.xls_files_tree.column("Errori", width=50, anchor='center')
        self.xls_files_tree.pack(fill=tk.BOTH, expand=True)

        details_pane = ttk.Frame(pane, padding=10)
        pane.add(details_pane, weight=2)
        self.errors_frame = ttk.LabelFrame(details_pane, text="Dettaglio Errori", padding=10)
        self.errors_frame.pack(fill=tk.BOTH, expand=True)
        self.correction_panel = ttk.LabelFrame(details_pane, text="Pannello di Correzione", padding=10)
        self.correction_panel.pack(fill=tk.X, pady=10)
        self.correction_panel.grid_columnconfigure(1, weight=1)

        files_with_errors = [res for res in self.analysis_results if res.is_valid and res.human_errors]
        self.xlsx_files_tree.delete(*self.xlsx_files_tree.get_children())
        self.xls_files_tree.delete(*self.xls_files_tree.get_children())
        for res in files_with_errors:
            if res.file_path.lower().endswith('.xlsx'):
                self.xlsx_files_tree.insert("", "end", iid=res.file_path, values=(res.base_filename, len(res.human_errors)))
            else:
                self.xls_files_tree.insert("", "end", iid=res.file_path, values=(res.base_filename, len(res.human_errors)))

        self.xlsx_files_tree.bind("<<TreeviewSelect>>", self._on_file_error_select)
        self.xls_files_tree.bind("<<TreeviewSelect>>", self._on_file_error_select)

    def _on_file_error_select(self, event):
        for widget in self.errors_frame.winfo_children(): widget.destroy()
        for widget in self.correction_panel.winfo_children(): widget.destroy()

        tree = event.widget
        selected_item = tree.focus()
        if not selected_item: return
        sheet_result = next((res for res in self.analysis_results if res.file_path == selected_item), None)
        if not sheet_result: return
        cols = ("Descrizione", "Cella", "Suggerimento")
        errors_tree = ttk.Treeview(self.errors_frame, columns=cols, show='headings')
        for col in cols: errors_tree.heading(col, text=col)
        errors_tree.pack(fill=tk.BOTH, expand=True)
        for i, error in enumerate(sheet_result.human_errors):
            errors_tree.insert("", "end", iid=str(i), values=(error.description, error.cell or 'N/A', error.suggestion or ''))
        errors_tree.bind("<<TreeviewSelect>>", partial(self._on_error_detail_select, sheet_result, errors_tree))

    def _on_error_detail_select(self, sheet_result, errors_tree, event):
        for widget in self.correction_panel.winfo_children(): widget.destroy()
        selected_item_id = errors_tree.focus()
        if not selected_item_id: return
        selected_error = sheet_result.human_errors[int(selected_item_id)]

        is_xlsx = sheet_result.file_path.lower().endswith('.xlsx')

        if is_xlsx:
            ttk.Label(self.correction_panel, text="Cella da modificare:").grid(row=0, column=0, sticky='w')
            ttk.Label(self.correction_panel, text=selected_error.cell or "N/A", font=('Segoe UI', 10, 'bold')).grid(row=0, column=1, sticky='w')
            ttk.Label(self.correction_panel, text="Nuovo Valore:").grid(row=1, column=0, sticky='w')
            entry = ttk.Entry(self.correction_panel)
            if selected_error.suggestion: entry.insert(0, selected_error.suggestion)
            entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)
            btn_correct = ttk.Button(self.correction_panel, text="Correggi e Rianalizza", style="Accent.TButton",
                                     command=lambda: self._apply_correction(sheet_result.file_path, selected_error.cell, entry.get()))
            btn_correct.grid(row=2, column=1, sticky='e', pady=5)
        else:
            ttk.Label(self.correction_panel, text="La modifica automatica è supportata solo per file .xlsx.").pack(pady=5)

        btn_open = ttk.Button(self.correction_panel, text="Apri Scheda",
                              command=lambda: self._on_file_click(sheet_result.file_path, sheet_result.base_filename, open_file_direct=True))
        btn_open.pack(pady=5)

    def _apply_correction(self, file_path, cell, value):
        if not cell:
            messagebox.showerror("Errore", "Nessuna cella specificata per questo errore.", parent=self.root)
            return
        if not file_path.lower().endswith('.xlsx'):
            messagebox.showwarning("Funzionalità Limitata", "La correzione automatica è supportata solo per i file .xlsx.", parent=self.root)
            return
        if excel_io.write_cell(file_path, cell, value):
            messagebox.showinfo("Successo", "Correzione applicata. Rianalisi del file in corso...", parent=self.root)
            self._reanalyze_single_file(file_path)
        else:
            messagebox.showerror("Errore", "Impossibile applicare la correzione. Controllare i log.", parent=self.root)

    def _reanalyze_single_file(self, file_path):
        self.progress_label['text'] = f"Rianalisi di {os.path.basename(file_path)}..."
        self.root.update_idletasks()
        try:
            raw_data = excel_io.read_instrument_sheet_raw_data(file_path)
            new_result = analysis.analyze_sheet_data(raw_data, self.strumenti_campione)
            # Find and replace the old result
            index_to_replace = -1
            for i, res in enumerate(self.analysis_results):
                if res.file_path == file_path:
                    index_to_replace = i
                    break
            if index_to_replace != -1:
                self.analysis_results[index_to_replace] = new_result
            else: # Should not happen if called from correction tab
                self.analysis_results.append(new_result)
        except Exception as e:
            logger.error(f"Errore durante la rianalisi del file {os.path.basename(file_path)}: {e}")
            messagebox.showerror("Errore Rianalisi", f"Impossibile rianalizzare il file: {e}", parent=self.root)

        self._process_final_results()
        self._populate_results_ui()
        self.progress_label['text'] = "Rianalisi completata."
        messagebox.showinfo("Completato", "Rianalisi completata. La vista è stata aggiornata.", parent=self.root)


    def _populate_suggerimenti_tab(self):
        for widget in self.suggerimenti_tab.winfo_children(): widget.destroy()
        input_frame = ttk.LabelFrame(self.suggerimenti_tab, text="Parametri Ricerca", padding=10)
        input_frame.pack(fill=tk.X, pady=5)
        ttk.Label(input_frame, text="ID Certificato (opz.):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.cert_id_sugg_entry = ttk.Entry(input_frame, width=25)
        self.cert_id_sugg_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        ttk.Label(input_frame, text="Range Richiesto:").grid(row=0, column=2, padx=(10,5), pady=5, sticky=tk.W)
        self.range_sugg_entry = ttk.Entry(input_frame, width=30)
        self.range_sugg_entry.grid(row=0, column=3, sticky=tk.EW, padx=5, pady=5)
        ttk.Label(input_frame, text="Data Rif. (gg/mm/aaaa):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.date_sugg_entry = ttk.Entry(input_frame, width=15)
        self.date_sugg_entry.insert(0, datetime.now().strftime('%d/%m/%Y'))
        self.date_sugg_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        ttk.Button(input_frame, text="Cerca Alternative", command=self._search_suggestions, style="Accent.TButton").grid(row=1, column=2, columnspan=2, padx=5, pady=5)
        self.sugg_results_text = tk.Text(self.suggerimenti_tab, wrap=tk.WORD, state=tk.DISABLED, font=("Consolas", 10))
        self.sugg_results_text.pack(fill='both', expand=True, pady=5)

    def _populate_config_tab(self):
        for widget in self.config_tab.winfo_children(): widget.destroy()
        self.config_entries = {}
        frame = ttk.LabelFrame(self.config_tab, text="Percorsi File", padding=10)
        frame.pack(fill=tk.X, padx=5, pady=5)
        def create_config_row(parent, label_text, config_key, row_index, is_folder=False):
            ttk.Label(parent, text=label_text).grid(row=row_index, column=0, sticky=tk.W, padx=5, pady=5)
            entry = ttk.Entry(parent, width=100)
            entry.grid(row=row_index, column=1, sticky=tk.EW, padx=5)
            current_value = getattr(config, config_key, "") or ""
            if current_value: entry.insert(0, current_value)
            self.config_entries[config_key] = entry
            browse_cmd = partial(self._browse_folder, entry) if is_folder else partial(self._browse_file, entry)
            ttk.Button(parent, text="Sfoglia...", command=browse_cmd).grid(row=row_index, column=2, padx=5)
        create_config_row(frame, "File Registro Strumenti:", 'FILE_REGISTRO_STRUMENTI', 0)
        create_config_row(frame, "Cartella Schede da Analizzare:", 'FOLDER_PATH_DEFAULT', 1, is_folder=True)
        create_config_row(frame, "File Dati Compilazione:", 'FILE_DATI_COMPILAZIONE_SCHEDE', 2)
        create_config_row(frame, "File Master Digitale (.xlsx):", 'FILE_MASTER_DIGITALE_XLSX', 3)
        create_config_row(frame, "File Master Analogico (.xlsx):", 'FILE_MASTER_ANALOGICO_XLSX', 4)
        frame.columnconfigure(1, weight=1)
        save_button = ttk.Button(self.config_tab, text="Salva Configurazione", command=self._save_config, style="Accent.TButton")
        save_button.pack(pady=10)

    def _update_cert_details_map(self):
        self.cert_details_map.clear()
        for usage in self.all_cert_usages:
            details = self.cert_details_map[usage.certificate_id]
            if not details.get('id'):
                details.update({'id': usage.certificate_id, 'utilizzi': 0, 'dettaglio_usi_list': [], 'date_utilizzo_obj_set': set(), 'range_su_scheda_counter': Counter(), 'tipologie_scheda_associate_counter': Counter(), 'usi_congrui': 0, 'usi_total_incongrui': 0, 'usi_prima_emissione': 0, 'usi_scaduti_puri': 0})
            details['utilizzi'] += 1
            details['dettaglio_usi_list'].append(usage)
            if usage.card_date: details['date_utilizzo_obj_set'].add(usage.card_date)
            details['range_su_scheda_counter'][usage.instrument_range_on_card] += 1
            details['tipologie_scheda_associate_counter'][usage.tipologia_strumento_scheda] += 1
            if usage.is_congruent: details['usi_congrui'] += 1
            elif usage.is_congruent is False: details['usi_total_incongrui'] += 1
            if usage.used_before_emission: details['usi_prima_emissione'] += 1
            elif usage.is_expired_at_use: details['usi_scaduti_puri'] += 1

    def _prepare_data_for_treeview(self) -> List[Dict]:
        tree_data = []
        for cert_id, details in self.cert_details_map.items():
            scad_rec = max(details['date_utilizzo_obj_set']).strftime('%d/%m/%Y') if details['date_utilizzo_obj_set'] else "N/D"
            range_p = details['range_su_scheda_counter'].most_common(1)[0][0] if details['range_su_scheda_counter'] else "N/D"
            tip_p = details['tipologie_scheda_associate_counter'].most_common(1)[0][0] if details['tipologie_scheda_associate_counter'] else "N/D"
            tree_data.append({
                "ID Certificato": cert_id, "Utilizzi": details['utilizzi'], "Tipologia Principale": tip_p,
                "Congrui": details['usi_congrui'], "Non Congrui": details['usi_total_incongrui'],
                "Prima Emiss.": details['usi_prima_emissione'], "Scaduti": details['usi_scaduti_puri'],
                "Scadenza Recente": scad_rec, "Range Principale": range_p
            })
        return sorted(tree_data, key=lambda x: (-x["Prima Emiss."], -x["Non Congrui"], -x["Utilizzi"]))

    def _on_tree_item_single_click(self, event):
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
        item_id = self.tree_cert.identify_row(event.y)
        if not item_id: return
        if self.tree_cert.parent(item_id):
            self._on_file_click(item_id, os.path.basename(item_id), open_file_direct=True)
        else:
            values = self.tree_cert.item(item_id, 'values')
            cert_id, range_val = values[0], values[8]
            self.notebook.select(self.suggerimenti_tab)
            self.cert_id_sugg_entry.delete(0, tk.END); self.cert_id_sugg_entry.insert(0, cert_id)
            self.range_sugg_entry.delete(0, tk.END); self.range_sugg_entry.insert(0, range_val)
            self._search_suggestions()

    def _search_suggestions(self):
        cert_id_target = self.cert_id_sugg_entry.get().strip()
        range_req = self.range_sugg_entry.get().strip()
        date_ref_str = self.date_sugg_entry.get().strip()
        date_ref = excel_io.parse_date_robust(date_ref_str)
        if not date_ref:
            messagebox.showerror("Errore Data", "Formato data non valido. Usare gg/mm/aaaa.", parent=self.root)
            return
        results = analysis.trova_strumenti_alternativi(range_req, date_ref, self.strumenti_campione)
        self.sugg_results_text.config(state=tk.NORMAL)
        self.sugg_results_text.delete("1.0", tk.END)
        if not results:
            self.sugg_results_text.insert(tk.END, "Nessuna alternativa valida trovata.")
        else:
            count = 0
            for res in results:
                if res.id_certificato == cert_id_target: continue
                count += 1
                scad_str = res.scadenza.strftime('%d/%m/%Y') if res.scadenza else 'N/D'
                self.sugg_results_text.insert(tk.END, f"ID: {res.id_certificato}, Modello: {res.modello_strumento}, Range: {res.range}, Scadenza: {scad_str}\n")
            if count == 0: self.sugg_results_text.insert(tk.END, "Nessuna alternativa valida trovata (escludendo il certificato di partenza).")
        self.sugg_results_text.config(state=tk.DISABLED)

    def _generate_report_word(self):
        logger.info("Preparazione dati per il report Word...")
        temporal_list, incongruent_list = [], []
        for usage in self.all_cert_usages:
            item = usage.__dict__.copy()
            item['card_date_str'] = usage.card_date.strftime('%d/%m/%Y') if usage.card_date else 'N/D'
            if usage.used_before_emission: item['alert_type'] = 'premature_emission'; temporal_list.append(item)
            elif usage.is_expired_at_use: item['alert_type'] = 'expired_at_use'; temporal_list.append(item)
            if usage.is_congruent is False and not usage.used_before_emission: incongruent_list.append(item)
        file_path = reporting.crea_e_apri_report_anomalie_word(self.human_errors_details, temporal_list, incongruent_list, self.candidate_files_count, self.validated_file_count)
        if file_path:
            messagebox.showinfo("Report Generato", f"Report Word generato e aperto:\n{file_path}", parent=self.root)
        else:
            messagebox.showwarning("Report non Generato", "Nessuna anomalia significativa trovata o si è verificato un errore.", parent=self.root)

    def _on_file_click(self, file_path, filename, open_file_direct=False):
        try:
            pyperclip.copy(file_path)
            action_text = "Aprire il file?" if open_file_direct else f"Aprire la cartella del file '{filename}'?"
            if messagebox.askyesno("Percorso Copiato", f"Percorso copiato negli appunti:\n{file_path}\n\n{action_text}", parent=self.root):
                target = file_path if open_file_direct else os.path.dirname(file_path)
                if sys.platform == "win32": os.startfile(target)
                else: subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", target])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il percorso: {e}", parent=self.root)

    def _browse_file(self, entry_widget):
        filepath = filedialog.askopenfilename(title="Seleziona File", filetypes=(("Excel Files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")))
        if filepath:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filepath)

    def _browse_folder(self, entry_widget):
        folderpath = filedialog.askdirectory(title="Seleziona Cartella")
        if folderpath:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folderpath)

    def _save_config(self):
        new_config_data = {key: entry.get() for key, entry in self.config_entries.items()}
        if excel_io.save_configuration(new_config_data):
            messagebox.showinfo("Successo", "Configurazione salvata con successo. Le modifiche saranno applicate alla prossima analisi.", parent=self.root)
        else:
            messagebox.showerror("Errore", "Impossibile salvare la configurazione. Controllare i log per i dettagli.", parent=self.root)

    def _sort_treeview(self, tree, col, reverse):
        pass

    def _on_close(self):
        if messagebox.askokcancel("Chiudi", "Vuoi davvero chiudere l'applicazione?"):
            self.root.destroy()
            logger.info("Applicazione chiusa dall'utente.")
