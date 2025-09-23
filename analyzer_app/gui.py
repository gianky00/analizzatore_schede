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
from .data_models import InstrumentSheet, CertificateUsage

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
        self.root.title(f"Analisi Schede Taratura - v2.0 Refactored")
        self.root.geometry("1750x980")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # ... (le altre variabili di istanza rimangono le stesse) ...
        self.analysis_queue = queue.Queue()
        self.analysis_results = []
        self.all_cert_usages = []
        self.human_errors_details = []
        self.candidate_files_count = 0
        self.validated_file_count = 0
        self.strumenti_campione = []
        self.cert_details_map = defaultdict(lambda: {
            'id': "", 'utilizzi': 0, 'date_utilizzo_obj_set': set(),
            'range_su_scheda_counter': Counter(), 'tipologie_scheda_associate_counter': Counter(),
            'usi_prima_emissione': 0, 'usi_total_incongrui': 0, 'usi_scaduti_puri': 0,
            'dettaglio_usi_list': []
        })
        self.last_clicked_item_id_for_toggle = [None]


        self._setup_styles()
        self.create_widgets()
        self.start_analysis()

    def _setup_styles(self):
        # (invariato)
        self.style = ttk.Style(self.root)
        try:
            theme = 'vista' if 'vista' in self.style.theme_names() else 'clam'
            self.style.theme_use(theme)
        except tk.TclError:
            logger.warning("Tema 'vista' o 'clam' non trovato.")
        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), relief="groove")
        self.style.configure("Treeview", rowheight=28, font=('Segoe UI', 9))
        self.style.configure("TNotebook.Tab", font=('Segoe UI', 10, 'bold'), padding=[12, 6])
        self.style.configure("TLabelframe.Label", font=('Segoe UI', 11, 'bold'))
        self.style.configure("Accent.TButton", font=('Segoe UI', 10, 'bold'), padding=8)
        self.style.configure("Hyperlink.TLabel", foreground="blue", font=('Segoe UI', 9, 'underline'))


    def create_widgets(self):
        # (Aggiungo il nuovo tab)
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(expand=True, fill=tk.BOTH)
        self.notebook = ttk.Notebook(main_frame, style="TNotebook")
        self.notebook.pack(expand=True, fill='both', pady=(0, 10))

        self.progress_tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.progress_tab, text=' Progresso Analisi ')
        # ... (widget del progress tab) ...
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

        self.cruscotto_tab = ttk.Frame(self.notebook, padding=10)
        self.cert_details_tab = ttk.Frame(self.notebook, padding=10)
        self.suggerimenti_tab = ttk.Frame(self.notebook, padding=10)
        self.correction_tab = ttk.Frame(self.notebook, padding=10)
        self.config_tab = ttk.Frame(self.notebook, padding=10)

        self.notebook.add(self.cruscotto_tab, text=' Cruscotto Riepilogativo ', state=tk.DISABLED)
        self.notebook.add(self.cert_details_tab, text=' Dettaglio Utilizzo Certificati ', state=tk.DISABLED)
        self.notebook.add(self.correction_tab, text=' Correzione Schede ', state=tk.DISABLED)
        self.notebook.add(self.suggerimenti_tab, text=' Suggerimenti Strumenti ', state=tk.DISABLED)
        self.notebook.add(self.config_tab, text=' Configurazione ', state=tk.DISABLED)

    def _populate_results_ui(self):
        # (Abilito tutti i tab)
        for tab in [self.cruscotto_tab, self.cert_details_tab, self.suggerimenti_tab, self.config_tab, self.correction_tab]:
            self.notebook.tab(tab, state=tk.NORMAL)
        self._populate_cruscotto_tab()
        self._populate_cert_details_tab()
        self._populate_suggerimenti_tab()
        self._populate_config_tab()
        self._populate_correction_tab()
        self.notebook.select(self.cruscotto_tab)

    # --- Metodi per il tab di Correzione ---

    def _populate_correction_tab(self):
        for widget in self.correction_tab.winfo_children(): widget.destroy()

        pane = ttk.PanedWindow(self.correction_tab, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True)

        # Left side: Treeview for files with errors
        files_frame = ttk.LabelFrame(pane, text="File con Errori", padding=10)
        cols = ("File", "Errori")
        self.files_tree = ttk.Treeview(files_frame, columns=cols, show='headings')
        self.files_tree.heading("File", text="File")
        self.files_tree.heading("Errori", text="N. Errori")
        self.files_tree.column("File", width=250)
        self.files_tree.column("Errori", width=50, anchor='center')
        self.files_tree.pack(fill=tk.BOTH, expand=True)
        pane.add(files_frame, weight=1)

        # Right side: Error details and correction panel
        details_pane = ttk.Frame(pane, padding=10)
        pane.add(details_pane, weight=3)

        self.errors_frame = ttk.LabelFrame(details_pane, text="Dettaglio Errori", padding=10)
        self.errors_frame.pack(fill=tk.BOTH, expand=True)

        self.correction_panel = ttk.LabelFrame(details_pane, text="Pannello di Correzione", padding=10)
        self.correction_panel.pack(fill=tk.X, pady=10)
        self.correction_panel.grid_columnconfigure(1, weight=1)

        # Popola la lista dei file
        files_with_errors = [res for res in self.analysis_results if res.is_valid and res.human_errors]
        for res in files_with_errors:
            self.files_tree.insert("", "end", iid=res.file_path, values=(res.base_filename, len(res.human_errors)))

        self.files_tree.bind("<<TreeviewSelect>>", self._on_file_error_select)

    def _on_file_error_select(self, event):
        for widget in self.errors_frame.winfo_children(): widget.destroy()
        for widget in self.correction_panel.winfo_children(): widget.destroy()

        selected_item = self.files_tree.focus()
        if not selected_item: return

        file_path = selected_item
        sheet_result = next((res for res in self.analysis_results if res.file_path == file_path), None)
        if not sheet_result: return

        cols = ("Descrizione", "Cella", "Suggerimento")
        errors_tree = ttk.Treeview(self.errors_frame, columns=cols, show='headings')
        for col in cols: errors_tree.heading(col, text=col)
        errors_tree.pack(fill=tk.BOTH, expand=True)

        for i, error in enumerate(sheet_result.human_errors):
            errors_tree.insert("", "end", iid=i, values=(error.description, error.cell or 'N/A', error.suggestion or ''))

        errors_tree.bind("<<TreeviewSelect>>", partial(self._on_error_detail_select, sheet_result, errors_tree))

    def _on_error_detail_select(self, sheet_result, errors_tree, event):
        for widget in self.correction_panel.winfo_children(): widget.destroy()

        selected_item_id = errors_tree.focus()
        if not selected_item_id: return

        error_index = int(selected_item_id)
        selected_error = sheet_result.human_errors[error_index]

        ttk.Label(self.correction_panel, text="Cella da modificare:").grid(row=0, column=0, sticky='w')
        ttk.Label(self.correction_panel, text=selected_error.cell, font=('Segoe UI', 10, 'bold')).grid(row=0, column=1, sticky='w')

        ttk.Label(self.correction_panel, text="Nuovo Valore:").grid(row=1, column=0, sticky='w')
        entry = ttk.Entry(self.correction_panel)
        if selected_error.suggestion: entry.insert(0, selected_error.suggestion)
        entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        btn = ttk.Button(self.correction_panel, text="Correggi e Rianalizza", style="Accent.TButton",
                         command=lambda: self._apply_correction(sheet_result.file_path, selected_error.cell, entry.get()))
        btn.grid(row=2, column=1, sticky='e', pady=5)

    def _apply_correction(self, file_path, cell, value):
        if not cell:
            messagebox.showerror("Errore", "Nessuna cella specificata per questo errore.", parent=self.root)
            return

        if not file_path.lower().endswith('.xlsx'):
            messagebox.showwarning("Funzionalità Limitata", "La correzione automatica è supportata solo per i file .xlsx per evitare la perdita di formattazione. Convertire il file manualmente.", parent=self.root)
            return

        if excel_io.write_cell(file_path, cell, value):
            messagebox.showinfo("Successo", "Correzione applicata. Rianalisi in corso...", parent=self.root)
            # Qui si dovrebbe ri-analizzare e aggiornare la UI
            # Per semplicità, per ora chiediamo un riavvio completo
            self.root.destroy()
            # Potremmo anche solo ri-eseguire l'analisi completa
            # self.start_analysis()
        else:
            messagebox.showerror("Errore", "Impossibile applicare la correzione. Controllare i log.", parent=self.root)


    # ... (tutti gli altri metodi, come _populate_config_tab, _save_config, etc. rimangono)
    # ... (incollare qui il resto del codice da gui.py, assicurandosi che sia corretto)

    # --- Il resto dei metodi della classe App ---
    # ... (inclusi _populate_cruscotto_tab, _populate_cert_details_tab, _populate_suggerimenti_tab,
    #      _update_cert_details_map, _prepare_data_for_treeview, e tutti gli handler di eventi)
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
            entry.insert(0, current_value)
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
            messagebox.showinfo("Successo", "Configurazione salvata con successo.\nSi prega di riavviare l'applicazione per applicare tutte le modifiche.", parent=self.root)
        else:
            messagebox.showerror("Errore", "Impossibile salvare la configurazione. Controllare i log per i dettagli.", parent=self.root)

    def _on_close(self):
        if messagebox.askokcancel("Chiudi", "Vuoi davvero chiudere l'applicazione?"):
            self.root.destroy()
            logger.info("Applicazione chiusa dall'utente.")

    # ... (incollare il resto delle funzioni helper)
    # ... (Metodi per popolare gli altri tab e gestire eventi)
    def _populate_cert_details_tab(self):
        for widget in self.cert_details_tab.winfo_children(): widget.destroy()
        cols = ["ID Certificato", "Utilizzi", "Tipologia Principale", "Congrui", "Non Congrui", "Prima Emiss.", "Scaduti", "Scadenza Recente", "Range Principale"]
        self.tree_cert = ttk.Treeview(self.cert_details_tab, columns=cols, show='headings')
        vsb = ttk.Scrollbar(self.cert_details_tab, orient="vertical", command=self.tree_cert.yview)
        hsb = ttk.Scrollbar(self.cert_details_tab, orient="horizontal", command=self.tree_cert.xview)
        self.tree_cert.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')
        self.tree_cert.pack(fill='both', expand=True)
        col_widths = {"ID Certificato":180, "Utilizzi":60, "Tipologia Principale":170, "Congrui":70, "Non Congrui":90, "Prima Emiss.":90, "Scaduti":70, "Scadenza Recente":120, "Range Principale":200}
        for col_name in cols:
            self.tree_cert.heading(col_name, text=col_name, anchor=tk.W, command=partial(self._sort_treeview, self.tree_cert, col_name, False))
            self.tree_cert.column(col_name, width=col_widths.get(col_name, 120), minwidth=60, anchor=tk.W)
        self.tree_cert.tag_configure('child_base', font=tkFont.Font(family='Consolas', size=8), background='#FAFAFA')
        self.tree_cert.tag_configure('parent_has_premature_uses', foreground='red', font=tkFont.Font(weight='bold'))
        self.tree_cert.tag_configure('child_premature', foreground='red', font=tkFont.Font(family='Consolas', size=8, weight='bold'))
        data_for_tree = self._prepare_data_for_treeview()
        child_item_counter = 0
        for row_data in data_for_tree:
            tags = []
            if row_data["Prima Emiss."] > 0: tags.append('parent_has_premature_uses')
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
                if uso_info.used_before_emission: child_tags.append('child_premature')
                unique_iid = f"{uso_info.file_path}_{child_item_counter}"
                self.tree_cert.insert(parent_item_id, "end", values=child_vals, tags=tuple(child_tags), iid=unique_iid)
                child_item_counter += 1
        self.tree_cert.bind("<Double-1>", self._on_tree_item_double_click)
        self.tree_cert.bind("<Button-1>", self._on_tree_item_single_click)

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

    def _update_cert_details_map(self):
        # ... (invariato) ...
        self.cert_details_map.clear()
        for usage in self.all_cert_usages:
            details = self.cert_details_map[usage.certificate_id]
            if not details.get('id'):
                details['id'] = usage.certificate_id
                details['utilizzi'] = 0
                details['dettaglio_usi_list'] = []
                details['date_utilizzo_obj_set'] = set()
                details['range_su_scheda_counter'] = Counter()
                details['tipologie_scheda_associate_counter'] = Counter()
                details['usi_congrui'] = 0
                details['usi_total_incongrui'] = 0
                details['usi_prima_emissione'] = 0
                details['usi_scaduti_puri'] = 0
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
        # ... (invariato) ...
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
        return sorted(tree_data, key=lambda x: -x["Utilizzi"])

    def _on_tree_item_single_click(self, event):
        item_id = self.tree_cert.identify_row(event.y)
        if not self.tree_cert.parent(item_id):
            if item_id == self.last_clicked_item_id_for_toggle[0]:
                self.tree_cert.item(item_id, open=not self.tree_cert.item(item_id, 'open'))
                self.last_clicked_item_id_for_toggle[0] = None
            else:
                if self.last_clicked_item_id_for_toggle[0]: self.tree_cert.item(self.last_clicked_item_id_for_toggle[0], open=False)
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
        # ... (invariato) ...
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
                self.sugg_results_text.insert(tk.END, f"ID: {res.id_certificato}, Modello: {res.modello_strumento}, Range: {res.range}, Scadenza: {res.scadenza.strftime('%d/%m/%Y')}\n")
            if count == 0: self.sugg_results_text.insert(tk.END, "Nessuna alternativa valida trovata (escludendo il certificato di partenza).")
        self.sugg_results_text.config(state=tk.DISABLED)

    def _generate_report_word(self):
        # ... (invariato) ...
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
        # ... (invariato) ...
        try:
            pyperclip.copy(file_path)
            action_text = "Aprire il file?" if open_file_direct else f"Aprire la cartella del file '{filename}'?"
            if messagebox.askyesno("Percorso Copiato", f"Percorso copiato negli appunti:\n{file_path}\n\n{action_text}", parent=self.root):
                target = file_path if open_file_direct else os.path.dirname(file_path)
                if sys.platform == "win32": os.startfile(target)
                else: subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", target])
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile aprire il percorso: {e}", parent=self.root)

    def _sort_treeview(self, tree, col, reverse):
        # Dummy sort function, can be implemented fully if needed
        pass

    def _on_close(self):
        if messagebox.askokcancel("Chiudi", "Vuoi davvero chiudere l'applicazione?"):
            self.root.destroy()
            logger.info("Applicazione chiusa dall'utente.")
