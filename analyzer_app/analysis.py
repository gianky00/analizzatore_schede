def normalize_sp_code(sp_code_raw) -> str:
    if pd.isna(sp_code_raw):
        return ""
    s_norm = str(sp_code_raw).strip().upper().replace("S.P.", "SP").replace(".", "").replace("-", "/")
    s_norm = " ".join(s_norm.split())
    return re.sub(r'\s*SP\s*(\d+)\s*/\s*(\d+)', r'SP \1/\2', s_norm)

def normalize_um(um_str_raw) -> str:
    if pd.isna(um_str_raw):
        return ""
    s_norm = str(um_str_raw).strip().lower()
    s_norm = " ".join(s_norm.split())
    for k, v in config.MAPPA_NORMALIZZAZIONE_UM.items():
        s_norm = s_norm.replace(k, v)
    return s_norm.replace(" ", "")

def normalize_range_string(range_str_raw) -> str:
    if pd.isna(range_str_raw):
        return ""
    if isinstance(range_str_raw, (int, float)):
        range_str_raw = str(int(range_str_raw)) if range_str_raw == int(range_str_raw) else str(range_str_raw)
    norm_str = str(range_str_raw).lower()
    norm_str = " ".join(norm_str.split())
    norm_str = re.sub(r'\s*[\/÷]\s*', '-', norm_str)
    norm_str = re.sub(r'\s*to\s*', '-', norm_str, flags=re.IGNORECASE)
    norm_str = re.sub(r'\s*-\s*', '-', norm_str)
    return re.sub(r'\s+', '', norm_str)

def is_cell_value_empty(cell_value) -> bool:
    if cell_value is None:
        return True
    if pd.isna(cell_value):
        return True
    if isinstance(cell_value, str) and not cell_value.strip():
        return True
    return str(cell_value).strip().lower() == "nan"

def is_um_pressione_valida(um: str) -> bool:
    return um in config.LISTA_UM_PRESSIONE_RICONOSCIUTE


def _determina_sottotipo_l9(modello_l9_norm: str) -> str | None:
    """Determina il sottotipo L9 normalizzato in base al modello."""
    if not modello_l9_norm:
        return None
    for l9_key, sottotipi_possibili in config.MAPPA_L9_SOTTOTIPO_NORMALIZZATA.items():
        if l9_key in modello_l9_norm or modello_l9_norm == l9_key:
            # Ritorna il primo sottotipo possibile (priorità)
            return sottotipi_possibili[0] if sottotipi_possibili else None
    return None


def _verifica_congruita_certificato(
    tipologia_strumento: str,
    modello_l9_normalizzato: str,
    modello_campione_usato: str
) -> tuple[bool | None, str]:
    """
    Verifica la congruità di un certificato campione rispetto alla tipologia strumento.

    Returns:
        tuple: (is_congruent, notes) dove is_congruent può essere True, False, o None (non verificabile)
    """
    if tipologia_strumento == "N/D" or not modello_campione_usato or modello_campione_usato == "N/D":
        return None, "Tipologia strumento o modello campione non disponibili per verifica"

    modello_campione_upper = modello_campione_usato.strip().upper()
    regole = config.REGOLE_CONGRUITA_CERTIFICATI_NORMALIZZATE.get(tipologia_strumento)

    if not regole:
        return None, f"Nessuna regola definita per tipologia '{tipologia_strumento}'"

    # Determina il sottotipo L9 se applicabile
    sottotipo_l9 = _determina_sottotipo_l9(modello_l9_normalizzato)

    # Controlla prima se è esplicitamente incongruo
    modelli_incongrui = regole.get('modelli_campione_incongrui', [])
    eccezioni_incongrui = regole.get('eccezioni_l9_incongrui', {})

    # Verifica eccezioni per sottotipo L9
    if sottotipo_l9 and sottotipo_l9 in eccezioni_incongrui:
        eccezioni_per_sottotipo = eccezioni_incongrui[sottotipo_l9]
        if any(m in modello_campione_upper or modello_campione_upper in m for m in eccezioni_per_sottotipo):
            return False, f"Modello '{modello_campione_usato}' non congruo per sottotipo '{sottotipo_l9}'"

    # Verifica se è nei modelli incongrui generali
    if any(m in modello_campione_upper or modello_campione_upper in m for m in modelli_incongrui):
        return False, f"Modello '{modello_campione_usato}' non congruo per tipologia '{tipologia_strumento}'"

    # Verifica se è nei modelli congrui per sottotipo
    if sottotipo_l9:
        sottotipi_regole = regole.get('sottotipi_l9', {})
        if sottotipo_l9 in sottotipi_regole:
            modelli_sottotipo = sottotipi_regole[sottotipo_l9]
            if any(m in modello_campione_upper or modello_campione_upper in m for m in modelli_sottotipo):
                return True, f"Congruo per sottotipo '{sottotipo_l9}'"

    # Verifica se è nei modelli congrui generali
    modelli_congrui = regole.get('modelli_campione_congrui', [])
    if any(m in modello_campione_upper or modello_campione_upper in m for m in modelli_congrui):
        return True, f"Congruo per tipologia '{tipologia_strumento}'"

    return None, f"Modello '{modello_campione_usato}' non classificato nelle regole"

def trova_strumenti_alternativi(
    range_richiesto_raw: str,
    data_riferimento_scheda: datetime,
    strumenti_campione_list: list[CalibrationStandard]
) -> list[CalibrationStandard]:
    """
    Trova strumenti alternativi validi per un dato range e data di riferimento.
    Restituisce solo strumenti con certificato valido (non scaduto) alla data di riferimento.
    """
    if not strumenti_campione_list or not data_riferimento_scheda:
        return []

    range_norm = normalize_range_string(range_richiesto_raw)
    risultati = []

    for strumento in strumenti_campione_list:
        # Verifica che il certificato non sia scaduto alla data di riferimento
        if strumento.scadenza and strumento.scadenza >= data_riferimento_scheda:
            # Verifica compatibilità range (match esatto o contiene)
            range_campione_norm = normalize_range_string(strumento.range)
            if range_norm and range_campione_norm and (range_norm == range_campione_norm or range_norm in range_campione_norm):
                # Match esatto o il range del campione copre il range richiesto
                risultati.append(strumento)

    # Ordina per scadenza (i più recenti prima)
    risultati.sort(key=lambda x: x.scadenza if x.scadenza else datetime.min, reverse=True)
    return risultati

def _validate_basic_compilation_fields(
    raw_data: dict,
    file_type: str | None,
    add_error_func: callable
) -> None:
    """Valida i campi obbligatori di compilazione (ODC, PDL, Contratto, etc.)."""
    fields_to_validate = {
        'analogico': {
            'odc': (config.KEY_COMP_ANA_ODC_MANCANTE, config.SCHEDA_ANA_CELL_ODC),
            'card_date': (config.KEY_COMP_ANA_DATA_COMP_MANCANTE, config.SCHEDA_ANA_CELL_DATA_COMPILAZIONE),
            'pdl': (config.KEY_COMP_ANA_PDL_MANCANTE, config.SCHEDA_ANA_CELL_PDL),
            'esecutore': (config.KEY_COMP_ANA_ESECUTORE_MANCANTE, config.SCHEDA_ANA_CELL_ESECUTORE),
            'supervisore': (config.KEY_COMP_ANA_SUPERVISORE_MANCANTE, config.SCHEDA_ANA_CELL_SUPERVISORE_ISAB),
            'contratto': (config.KEY_COMP_ANA_CONTRATTO_MANCANTE, config.SCHEDA_ANA_CELL_CONTRATTO_COEMI),
            'sp_code': (config.KEY_SP_VUOTO, config.SCHEDA_ANA_CELL_TIPOLOGIA_STRUM),
            'modello_l9': (config.KEY_L9_VUOTO, config.SCHEDA_ANA_CELL_MODELLO_STRUM),
            'range_ing': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_RANGE_INGRESSO),
            'um_ing': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_UM_INGRESSO),
            'range_usc': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_RANGE_USCITA),
            'um_usc': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_UM_USCITA),
            'range_dcs': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_RANGE_DCS),
            'um_dcs': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_ANA_CELL_UM_DCS),
        },
        'digitale': {
            'odc': (config.KEY_COMP_DIG_ODC_MANCANTE, config.SCHEDA_DIG_CELL_ODC),
            'card_date': (config.KEY_COMP_DIG_DATA_COMP_MANCANTE, config.SCHEDA_DIG_CELL_DATA_COMPILAZIONE),
            'pdl': (config.KEY_COMP_DIG_PDL_MANCANTE, config.SCHEDA_DIG_CELL_PDL),
            'esecutore': (config.KEY_COMP_DIG_ESECUTORE_MANCANTE, config.SCHEDA_DIG_CELL_ESECUTORE),
            'supervisore': (config.KEY_COMP_DIG_SUPERVISORE_MANCANTE, config.SCHEDA_DIG_CELL_SUPERVISORE_ISAB),
            'contratto': (config.KEY_COMP_DIG_CONTRATTO_MANCANTE, config.SCHEDA_DIG_CELL_CONTRATTO_COEMI),
            'sp_code': (config.KEY_SP_VUOTO, config.SCHEDA_DIG_CELL_TIPOLOGIA_STRUM),
            'range_um_processo': (config.KEY_CELL_RANGE_UM_NON_LEGGIBILE, config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO),
        }
    }

    tipi_da_controllare = ['analogico', 'digitale'] if not file_type else [file_type]
    for tipo in tipi_da_controllare:
        for field_name, (key, cell) in fields_to_validate[tipo].items():
            raw_value = raw_data.get(field_name.lower())
            if raw_value == "#FORMULA_ERROR#":
                add_error_func(config.KEY_FORMULA_ERROR, cell)
            elif is_cell_value_empty(raw_value):
                add_error_func(key, cell)

    # Validazione Contratto
    contratto_val = raw_data.get('contratto')
    if not is_cell_value_empty(contratto_val):
        contratto_str = str(contratto_val).strip()
        if not (contratto_str == config.VALORE_ATTESO_CONTRATTO_COEMI or
                contratto_str == config.VALORE_ATTESO_CONTRATTO_COEMI_VARIANTE_NUMERICA):
            key_diverso = config.KEY_COMP_ANA_CONTRATTO_DIVERSO if file_type == 'analogico' else config.KEY_COMP_DIG_CONTRATTO_DIVERSO
            cell_contratto = config.SCHEDA_ANA_CELL_CONTRATTO_COEMI if file_type == 'analogico' else config.SCHEDA_DIG_CELL_CONTRATTO_COEMI
            add_error_func(key_diverso, cell_contratto, suggestion=config.VALORE_ATTESO_CONTRATTO_COEMI)


def _apply_dynamic_validation_rules(
    raw_data: dict,
    tipologia_strumento: str,
    modello_l9_norm: str,
    add_error_func: callable
) -> None:
    """Applica le regole di validazione dinamiche definite nel config."""
    if not config.VALIDATION_RULES:
        return

    normalized_values = {
        'um_ing': normalize_um(raw_data.get('um_ing')),
        'um_usc': normalize_um(raw_data.get('um_usc')),
        'um_dcs': normalize_um(raw_data.get('um_dcs')),
        'range_ing': normalize_range_string(raw_data.get('range_ing')),
        'range_usc': normalize_range_string(raw_data.get('range_usc')),
        'range_dcs': normalize_range_string(raw_data.get('range_dcs')),
        'modello_l9': modello_l9_norm,
        'range_um_processo': raw_data.get('range_um_processo', "")
    }

    for rule in config.VALIDATION_RULES:
        if (rule['TipologiaStrumento'] != '*' and rule['TipologiaStrumento'] != tipologia_strumento):
            continue
        if (rule['ModelloL9'] != '*' and rule['ModelloL9'] != modello_l9_norm):
            continue

        valore_a = (raw_data.get(rule['CampoA'].lower())
                    if rule['CampoA'] not in normalized_values
                    else normalized_values.get(rule['CampoA']))

        valore_b_raw = rule['CampoB_o_Costante']
        valore_b = (raw_data.get(valore_b_raw)
                    if valore_b_raw in raw_data
                    else (normalized_values.get(valore_b_raw)
                          if valore_b_raw in normalized_values
                          else valore_b_raw))

        triggered = False
        op = rule['Operatore']

        if op == 'is_empty':
            if is_cell_value_empty(valore_a):
                triggered = True
        elif op == 'is_not_empty':
            if not is_cell_value_empty(valore_a):
                triggered = True
        elif op == '==' and not is_cell_value_empty(valore_a):
            if str(valore_a) == str(valore_b):
                triggered = True
        elif op == '!=' and not is_cell_value_empty(valore_a):
            if str(valore_a) != str(valore_b):
                triggered = True
        elif op == 'in' and not is_cell_value_empty(valore_a):
            lista_valori = [v.strip() for v in valore_b.split(',')]
            if str(valore_a) in lista_valori:
                triggered = True
        elif op == 'not_in' and not is_cell_value_empty(valore_a):
            lista_valori = [v.strip() for v in valore_b.split(',')]
            if str(valore_a) not in lista_valori:
                triggered = True

        if triggered:
            add_error_func(rule['ChiaveErrore'])


def _validate_hardware_specific_rules(
    raw_data: dict,
    file_type: str,
    tipologia_strumento: str,
    modello_l9_norm: str,
    human_errors: list[SheetError],
    add_error_func: callable
) -> None:
    """Applica la logica di validazione hardcoded per specifiche tipologie di strumenti."""
    if file_type == "analogico":
        if modello_l9_norm == "SKIN POINT":
            add_error_func(config.KEY_L9_SKINPOINT_INCOMPLETO, cell=config.SCHEDA_ANA_CELL_MODELLO_STRUM)

        range_ing_norm = normalize_range_string(raw_data.get('range_ing'))
        um_ing_norm = normalize_um(raw_data.get('um_ing'))
        range_usc_norm = normalize_range_string(raw_data.get('range_usc'))
        um_usc_norm = normalize_um(raw_data.get('um_usc'))
        range_dcs_norm = normalize_range_string(raw_data.get('range_dcs'))
        um_dcs_norm = normalize_um(raw_data.get('um_dcs'))

        if tipologia_strumento == "TEMPERATURA":
            if modello_l9_norm == "CONVERTITORE":
                # Verifica UM
                if (not any(e.cell == config.SCHEDA_ANA_CELL_UM_INGRESSO or e.cell == config.SCHEDA_ANA_CELL_UM_DCS
                            for e in human_errors) and um_ing_norm != um_dcs_norm):
                    add_error_func(config.KEY_ERR_ANA_TEMP_CONV_C9F9_UM_DIVERSE)

                if (not any(e.cell == config.SCHEDA_ANA_CELL_UM_USCITA for e in human_errors) and
                        um_usc_norm != config.UM_MA_NORMALIZZATA):
                    add_error_func(config.KEY_ERR_ANA_TEMP_CONV_F12_UM_NON_MA)

                # Verifica Range
                if (not any(e.cell == config.SCHEDA_ANA_CELL_RANGE_INGRESSO or e.cell == config.SCHEDA_ANA_CELL_RANGE_DCS
                            for e in human_errors) and range_ing_norm != range_dcs_norm):
                    add_error_func(config.KEY_ERR_ANA_TEMP_CONV_A9D9_RANGE_DIVERSI)

                if (not any(e.cell == config.SCHEDA_ANA_CELL_RANGE_USCITA for e in human_errors) and
                        range_usc_norm != config.RANGE_4_20_NORMALIZZATO):
                    add_error_func(config.KEY_ERR_ANA_TEMP_CONV_D12_RANGE_NON_4_20)

            elif modello_l9_norm != "":
                if not (um_ing_norm == um_dcs_norm and um_dcs_norm == um_usc_norm):
                    add_error_func(config.KEY_ERR_ANA_TEMP_NOCONV_UM_NON_COINCIDENTI)
                if not (range_ing_norm == range_dcs_norm and range_dcs_norm == range_usc_norm):
                    add_error_func(config.KEY_ERR_ANA_TEMP_NOCONV_RANGE_NON_COINCIDENTI)

    elif file_type == "digitale":
        range_um_proc_raw = raw_data.get('range_um_processo', "")
        if not is_cell_value_empty(range_um_proc_raw):
            um_proc_norm = normalize_um(range_um_proc_raw)
            if tipologia_strumento == "PRESSIONE":
                if not is_um_pressione_valida(um_proc_norm):
                    add_error_func(config.KEY_ERR_DIG_PRESS_D22_UM_NON_PRESSIONE, config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)
            elif tipologia_strumento == "LIVELLO" and config.UM_PERCENTO_NORMALIZZATA not in um_proc_norm:
                add_error_func(config.KEY_ERR_DIG_LIVELLO_D22_UM_NON_PERCENTO, config.SCHEDA_DIG_CELL_RANGE_UM_PROCESSO)


def _process_certificates_usages(
    raw_data: dict,
    file_type: str,
    base_filename: str,
    file_path: str,
    card_date: datetime | None,
    tipologia_strumento: str,
    modello_l9_norm: str,
    strumenti_campione_list: list[CalibrationStandard]
) -> list[CertificateUsage]:
    """Processa i certificati elencati nella scheda e ne verifica la validità."""
    extracted_certs_data: list[CertificateUsage] = []

    cert_ids_raw = raw_data.get('cert_ids', [])
    cert_expiries_raw = raw_data.get('cert_expiries', [])
    cert_models_raw = raw_data.get('cert_models', [])
    cert_ranges_raw = raw_data.get('cert_ranges', [])

    # Processa ogni certificato
    for idx in range(len(cert_ids_raw)):
        cert_id_val = cert_ids_raw[idx] if idx < len(cert_ids_raw) else None
        if is_cell_value_empty(cert_id_val):
            continue

        cert_id_str = str(cert_id_val).strip()
        if not cert_id_str:
            continue

        # Estrai dati del certificato
        cert_expiry_raw = cert_expiries_raw[idx] if idx < len(cert_expiries_raw) else None
        cert_model_raw = cert_models_raw[idx] if idx < len(cert_models_raw) else None
        cert_range_raw = cert_ranges_raw[idx] if idx < len(cert_ranges_raw) else None

        cert_expiry_dt = parse_date_robust(cert_expiry_raw, base_filename)
        cert_model_str = str(cert_model_raw).strip().upper() if not is_cell_value_empty(cert_model_raw) else "N/D"
        cert_range_str = str(cert_range_raw).strip() if not is_cell_value_empty(cert_range_raw) else "N/D"

        # Trova lo strumento campione corrispondente nel registro
        strumento_campione_trovato: CalibrationStandard | None = None
        for strum in strumenti_campione_list:
            if strum.id_certificato == cert_id_str:
                strumento_campione_trovato = strum
                break

        modello_campione_usato = (strumento_campione_trovato.modello_strumento
                                  if strumento_campione_trovato else cert_model_str)

        # Verifica scadenza
        is_expired = False
        if card_date and cert_expiry_dt:
            is_expired = card_date > cert_expiry_dt

        # Verifica uso prima dell'emissione
        used_before_emission = False
        data_emissione_presunta = None
        if strumento_campione_trovato and strumento_campione_trovato.data_emissione:
            data_emissione_presunta = strumento_campione_trovato.data_emissione
            if card_date and isinstance(data_emissione_presunta, (datetime, pd.Timestamp)):
                if isinstance(data_emissione_presunta, pd.Timestamp):
                    data_emissione_presunta = data_emissione_presunta.to_pydatetime()
                used_before_emission = card_date < data_emissione_presunta

        # Verifica congruità
        is_congruent, congruency_notes = _verifica_congruita_certificato(
            tipologia_strumento,
            modello_l9_norm,
            modello_campione_usato
        )

        # Crea oggetto CertificateUsage
        cert_usage = CertificateUsage(
            file_name=base_filename,
            file_path=file_path,
            card_type=file_type,
            card_date=card_date,
            certificate_id=cert_id_str,
            certificate_expiry_raw=str(cert_expiry_raw) if cert_expiry_raw else "N/D",
            certificate_expiry=cert_expiry_dt,
            instrument_model_on_card=cert_model_str,
            instrument_range_on_card=cert_range_str,
            is_expired_at_use=is_expired,
            tipologia_strumento_scheda=tipologia_strumento,
            modello_l9_scheda=modello_l9_norm,
            modello_strumento_campione_usato=modello_campione_usato,
            is_congruent=is_congruent,
            congruency_notes=congruency_notes,
            used_before_emission=used_before_emission
        )
        extracted_certs_data.append(cert_usage)

    return extracted_certs_data
