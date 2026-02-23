# 7 — Release Engineering: Offuscamento & Packaging (PyArmor + Inno Setup)

> Prompt universale per implementare un workflow di rilascio professionale su progetti Python.
> Include l'offuscamento del codice sorgente (PyArmor), la creazione di un pacchetto 
> standalone (PyInstaller) e la generazione di un installer Windows (.exe) con Inno Setup.
>
> Compatibile con qualsiasi LLM che abbia accesso al filesystem e al terminale.

---

## Prompt

```
Ruolo: Agisci come un Release Engineer specializzato in software distribution per Python su Windows.

Contesto: Devo proteggere il codice sorgente e creare un installer professionale per un progetto Python.
Il workflow deve includere: offuscamento con PyArmor, compilazione con PyInstaller e creazione setup con Inno Setup.

Obiettivo: Analizza il progetto, configura gli script di build e genera l'installer finale.

=============================================================================
FASE 0 — ANALISI DEL PROGETTO
=============================================================================

1. IDENTIFICA i componenti chiave:
   - Entry point: (es. main.py, app.py)
   - Cartella sorgente: (es. src/, lib/)
   - Cartella assets: (es. assets/, data/, icons/)
   - Versione corrente: (cerca in version.py o pyproject.toml)

2. VERIFICA i prerequisiti:
   - Python 3.10+ installato
   - Inno Setup 6+ installato nel sistema (C:\Program Files (x86)\Inno Setup 6\ISCC.exe)
   - Virtual environment attivo

3. IDENTIFICA le dipendenze critiche (hidden imports):
   - Cerca librerie che usano dynamic import (es. win32com, logging.handlers, sqlalchemy)

=============================================================================
FASE 1 — INSTALLAZIONE TOOL DI RELEASE
=============================================================================

Installa i tool necessari nel virtual environment:

```bash
pip install pyinstaller pyarmor
```

Verifica le versioni:
```bash
pyinstaller --version
pyarmor --version
```

=============================================================================
FASE 2 — PYARMOR: OFFUSCAMENTO (Security First)
=============================================================================

L'offuscamento protegge la proprietà intellettuale rendendo il bytecode illeggibile.

1. ESEGUI la generazione del codice offuscato:
   ```bash
   # Crea una copia offuscata di src/ e main.py nella cartella build/obf/
   pyarmor gen --output build/obf --recursive src/ main.py
   ```

2. REGOLE CRITICHE:
   - Mantieni la struttura delle directory identica all'originale.
   - Verifica che `build/obf/` contenga il file `main.py` e la cartella `src/`.
   - Testa il codice offuscato prima di procedere:
     ```bash
     python build/obf/main.py
     ```

=============================================================================
FASE 3 — PYINSTALLER: BUNDLING (Portable App)
=============================================================================

Trasforma il codice offuscato in una directory standalone con tutte le dipendenze.

1. CONFIGURA il comando PyInstaller puntando al codice offuscato:
   ```bash
   pyinstaller --name "NOME_APP" 
               --onedir 
               --windowed 
               --noconfirm 
               --clean 
               --distpath "dist" 
               --workpath "build" 
               --icon "assets/app.ico" 
               --add-data "build/obf/src;src" 
               --add-data "assets;assets" 
               build/obf/main.py
   ```

2. REGOLE PER GLI IMPORT NASCOSTI:
   - Se l'app crasha all'avvio con "ModuleNotFoundError", aggiungi:
     `--hidden-import modulo_mancante`
   - Se usi librerie pesanti (matplotlib, pandas, selenium), usa:
     `--collect-all nome_libreria`

=============================================================================
FASE 4 — INNO SETUP: INSTALLER (.EXE)
=============================================================================

Crea un file `.iss` (Inno Setup Script) per generare l'installer finale.

1. CREA `installer_config.iss` nella root (o admin/):
```iss
[Setup]
AppId={{GENERA-UN-GUID-UNICO}}
AppName=NOME_APP
AppVersion=1.0.0
DefaultDirName={autopf}\NOME_APP
DefaultGroupName=NOME_APP
OutputDir=Setup
OutputBaseFilename=NOME_APP_Setup
SetupIconFile=assets\setup.ico
Compression=lzma2/ultra64
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "italian"; MessagesFile: "compiler:Languages\Italian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; Copia i file generati da PyInstaller
Source: "dist\NOME_APP\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\NOME_APP"; Filename: "{app}\NOME_APP.exe"
Name: "{autodesktop}\NOME_APP"; Filename: "{app}\NOME_APP.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\NOME_APP.exe"; Description: "{cm:LaunchProgram,NOME_APP}"; Flags: nowait postinstall skipifsilent
```

2. COMPILA l'installer:
   ```cmd
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_config.iss
   ```

=============================================================================
FASE 5 — SCRIPT DI AUTOMAZIONE (build_all.py)
=============================================================================

Crea uno script Python che automatizza l'intero processo per evitare errori manuali.

```python
import os
import subprocess
import shutil

def run(cmd):
    print(f"Executing: {cmd}")
    subprocess.run(cmd, shell=True, check=True)

def build():
    # 1. Cleanup
    for folder in ['build', 'dist', 'Setup']:
        if os.path.exists(folder): shutil.rmtree(folder)
    
    # 2. PyArmor
    run("pyarmor gen --output build/obf --recursive src/ main.py")
    
    # 3. PyInstaller
    run("pyinstaller --onedir --windowed --noconfirm --add-data "build/obf/src;src" build/obf/main.py")
    
    # 4. Inno Setup
    iscc = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if os.path.exists(iscc):
        run(f""{iscc}" installer_config.iss")

if __name__ == "__main__":
    build()
```

=============================================================================
FASE 6 — VERIFICA FINALE
=============================================================================

1. SMOKE TEST:
   - Installa l'applicazione tramite il setup generato.
   - Avvia l'app dal menu Start.
   - Verifica che tutti gli assets (icone, stili) siano caricati correttamente.
   - Controlla i log per eventuali errori di import.

2. VERIFICA OFFUSCAMENTO:
   - Vai nella cartella di installazione.
   - Prova ad aprire un file `.py` nella cartella `src/`.
   - Deve apparire offuscato (codice non leggibile, header PyArmor).

=============================================================================
REPORT FINALE
=============================================================================

Produci un report con questo formato:

```
╔══════════════════════════════════════════════════════════════╗
║             RELEASE REPORT — PACKAGING COMPLETED            ║
╠══════════════════════════════════════════════════════════════╣
║ App Name:           <nome>                                  ║
║ Version:            <versione>                              ║
║ Obfuscation:        ✅ PyArmor (recursive)                  ║
║ Bundling:           ✅ PyInstaller (onedir)                 ║
║ Installer:          ✅ Inno Setup (.exe)                    ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║ ARTEFATTI GENERATI:                                          ║
║  - build/obf/        (Codice offuscato)                     ║
║  - dist/<nome>/      (Portable folder)                      ║
║  - Setup/<nome>.exe  (Installer finale)                     ║
║                                                              ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║ DIMENSIONI:                                                  ║
║  Portable Folder:    N MB                                    ║
║  Installer EXE:      N MB                                    ║
║                                                              ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║ NOTE TECNICHE:                                               ║
║  - Hidden imports aggiunti: [X, Y, Z]                        ║
║  - Assets inclusi: [cartella_assets, icone]                  ║
║  - Inno Setup GUID: {XXXXXXXX-XXXX-...}                     ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
```

=============================================================================
PRINCIPI GUIDA
=============================================================================

1. ISOLAMENTO: Il bundling deve includere TUTTE le dipendenze. L'utente
   non deve avere Python installato per far girare l'app.

2. PROTEZIONE: L'offuscamento è obbligatorio per software commerciale/IP.
   Verifica sempre che il codice in `dist/` sia effettivamente offuscato.

3. ATOMICITÀ: Lo script di build deve fallire se uno step fallisce.
   Non generare un installer corrotto se PyInstaller ha dato errori.

4. USER EXPERIENCE: L'installer deve essere pulito, con icone corrette
   e collegamenti al menu Start/Desktop.

5. MANUTENIBILITÀ: Usa variabili nel file `.iss` per la versione, in modo
   da aggiornarla in un unico punto o passarla via riga di comando.
```
