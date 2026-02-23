"""
Script per creare l'eseguibile standalone con PyInstaller.
Eseguire: python build_exe.py
"""
# ruff: noqa: T201
import subprocess
import sys


def main():
    print("=" * 60)
    print("  BUILD ANALIZZATORE SCHEDE TARATURA")
    print("=" * 60)
    print()

    # Installa PyInstaller se necessario
    print("[1/3] Verifica PyInstaller...")
    try:
        import PyInstaller  # noqa: F401
        print("      PyInstaller trovato.")

    except ImportError:
        print("      Installazione PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller", "-q"])
        print("      PyInstaller installato.")

    # Installa dipendenze
    print("[2/3] Verifica dipendenze...")
    deps = ["pandas", "openpyxl", "xlrd", "pyperclip", "python-docx"]
    for dep in deps:
        try:
            __import__(dep.replace("-", "_"))
        except ImportError:
            print(f"      Installazione {dep}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", dep, "-q"])
    print("      Dipendenze OK.")

    # Build con PyInstaller
    print("[3/3] Creazione eseguibile...")
    print()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                    # Singolo file exe
        "--windowed",                   # No console
        "--name", "AnalizzatoreSchede", # Nome exe
        "--clean",                      # Pulizia build precedenti
        "--noconfirm",                  # Non chiedere conferma
        "main.py"
    ]

    result = subprocess.run(cmd)

    if result.returncode == 0:
        print()
        print("=" * 60)
        print("  BUILD PYINSTALLER COMPLETATO!")
        print("=" * 60)
        print()
        
        # Tentativo di build Installer con Inno Setup
        print("[EXTRA] Ricerca Inno Setup Compiler...")
        iscc_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
        if not os.path.exists(iscc_path):
            iscc_path = r"C:\Program Files\Inno Setup 6\ISCC.exe"
            
        if os.path.exists(iscc_path):
            print(f"      Trovato: {iscc_path}")
            print("      Generazione Installer in corso...")
            try:
                iss_result = subprocess.run([iscc_path, "installer_setup.iss"], capture_output=True, text=True)
                if iss_result.returncode == 0:
                    print("      INSTALLER GENERATO CON SUCCESSO!")
                    print("      Percorso: installer_output\AnalizzatoreSchede_Setup_v8.1.exe")
                else:
                    print(f"      ERRORE Inno Setup:\n{iss_result.stderr}")
            except Exception as e:
                print(f"      Errore durante l'esecuzione di Inno Setup: {e}")
        else:
            print("      Inno Setup non trovato (percorso standard).")
            print("      Puoi compilare manualmente 'installer_setup.iss'.")

        print()
        print("  L'eseguibile standalone si trova in: dist/AnalizzatoreSchede.exe")

        print()
        print("  Per distribuire l'applicazione:")
        print("  1. Copia 'dist/AnalizzatoreSchede.exe' dove preferisci")
        print("  2. Al primo avvio, configura i percorsi")
        print("  3. La configurazione viene salvata in 'config.json'")
        print()
    else:
        print()
        print("ERRORE durante la build!")
        print("Controlla i messaggi sopra per dettagli.")

    input("Premi INVIO per chiudere...")

if __name__ == "__main__":
    main()

