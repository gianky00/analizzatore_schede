"""
Script per creare l'eseguibile standalone con PyInstaller.
Eseguire: python build_exe.py
"""
import subprocess
import sys
import os

def main():
    print("=" * 60)
    print("  BUILD ANALIZZATORE SCHEDE TARATURA")
    print("=" * 60)
    print()
    
    # Installa PyInstaller se necessario
    print("[1/3] Verifica PyInstaller...")
    try:
        import PyInstaller
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
        print("  BUILD COMPLETATO!")
        print("=" * 60)
        print()
        print("  L'eseguibile si trova in: dist/AnalizzatoreSchede.exe")
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

