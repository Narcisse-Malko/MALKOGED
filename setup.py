import sys
import os
from cx_Freeze import setup, Executable

# Détection automatique du dossier CustomTkinter pour inclure les thèmes
import customtkinter
ctk_path = os.path.dirname(customtkinter.__file__)

# Fichiers à inclure dans le package final
files = [
    "ged_enterprise_config.json",
    "ged_file_index.json",
    (ctk_path, "customtkinter") # Indispensable pour l'interface graphique
]

# On s'assure que les fichiers JSON existent avant la compilation pour éviter une erreur
for json_file in ["ged_enterprise_config.json", "ged_file_index.json"]:
    if not os.path.exists(json_file):
        with open(json_file, "w") as f:
            f.write("{}")

build_exe_options = {
    "packages": ["os", "json", "shutil", "hashlib", "threading", "requests", "re", 
             "pdfplumber", "mutagen", "docx", "openpyxl", "pptx"],
}
# Remplacer l'ancien bloc "base = None..." par celui-ci :
base = None
if sys.platform == "win32":
    # On essaie d'utiliser "gui" (recommandé par l'erreur) ou "Win32GUI"
    base = "gui"
setup(
    name="MALKOGED_AI",
    version="4.0",
    description="Solution GED Immobilière Intégrale",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, icon=None)] # Remplacez None par "logo.ico" si vous en avez un
)
