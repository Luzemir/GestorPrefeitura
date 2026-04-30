import os
import shutil

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# Estrutura de diretórios principal
DIRS = {
    "core": os.path.join(ROOT_DIR, "core"),
    "scripts": os.path.join(ROOT_DIR, "scripts"),
    "testes": os.path.join(ROOT_DIR, "testes"),
    "data": os.path.join(ROOT_DIR, "data"),
}

def create_dirs():
    """Garante que as pastas principais existem."""
    for path in DIRS.values():
        if not os.path.exists(path):
            os.makedirs(path)

def organize_files():
    """Movo arquivos da raiz (bagunçados) para as pastas corretas."""
    for file in os.listdir(ROOT_DIR):
        file_path = os.path.join(ROOT_DIR, file)
        
        # Ignorar pastas
        if not os.path.isfile(file_path):
            continue

        # Arquivos que DEVEM ficar na raiz
        if file in ["organize.py", "requirements.txt", "README.md", "main.py"] or file.startswith("."):
            continue

        # Regras de organização
        if file.endswith(".py"):
            if file.startswith("test_"):
                shutil.move(file_path, os.path.join(DIRS["testes"], file))
                print(f"Movido para testes/: {file}")
            else:
                shutil.move(file_path, os.path.join(DIRS["scripts"], file))
                print(f"Movido para scripts/: {file}")
                
        elif file.endswith((".txt", ".xlsx", ".png", ".html", ".csv", ".json")):
            shutil.move(file_path, os.path.join(DIRS["data"], file))
            print(f"Movido para data/: {file}")

if __name__ == "__main__":
    print("=== Regra Local: Organizando projeto GestorPrefeitura ===")
    create_dirs()
    organize_files()
    print("✅ Projeto organizado com sucesso de acordo com a estrutura!")
