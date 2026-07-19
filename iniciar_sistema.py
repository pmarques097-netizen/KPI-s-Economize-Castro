from __future__ import annotations

import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

HOST = "127.0.0.1"
PORT = 8501
URL = f"http://localhost:{PORT}"
BASE_DIR = Path(__file__).resolve().parent
APP = BASE_DIR / "app.py"


def porta_respondendo(host: str, port: int, timeout: float = 0.5) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def main() -> int:
    if not APP.exists():
        print(f"[ERRO] app.py não encontrado em: {APP}")
        input("Pressione ENTER para fechar...")
        return 1

    comando = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(APP),
        "--server.port",
        str(PORT),
        "--server.address",
        "localhost",
        "--server.headless",
        "true",
        "--browser.gatherUsageStats",
        "false",
    ]

    print("=" * 52)
    print(" REDE ECONOMIZE - KPI COMERCIAL")
    print("=" * 52)
    print(f"Iniciando em {URL}")
    print("Não feche esta janela enquanto estiver usando o sistema.\n")

    processo = subprocess.Popen(comando, cwd=BASE_DIR)

    try:
        abriu = False
        for _ in range(120):
            if processo.poll() is not None:
                print("\n[ERRO] O Streamlit foi encerrado antes de iniciar.")
                input("Pressione ENTER para fechar...")
                return processo.returncode or 1

            if porta_respondendo(HOST, PORT):
                time.sleep(1.0)
                abriu = webbrowser.open_new_tab(URL)
                print(f"\nSistema disponível em: {URL}")
                if not abriu:
                    print("O navegador não abriu automaticamente. Use o endereço acima.")
                break

            time.sleep(0.5)
        else:
            print("\n[AVISO] O sistema demorou mais que o esperado para iniciar.")
            print(f"Tente abrir manualmente: {URL}")

        return processo.wait()

    except KeyboardInterrupt:
        print("\nEncerrando o sistema...")
        processo.terminate()
        try:
            processo.wait(timeout=8)
        except subprocess.TimeoutExpired:
            processo.kill()
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
