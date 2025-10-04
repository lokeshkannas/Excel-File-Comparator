\
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REQ = os.path.join(HERE, "requirements.txt")

def run_cmd(cmd):
    print(">", " ".join(cmd), flush=True)
    return subprocess.call(cmd)

def ensure_deps():
    # Use current python to install requirements locally (user mode)
    run_cmd([sys.executable, "-m", "pip", "install", "-r", REQ])

def run_app():
    run_cmd([sys.executable, os.path.join(HERE, "excel_comparator.py")])

if __name__ == "__main__":
    ensure_deps()
    run_app()
