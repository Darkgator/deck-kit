"""Ponto de entrada: python deck.py <comando>. Ver `python deck.py --help`."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

if sys.version_info < (3, 8):
    sys.exit("deckkit precisa de Python 3.8 ou mais novo")

try:
    from deckkit.cli import main
except ModuleNotFoundError as e:  # python-pptx / lxml / Pillow
    sys.exit(f"falta uma dependência ({e.name}). Instale com: pip install python-pptx")

if __name__ == "__main__":
    sys.exit(main())
