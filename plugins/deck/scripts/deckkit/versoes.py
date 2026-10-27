"""Versões e proteção de edição manual: cada geração cria `<nome> vN.pptx` e registra a assinatura do arquivo.

Se alguém editar o .pptx à mão, `deckkit verificar` avisa, e a próxima versão parte de uma cópia, nunca por cima.
"""
from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime
from pathlib import Path

REGISTRO = ".deckkit-versoes.json"
NOME_MAX = 120  # nomes longos quebram a abertura no Office quando o caminho passa de ~255 caracteres


def limpar_nome(nome: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r'[<>:"/\\|?*\[\]]+', " ", nome)).strip()


def proximo_destino(pasta: Path, base: str, extensao: str = ".pptx") -> Path:
    pasta.mkdir(parents=True, exist_ok=True)
    base = limpar_nome(base)[:NOME_MAX].rstrip()
    n = 1
    while (pasta / f"{base} v{n}{extensao}").exists():
        n += 1
    return pasta / f"{base} v{n}{extensao}"


def sha(arquivo: Path) -> str:
    return hashlib.sha256(Path(arquivo).read_bytes()).hexdigest()


def _ler(pasta: Path) -> dict:
    arq = pasta / REGISTRO
    if arq.is_file():
        try:
            return json.loads(arq.read_text(encoding="utf-8"))
        except ValueError:
            return {}
    return {}


def registrar(destino: Path, slides: int, origem: str = "") -> None:
    reg = _ler(destino.parent)
    reg[destino.name] = {"sha256": sha(destino), "gerado_em": datetime.now().isoformat(timespec="seconds"),
                         "slides": slides, "origem": origem}
    (destino.parent / REGISTRO).write_text(json.dumps(reg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def verificar(pasta: Path) -> list[dict]:
    """Arquivos gerados que sumiram ou foram editados à mão depois de gerados."""
    saida = []
    for nome, item in sorted(_ler(pasta).items()):
        arq = pasta / nome
        if not arq.exists():
            saida.append({"arquivo": nome, "situacao": "sumiu"})
        elif sha(arq) != item["sha256"]:
            saida.append({"arquivo": nome, "situacao": "EDITADO À MÃO depois de gerado: não regerar por cima; "
                                                       "mudança vira edição pontual, a partir de uma cópia"})
    return saida
