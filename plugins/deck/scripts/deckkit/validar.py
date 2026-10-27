"""Validação de um deck (JSON) antes de gerar: erros que travam e avisos que o agente deve ler.

Erros: tipo desconhecido, campo faltando, texto acima do limite, termo vedado pela marca, travessão (se a marca proíbe).
Avisos: três slides seguidos do mesmo tipo, slide denso, deck só de texto, número sem fonte (se a marca exige).
"""
from __future__ import annotations

import re
import tempfile
from pathlib import Path

from .deck import Deck
from .marca import Marca
from .tipos import TIPOS, ErroDeSlide

TIPOS_SO_TEXTO = {"texto", "lista"}
TRAVESSOES = ("—", "–")
LIMITE_PALAVRAS = {"tabela": 400, "implantacao": 140, "recomendacao": 130, "cronograma": 110, "passos": 110, "texto": 110}


def _textos(v, saida: list) -> None:
    if isinstance(v, str):
        saida.append(v)
    elif isinstance(v, dict):
        for k, x in v.items():
            if k not in ("tipo", "foto", "estilo", "formato", "tom", "chave", "estado", "cor"):
                _textos(x, saida)
    elif isinstance(v, list):
        for x in v:
            _textos(x, saida)


def _palavras(s: dict) -> int:
    """Palavras visíveis no slide (as notas do apresentador não contam)."""
    t: list = []
    _textos({k: v for k, v in s.items() if k not in ("notas",)}, t)
    return sum(len(x.split()) for x in t)


def _tem_numero(s: dict) -> bool:
    t: list = []
    _textos({k: v for k, v in s.items() if k not in ("notas", "titulo", "eyebrow")}, t)
    return any(re.search(r"\d{2,}|\d+\s?%|R\$", x) for x in t)


def validar(deck: dict, marca: Marca, modelo: Path = None) -> dict:
    """Devolve {'erros': [...], 'avisos': [...], 'sugestoes': [...]}; cada item é um texto com o número do slide."""
    erros: list = []
    avisos: list = []
    sugestoes: list = []
    slides = deck.get("slides")
    if not isinstance(slides, list) or not slides:
        return {"erros": ["o deck precisa de uma lista 'slides' com pelo menos um slide"], "avisos": [], "sugestoes": []}

    # 1) cada slide gerado de verdade num deck temporário: pega erro de campo, limite e layout
    try:
        d = Deck(marca, projeto="", modelo=modelo)
    except ErroDeSlide as e:
        return {"erros": [str(e)], "avisos": [], "sugestoes": []}
    for i, s in enumerate(slides, 1):
        try:
            d.adicionar(s)
        except ErroDeSlide as e:
            erros.append(f"slide {i}: {e}")
        except (KeyError, TypeError, IndexError, AttributeError) as e:
            tipo = s.get("tipo") if isinstance(s, dict) else "?"
            erros.append(f"slide {i} ({tipo}): campo ausente ou no formato errado ({type(e).__name__}: {e}). "
                         f"Veja `deckkit tipos {tipo}`")

    # 2) regras da marca sobre o texto
    for i, s in enumerate(slides, 1):
        t: list = []
        _textos(s, t)
        corpo = "\n".join(t)
        for termo in marca.vedados:
            if re.search(rf"\b{re.escape(termo)}\b", corpo, re.IGNORECASE):
                erros.append(f"slide {i}: termo vedado pela marca: '{termo}'")
        if marca.sem_travessao and any(c in corpo for c in TRAVESSOES):
            erros.append(f"slide {i}: a marca não aceita travessão nem meia-risca; troque por vírgula, dois-pontos ou 'a'")

    # 3) avisos de qualidade (não travam)
    tipos = [s.get("tipo") for s in slides]
    for i in range(2, len(tipos)):
        if tipos[i] == tipos[i - 1] == tipos[i - 2]:
            avisos.append(f"slides {i - 1} a {i + 1}: três do mesmo tipo ('{tipos[i]}') em sequência; varie o formato")
    if len(slides) >= 4 and sum(t in TIPOS_SO_TEXTO for t in tipos) / len(tipos) > 0.4:
        avisos.append("mais de 40% dos slides são só texto (texto ou lista); prefira número, fluxo, tabela ou cartões")
    for i, s in enumerate(slides, 1):
        p = _palavras(s)
        teto = LIMITE_PALAVRAS.get(s.get("tipo"), 90)
        if p > teto:
            avisos.append(f"slide {i}: {p} palavras (o tipo '{s.get('tipo')}' costuma ficar abaixo de {teto}); "
                          f"mova o detalhe para `notas` (campo do apresentador)")
        if marca.exigir_fonte and _tem_numero(s) and not (s.get("fonte") or "fonte" in str(s.get("nota", "")).lower()):
            avisos.append(f"slide {i}: tem número mas não tem `fonte`; diga de onde vem e o período")
    if tipos[0] != "capa":
        sugestoes.append("o primeiro slide costuma ser `capa`")
    tipo_deck = str(deck.get("tipo_deck", "")).lower()
    if tipo_deck in ("diagnostico", "diagnóstico", "business-case", "kickoff", "kick-off") and tipos[-1] not in ("decisao", "cronograma"):
        sugestoes.append("este tipo de deck costuma terminar em `decisao` (ou `cronograma`, no kick-off)")
    if not any(s.get("notas") for s in slides) and len(slides) > 6:
        sugestoes.append("nenhum slide tem `notas`; o que não coube na tela pode ir para as notas do apresentador")
    return {"erros": erros, "avisos": avisos, "sugestoes": sugestoes}
