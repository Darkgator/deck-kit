"""Linha de comando do deckkit. Pensada para um agente de código chamar e ler a saída.

    deck.py doctor                         confere Python, python-pptx, PowerPoint/LibreOffice e as marcas
    deck.py marcas                         lista as marcas (e a padrão)
    deck.py marca-instalar <pasta>         instala um pacote de marca em ~/.deckkit/marcas
    deck.py marca-padrao <nome>            define a marca usada quando nada é dito
    deck.py config [--sem-travessao sim] [--autor NOME]   preferências pessoais (valem sobre a marca)
    deck.py modelo [marca]                 (re)gera modelo.pptx e modelo.potx da marca
    deck.py tipos [tipo]                   tipos de slide, campos e limites
    deck.py guia                           guia de escolha do tipo de slide
    deck.py exemplo [nome]                 exemplos prontos (kickoff, business-case, diagnostico)
    deck.py validar deck.json              erros, avisos e sugestões, sem gerar
    deck.py gerar deck.json [--previa]     valida, gera `<nome> vN.pptx` e (se pedido) a prévia
    deck.py prever arquivo.pptx            um PNG por slide e uma folha de contato
    deck.py revisar arquivo.pptx           revisão estrutural do arquivo gerado
    deck.py verificar [pasta]              arquivos editados à mão depois de gerados
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

from . import __version__, marca as mod_marca
from .marca import ErroDeMarca, RAIZ_PLUGIN
from .tipos import LIMITES, TIPOS

REFERENCIAS = RAIZ_PLUGIN / "skills" / "deck" / "references"


def _saida_utf8() -> None:
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def _ler_deck(caminho: Path) -> dict:
    try:
        return json.loads(Path(caminho).read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        raise SystemExit(f"arquivo não encontrado: {caminho}")
    except ValueError as e:
        raise SystemExit(f"o JSON do deck está inválido: {e}")


def _marca(args, deck: dict = None):
    ref = getattr(args, "marca", None) or (deck or {}).get("marca")
    return mod_marca.localizar(ref)


def _modelo_da(m) -> Path:
    """Modelo da marca; se ainda não existe, gera na hora."""
    from .construtor import construir
    if not m.modelo.is_file():
        try:
            construir(m)
        except OSError:  # pasta somente leitura: gera no cache do usuário
            return construir(m, mod_marca.CASA / "cache" / f"{m.nome}.pptx")
    return m.modelo


def _imprimir(titulo: str, itens: list) -> None:
    if itens:
        print(f"{titulo} ({len(itens)}):")
        for i in itens:
            print(f"  - {i}")


# ---------------------------------------------------------------- comandos
def cmd_doctor(args) -> int:
    import platform
    print(f"deckkit {__version__} | Python {platform.python_version()} | {platform.system()}")
    ok = True
    for modulo, pacote in (("pptx", "python-pptx"), ("lxml", "lxml"), ("PIL", "Pillow")):
        try:
            __import__(modulo)
            print(f"  ok   {pacote}")
        except ImportError:
            ok = False
            print(f"  FALTA {pacote}: pip install {pacote}")
    from .render import _libreoffice
    import os
    ppt = os.name == "nt"
    lo = bool(shutil.which("soffice") or shutil.which("libreoffice")) and bool(shutil.which("pdftoppm"))
    print(f"  {'ok  ' if ppt else '--  '} PowerPoint (prévia, só Windows): {'disponível se o Office estiver instalado' if ppt else 'não se aplica'}")
    print(f"  {'ok  ' if lo else '--  '} LibreOffice + pdftoppm (prévia alternativa): {'encontrados' if lo else 'não encontrados'}")
    if not (ppt or lo):
        print("  AVISO: sem como gerar prévia; o agente não poderá olhar os slides e deve avisar você.")
    marcas = mod_marca.listar()
    print(f"  marcas: {', '.join(n for n, _ in marcas) or 'nenhuma'} | padrão: {mod_marca.padrao_configurado() or 'neutro'}")
    return 0 if ok else 2


def cmd_marcas(args) -> int:
    padrao = mod_marca.padrao_configurado() or "neutro"
    for nome, pasta in mod_marca.listar():
        print(f"{'*' if nome == padrao else ' '} {nome}  ({pasta})")
    print("* = padrão")
    return 0


def cmd_marca_instalar(args) -> int:
    destino = mod_marca.instalar(Path(args.pasta))
    print(f"marca instalada em {destino}")
    m = mod_marca.carregar(destino)
    if not m.modelo.is_file():
        from .construtor import construir
        construir(m)
        print("modelo.pptx gerado")
    print(f"para usar sempre: deck.py marca-padrao {destino.name}")
    return 0


def cmd_marca_padrao(args) -> int:
    mod_marca.localizar(args.nome)  # confere que existe
    mod_marca.definir_padrao(args.nome)
    print(f"marca padrão: {args.nome}")
    return 0


def cmd_config(args) -> int:
    """Preferências pessoais em ~/.deckkit/config.json: valem sobre a marca."""
    mod_marca.CASA.mkdir(parents=True, exist_ok=True)
    cfg = mod_marca._config()
    est = cfg.setdefault("estilo", {})
    if args.sem_travessao is not None:
        est["sem_travessao"] = args.sem_travessao == "sim"
    if args.exigir_fonte is not None:
        est["exigir_fonte"] = args.exigir_fonte == "sim"
    if args.autor is not None:
        cfg["autoria"] = {"autor": args.autor, "cedido_a": args.cedido_a or ""} if args.autor else {}
    (mod_marca.CASA / "config.json").write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(cfg, ensure_ascii=False, indent=2))
    return 0


def cmd_modelo(args) -> int:
    from .construtor import construir
    m = mod_marca.localizar(args.marca)
    destino = construir(m)
    print(f"modelo gerado: {destino} (e {destino.with_suffix('.potx').name})")
    return 0


def cmd_tipos(args) -> int:
    nomes = [args.tipo] if args.tipo else list(TIPOS)
    for n in nomes:
        if n not in TIPOS:
            raise SystemExit(f"tipo desconhecido: {n}. Tipos: {', '.join(TIPOS)}")
        t = TIPOS[n]
        print(f"{n}: {t['uso']}")
        print(f"   obrigatórios: {', '.join(t['obrigatorios'])}")
        print(f"   opcionais:    {', '.join(t['opcionais'])}")
    if not args.tipo:
        print("\nLimites de caracteres por campo: " + ", ".join(f"{k} {v}" for k, v in LIMITES.items()))
        print("Detalhes e exemplos: deck.py guia  |  deck.py exemplo")
    return 0


def cmd_guia(args) -> int:
    arq = REFERENCIAS / "guia-de-escolha.md"
    print(arq.read_text(encoding="utf-8") if arq.is_file() else "guia não encontrado")
    return 0


def cmd_exemplo(args) -> int:
    pasta = REFERENCIAS / "exemplos"
    achados = sorted(pasta.glob("*.json")) if pasta.is_dir() else []
    if not args.nome:
        for p in achados:
            print(p.stem)
        return 0
    for p in achados:
        if p.stem == args.nome:
            print(p.read_text(encoding="utf-8"))
            return 0
    raise SystemExit(f"exemplo não encontrado: {args.nome}")


def _resultado_validacao(r: dict, como_json: bool) -> None:
    if como_json:
        print(json.dumps(r, ensure_ascii=False, indent=2))
        return
    _imprimir("ERROS (travam a geração)", r["erros"])
    _imprimir("AVISOS (leia e decida)", r["avisos"])
    _imprimir("SUGESTÕES", r["sugestoes"])
    if not (r["erros"] or r["avisos"] or r["sugestoes"]):
        print("deck válido, sem avisos")


def cmd_validar(args) -> int:
    from .validar import validar
    deck = _ler_deck(args.deck)
    m = _marca(args, deck)
    r = validar(deck, m, _modelo_da(m))
    _resultado_validacao(r, args.json)
    return 1 if r["erros"] else 0


def cmd_gerar(args) -> int:
    from . import versoes
    from .deck import Deck
    from .tipos import ErroDeSlide
    from .validar import validar
    deck = _ler_deck(args.deck)
    m = _marca(args, deck)
    modelo = _modelo_da(m)
    r = validar(deck, m, modelo)
    if r["erros"]:
        print("Deck NÃO gerado. Corrija os erros e rode de novo:")
        _resultado_validacao(r, False)
        return 1
    d = Deck(m, projeto=deck.get("projeto") or deck.get("nome_projeto", ""), modelo=modelo, autoria=not args.sem_autoria)
    try:
        for s in deck["slides"]:
            d.adicionar(s)
    except ErroDeSlide as e:
        print(f"Deck NÃO gerado: {e}")
        return 1
    pasta = Path(args.saida) if args.saida else Path.cwd()
    base = (deck.get("nome") or "Apresentação")
    projeto = deck.get("projeto") or deck.get("nome_projeto")
    if projeto and not args.sem_projeto_no_nome:
        base = f"[{projeto}] {base}"
    destino = versoes.proximo_destino(pasta, base)
    d.salvar(destino)
    versoes.registrar(destino, len(deck["slides"]), origem=str(Path(args.deck).name))
    print(f"Deck gerado: {destino}")
    print(f"  {len(deck['slides'])} slides, marca '{m.nome}'")
    _imprimir("AVISOS", r["avisos"])
    _imprimir("SUGESTÕES", r["sugestoes"])
    if len(str(destino)) > 240:
        print(f"  AVISO: o caminho tem {len(str(destino))} caracteres; o Office pode não abrir acima de ~255. Use --saida com uma pasta mais curta.")
    if args.previa:
        _previa(destino)
    return 0


def _previa(pptx: Path) -> None:
    from .render import folha, prever
    imgs = prever(pptx)
    if not imgs:
        print("  PRÉVIA: não foi possível gerar (sem PowerPoint nem LibreOffice). Avise o usuário que os slides não foram vistos.")
        return
    f = folha(imgs, imgs[0].parent / "folha.png")
    print(f"  prévia: {imgs[0].parent} ({len(imgs)} imagens, slide-01.png ...)")
    print(f"  folha de contato: {f}")


def cmd_prever(args) -> int:
    _previa(Path(args.arquivo))
    return 0


def cmd_revisar(args) -> int:
    from .revisar import revisar
    r = revisar(Path(args.arquivo))
    if args.json:
        print(json.dumps(r, ensure_ascii=False, indent=2))
    else:
        print(f"{r['slides']} slides")
        _imprimir("ERROS", r["erros"])
        _imprimir("AVISOS", r["avisos"])
        if not (r["erros"] or r["avisos"]):
            print("revisão estrutural sem problemas (falta olhar os slides renderizados)")
        if r["slides_sem_notas"]:
            print(f"slides sem notas do apresentador: {r['slides_sem_notas']}")
    return 1 if r["erros"] else 0


def cmd_verificar(args) -> int:
    from .versoes import verificar
    avisos = verificar(Path(args.pasta))
    for a in avisos:
        print(f"{a['arquivo']}: {a['situacao']}")
    print(f"{len(avisos)} aviso(s)." if avisos else "Nenhum arquivo editado à mão.")
    return 0


# ---------------------------------------------------------------- entrada
def main(argv=None) -> int:
    _saida_utf8()
    ap = argparse.ArgumentParser(prog="deck.py", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--versao", action="version", version=__version__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    sub.add_parser("doctor").set_defaults(f=cmd_doctor)
    sub.add_parser("marcas").set_defaults(f=cmd_marcas)
    p = sub.add_parser("marca-instalar")
    p.add_argument("pasta")
    p.set_defaults(f=cmd_marca_instalar)
    p = sub.add_parser("marca-padrao")
    p.add_argument("nome")
    p.set_defaults(f=cmd_marca_padrao)
    p = sub.add_parser("config")
    p.add_argument("--sem-travessao", choices=["sim", "nao"], help="recusa travessão e meia-risca no texto")
    p.add_argument("--exigir-fonte", choices=["sim", "nao"], help="avisa número sem fonte")
    p.add_argument("--autor", help="nome na nota de autoria do rodapé (vazio remove)")
    p.add_argument("--cedido-a", help="quem recebe os direitos, na nota de autoria")
    p.set_defaults(f=cmd_config)
    p = sub.add_parser("modelo")
    p.add_argument("marca", nargs="?")
    p.set_defaults(f=cmd_modelo)
    p = sub.add_parser("tipos")
    p.add_argument("tipo", nargs="?")
    p.set_defaults(f=cmd_tipos)
    sub.add_parser("guia").set_defaults(f=cmd_guia)
    p = sub.add_parser("exemplo")
    p.add_argument("nome", nargs="?")
    p.set_defaults(f=cmd_exemplo)

    p = sub.add_parser("validar")
    p.add_argument("deck")
    p.add_argument("--marca")
    p.add_argument("--json", action="store_true")
    p.set_defaults(f=cmd_validar)

    p = sub.add_parser("gerar")
    p.add_argument("deck")
    p.add_argument("--marca")
    p.add_argument("--saida", help="pasta de destino (padrão: a pasta atual)")
    p.add_argument("--previa", action="store_true", help="gera também um PNG por slide e a folha de contato")
    p.add_argument("--sem-autoria", action="store_true", help="não escreve a nota de autoria no rodapé")
    p.add_argument("--sem-projeto-no-nome", action="store_true", help="não prefixa o nome do arquivo com [projeto]")
    p.set_defaults(f=cmd_gerar)

    p = sub.add_parser("prever")
    p.add_argument("arquivo")
    p.set_defaults(f=cmd_prever)
    p = sub.add_parser("revisar")
    p.add_argument("arquivo")
    p.add_argument("--json", action="store_true")
    p.set_defaults(f=cmd_revisar)
    p = sub.add_parser("verificar")
    p.add_argument("pasta", nargs="?", default=".")
    p.set_defaults(f=cmd_verificar)

    a = ap.parse_args(argv)
    try:
        return a.f(a)
    except ErroDeMarca as e:
        print(f"erro de marca: {e}", file=sys.stderr)
        return 2
