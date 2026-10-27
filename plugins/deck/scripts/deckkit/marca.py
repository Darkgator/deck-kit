"""Pacotes de marca: cores, fonte, logotipo e o modelo .pptx de uma organização.

Um pacote é uma pasta com `marca.json` (e, se houver, imagens). O modelo `modelo.pptx` é gerado a partir
dele por `deckkit modelo`. A marca "neutro" vem embutida; outras ficam em `~/.deckkit/marcas/<nome>/`,
em `./.deckkit/marcas/<nome>/` (por projeto) ou em qualquer pasta informada por caminho.
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path

RAIZ_PLUGIN = Path(__file__).resolve().parent.parent.parent
EMBUTIDAS = RAIZ_PLUGIN / "marcas"
CASA = Path.home() / ".deckkit"

PADRAO = {
    "fonte": "Arial",
    "cores": {
        "primaria": "1F3A5F", "secundaria": "2F6DB5", "destaque": "D9480F", "aviso": "F08C00",
        "petalas": ["C92A2A", "D9480F", "F08C00", "F2C14E"],
        "fundo": "F3F1EA", "tinta": "1E1E1C", "cinza": "5B6475",
    },
    "estilo": {"sem_travessao": False, "exigir_fonte": False, "vedados": []},
}


class ErroDeMarca(ValueError):
    pass


def mistura(hexa: str, fator: float) -> str:
    """Mistura a cor com branco: 0 mantém, 1 vira branco."""
    h = hexa.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    r, g, b = (round(v + (255 - v) * fator) for v in (r, g, b))
    return f"{r:02X}{g:02X}{b:02X}"


def _hex(v: str) -> str:
    v = v.lstrip("#").upper()
    if len(v) != 6 or any(c not in "0123456789ABCDEF" for c in v):
        raise ErroDeMarca(f"cor inválida: {v!r} (use 6 dígitos hexadecimais, como 1F3A5F)")
    return v


class Marca:
    """Cores derivadas, fonte e ativos de uma marca."""

    def __init__(self, pasta: Path, dados: dict):
        self.pasta = Path(pasta)
        self.nome = dados.get("nome") or self.pasta.name
        self.fonte = dados.get("fonte", PADRAO["fonte"])
        c = {**PADRAO["cores"], **dados.get("cores", {})}
        self.primaria, self.secundaria = _hex(c["primaria"]), _hex(c["secundaria"])
        self.destaque, self.aviso = _hex(c["destaque"]), _hex(c["aviso"])
        pet = c["petalas"]
        if len(pet) != 4:
            raise ErroDeMarca("cores.petalas precisa de 4 cores")
        self.petalas = [_hex(p) for p in pet]
        self.fundo, self.tinta, self.cinza = _hex(c["fundo"]), _hex(c["tinta"]), _hex(c["cinza"])
        # derivadas (podem ser sobrescritas em marca.json)
        self.painel = _hex(c.get("painel") or mistura(self.secundaria, 0.88))
        self.suave = _hex(c.get("suave") or mistura(self.primaria, 0.96))
        self.palido = _hex(c.get("palido") or mistura(self.primaria, 0.92))
        self.linha = _hex(c.get("linha") or mistura(self.primaria, 0.82))
        self.sucesso = _hex(c.get("sucesso", "2E9E5B"))
        self.sucesso_esc = _hex(c.get("sucesso_esc", "1F7A44"))
        self.sucesso_fundo = _hex(c.get("sucesso_fundo", "EAF6EF"))
        self.alerta_txt = _hex(c.get("alerta_txt", "C2410C"))
        self.atual_fundo = _hex(c.get("atual_fundo", mistura(self.aviso, 0.9)))
        self.atual_marcado = _hex(c.get("atual_marcado", mistura(self.destaque, 0.85)))
        self.cinza_claro = _hex(c.get("cinza_claro", "A5ADBD"))
        est = {**PADRAO["estilo"], **dados.get("estilo", {})}
        self.sem_travessao = bool(est["sem_travessao"])
        self.exigir_fonte = bool(est["exigir_fonte"])
        self.vedados = [str(t).lower() for t in est["vedados"]]
        self.classificacao = dados.get("classificacao")  # ex.: "INTERNO"
        self.autoria = dados.get("autoria") or {}  # {"autor": "...", "cedido_a": "..."}
        self.foto_corte = dados.get("foto_corte")  # [esq, topo, dir, base] em fração, ou None
        self.descricao_foto = dados.get("descricao_foto", "Fotografia institucional")
        self._ativos = {k: dados.get(k) for k in ("logo", "marca_dagua", "ornamento", "foto_capa")}

    def ativo(self, chave: str):
        rel = self._ativos.get(chave)
        if not rel:
            return None
        p = (self.pasta / rel).resolve()
        if not p.is_file():
            raise ErroDeMarca(f"marca '{self.nome}': o arquivo de {chave} não existe: {p}")
        return p

    @property
    def modelo(self) -> Path:
        return self.pasta / "modelo.pptx"

    def rodape_autoria(self) -> str:
        if not self.autoria.get("autor"):
            return ""
        from datetime import date
        ced = f" Cedido a {self.autoria['cedido_a']}." if self.autoria.get("cedido_a") else ""
        return f"{chr(0xA9)} {date.today().year} {self.autoria['autor']}.{ced}"


def carregar(pasta: Path) -> Marca:
    arq = Path(pasta) / "marca.json"
    if not arq.is_file():
        raise ErroDeMarca(f"não achei marca.json em {pasta}")
    return Marca(Path(pasta), json.loads(arq.read_text(encoding="utf-8-sig")))


def _pastas_de_busca():
    yield Path.cwd() / ".deckkit" / "marcas"
    yield CASA / "marcas"
    yield EMBUTIDAS


def listar() -> list[tuple[str, Path]]:
    achadas = {}
    for base in _pastas_de_busca():
        if base.is_dir():
            for p in sorted(base.iterdir()):
                if (p / "marca.json").is_file():
                    achadas.setdefault(p.name, p)
    return sorted(achadas.items())


def padrao_configurado() -> str | None:
    cfg = CASA / "config.json"
    if cfg.is_file():
        try:
            return json.loads(cfg.read_text(encoding="utf-8")).get("marca_padrao")
        except (OSError, ValueError):
            return None
    return None


def definir_padrao(nome: str) -> None:
    CASA.mkdir(parents=True, exist_ok=True)
    cfg = CASA / "config.json"
    dados = {}
    if cfg.is_file():
        try:
            dados = json.loads(cfg.read_text(encoding="utf-8"))
        except ValueError:
            dados = {}
    dados["marca_padrao"] = nome
    cfg.write_text(json.dumps(dados, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def localizar(nome_ou_caminho: str | None = None) -> Marca:
    """Marca pelo nome, por caminho de pasta ou pela marca padrão (config), senão a neutra."""
    ref = nome_ou_caminho or padrao_configurado() or "neutro"
    p = Path(ref)
    if p.is_dir() and (p / "marca.json").is_file():
        return _aplicar_config(carregar(p))
    for nome, pasta in listar():
        if nome.lower() == ref.lower():
            return _aplicar_config(carregar(pasta))
    disponiveis = ", ".join(n for n, _ in listar()) or "nenhuma"
    raise ErroDeMarca(f"marca '{ref}' não encontrada. Disponíveis: {disponiveis}. "
                      f"Instale uma com: deckkit marca-instalar <pasta>")


def instalar(origem: Path) -> Path:
    """Copia um pacote de marca para ~/.deckkit/marcas/<nome>/."""
    marca = carregar(origem)
    destino = CASA / "marcas" / Path(origem).name
    if destino.exists():
        shutil.rmtree(destino)
    shutil.copytree(origem, destino)
    return destino


def _config() -> dict:
    cfg = CASA / "config.json"
    if cfg.is_file():
        try:
            return json.loads(cfg.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}
    return {}


def _aplicar_config(m: Marca) -> Marca:
    """Preferências pessoais em ~/.deckkit/config.json valem sobre a marca: `estilo` e `autoria`."""
    cfg = _config()
    est = cfg.get("estilo") or {}
    if "sem_travessao" in est:
        m.sem_travessao = bool(est["sem_travessao"])
    if "exigir_fonte" in est:
        m.exigir_fonte = bool(est["exigir_fonte"])
    if est.get("vedados"):
        m.vedados = list(m.vedados) + [str(t).lower() for t in est["vedados"]]
    if "autoria" in cfg:
        m.autoria = cfg["autoria"] or {}
    return m
