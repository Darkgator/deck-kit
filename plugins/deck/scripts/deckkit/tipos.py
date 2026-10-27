"""Contrato dos tipos de slide: campos, limites de texto e layouts do modelo."""
from __future__ import annotations

# Nomes dos layouts do modelo (um por papel). O construtor cria, o gerador usa.
LAYOUTS = {
    "capa": "Capa",
    "numero": "Número",
    "painel": "Painel e conteúdo",
    "decisao": "Decisão",
    "conteudo": "Conteúdo",
    "numeros": "Números",
    "cartoes": "Quatro cartões",
    "recomendacao": "Recomendação",
    "tabela": "Tabela e status",
}

# Limites de caracteres por campo: o que passa disso não cabe no desenho. Erro de conteúdo: reescreva mais curto.
LIMITES = {"eyebrow": 48, "titulo": 70, "subtitulo": 110, "numero": 14, "rotulo": 95, "item": 120, "corpo": 330,
           "nota": 170, "chip": 28, "celula": 60, "fato": 210, "recomendacao": 230, "impacto": 90, "passo": 60}

# Cores de etapa (cronograma) quando a etapa traz `chave` conhecida; senão, as pétalas da marca.
CORES_ETAPAS = {"entender": "0A2540", "mapear": "2979FF", "diagnosticar": "607D8B", "redesenhar": "00ACC1",
                "implementar": "FF7043", "operar": "8BC34A"}

COMUNS = ["eyebrow", "nota", "notas", "fonte"]

TIPOS = {
    "capa": {
        "uso": "Abertura: título, subtítulo e data, com foto da marca se houver.",
        "obrigatorios": ["titulo"], "opcionais": ["eyebrow", "subtitulo", "data", "foto"],
        "layout": "capa",
    },
    "numero": {
        "uso": "Um número que carrega o slide, com até 3 dados de apoio e uma frase.",
        "obrigatorios": ["numero", "rotulo"], "opcionais": ["eyebrow", "apoios", "frase", "notas", "fonte"],
        "layout": "numero",
    },
    "texto": {
        "uso": "Até 3 blocos de texto curto, cada um com rótulo. Último recurso: prefira um tipo visual.",
        "obrigatorios": ["titulo", "blocos"], "opcionais": COMUNS, "layout": "conteudo",
    },
    "lista": {
        "uso": "Até 6 itens numerados e curtos.",
        "obrigatorios": ["titulo", "itens"], "opcionais": COMUNS, "layout": "conteudo",
    },
    "divisor": {
        "uso": "Abertura de seção, com número, subtítulo e até 4 itens.",
        "obrigatorios": ["titulo"], "opcionais": ["numero", "subtitulo", "itens", "notas"], "layout": "painel",
    },
    "cartoes": {
        "uso": "De 2 a 4 números lado a lado (4 cartões e uma faixa usam o layout de painel de números).",
        "obrigatorios": ["titulo", "cartoes"], "opcionais": COMUNS + ["faixa", "alerta"], "layout": "numeros",
    },
    "recomendacao": {
        "uso": "Um achado, o que ele custa e o que fazer: três blocos num slide.",
        "obrigatorios": ["titulo", "fato", "recomendacao"],
        "opcionais": COMUNS + ["tipo_rec", "prioridade", "altera_fluxo", "impactos"], "layout": "recomendacao",
    },
    "decisao": {
        "uso": "Decisão em até 3 opções (Seguir, Ajustar, Não seguir, por padrão).",
        "obrigatorios": ["titulo"], "opcionais": COMUNS + ["subtitulo", "opcoes"], "layout": "decisao",
    },
    "tabela": {
        "uso": "Tabela nativa e editável: até 7 linhas e 6 colunas.",
        "obrigatorios": ["titulo", "cabecalho", "linhas"], "opcionais": COMUNS + ["larguras"], "layout": "tabela",
    },
    "cronograma": {
        "uso": "De 2 a 6 etapas em faixas coloridas com período e entrega, e próximos passos.",
        "obrigatorios": ["titulo", "etapas"], "opcionais": COMUNS + ["proximos"], "layout": "conteudo",
    },
    "passos": {
        "uso": "Etapas em caixas (feito, aqui, a fazer); aceita de 2 a 10, como um calendário D1 a D10.",
        "obrigatorios": ["titulo", "etapas"], "opcionais": COMUNS, "layout": "conteudo",
    },
    "fluxo": {
        "uso": "Processo em caixas ligadas por setas, de 1 a 3 trilhas (hoje, piloto, depois), com destaque opcional.",
        "obrigatorios": ["titulo", "trilhas"], "opcionais": COMUNS + ["marca", "destaque", "estilo"], "layout": "painel",
    },
    "frentes": {
        "uso": "Até 4 frentes ou recomendações, em lista (padrão) ou em quatro cartões (`formato: cartoes`).",
        "obrigatorios": ["titulo", "frentes"], "opcionais": COMUNS + ["formato"], "layout": "painel",
    },
    "implantacao": {
        "uso": "Como vamos implantar: até 4 fatos, até 5 etapas em cartões ligados por setas, o que precisa e o que não muda.",
        "obrigatorios": ["titulo", "fatos", "etapas"],
        "opcionais": COMUNS + ["precisa", "nao_muda", "nao_muda_rotulo"], "layout": "conteudo",
    },
}


class ErroDeSlide(ValueError):
    pass


def conferir(slide: dict, campo: str, valor, limite_chave: str = None) -> str:
    """Texto dentro do limite do campo; erro com o slide e o campo se passar."""
    texto = "" if valor is None else str(valor)
    limite = LIMITES[limite_chave or campo]
    if len(texto) > limite:
        ref = slide.get("titulo") or slide.get("rotulo") or slide.get("numero") or ""
        raise ErroDeSlide(f"slide '{slide.get('tipo')}' ({ref}): '{campo}' tem {len(texto)} caracteres e o limite é "
                          f"{limite}: {texto[:60]}...")
    return texto
