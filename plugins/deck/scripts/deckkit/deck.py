"""Gerador de slides sobre o modelo da marca: o texto vai para os placeholders do layout; só o que o modelo
não tem (trilhas, listas de frentes, cronograma, etapas) é desenhado com formas nativas editáveis.

Texto que não cabe é erro de conteúdo: reescreva mais curto, nunca reduza a fonte. Quando um tipo não cabe no layout
(poucos cartões, opções longas), cai num desenho sobre o layout de conteúdo, sempre com base clara.
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches as In, Pt

from .marca import Marca, mistura
from .tipos import CORES_ETAPAS, LAYOUTS, TIPOS, ErroDeSlide, conferir

MARGEM = 0.70
UTIL = 13.333 - 2 * MARGEM
BRANCO = "FFFFFF"


def _rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def _tam(texto: str, cortes, resto: float) -> float:
    """Tamanho de fonte pelo comprimento do texto: [(até N caracteres, tamanho), ...]."""
    for limite, tam in cortes:
        if len(texto) <= limite:
            return tam
    return resto


def _sobre(fundo: str, escuro: str, claro: str = BRANCO) -> str:
    """Cor de texto legível sobre o fundo: escuro em fundo claro, branco em fundo escuro."""
    r, g, b = (int(fundo[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return escuro if 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.45 else claro


class Deck:
    def __init__(self, marca: Marca, projeto: str = "", modelo: Path = None, foto: Path = None, autoria: bool = True):
        self.m = marca
        self.projeto = projeto
        self.autoria = autoria
        modelo = Path(modelo) if modelo else marca.modelo
        if not modelo.is_file():
            raise ErroDeSlide(f"o modelo da marca '{marca.nome}' não existe ({modelo}). Gere com: deckkit modelo {marca.nome}")
        self.prs = Presentation(str(modelo))
        self.lay = {l.name: l for l in self.prs.slide_layouts}
        faltam = [n for n in LAYOUTS.values() if n not in self.lay]
        if faltam:
            raise ErroDeSlide(f"o modelo {modelo.name} não tem os layouts: {', '.join(faltam)}. Gere de novo com: deckkit modelo")
        self.foto = Path(foto) if foto else marca.ativo("foto_capa")
        self.n = 0

    # ------------------------------------------------------------ primitivas de desenho
    def caixa(self, sl, x, y, w, h, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE):
        s = sl.shapes.add_shape(shape, In(x), In(y), In(w), In(h))
        s.shadow.inherit = False
        if fill is None:
            s.fill.background()
        else:
            s.fill.solid()
            s.fill.fore_color.rgb = _rgb(fill)
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = _rgb(line)
            s.line.width = Pt(lw)
        return s

    def texto(self, sl, x, y, w, h, conteudo, size=14, cor=None, bold=False, align=PP_ALIGN.LEFT, espaco=1.18,
              italic=False, anchor=MSO_ANCHOR.TOP):
        tb = sl.shapes.add_textbox(In(x), In(y), In(w), In(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        linhas = conteudo if isinstance(conteudo, list) else [conteudo]
        for i, ln in enumerate(linhas):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = espaco
            r = p.add_run()
            r.text = str(ln)
            f = r.font
            f.name, f.size, f.bold, f.italic = self.m.fonte, Pt(size), bold, italic
            f.color.rgb = _rgb(cor or self.m.cinza)
        return tb

    def faixa(self, sl, x, y, w, h, conteudo, fundo=None, barra=None, size=14, cor=None):
        self.caixa(sl, x, y, w, h, fill=fundo or self.m.palido)
        self.caixa(sl, x, y, 0.05, h, fill=barra or mistura(self.m.secundaria, 0.25))
        self.texto(sl, x + 0.26, y + 0.14, w - 0.5, h - 0.24, conteudo, size=size, cor=cor or self.m.tinta, espaco=1.3)

    def alerta(self, sl, x, y, w, h, conteudo):
        self.faixa(sl, x, y, w, h, conteudo, fundo=self.m.atual_marcado, barra=self.m.destaque)

    def nota(self, sl, y, conteudo):
        self.texto(sl, MARGEM, y, UTIL, 0.3, conteudo, size=12, cor=self.m.cinza_claro, italic=True, espaco=1.2)

    def linha(self, sl, x1, y1, x2, y2, cor, w):
        c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, In(x1), In(y1), In(x2), In(y2))
        c.line.color.rgb = _rgb(cor)
        c.line.width = Pt(w)

    def seta(self, sl, x1, y1, x2, y2, cor=None, w=1.75):
        c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, In(x1), In(y1), In(x2), In(y2))
        c.line.color.rgb = _rgb(cor or self.m.cinza)
        c.line.width = Pt(w)
        fim = etree.SubElement(c.line._get_or_add_ln(), qn("a:tailEnd"))
        fim.set("type", "triangle")
        fim.set("w", "med")
        fim.set("len", "med")

    def bola(self, sl, cx, cy, r, fill, texto, cor_texto, tam=16):
        o = sl.shapes.add_shape(MSO_SHAPE.OVAL, In(cx - r), In(cy - r), In(2 * r), In(2 * r))
        o.shadow.inherit = False
        o.fill.solid()
        o.fill.fore_color.rgb = _rgb(fill)
        o.line.fill.background()
        tf = o.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = texto
        run.font.size, run.font.bold, run.font.name = Pt(tam), True, self.m.fonte
        run.font.color.rgb = _rgb(cor_texto)

    def caixa_passo(self, sl, x, y, w, h, texto, fill, linha, lw, tam, forma=MSO_SHAPE.ROUNDED_RECTANGLE):
        o = sl.shapes.add_shape(forma, In(x), In(y), In(w), In(h))
        o.shadow.inherit = False
        if forma == MSO_SHAPE.ROUNDED_RECTANGLE:
            o.adjustments[0] = 0.08
        o.fill.solid()
        o.fill.fore_color.rgb = _rgb(fill)
        o.line.color.rgb = _rgb(linha)
        o.line.width = Pt(lw)
        tf = o.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = In(0.1)
        tf.margin_top = tf.margin_bottom = In(0.05)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = texto
        r.font.size, r.font.name = Pt(tam), self.m.fonte
        r.font.color.rgb = _rgb(self.m.tinta)

    # ------------------------------------------------------------ primitivas do modelo
    def _novo(self, chave: str):
        sl = self.prs.slides.add_slide(self.lay[LAYOUTS[chave]])
        self.n += 1
        return sl

    @staticmethod
    def _ph(sl, idx: int):
        for p in sl.placeholders:
            if p.placeholder_format.idx == idx:
                return p
        raise ErroDeSlide(f"o layout '{sl.slide_layout.name}' não tem o placeholder {idx}")

    @staticmethod
    def _tira(sl, *idxs: int) -> None:
        for p in list(sl.placeholders):
            if p.placeholder_format.idx in idxs:
                p._element.getparent().remove(p._element)

    def _p(self, ph, linhas, tam=None, cor=None, neg=None, depois=None):
        """Escreve nos parágrafos do placeholder, mantendo a formatação do layout salvo o que for dado.
        `linhas`: str, ou lista de linhas; cada linha é str, (str, opções) ou lista de (str, opções) no mesmo parágrafo."""
        tf = ph.text_frame
        linhas = linhas if isinstance(linhas, list) else [linhas]
        for p in list(tf.paragraphs)[1:]:
            p._p.getparent().remove(p._p)
        p0 = tf.paragraphs[0]
        for r in list(p0.runs):
            r._r.getparent().remove(r._r)
        for i, ln in enumerate(linhas):
            p = p0 if i == 0 else tf.add_paragraph()
            segs = ln if isinstance(ln, list) else [ln]
            primeiro = segs[0][1] if segs and not isinstance(segs[0], str) else {}
            if (primeiro.get("depois") or depois) and i < len(linhas) - 1:
                p.space_after = Pt(primeiro.get("depois", depois))
            for seg in segs:
                t, o = (seg, {}) if isinstance(seg, str) else seg
                r = p.add_run()
                r.text = t
                f = r.font
                if o.get("tam", tam):
                    f.size = Pt(o.get("tam", tam))
                if o.get("neg", neg) is not None:
                    f.bold = o.get("neg", neg)
                if o.get("cor", cor):
                    f.color.rgb = _rgb(o.get("cor", cor))

    def _rodape(self, sl) -> None:
        if self.n < 2:
            return
        partes = [self.projeto, self.m.rodape_autoria() if self.autoria else ""]
        rod = "   |   ".join(p for p in partes if p)
        self.texto(sl, 0.7, 7.2, 10.2, 0.22, f"{self.n}   {rod}".rstrip(), size=10, cor=self.m.cinza_claro)

    def _titulo(self, sl, s, cortes=None, resto=24):
        t = conferir(s, "titulo", s["titulo"])
        self._p(sl.shapes.title, t, tam=_tam(t, cortes, resto) if cortes else None)

    def _chamada(self, sl, idx: int, s) -> None:
        e = conferir(s, "eyebrow", s.get("eyebrow", ""))
        if e:
            self._p(self._ph(sl, idx), e.upper())
        else:
            self._tira(sl, idx)

    def _nota_ph(self, sl, idx: int, s) -> None:
        if s.get("nota"):
            self._p(self._ph(sl, idx), conferir(s, "nota", s["nota"]))
        else:
            self._tira(sl, idx)

    def base(self, eyebrow: str, titulo: str):
        """Layout de conteúdo com título e chamada, sem o corpo (as formas são desenhadas por cima)."""
        sl = self._novo("conteudo")
        if eyebrow:
            self._p(self._ph(sl, 12), eyebrow.upper())
        else:
            self._tira(sl, 12)
        self._p(sl.shapes.title, titulo, tam=_tam(titulo, [(52, 28)], 24))
        self._tira(sl, 10, 11)
        self._rodape(sl)
        return sl

    # ------------------------------------------------------------ tipos com placeholders
    def capa(self, s):
        sl = self._novo("capa")
        t = conferir(s, "titulo", s["titulo"])
        self._p(sl.shapes.title, t, tam=_tam(t, [(26, 46), (42, 36)], 30))
        self._p(self._ph(sl, 1), conferir(s, "subtitulo", s.get("subtitulo", "")))
        apoio = [x for x in (s.get("eyebrow"), conferir(s, "nota", s.get("data", ""), "subtitulo")) if x]
        if apoio:
            self._p(self._ph(sl, 14), apoio)
        else:
            self._tira(sl, 14)
        foto = Path(s["foto"]) if s.get("foto") else self.foto
        if foto and foto.is_file():
            f = self._ph(sl, 13).insert_picture(str(foto))
            if self.m.foto_corte and not s.get("foto"):
                f.crop_left, f.crop_top, f.crop_right, f.crop_bottom = self.m.foto_corte
            f._element.nvPicPr.cNvPr.set("descr", self.m.descricao_foto)
        else:
            self._tira(sl, 13)

    def divisor(self, s):
        sl = self._novo("painel")
        self._p(self._ph(sl, 10), conferir(s, "numero", s.get("numero", "")))
        self._titulo(sl, s, cortes=[(40, 28)])
        sub = conferir(s, "subtitulo", s.get("subtitulo", ""))
        itens = [chr(0x2022) + " " + conferir(s, "item", it) for it in s.get("itens", [])[:4]]
        corpo = ([(sub, {"tam": 20, "cor": self.m.primaria})] if sub else []) + itens
        if corpo:
            self._p(self._ph(sl, 11), corpo, tam=18, depois=10)
        else:
            self._tira(sl, 11)
        self._tira(sl, 12)
        self._rodape(sl)

    def texto_blocos(self, s):
        sl = self._novo("conteudo")
        self._chamada(sl, 12, s)
        self._titulo(sl, s, cortes=[(52, 28)])
        linhas: list = []
        for b in s["blocos"][:3]:
            linhas.append((conferir(s, "rotulo", b["rotulo"]).upper(),
                           {"tam": 12, "neg": True, "cor": self.m.alerta_txt, "depois": 3}))
            linhas.append((conferir(s, "corpo", b["corpo"]), {"tam": 16, "depois": 22}))
        self._p(self._ph(sl, 10), linhas)
        self._nota_ph(sl, 11, s)
        self._rodape(sl)

    def lista(self, s):
        sl = self._novo("conteudo")
        self._chamada(sl, 12, s)
        self._titulo(sl, s, cortes=[(52, 28)])
        itens = [[(f"{i + 1:02d}   ", {"neg": True, "cor": self.m.primaria}), (conferir(s, "item", it), {})]
                 for i, it in enumerate(s["itens"][:6])]
        self._p(self._ph(sl, 10), itens, tam=18, depois=14)
        self._nota_ph(sl, 11, s)
        self._rodape(sl)

    def cartoes(self, s):
        cs = s["cartoes"][:4]
        if len(cs) != 4 or not s.get("faixa"):
            return self._cartoes_n(s, cs)
        sl = self._novo("numeros")
        self._chamada(sl, 9, s)
        self._titulo(sl, s, cortes=[(52, 30)])
        n0 = conferir(s, "numero", cs[0]["numero"])
        self._p(self._ph(sl, 10), n0, tam=_tam(n0, [(8, 100), (11, 72)], 54))
        self._p(self._ph(sl, 11), conferir(s, "rotulo", cs[0]["rotulo"]), tam=_tam(cs[0]["rotulo"], [(40, 24)], 18))
        self._tira(sl, 12)
        for i, c in enumerate(cs[1:]):
            n = conferir(s, "numero", c["numero"])
            self._p(self._ph(sl, 13 + 2 * i), n, tam=_tam(n, [(2, 40), (4, 30), (5, 26)], 18))
            self._p(self._ph(sl, 14 + 2 * i), conferir(s, "rotulo", c["rotulo"]), tam=_tam(c["rotulo"], [(60, 16)], 13))
        self._p(self._ph(sl, 19), conferir(s, "corpo", s["faixa"]), tam=_tam(s["faixa"], [(150, 16)], 14))
        notas = [x for x in (s.get("alerta"), s.get("nota")) if x]
        if notas:
            self._p(self._ph(sl, 20), conferir(s, "nota", " · ".join(notas)[:170]))
        else:
            self._tira(sl, 20)
        self._rodape(sl)

    def _cartoes_n(self, s, cs):
        """Dois a quatro cartões de número, desenhados sobre o layout de conteúdo."""
        sl = self.base(conferir(s, "eyebrow", s.get("eyebrow", "")), conferir(s, "titulo", s["titulo"]))
        gap = 0.3
        w = (UTIL - gap * (len(cs) - 1)) / len(cs)
        for i, c in enumerate(cs):
            x = MARGEM + i * (w + gap)
            n = conferir(s, "numero", c["numero"])
            self.caixa(sl, x, 1.75, w, 2.6, fill=BRANCO, line=self.m.linha)
            self.caixa(sl, x, 1.75, w, 0.1, fill=self.m.petalas[i])
            self.texto(sl, x + 0.3, 2.1, w - 0.6, 0.9, n, size=_tam(n, [(5, 48), (8, 36)], 28), cor=self.m.primaria,
                       bold=True, espaco=1.0)
            self.texto(sl, x + 0.3, 3.15, w - 0.6, 1.1, conferir(s, "rotulo", c["rotulo"]), size=16, cor=self.m.tinta)
        y = 4.7
        if s.get("faixa"):
            self.faixa(sl, MARGEM, y, UTIL, 1.0, conferir(s, "corpo", s["faixa"]), size=16)
            y += 1.2
        if s.get("alerta"):
            self.alerta(sl, MARGEM, y, UTIL, 0.8, conferir(s, "nota", s["alerta"]))
        if s.get("nota"):
            self.nota(sl, 6.6, conferir(s, "nota", s["nota"]))

    def recomendacao(self, s):
        sl = self._novo("recomendacao")
        self._chamada(sl, 9, s)
        self._titulo(sl, s, cortes=[(52, 30)])
        chips = [conferir(s, "chip", x) for x in (s.get("tipo_rec"), s.get("prioridade"), s.get("altera_fluxo")) if x]
        if chips:
            self._p(self._ph(sl, 10), " · ".join(chips).upper())
        else:
            self._tira(sl, 10)
        self._p(self._ph(sl, 11), conferir(s, "fato", s["fato"]), tam=_tam(s["fato"], [(120, 16)], 14))
        imp = [chr(0x2022) + " " + conferir(s, "impacto", x) for x in s.get("impactos", [])[:3]]
        if imp:
            self._p(self._ph(sl, 12), imp, depois=6)
        else:
            self._tira(sl, 12)
        self._p(self._ph(sl, 13), conferir(s, "recomendacao", s["recomendacao"]),
                tam=_tam(s["recomendacao"], [(120, 18)], 16))
        self._nota_ph(sl, 14, s)
        self._rodape(sl)

    def decisao(self, s):
        opcoes = s.get("opcoes") or [{"rotulo": "Seguir", "corpo": ""}, {"rotulo": "Ajustar", "corpo": ""},
                                     {"rotulo": "Não seguir", "corpo": ""}]
        if len(opcoes) != 3 or any(len(o["rotulo"]) > 18 or len(o.get("corpo", "")) > 64 for o in opcoes):
            return self._decisao_cartoes(s, opcoes)
        sl = self._novo("decisao")
        self._titulo(sl, s, cortes=[(34, 52), (55, 40)], resto=32)
        sub = s.get("subtitulo") or s.get("eyebrow")
        if sub:
            self._p(self._ph(sl, 10), conferir(s, "subtitulo", sub))
        else:
            self._tira(sl, 10)
        for i, o in enumerate(opcoes):
            self._p(self._ph(sl, 11 + i), conferir(s, "rotulo", o["rotulo"]))
            if o.get("corpo"):
                self._p(self._ph(sl, 14 + i), conferir(s, "corpo", o["corpo"]))
            else:
                self._tira(sl, 14 + i)
        if s.get("nota"):
            self.nota(sl, 6.5, conferir(s, "nota", s["nota"]))
        self._rodape(sl)

    def _decisao_cartoes(self, s, opcoes):
        """Opções longas: cartões grandes, que comportam mais texto."""
        sl = self.base(conferir(s, "eyebrow", s.get("eyebrow", "Decisão")), conferir(s, "titulo", s["titulo"]))
        m = self.m
        cores = [(m.sucesso, m.sucesso_fundo), (m.secundaria, mistura(m.secundaria, 0.9)), (m.destaque, m.atual_marcado)]
        gap = 0.3
        w = (UTIL - gap * (len(opcoes) - 1)) / len(opcoes)
        for i, o in enumerate(opcoes[:3]):
            barra, fundo = cores[i % 3]
            x = MARGEM + i * (w + gap)
            self.caixa(sl, x, 2.1, w, 3.4, fill=fundo)
            self.caixa(sl, x, 2.1, w, 0.08, fill=barra)
            self.texto(sl, x + 0.3, 2.45, w - 0.6, 0.5, conferir(s, "rotulo", o["rotulo"]), size=24, cor=_sobre(fundo, barra, barra),
                       bold=True)
            self.texto(sl, x + 0.3, 3.15, w - 0.6, 2.2, conferir(s, "corpo", o.get("corpo", "")), size=14, cor=m.tinta, espaco=1.3)
        if s.get("nota"):
            self.nota(sl, 5.9, conferir(s, "nota", s["nota"]))

    def tabela(self, s):
        sl = self._novo("tabela")
        self._chamada(sl, 9, s)
        self._titulo(sl, s, cortes=[(52, 30)])
        cab, linhas = s["cabecalho"], s["linhas"][:7]
        cols = len(cab)
        ph = self._ph(sl, 10)
        largura = ph.width
        gf = ph.insert_table(rows=len(linhas) + 1, cols=cols)
        gf.width = largura
        tb = gf.table
        larguras = s.get("larguras") or [1] * cols
        total = sum(larguras)
        for j, l in enumerate(larguras):
            tb.columns[j].width = int(largura * l / total)
        for linha_t in tb.rows:
            linha_t.height = In(0.6)
        for i, linha in enumerate([cab] + linhas):
            for j, v in enumerate(linha):
                c = tb.cell(i, j)
                c.text = ""
                r = c.text_frame.paragraphs[0].add_run()
                r.text = conferir(s, "celula", v)
                r.font.size = Pt(12)
                r.font.name = self.m.fonte
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
        self._nota_ph(sl, 11, s)
        self._rodape(sl)

    def numero(self, s):
        sl = self._novo("numero")
        self._chamada(sl, 10, s)
        n = conferir(s, "numero", s["numero"])
        self._p(self._ph(sl, 11), n, tam=_tam(n, [(7, 120), (9, 96)], 72))
        self._p(self._ph(sl, 12), conferir(s, "rotulo", s["rotulo"]), tam=_tam(s["rotulo"], [(60, 22)], 18))
        ap = s.get("apoios", [])[:3]
        for i in range(3):
            if i < len(ap):
                nn = conferir(s, "numero", ap[i]["numero"])
                self._p(self._ph(sl, 13 + i), [[(nn, {"tam": 36, "neg": True, "cor": self.m.primaria})],
                                                  [(conferir(s, "rotulo", ap[i]["rotulo"]), {})]])
            else:
                self._tira(sl, 13 + i)
        if s.get("frase"):
            self._p(self._ph(sl, 16), conferir(s, "corpo", s["frase"]), tam=_tam(s["frase"], [(70, 22)], 18))
        else:
            self._tira(sl, 16)
        self._rodape(sl)

    # ------------------------------------------------------------ tipos desenhados
    def cronograma(self, s):
        sl = self.base(conferir(s, "eyebrow", s.get("eyebrow", "")), conferir(s, "titulo", s["titulo"]))
        etapas = s["etapas"]
        n = len(etapas)
        gap = 0.08
        w = (UTIL - gap * (n - 1)) / n
        for i, e in enumerate(etapas):
            hexa = (e.get("cor") or CORES_ETAPAS.get(str(e.get("chave", "")).lower()) or self.m.petalas[i % 4]).lstrip("#")
            x = MARGEM + i * (w + gap)
            self.texto(sl, x, 2.0, w, 0.3, conferir(s, "chip", e.get("periodo", "")), size=14, cor=self.m.tinta, bold=True,
                       align=PP_ALIGN.CENTER)
            self.caixa(sl, x, 2.4, w, 1.3, fill=hexa)
            self.texto(sl, x + 0.1, 2.4, w - 0.2, 1.3, conferir(s, "chip", e["etapa"]), size=16,
                       cor=_sobre(hexa, self.m.tinta), bold=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self.texto(sl, x + 0.05, 3.85, w - 0.1, 1.2, conferir(s, "rotulo", e.get("entrega", "")), size=14,
                       align=PP_ALIGN.CENTER, cor=self.m.tinta)
        if s.get("proximos"):
            self.texto(sl, MARGEM, 5.2, UTIL, 0.3, "PRÓXIMOS PASSOS", size=12, cor=self.m.alerta_txt, bold=True)
            for i, p in enumerate(s["proximos"][:3]):
                self.texto(sl, MARGEM, 5.55 + i * 0.4, UTIL, 0.35, chr(0x2022) + " " + conferir(s, "item", p), size=16,
                           cor=self.m.tinta)
        if s.get("nota"):
            self.nota(sl, 6.7, conferir(s, "nota", s["nota"]))

    def passos(self, s):
        sl = self.base(conferir(s, "eyebrow", s.get("eyebrow", "")), conferir(s, "titulo", s["titulo"]))
        etapas = s["etapas"]
        n = len(etapas)
        gap = 0.18 if n < 8 else 0.08
        w = (UTIL - gap * (n - 1)) / n
        m = self.m
        for i, e in enumerate(etapas):
            x = MARGEM + i * (w + gap)
            aqui, feito = e.get("estado") == "aqui", e.get("estado") == "feito"
            fundo = m.secundaria if aqui else (m.sucesso_fundo if feito else BRANCO)
            cor = BRANCO if aqui else (m.sucesso_esc if feito else m.tinta)
            self.caixa(sl, x, 2.6, w, 1.4, fill=fundo, line=None if aqui or feito else m.linha)
            self.texto(sl, x + 0.1, 2.6, w - 0.2, 1.4, conferir(s, "chip", e["etapa"]), size=15, cor=cor, bold=True,
                       align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            if aqui:
                self.texto(sl, x - 0.2, 2.15, w + 0.4, 0.3, "ESTAMOS AQUI", size=12, cor=m.alerta_txt, bold=True,
                           align=PP_ALIGN.CENTER)
            self.texto(sl, x if n >= 8 else x + 0.05, 4.15, w if n >= 8 else w - 0.1, 1.2,
                       conferir(s, "rotulo", e.get("detalhe", "")), size=12 if n >= 8 else 13,
                       align=PP_ALIGN.CENTER, cor=m.tinta)
        if s.get("nota"):
            self.nota(sl, 6.62, conferir(s, "nota", s["nota"]))

    def _trilha(self, sl, y, rotulo, passos, tom, marcados, compacto=False, estilo="caixas"):
        """Uma trilha de passos em caixas ligadas por setas (ou em chevrons). Devolve os centros e a base das caixas."""
        m = self.m
        self.texto(sl, 5.4, y, 5, 0.3, rotulo, size=14, cor=m.primaria, bold=True)
        w, gap = 1.6, 0.33
        h = 0.85 if compacto else 1.15
        yb = y + 0.4
        xs = []
        for i, tx in enumerate(passos):
            marcado = (i + 1) in marcados
            if tom == "novo":
                fill, linha, lw = m.palido, m.primaria, 1.5
            else:
                fill, linha, lw = (m.atual_marcado, m.destaque, 2.25) if marcado else (m.atual_fundo, m.aviso, 1.5)
            tam = 13 if compacto else 14
            if estilo == "chevron":
                x = 5.4 + i * (w + gap - 0.02)
                self.caixa_passo(sl, x, yb, w + gap, h, tx, fill, linha, lw, tam, forma=MSO_SHAPE.PENTAGON)
                xs.append(x + (w + gap) / 2)
            else:
                x = 5.4 + i * (w + gap)
                self.caixa_passo(sl, x, yb, w, h, tx, fill, linha, lw, tam)
                if i < len(passos) - 1:
                    self.seta(sl, x + w + 0.03, yb + h / 2, x + w + gap - 0.03, yb + h / 2)
                xs.append(x + w / 2)
        return xs, yb + h

    def fluxo(self, s):
        sl = self._novo("painel")
        self._chamada(sl, 10, s)
        self._titulo(sl, s, cortes=[(75, 28)])
        self._tira(sl, 11)
        self._nota_ph(sl, 12, s)
        trilhas = s["trilhas"][:3]
        compacto = len(trilhas) == 3
        ys = [0.5, 2.75, 5.0] if compacto else [1.0, 3.95]
        marca = s.get("marca")
        estilo = s.get("estilo", "caixas")
        for k, tr in enumerate(trilhas):
            passos = list(tr["passos"][:4])
            for p in passos:
                conferir(s, "passo", p, "passo")
            marcados = set()
            if marca and marca.get("de") and tr.get("tom") != "novo":
                marcados = set(range(marca["de"], marca.get("ate", marca["de"]) + 1))
            xs, base = self._trilha(sl, ys[k], conferir(s, "rotulo", tr["rotulo"]).upper(), passos,
                                    "novo" if tr.get("tom") == "novo" else "atual", marcados, compacto, estilo)
            if k == 0 and marca and marca.get("de"):
                x1, x2 = xs[marca["de"] - 1] - 0.8, xs[min(marca["ate"], len(xs)) - 1] + 0.8
                self.linha(sl, x1, base + 0.1, x2, base + 0.1, self.m.destaque, 3)
                self.texto(sl, x1, base + 0.18, x2 - x1, 0.3, conferir(s, "rotulo", marca["texto"]), size=14,
                           cor=self.m.alerta_txt, bold=True, align=PP_ALIGN.CENTER)
        dest = s.get("destaque")
        if dest:
            if compacto or (len(trilhas) > 1 and len(trilhas[1]["passos"]) > 3):
                raise ErroDeSlide("fluxo: o destaque só cabe com duas trilhas e a segunda com até 3 passos")
            self.texto(sl, 11.1, 4.35, 1.9, 0.6, conferir(s, "numero", dest["valor"]), size=32, cor=self.m.primaria,
                       bold=True, align=PP_ALIGN.CENTER)
            self.texto(sl, 11.1, 5.0, 1.9, 0.8, conferir(s, "rotulo", dest["rotulo"]), size=13, cor=self.m.cinza,
                       align=PP_ALIGN.CENTER)
        self._rodape(sl)

    def frentes(self, s):
        itens = s["frentes"][:4]
        if s.get("formato") == "cartoes":
            return self._frentes_cartoes(s, itens)
        sl = self._novo("painel")
        m = self.m
        self._chamada(sl, 10, s)
        self._titulo(sl, s, cortes=[(75, 28)])
        self._tira(sl, 11)
        self._nota_ph(sl, 12, s)
        for i, it in enumerate(itens):
            y0 = 0.75 + i * 1.5
            if i:
                self.linha(sl, 5.4, y0 - 0.1, 12.75, y0 - 0.1, m.linha, 1)
            fundo = m.petalas[i]
            self.bola(sl, 5.4 + 0.425, y0 + 0.2 + 0.425, 0.425, fundo, conferir(s, "numero", it["numero"]),
                      _sobre(fundo, m.tinta), tam=22)
            self.texto(sl, 6.55, y0 + 0.15, 3.7, 1.1, conferir(s, "rotulo", it["titulo"]), size=20, cor=m.tinta,
                       bold=True, anchor=MSO_ANCHOR.MIDDLE)
            self.texto(sl, 10.5, y0 + 0.2, 2.3, 0.5, [conferir(s, "chip", it.get("area", "")).upper(),
                                                      conferir(s, "chip", it.get("tipo", "")).upper()],
                       size=12, cor=m.cinza, bold=True)
            self.texto(sl, 10.5, y0 + 0.8, 2.3, 0.4, conferir(s, "chip", it.get("situacao", "")), size=14,
                       cor=m.sucesso_esc if it.get("ok") else m.cinza, bold=True)
        self._rodape(sl)

    def _frentes_cartoes(self, s, itens):
        sl = self._novo("cartoes")
        self._chamada(sl, 9, s)
        self._titulo(sl, s, cortes=[(52, 30)])
        for i in range(4):
            base = 10 + i * 5
            if i < len(itens):
                it = itens[i]
                self._p(self._ph(sl, base), conferir(s, "numero", it["numero"]))
                self._p(self._ph(sl, base + 1), conferir(s, "rotulo", it["titulo"]), tam=_tam(it["titulo"], [(50, 18)], 16))
                self._p(self._ph(sl, base + 2), conferir(s, "chip", it.get("area", "")).upper())
                self._p(self._ph(sl, base + 3), conferir(s, "chip", it.get("tipo", "")).upper())
                self._p(self._ph(sl, base + 4), conferir(s, "chip", it.get("situacao", "")),
                        cor=self.m.sucesso_esc if it.get("ok") else self.m.cinza)
            else:
                self._tira(sl, *range(base, base + 5))
        if s.get("nota"):
            self.nota(sl, 6.6, conferir(s, "nota", s["nota"]))
        self._rodape(sl)

    def implantacao(self, s):
        sl = self.base(conferir(s, "eyebrow", s.get("eyebrow", "")), conferir(s, "titulo", s["titulo"]))
        m = self.m
        fatos = s["fatos"][:4]
        gap = 0.25
        w = (UTIL - gap * (len(fatos) - 1)) / len(fatos)
        for i, f in enumerate(fatos):
            x = MARGEM + i * (w + gap)
            self.caixa(sl, x, 1.5, w, 1.2, fill=m.palido)
            self.caixa(sl, x, 1.5, 0.06, 1.2, fill=m.secundaria)
            self.texto(sl, x + 0.25, 1.56, w - 0.4, 0.25, conferir(s, "chip", f["rotulo"]).upper(), size=12, cor=m.cinza, bold=True)
            self.texto(sl, x + 0.25, 1.82, w - 0.4, 0.4, conferir(s, "chip", f["valor"]),
                       size=20 if len(f["valor"]) <= 14 else 16, cor=m.primaria, bold=True)
            self.texto(sl, x + 0.25, 2.2, w - 0.4, 0.45, conferir(s, "rotulo", f.get("detalhe", "")), size=12, cor=m.cinza)
        et = s["etapas"][:5]
        n = len(et)
        gap = 0.3
        w = (UTIL - gap * (n - 1)) / n
        for i, e in enumerate(et):
            depois = bool(e.get("depois"))
            x = MARGEM + i * (w + gap)
            self.caixa(sl, x, 2.9, w, 2.5, fill=BRANCO, line=m.cinza_claro if depois else m.primaria, lw=1.5)
            self.texto(sl, x + 0.12, 3.0, w - 0.24, 0.3, conferir(s, "chip", e.get("periodo", "")), size=14,
                       cor=m.cinza if depois else m.primaria, bold=True)
            self.texto(sl, x + 0.12, 3.32, w - 0.24, 0.6, conferir(s, "chip", e["etapa"]) if len(e["etapa"]) <= 28 else
                       conferir(s, "rotulo", e["etapa"]), size=16, cor=m.tinta, bold=True)
            self.texto(sl, x + 0.12, 3.98, w - 0.24, 1.4, conferir(s, "item", e.get("detalhe", "")), size=12, cor=m.cinza)
            if i < n - 1:
                self.seta(sl, x + w + 0.04, 4.15, x + w + gap - 0.04, 4.15)
            if e.get("aprovacao"):
                self.texto(sl, x, 5.45, w, 0.22, "APROVAÇÃO", size=12, cor=m.alerta_txt, bold=True)
                self.texto(sl, x, 5.68, w, 0.45, conferir(s, "chip", e["aprovacao"]) if len(e["aprovacao"]) <= 28
                           else conferir(s, "rotulo", e["aprovacao"]), size=12, cor=m.tinta, bold=True)
        caixas = [("Para virar, precisa", s.get("precisa")), (s.get("nao_muda_rotulo", "O que não muda"), s.get("nao_muda"))]
        caixas = [c for c in caixas if c[1]]
        if caixas:
            w2 = (UTIL - 0.25 * (len(caixas) - 1)) / len(caixas)
            for i, (rot, tx) in enumerate(caixas):
                x = MARGEM + i * (w2 + 0.25)
                self.caixa(sl, x, 6.25, w2, 0.7, fill=m.suave)
                self.texto(sl, x + 0.2, 6.28, w2 - 0.4, 0.2, rot.upper(), size=12, cor=m.alerta_txt, bold=True)
                self.texto(sl, x + 0.2, 6.5, w2 - 0.4, 0.42, conferir(s, "nota", tx), size=12, cor=m.tinta)
        self._rodape(sl)

    # ------------------------------------------------------------ entrada
    METODOS = {"capa": "capa", "divisor": "divisor", "texto": "texto_blocos", "lista": "lista", "cartoes": "cartoes",
               "recomendacao": "recomendacao", "decisao": "decisao", "tabela": "tabela", "numero": "numero",
               "cronograma": "cronograma", "passos": "passos", "fluxo": "fluxo", "frentes": "frentes",
               "implantacao": "implantacao"}

    def adicionar(self, s: dict) -> None:
        tipo = s.get("tipo")
        if tipo not in self.METODOS:
            raise ErroDeSlide(f"tipo de slide desconhecido: {tipo!r}. Tipos: {', '.join(self.METODOS)}")
        faltam = [c for c in TIPOS[tipo]["obrigatorios"] if c not in s or s[c] in (None, "", [])]
        if faltam:
            raise ErroDeSlide(f"slide '{tipo}' ({s.get('titulo', '')}): faltam os campos {', '.join(faltam)}")
        s = dict(s)
        if s.get("fonte"):  # a fonte entra na nota do slide
            s["nota"] = (s["nota"] + " " if s.get("nota") else "") + "Fonte: " + s["fonte"]
        getattr(self, self.METODOS[tipo])(s)
        if s.get("notas"):  # o que não coube no slide fica nas notas do apresentador
            self.prs.slides[len(self.prs.slides) - 1].notes_slide.notes_text_frame.text = s["notas"]

    def salvar(self, destino: Path) -> None:
        from datetime import datetime
        cp = self.prs.core_properties
        cp.title = self.projeto or Path(destino).stem
        cp.author = self.m.autoria.get("autor", "") if self.autoria else ""
        cp.last_modified_by = cp.author
        cp.comments = ""
        cp.created = cp.modified = datetime.now()
        self.prs.save(str(destino))
