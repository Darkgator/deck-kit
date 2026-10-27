"""Constrói o modelo .pptx de uma marca: 9 layouts reais (placeholders), tema com as cores e a fonte da marca.

Base clara em todos os layouts. Imagens (logotipo, marca d'água, ornamento) e selo de classificação são opcionais.
Gera `modelo.pptx` e `modelo.potx` na pasta da marca.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from .marca import Marca
from .tipos import LAYOUTS

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
WHITE = "FFFFFF"


def E(v: float) -> int:
    return int(round(v * 914400))


def _xml(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def parse(xml: str):
    return etree.fromstring(xml.replace("<p:sp>", f"<p:sp {NS}>", 1).replace("<p:pic>", f"<p:pic {NS}>", 1))


class Camada:
    """Acumula as formas de um layout e entrega os elementos XML na ordem."""

    def __init__(self, part, m: Marca):
        self.part, self.m, self.itens, self.n = part, m, [], 1

    def _id(self) -> int:
        self.n += 1
        return self.n

    def retangulo(self, nome, x, y, w, h, fill, alpha=None, linha=None):
        a = f'<a:alpha val="{int(alpha * 100000)}"/>' if alpha else ""
        ln = (f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{linha}"/></a:solidFill></a:ln>' if linha
              else '<a:ln><a:noFill/></a:ln>')
        self.itens.append(parse(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self._id()}" name="{nome}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}">{a}</a:srgbClr></a:solidFill>'
            f'{ln}</p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="pt-BR"/></a:p></p:txBody></p:sp>'))

    def texto_fixo(self, nome, x, y, w, h, texto, tam, cor, neg=True, spc=200, al="ctr"):
        self.itens.append(parse(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self._id()}" name="{nome}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{al}"/><a:r><a:rPr lang="pt-BR" sz="{int(tam * 100)}" b="{int(neg)}" spc="{spc}">'
            f'<a:solidFill><a:srgbClr val="{cor}"/></a:solidFill><a:latin typeface="{self.m.fonte}"/></a:rPr>'
            f'<a:t>{_xml(texto)}</a:t></a:r></a:p></p:txBody></p:sp>'))

    def imagem(self, nome, arq: Path, x, y, w, h, descr=""):
        _, rid = self.part.get_or_add_image_part(str(arq))
        self.itens.append(parse(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{self._id()}" name="{nome}" descr="{_xml(descr)}"/>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
            f'<p:nvPr userDrawn="1"/></p:nvPicPr><p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'))

    def _proporcao(self, arq: Path) -> float:
        from PIL import Image
        with Image.open(arq) as im:
            return im.height / im.width

    def logo(self, x, y, w):
        arq = self.m.ativo("logo")
        if arq:
            self.imagem("Logotipo", arq, x, y, w, w * self._proporcao(arq), f"Logotipo {self.m.nome}")

    def marca_dagua(self, x, y, w):
        arq = self.m.ativo("marca_dagua")
        if arq:
            self.imagem("Marca d'água", arq, x, y, w, w * self._proporcao(arq), "")

    def ornamento(self, x, y, w):
        arq = self.m.ativo("ornamento")
        if arq:
            self.imagem("Ornamento", arq, x, y, w, w * self._proporcao(arq), "")

    def classificacao(self, direita=False):
        if not self.m.classificacao:
            return
        if direita:  # layouts com chamada longa no alto à esquerda: o selo vai para o canto direito
            self.texto_fixo("Classificação", 11.1, 0.12, 1.9, 0.22, self.m.classificacao, 11, self.m.petalas[0], al="r")
        else:
            self.texto_fixo("Classificação", 5.67, 0.12, 2.0, 0.22, self.m.classificacao, 11, self.m.petalas[0])

    def ph(self, nome, tipo, idx, x, y, w, h, tam=16, neg=False, cor=None, al="l", anc="t", it=False,
           fill=None, linha=None, prompt="Clique para editar o texto", spc=None, ls=None, margem=0):
        cor = cor or self.m.tinta
        attr = f'type="{tipo}"' + (f' idx="{idx}"' if idx is not None else "")
        if tipo in ("body", "pic") and idx is not None:
            attr += ' sz="quarter"'
        sp = ""
        if fill:
            sp += f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        if linha:
            sp += f'<a:ln w="28575"><a:solidFill><a:srgbClr val="{linha}"/></a:solidFill></a:ln>'
        geom = ('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>'
                if fill or linha else "")
        ins = E(margem)
        if tipo in ("pic", "tbl"):
            corpo = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="pt-BR"/></a:p></p:txBody>'
        else:
            corpo = (f'<p:txBody><a:bodyPr wrap="square" lIns="{ins}" tIns="{ins}" rIns="{ins}" bIns="{ins}" anchor="{anc}"><a:normAutofit/></a:bodyPr>'
                     f'<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="{al}">'
                     + (f'<a:lnSpc><a:spcPct val="{int(ls * 100000)}"/></a:lnSpc>' if ls else "")
                     + f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/>'
                     f'<a:defRPr lang="pt-BR" sz="{int(tam * 100)}" b="{int(neg)}" i="{int(it)}"'
                     + (f' spc="{spc}"' if spc else "") +
                     f'><a:solidFill><a:srgbClr val="{cor}"/></a:solidFill><a:latin typeface="{self.m.fonte}"/></a:defRPr>'
                     f'</a:lvl1pPr></a:lstStyle>'
                     f'<a:p><a:r><a:rPr lang="pt-BR"/><a:t>{_xml(prompt)}</a:t></a:r></a:p></p:txBody>')
        self.itens.append(parse(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{self._id()}" name="{nome}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {attr}/></p:nvPr></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{E(x)}" y="{E(y)}"/><a:ext cx="{E(w)}" cy="{E(h)}"/></a:xfrm>{geom}{sp}</p:spPr>'
            f'{corpo}</p:sp>'))


def _limpar(layout, nome: str, fundo: str):
    sld = layout._element
    if "type" in sld.attrib:
        del sld.attrib["type"]
    pns = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    c = sld.find(f"{pns}cSld")
    c.set("name", nome)
    for bg in c.findall(f"{pns}bg"):
        c.remove(bg)
    c.insert(0, etree.fromstring(
        f'<p:bg {NS}><p:bgPr><a:solidFill><a:srgbClr val="{fundo}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'))
    arv = c.find(f"{pns}spTree")
    for filho in list(arv):
        if not filho.tag.endswith("}nvGrpSpPr") and not filho.tag.endswith("}grpSpPr"):
            arv.remove(filho)
    return arv


class Construtor:
    def __init__(self, m: Marca):
        self.m = m

    # ---- família F: contar
    def capa(self, c: Camada):
        m = self.m
        c.logo(0.7, 0.6, 2.3)
        c.retangulo("Bloco da foto", 5.6, 0, 7.733, 7.5, m.painel)  # aparece quando a capa não tem foto
        c.ornamento(4.0, 5.9, 1.4)
        c.ph("Título", "ctrTitle", None, 0.7, 2.0, 4.8, 2.6, 46, True, m.primaria, ls=0.98, prompt="Título da apresentação")
        c.ph("Subtítulo", "subTitle", 1, 0.7, 4.75, 4.6, 1.0, 16, False, m.cinza, prompt="Subtítulo")
        c.ph("Área, projeto e data", "body", 14, 0.7, 6.35, 3.2, 0.85, 12, False, m.cinza, prompt="Área, projeto e data")
        c.ph("Foto", "pic", 13, 5.6, 0, 7.733, 7.5)

    def numero(self, c: Camada):
        m = self.m
        c.marca_dagua(7.4, 0.2, 6.2)
        c.classificacao()
        c.logo(11.2, 6.85, 1.5)
        c.retangulo("Filete", 0.8, 5.85, 1.2, 0.07, m.petalas[3])
        c.ph("Chamada", "body", 10, 0.8, 0.7, 8, 0.3, 12, True, m.cinza, spc=150, prompt="CHAMADA EM CAIXA ALTA")
        c.ph("Número", "body", 11, 0.8, 1.2, 9, 2.1, 120, True, m.primaria, prompt="99.353")
        c.ph("Legenda do número", "body", 12, 0.8, 3.3, 11, 0.5, 22, False, m.tinta, prompt="O que o número mede e o período")
        for i in range(3):
            c.ph(f"Dado de apoio {i + 1}", "body", 13 + i, 0.8 + i * 3.6, 4.55, 3.2, 0.9, 14, False, m.cinza,
                 prompt="4  descrição do dado")
        c.ph("Frase de impacto", "body", 16, 0.8, 6.05, 11.8, 0.5, 22, False, m.primaria, it=True, prompt="Frase de impacto")

    def painel(self, c: Camada):
        m = self.m
        c.retangulo("Painel", 0, 0, 4.7, 7.5, m.painel)
        c.marca_dagua(-1.3, 4.1, 4.9)
        c.logo(0.7, 0.6, 2.0)
        c.classificacao()
        c.logo(11.3, 6.95, 1.5)
        c.ph("Chamada", "body", 10, 0.7, 1.9, 3.5, 0.3, 12, True, m.cinza, spc=120, prompt="CHAMADA EM CAIXA ALTA")
        c.ph("Título", "title", None, 0.7, 2.35, 3.5, 3.2, 28, True, m.primaria, ls=1.05, prompt="Frase que afirma o ponto do slide")
        c.ph("Nota do painel", "body", 12, 0.7, 6.0, 3.5, 0.9, 14, False, m.cinza, prompt="Nota de rodapé do painel (opcional)")
        c.ph("Conteúdo", "body", 11, 5.4, 0.8, 7.4, 5.9, 16, False, m.tinta, prompt="Conteúdo: texto, passos, lista ou diagrama")

    def decisao(self, c: Camada):
        m = self.m
        c.logo(0.8, 6.85, 1.5)
        c.classificacao()
        c.ph("Título", "title", None, 0.8, 1.5, 11, 1.7, 52, True, m.primaria, ls=1.0, prompt="Pergunta ou decisão")
        c.ph("Subtítulo", "body", 10, 0.8, 3.45, 10, 0.5, 20, False, m.cinza, prompt="Contexto da decisão")
        for i in range(3):
            c.ph(f"Opção {i + 1}", "body", 11 + i, 0.8 + i * 4.0, 4.5, 3.5, 0.95, 24, True, m.primaria, al="ctr", anc="ctr",
                 fill=WHITE, linha=m.primaria, prompt=f"Opção {i + 1}", margem=0.08)
            c.ph(f"Explicação {i + 1}", "body", 14 + i, 1.0 + i * 4.0, 5.65, 3.1, 0.9, 16, False, m.tinta, al="ctr",
                 prompt="O que acontece")

    def conteudo(self, c: Camada):
        m = self.m
        c.retangulo("Filete lateral", 0, 0, 0.18, 7.5, m.destaque)
        c.classificacao(direita=True)
        c.logo(11.2, 6.98, 1.5)
        c.ph("Chamada", "body", 12, 0.7, 0.2, 9.5, 0.25, 12, True, m.cinza, spc=120, prompt="CHAMADA EM CAIXA ALTA")
        c.ph("Título", "title", None, 0.7, 0.5, 11.8, 0.9, 28, True, m.primaria, prompt="Título que afirma o ponto do slide")
        c.ph("Conteúdo", "body", 10, 0.7, 1.55, 11.9, 5.2, 18, False, m.tinta, prompt="Conteúdo")
        c.ph("Fonte e nota", "body", 11, 0.7, 6.95, 9.5, 0.3, 12, False, m.cinza, prompt="Fonte e período dos dados")

    # ---- família B: painéis operacionais
    def _chrome_b(self, c: Camada, titulo_prompt="Título que afirma o ponto do slide"):
        m = self.m
        c.classificacao(direita=True)
        c.retangulo("Filete do rodapé", 0, 7.08, 13.333, 0.04, m.primaria)
        c.logo(11.2, 7.1, 1.5)
        c.ph("Chamada", "body", 9, 0.6, 0.14, 9.5, 0.25, 12, True, m.cinza, spc=120, prompt="CHAMADA EM CAIXA ALTA")
        c.ph("Título", "title", None, 0.6, 0.45, 12.1, 0.7, 30, True, m.primaria, prompt=titulo_prompt)

    def numeros(self, c: Camada):
        m = self.m
        self._chrome_b(c)
        c.retangulo("Cartão destaque", 0.6, 1.45, 7.4, 3.65, WHITE, linha=m.linha)
        c.retangulo("Barra destaque", 0.6, 1.45, 0.12, 3.65, m.primaria)
        c.ph("Número de destaque", "body", 10, 1.1, 1.85, 6.7, 1.9, 100, True, m.primaria, prompt="99.353")
        c.ph("Rótulo do destaque", "body", 11, 1.1, 3.85, 6.6, 0.5, 24, False, m.tinta, prompt="O que o número mede")
        c.ph("Período do destaque", "body", 12, 1.1, 4.4, 6.6, 0.4, 18, False, m.cinza, prompt="Período e fonte")
        for i in range(3):
            y = 1.45 + i * 1.25
            c.retangulo(f"Cartão de apoio {i + 1}", 8.3, y, 4.4, 1.15, WHITE, linha=m.linha)
            c.retangulo(f"Barra de apoio {i + 1}", 8.3, y, 0.12, 1.15, m.petalas[i])
            c.ph(f"Número de apoio {i + 1}", "body", 13 + 2 * i, 8.65, y + 0.15, 1.1, 0.85, 40, True, m.primaria, prompt="4")
            c.ph(f"Rótulo de apoio {i + 1}", "body", 14 + 2 * i, 9.85, y + 0.2, 2.7, 0.8, 16, False, m.tinta,
                 prompt="O que o número mede")
        c.retangulo("Faixa", 0.6, 5.3, 12.1, 0.95, m.palido)
        c.retangulo("Barra da faixa", 0.6, 5.3, 0.06, 0.95, m.secundaria)
        c.ph("Faixa de leitura", "body", 19, 0.95, 5.4, 11.5, 0.75, 16, False, m.tinta, anc="ctr", prompt="Leitura em uma frase")
        c.ph("Fonte e nota", "body", 20, 0.6, 6.45, 12.1, 0.5, 12, False, m.cinza, it=True, prompt="Fonte e período dos dados")

    def cartoes(self, c: Camada):
        m = self.m
        self._chrome_b(c, "Título da lista de frentes")
        for i in range(4):
            x = 0.6 + i * 3.1
            base = 10 + i * 5
            c.retangulo(f"Cartão {i + 1}", x, 1.45, 2.9, 5.0, WHITE, linha=m.linha)
            c.retangulo(f"Barra {i + 1}", x, 1.45, 2.9, 0.12, m.petalas[i])
            c.ph(f"Número {i + 1}", "body", base, x + 0.3, 1.8, 2.3, 0.7, 40, True, m.primaria, prompt="01")
            c.ph(f"Título {i + 1}", "body", base + 1, x + 0.3, 2.65, 2.4, 1.6, 18, True, m.tinta, prompt="Título da frente")
            c.ph(f"Etiqueta A {i + 1}", "body", base + 2, x + 0.3, 4.4, 2.3, 0.36, 12, True, WHITE, al="ctr", anc="ctr",
                 fill=m.primaria, prompt="ÁREA")
            c.ph(f"Etiqueta B {i + 1}", "body", base + 3, x + 0.3, 4.88, 2.3, 0.36, 12, True, m.primaria, al="ctr", anc="ctr",
                 fill=m.palido, prompt="TIPO")
            c.ph(f"Situação {i + 1}", "body", base + 4, x + 0.3, 5.6, 2.3, 0.5, 14, True, m.cinza, prompt="Situação")

    def recomendacao(self, c: Camada):
        m = self.m
        self._chrome_b(c, "Título da recomendação")
        c.ph("Etiquetas", "body", 10, 0.6, 1.3, 12.1, 0.4, 12, True, m.primaria, spc=100,
             prompt="PROCESSO · PRIORIDADE ALTA · ALTERA O FLUXO")
        c.texto_fixo("Rótulo achado", 0.6, 2.0, 5.8, 0.25, "O QUE ENCONTRAMOS", 12, m.alerta_txt, spc=100, al="l")
        c.ph("O que encontramos", "body", 11, 0.6, 2.35, 5.8, 1.8, 16, False, m.tinta, prompt="Fato com número, fonte e período")
        c.texto_fixo("Rótulo custo", 0.6, 4.3, 5.8, 0.25, "O QUE ISSO CUSTA", 12, m.alerta_txt, spc=100, al="l")
        c.ph("O que isso custa", "body", 12, 0.6, 4.65, 5.8, 1.7, 14, False, m.tinta, prompt="Até três impactos")
        c.retangulo("Cartão da recomendação", 6.9, 2.0, 5.8, 4.35, m.sucesso_fundo)
        c.retangulo("Barra da recomendação", 6.9, 2.0, 0.07, 4.35, m.sucesso)
        c.texto_fixo("Rótulo recomendação", 7.3, 2.25, 5.0, 0.25, "RECOMENDAÇÃO", 12, m.sucesso_esc, spc=100, al="l")
        c.ph("Recomendação", "body", 13, 7.3, 2.75, 5.1, 3.4, 18, False, m.tinta, prompt="O que fazer, em ordem")
        c.ph("Fonte e nota", "body", 14, 0.6, 6.5, 12.1, 0.45, 12, False, m.cinza, it=True, prompt="Fonte e período dos dados")

    def tabela(self, c: Camada):
        m = self.m
        self._chrome_b(c, "Título da tabela")
        c.ph("Tabela", "tbl", 10, 0.6, 1.45, 12.1, 4.7)
        c.ph("Fonte e nota", "body", 11, 0.6, 6.4, 12.1, 0.5, 12, False, m.cinza, it=True, prompt="Fonte e período dos dados")


def _tema(prs, m: Marca) -> None:
    parte = prs.slide_master.part.part_related_by(RT.THEME)
    x = parte.blob.decode("utf-8")
    cores = {"dk2": m.primaria, "lt2": m.fundo, "accent1": m.secundaria, "accent2": m.destaque, "accent3": m.aviso,
             "accent4": m.petalas[3], "accent5": m.petalas[0], "accent6": m.sucesso, "hlink": m.secundaria,
             "folHlink": m.cinza}
    for k, v in cores.items():
        x = re.sub(rf"(<a:{k}>)<a:(?:srgbClr|sysClr)[^>]*?/?>(?:</a:(?:srgbClr|sysClr)>)?(</a:{k}>)",
                   rf'\1<a:srgbClr val="{v}"/>\2', x)
    x = re.sub(r'(<a:(?:major|minor)Font><a:latin typeface=")[^"]*"', rf'\1{m.fonte}"', x)
    nome = re.sub(r'[^A-Za-z0-9 ]', "", m.nome)[:30] or "Marca"
    x = re.sub(r'<a:clrScheme name="[^"]*"', f'<a:clrScheme name="{nome}"', x)
    x = re.sub(r'<a:fontScheme name="[^"]*"', f'<a:fontScheme name="{nome}"', x)
    parte._blob = x.encode("utf-8")


def construir(m: Marca, saida: Path = None) -> Path:
    """Gera modelo.pptx e modelo.potx; devolve o caminho do .pptx."""
    saida = Path(saida) if saida else m.modelo
    prs = Presentation()
    prs.slide_width, prs.slide_height = E(13.333), E(7.5)
    lays = list(prs.slide_layouts)
    b = Construtor(m)
    plano = [("capa", m.fundo, b.capa), ("numero", m.fundo, b.numero), ("painel", WHITE, b.painel),
             ("decisao", m.fundo, b.decisao), ("conteudo", WHITE, b.conteudo), ("numeros", m.suave, b.numeros),
             ("cartoes", m.suave, b.cartoes), ("recomendacao", m.suave, b.recomendacao), ("tabela", m.suave, b.tabela)]
    for lay, (chave, fundo, fn) in zip(lays, plano):
        arv = _limpar(lay, LAYOUTS[chave], fundo)
        cam = Camada(lay.part, m)
        fn(cam)
        for el in cam.itens:
            arv.append(el)
    for lay in lays[len(plano):]:
        prs.slide_layouts.remove(lay)
    _tema(prs, m)
    from datetime import datetime
    cp = prs.core_properties
    cp.title = f"Modelo {m.nome}"
    cp.author = (m.autoria or {}).get("autor", "") or m.nome
    cp.last_modified_by = cp.author
    cp.comments = ""
    cp.created = cp.modified = datetime.now()
    saida.parent.mkdir(parents=True, exist_ok=True)
    prs.save(saida)
    potx = saida.with_suffix(".potx")
    with zipfile.ZipFile(saida) as zi, zipfile.ZipFile(potx, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            dados = zi.read(item.filename)
            if item.filename == "[Content_Types].xml":
                dados = dados.replace(b"presentationml.presentation.main+xml", b"presentationml.template.main+xml")
            zo.writestr(item, dados)
    return saida
